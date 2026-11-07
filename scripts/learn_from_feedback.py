# -*- coding: utf-8 -*-
"""learn_from_feedback.py — 用户习惯自主学习与技能演进引擎

功能：
1. record_feedback: 接收用户自然语言反馈，抽取风格/句式/词汇偏好，更新偏好库
2. diff_learning: 比对“技能生成版 PPTX”与“用户手动修改版 PPTX”，自动发现用户的改稿规律并吸纳为持久化规则
3. inject_profile: 在新项目初始化时注入用户偏好，让技能越用越懂当前用户
"""
from collections import Counter
from datetime import datetime
from pathlib import Path
import argparse
import difflib
import json
import re
import sys

from pptx import Presentation

BASE = Path(__file__).resolve().parent.parent
PROFILE_PATH = BASE / "user_profile.json"
LEARNED_MD_PATH = BASE / "references" / "user_learned_preferences.md"

DEFAULT_PROFILE = {
    "version": "1.0",
    "updated_at": "",
    "user_identity": "高校行政核心用户",
    "style_preferences": {
        "title_pattern": "优先八字对仗（动宾+动宾），拒绝中性标题",
        "structure_preference": "严格三点式/四点式闭环，成效与问题、计划1:1呼应",
        "data_style": "重增量轻存量，偏好百分点和增长率",
        "tone_rule": "稳妥工整，绝对避讳负面定性与口语词",
    },
    "custom_vocabulary_substitutions": {},
    "banned_phrases_extra": [],
    "preferred_framework_terms": [
        "五库五链", "育人智链", "三位一体", "协同升维", "数智赋能"
    ],
    "learning_history": []
}


def load_profile() -> dict:
    if PROFILE_PATH.is_file():
        try:
            return json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return DEFAULT_PROFILE.copy()


def save_profile(profile: dict):
    profile["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    PROFILE_PATH.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")
    sync_markdown_reference(profile)


def sync_markdown_reference(profile: dict):
    lines = [
        "# 用户专属偏好与自主演进规则库 (User Learned Preferences)",
        "",
        "> 本文档由学习引擎根据用户的日常反馈与 PPTX 改稿记录自动沉淀，生成新 PPT 时优先遵从。",
        f"> 最近更新时间：{profile.get('updated_at', '未记录')}",
        "",
        "## 一、核心风格偏好",
    ]
    for k, v in profile.get("style_preferences", {}).items():
        lines.append(f"- **{k}**: {v}")
    
    lines.append("\n## 二、专有名词与词汇升格替换习惯 (自学习沉淀)")
    subs = profile.get("custom_vocabulary_substitutions", {})
    if subs:
        lines.append("| 常见原词 | 用户偏好使用词 | 沉淀来源 |")
        lines.append("|---|---|---|")
        for orig, info in subs.items():
            if isinstance(info, dict):
                lines.append(f"| {orig} | **{info.get('preferred')}** | {info.get('source', '用户改动')} |")
            else:
                lines.append(f"| {orig} | **{info}** | 用户设定 |")
    else:
        lines.append("*暂未记录到特定词汇替换，通过 `wut_svg.py learn` 或 `feedback` 持续积累。*")

    lines.append("\n## 三、用户高频偏好体系词")
    for term in profile.get("preferred_framework_terms", []):
        lines.append(f"- `{term}`")

    lines.append("\n## 四、最近自主学习记录")
    history = profile.get("learning_history", [])[-5:]
    if history:
        for item in reversed(history):
            lines.append(f"- **[{item.get('time')}]** {item.get('summary')}")
    else:
        lines.append("*暂无学习历史。*")

    LEARNED_MD_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def extract_all_texts_from_pptx(pptx_path: Path) -> list[str]:
    prs = Presentation(pptx_path)
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    texts.append(txt)
    return texts


def record_user_feedback(feedback_text: str):
    profile = load_profile()
    
    # 规则抽取启发式分析
    insights = []
    
    # 提取对仗习惯
    if any(k in feedback_text for k in ("八字", "对仗", "对偶", "工整")):
        profile["style_preferences"]["title_pattern"] = "严格八字对仗，字数必须绝对对齐"
        insights.append("强化了八字对仗标题要求")
        
    # 提取三点式偏好
    if any(k in feedback_text for k in ("三点", "三个", "不要多", "聚焦")):
        profile["style_preferences"]["structure_preference"] = "严格每板块不超过3点，提炼闭环"
        insights.append("确立了严格三点式结构偏好")

    # 提取数据/增量偏好
    if any(k in feedback_text for k in ("数据", "百分比", "增长率", "同比")):
        profile["style_preferences"]["data_style"] = "必须呈现增长率与同比数据，弱化单一体量"
        insights.append("强调了增量与百分比优先的数据表达风格")

    # 提取名词替换指令 (例如: "把XX改成YY", "喜欢用YY而不是XX")
    replace_match = re.search(r"把[“\"']?([^“”\"'\s]+)[”\"']?改[成为][“\"']?([^“”\"'\s]+)[”\"']?", feedback_text)
    if replace_match:
        orig, prefer = replace_match.group(1), replace_match.group(2)
        profile["custom_vocabulary_substitutions"][orig] = {
            "preferred": prefer,
            "source": f"用户口头指令反馈: {feedback_text[:40]}"
        }
        insights.append(f"学到了词汇替换：【{orig}】→【{prefer}】")

    summary_str = "; ".join(insights) if insights else f"记录了用户综合反馈：{feedback_text[:50]}"
    profile["learning_history"].append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "type": "verbal_feedback",
        "raw": feedback_text,
        "summary": summary_str
    })
    
    save_profile(profile)
    print(f"\n[OK] 成功吸收用户反馈并进化技能配置！")
    print(f"  学习成果：{summary_str}")
    print(f"  规则已同步到 references/user_learned_preferences.md")


def diff_and_learn_from_deck(original_pptx: Path, modified_pptx: Path):
    if not original_pptx.is_file() or not modified_pptx.is_file():
        raise SystemExit(f"PPTX 文件不存在，请确认路径:\n  原版: {original_pptx}\n  改版: {modified_pptx}")

    profile = load_profile()
    orig_texts = extract_all_texts_from_pptx(original_pptx)
    mod_texts = extract_all_texts_from_pptx(modified_pptx)

    diff = list(difflib.ndiff(orig_texts, mod_texts))
    removals = [line[2:].strip() for line in diff if line.startswith("- ")]
    additions = [line[2:].strip() for line in diff if line.startswith("+ ")]

    learned_subs = {}
    # 查找可能的词句对齐替换
    for r in removals:
        for a in additions:
            s = difflib.SequenceMatcher(None, r, a)
            if 0.4 < s.ratio() < 0.95:  # 存在部分修改
                # 识别微调
                learned_subs[r] = {
                    "preferred": a,
                    "source": f"用户改稿实际比对（相似度 {s.ratio():.0%}）"
                }

    profile["custom_vocabulary_substitutions"].update(learned_subs)
    
    summary = f"通过比对用户改稿学习到 {len(learned_subs)} 处用词与表述精修偏好"
    profile["learning_history"].append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "type": "pptx_diff",
        "original_deck": original_pptx.name,
        "modified_deck": modified_pptx.name,
        "summary": summary,
        "substitutions_learned": len(learned_subs)
    })
    
    save_profile(profile)
    print(f"\n[OK] 完成用户改稿自主分析！共发现并吸收 {len(learned_subs)} 项精修表达习惯。")
    for k, v in list(learned_subs.items())[:5]:
        print(f"  原版: {k[:24]}...")
        print(f"  用户修改为: {v['preferred'][:24]}...")
    print(f"已持久化沉淀到 references/user_learned_preferences.md")


def main():
    parser = argparse.ArgumentParser(description="wut-ppt 用户习惯自主学习引擎")
    sub = parser.add_subparsers(dest="action", required=True)
    
    p_feedback = sub.add_parser("feedback", help="录入用户日常习惯与意见，自适应更新规则")
    p_feedback.add_argument("text", help="用户的具体修改要求或习惯描述")
    
    p_diff = sub.add_parser("diff", help="比对技能生成的 PPTX 与用户手工修改后的 PPTX")
    p_diff.add_argument("original", help="技能导出的 PPTX 路径")
    p_diff.add_argument("modified", help="用户自己手工修改后的最终 PPTX 路径")

    p_show = sub.add_parser("show", help="查看当前积累的用户画像与进化规则")

    args = parser.parse_args()
    if args.action == "feedback":
        record_user_feedback(args.text)
    elif args.action == "diff":
        diff_and_learn_from_deck(Path(args.original), Path(args.modified))
    elif args.action == "show":
        profile = load_profile()
        print(json.dumps(profile, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
