# -*- coding: utf-8 -*-
"""check_content_compliance.py — 高校行政公文合规与避讳词扫描器

检查规则（基于 references/content_writing_guide.md）：
1. 绝对避讳词拦截（负面定性、情绪化、示弱、口语词）
2. 称谓规范检查（自称、上级、兄弟单位、师生）
3. 标题结论先行检查（禁止“工作概况/总结汇报/情况介绍”等中性无观点标题）
4. 动作词低阶警告（提示升格为体制内行政动词）

用法:
  python scripts/check_content_compliance.py <project_path>
"""
from pathlib import Path
import argparse
import re
import sys
import xml.etree.ElementTree as ET

# 绝对避讳词（一旦出现直接报错拦截）
BANNED_WORDS = {
    # 负面定性
    "失败": "应用“仍有提升空间/成效未充分释放”",
    "错误": "应用“偏差/薄弱环节”",
    "严重问题": "应用“突出瓶颈/关键阻碍”",
    "搞砸": "严禁口语负面词",
    "严重失误": "严禁绝对负面定性",
    "漏洞百出": "严禁情绪化贬义词",
    # 情绪化表达
    "怨声载道": "严禁极端情绪化表达",
    "不满意": "应用“与师生期待仍有差距”",
    "意见很大": "应用“反馈意见较为集中”",
    # 极端示弱
    "经费不足": "应用“投入保障机制仍需健全/资金配置效率有待提高”",
    "缺钱": "严禁使用口语化示弱词汇",
    "人手不够": "应用“队伍梯队建设仍需完善/专业力量有待充实”",
    "没人愿意干": "严禁消极定性",
    "条件太差": "应用“基础支撑能力仍待强化”",
    # 口语粗俗
    "搞定": "应用“落实到位/取得实效/圆满完成”",
    "瞎搞": "严禁非正式用语",
    "随便弄": "严禁口语表达",
    "我们搞了": "严禁口语汇报表述",
    "咱们": "自称请用“我校/我院/本部门/本单位”",
}

# 称谓违规检查
PRONOUN_RULES = {
    "咱们": "自称统一使用“我校/我院/本单位/本部门”",
    "别的学校": "应用“兄弟高校/外校”",
    "其他单位": "应用“兄弟单位/相关协作单位”",
    "上面领导": "应用“上级主管部门/学校党委行政”",
    "上面要求": "应用“上位政策要求/主管部门部署”",
}

# 中性无判断标题拦截（标题必须结论先行）
NEUTRAL_TITLE_PATTERNS = [
    r"^.*工作概况$",
    r"^.*建设情况$",
    r"^.*情况介绍$",
    r"^.*工作总结$",
    r"^.*建设进展$",
    r"^.*相关汇报$",
]


def extract_svg_texts(svg_path: Path) -> list[tuple[int, str]]:
    """提取 SVG 中所有可见文本内容及大致行号"""
    results = []
    try:
        content = svg_path.read_text(encoding="utf-8")
        root = ET.fromstring(content)
        for text_el in root.iter():
            if text_el.tag.rsplit("}", 1)[-1] in ("text", "tspan"):
                txt = "".join(text_el.itertext()).strip()
                if txt:
                    results.append(txt)
    except Exception as exc:
        results.append(f"[XML解析异常: {exc}]")
    return results


def check_spec_content(spec_path: Path):
    errors, warnings = [], []
    if not spec_path.is_file():
        return errors, warnings
    
    text = spec_path.read_text(encoding="utf-8")
    
    # 检查避讳词
    for word, tip in BANNED_WORDS.items():
        if word in text:
            errors.append(f"design_spec.md 出现绝对避讳词【{word}】→ {tip}")
            
    # 检查称谓
    for word, tip in PRONOUN_RULES.items():
        if word in text:
            warnings.append(f"design_spec.md 出现非标准称谓【{word}】→ {tip}")

    # 检查标题结论先行
    slide_titles = re.findall(r"(?m)^-\s+\*\*Title\*\*:\s*(.+?)\s*$", text)
    for title in slide_titles:
        for pat in NEUTRAL_TITLE_PATTERNS:
            if re.match(pat, title.strip()):
                errors.append(
                    f"页面标题【{title}】属于中性概况标题！必须“结论先行”，改为有观点、有政绩动词的提炼句"
                )

    return errors, warnings


def check_svg_files(project: Path):
    errors, warnings = [], []
    svgs = sorted((project / "svg_output").glob("*.svg"))
    
    for svg in svgs:
        texts = extract_svg_texts(svg)
        full_text = " ".join(texts)
        
        for word, tip in BANNED_WORDS.items():
            if word in full_text:
                errors.append(f"{svg.name} 出现公文绝对避讳词【{word}】→ {tip}")
                
        for word, tip in PRONOUN_RULES.items():
            if word in full_text:
                warnings.append(f"{svg.name} 出现口语化称谓【{word}】→ {tip}")

    return errors, warnings


def main():
    parser = argparse.ArgumentParser(description="高校行政公文合规检查器")
    parser.add_argument("project", help="项目根目录")
    args = parser.parse_args()
    
    project = Path(args.project)
    spec_errors, spec_warnings = check_spec_content(project / "design_spec.md")
    svg_errors, svg_warnings = check_svg_files(project)
    
    all_warnings = spec_warnings + svg_warnings
    all_errors = spec_errors + svg_errors
    
    for w in all_warnings:
        print(f"[WARN-公文规范] {w}")
    for e in all_errors:
        print(f"[ERROR-公文红线] {e}")
        
    if all_errors:
        print(f"\n[FAIL] 发现 {len(all_errors)} 处公文写作红线违规，禁止直接向校领导汇报交付，请修改后重试！")
        sys.exit(1)
    else:
        print(f"[OK] 高校行政公文合规审查通过：无绝对避讳词、无中性空洞标题、称谓规范。")


if __name__ == "__main__":
    main()
