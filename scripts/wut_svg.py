# -*- coding: utf-8 -*-
"""wut-ppt SVG-first command entry.

Usage:
  python scripts/wut_svg.py init <project_name> [--dir <parent>]
  python scripts/wut_svg.py preset <preset_shape_svg.py arguments...>
  python scripts/wut_svg.py check <project_path>
  python scripts/wut_svg.py export <project_path>
"""
from pathlib import Path
import argparse
import re
import subprocess
import sys
import zipfile

BASE = Path(__file__).resolve().parent.parent
ENGINE = BASE / ".engine" / "ppt-master"
SCRIPTS = ENGINE / "scripts"
TEMPLATE = BASE / "assets" / "template.pptx"
MIN_SVG_FONT_PX = 16.0  # 1280x720 canvas: 16px * 0.75 = 12pt


def require_engine() -> None:
    if not (SCRIPTS / "svg_to_pptx.py").is_file():
        raise SystemExit("缺少 PPT Master 引擎，先运行: python scripts/engine_setup.py")


def run(script: str, args: list[str], capture: bool = False) -> subprocess.CompletedProcess:
    require_engine()
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, args)],
        check=True,
        capture_output=capture,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def init_project(args: argparse.Namespace) -> None:
    cmd = ["init", args.project_name, "--format", "ppt169"]
    if args.parent:
        cmd += ["--dir", args.parent]
    result = run("project_manager.py", cmd, capture=True)
    print(result.stdout, end="")
    match = re.search(r"Project created:\s*(.+)", result.stdout)
    if not match:
        raise SystemExit("无法识别新建项目路径")
    project = Path(match.group(1).strip())
    run("project_manager.py", ["scaffold-spec", project])
    run("project_manager.py", ["scaffold-lock", project])
    spec_path = project / "design_spec.md"
    profile_memo = ""
    profile_file = BASE / "user_profile.json"
    if profile_file.is_file():
        try:
            p_data = json.loads(profile_file.read_text(encoding="utf-8"))
            style_p = p_data.get("style_preferences", {})
            subs = p_data.get("custom_vocabulary_substitutions", {})
            sub_list = [f"{k}→{v.get('preferred') if isinstance(v, dict) else v}" for k, v in list(subs.items())[:5]]
            profile_memo = (
                f"\n> **已加载当前用户专属画像偏好**：\n"
                f"> - 标题风格：{style_p.get('title_pattern', '八字对仗')}\n"
                f"> - 结构偏好：{style_p.get('structure_preference', '三点式闭环')}\n"
                f"> - 数据风格：{style_p.get('data_style', '增量百分比优先')}\n"
            )
            if sub_list:
                profile_memo += f"> - 专属词汇沉淀：{', '.join(sub_list)}\n"
        except Exception:
            pass

    with spec_path.open("a", encoding="utf-8") as spec:
        spec.write(f"""

## XI. 高校行政公文内容设计与全稿视觉地图（强制约束）
{profile_memo}
> 1. 内容写作严格遵从 `wut-ppt/references/content_writing_guide.md` 五条铁律（工整优先、稳妥优先、结论先行、虚实结合、闭环结构）；
> 2. 标题优先八字对仗（动宾+动宾），禁中性概况标题；正文三点式/四点式起头（一是/二是/三是）；严禁口语与绝对避讳词；
> 3. 同一套 PPT 中正文 Layout ID、Geometry signature、Macro composition 均不得重复；正文必须占满，不得留下连续空白；
> 4. 使用箭头时必须在 Arrow semantics 中说明“来源→目标＋业务语义”；表格无真实交接/映射严禁添加箭头。

每个 `#### Slide NN` 块在 Relationships 后补充：

- **Layout ID**: L01（全稿唯一）
- **Geometry signature**: 页面任务 → 宏观构图 → 阅读方向 → 视觉中心 → 渲染语言（全稿唯一）
- **Macro composition**: 构图知识库 ID 或 custom-具体名称（全稿唯一）
- **Reading direction**: 例如自上而下 / 中心向四周
- **Visual center**: 页面第一视觉落点
- **Rendering language**: 行政编辑 / 数据新闻 / 工程蓝图 / 等距立体等
- **Distinction**: 与此前所有页面在宏观构图上的本质区别，不得只写换内容、颜色或图标
- **Arrow semantics**: 使用箭头时必填“来源→目标＋业务语义”；表格无真实交接/映射不得加箭头
""")
    (project / "deck.json").write_text(
        '{\n  "title": "汇报标题",\n  "speaker": "汇报人",\n'
        '  "chapters": [{"name": "第一章", "pages": []}]\n}\n',
        encoding="utf-8",
    )
    with zipfile.ZipFile(TEMPLATE) as zf:
        for info in zf.infolist():
            if info.filename.startswith("ppt/media/") and not info.is_dir():
                target = project / "images" / Path(info.filename).name
                target.write_bytes(zf.read(info))
    print(f"[OK] 张林模板媒体已提取到 {project / 'images'}")
    print("下一步: 完善 design_spec.md / spec_lock.md / deck.json，再逐页手写正文 svg_output/*.svg")


def check_svg_fonts(project: Path) -> list[str]:
    errors = []
    for svg in sorted((project / "svg_output").glob("*.svg")):
        text = svg.read_text(encoding="utf-8")
        values = [float(v) for v in re.findall(r"font-size\s*=\s*[\"']\s*(\d+(?:\.\d+)?)", text)]
        values += [float(v) for v in re.findall(r"font-size\s*:\s*(\d+(?:\.\d+)?)(?:px)?", text)]
        small = [v for v in values if v < MIN_SVG_FONT_PX]
        if small:
            errors.append(f"{svg.name}: font-size {min(small):g}px < 16px（12pt）")
    return errors


def check_project(args: argparse.Namespace) -> None:
    project = Path(args.project)
    run("project_manager.py", ["validate", project])
    run("svg_quality_checker.py", [project, "--stage", "final", "--json"])
    errors = check_svg_fonts(project)
    if errors:
        for error in errors:
            print(f"[ERROR] {error}")
        raise SystemExit(1)
    print("[OK] wut-ppt 字号检查：所有显式 SVG font-size 均不小于 16px（12pt）")
    subprocess.run([sys.executable, str(BASE / "scripts" / "check_content_compliance.py"), str(project)], check=True)
    subprocess.run([sys.executable, str(BASE / "scripts" / "deck_visual_audit.py"), str(project)], check=True)
    subprocess.run([sys.executable, str(BASE / "scripts" / "check_visual_density.py"), str(project)], check=True)


def export_project(args: argparse.Namespace) -> None:
    project = Path(args.project)
    check_project(args)
    if (project / "notes" / "total.md").is_file():
        run("total_md_split.py", [project])
    run("finalize_svg.py", [project])
    run("svg_to_pptx.py", [project])
    decks = sorted(
        (p for p in (project / "exports").glob("*.pptx") if not p.name.endswith("_assembled.pptx")),
        key=lambda p: p.stat().st_mtime,
    )
    if not decks:
        raise SystemExit("导出结束但未找到 PPTX")
    latest = decks[-1]
    subprocess.run([sys.executable, str(BASE / "scripts" / "check_deck.py"), str(latest)], check=True)
    deck_config = project / "deck.json"
    if deck_config.is_file():
        assembled = project / "exports" / f"{project.name}_assembled.pptx"
        subprocess.run([
            sys.executable, str(BASE / "scripts" / "assemble_deck.py"),
            str(deck_config), str(latest), str(assembled),
        ], check=True)
        latest = assembled
        subprocess.run([sys.executable, str(BASE / "scripts" / "check_deck.py"), str(latest)], check=True)
    print(f"[OK] 最终文件: {latest}")
    print(f"请继续一键视觉验证: python scripts/wut_svg.py verify \"{project}\"")


def verify_project(args: argparse.Namespace) -> None:
    project = Path(args.project)
    check_project(args)
    assembled = project / "exports" / f"{project.name}_assembled.pptx"
    if assembled.is_file():
        deck = assembled
    else:
        decks = sorted(
            (p for p in (project / "exports").glob("*.pptx") if not p.name.endswith("_assembled.pptx")),
            key=lambda p: p.stat().st_mtime,
        )
        if not decks:
            raise SystemExit("未找到导出的 PPTX，先运行 wut_svg.py export")
        deck = decks[-1]
    preview = project / "_preview"
    subprocess.run([sys.executable, str(BASE / "scripts" / "export_preview.py"), str(deck), str(preview)], check=True)
    subprocess.run([sys.executable, str(BASE / "scripts" / "make_contact_sheet.py"), str(preview)], check=True)
    density_args = [sys.executable, str(BASE / "scripts" / "check_visual_density.py"), str(preview)]
    if (project / "deck.json").is_file():
        density_args += ["--deck", str(project / "deck.json")]
    subprocess.run(density_args, check=True)
    print(f"[OK] 技术与视觉验证完成，请打开 {preview / 'contact_sheet.png'} 做最终人工审查")


def main() -> None:
    parser = argparse.ArgumentParser(description="wut-ppt SVG-first pipeline")
    sub = parser.add_subparsers(dest="command", required=True)
    p_init = sub.add_parser("init", help="初始化 SVG 项目并提取张林模板媒体")
    p_init.add_argument("project_name")
    p_init.add_argument("--dir", dest="parent")
    p_init.set_defaults(func=init_project)
    p_preset = sub.add_parser("preset", help="调用 PPT Master 原生预设形状工具")
    p_preset.add_argument("args", nargs=argparse.REMAINDER)
    p_preset.set_defaults(func=lambda a: run("preset_shape_svg.py", a.args))
    p_check = sub.add_parser("check", help="检查项目、SVG 兼容性、字号、版式、箭头和静态密度")
    p_check.add_argument("project")
    p_check.set_defaults(func=check_project)
    p_export = sub.add_parser("export", help="检查并导出原生可编辑 PPTX")
    p_export.add_argument("project")
    p_export.set_defaults(func=export_project)
    p_verify = sub.add_parser("verify", help="导出 PNG、全稿总览并检查正文空白区域")
    p_verify.add_argument("project")
    p_verify.set_defaults(func=verify_project)
    p_feedback = sub.add_parser("feedback", help="录入用户日常习惯与意见，自适应更新偏好")
    p_feedback.add_argument("text", help="具体修改要求或习惯描述")
    p_feedback.set_defaults(func=lambda a: subprocess.run([sys.executable, str(BASE / "scripts" / "learn_from_feedback.py"), "feedback", a.text], check=True))
    p_learn = sub.add_parser("learn", help="比对原 PPTX 与用户手工修改后的 PPTX，自主学习改动规律")
    p_learn.add_argument("original", help="技能导出的 PPTX 路径")
    p_learn.add_argument("modified", help="用户手工修改后的 PPTX 路径")
    p_learn.set_defaults(func=lambda a: subprocess.run([sys.executable, str(BASE / "scripts" / "learn_from_feedback.py"), "diff", a.original, a.modified], check=True))
    p_profile = sub.add_parser("profile", help="查看当前积累的用户专属画像与进化规则")
    p_profile.set_defaults(func=lambda a: subprocess.run([sys.executable, str(BASE / "scripts" / "learn_from_feedback.py"), "show"], check=True))
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
