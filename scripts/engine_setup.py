# -*- coding: utf-8 -*-
"""Download the pinned PPT Master engine into wut-ppt/.engine (local only)."""
from pathlib import Path
import shutil
import tempfile
import urllib.request
import zipfile

BASE = Path(__file__).resolve().parent.parent
ENGINE_ROOT = BASE / ".engine"
TARGET = ENGINE_ROOT / "ppt-master"
VERSION = "v6.7.3"
URL = f"https://github.com/hugohe3/ppt-master/releases/download/{VERSION}/ppt-master-skill-{VERSION}.zip"
REQUIRED = ("svg_to_pptx.py", "svg_quality_checker.py", "preset_shape_svg.py", "project_manager.py")


def valid(path: Path) -> bool:
    return all((path / "scripts" / name).is_file() for name in REQUIRED)


def main() -> None:
    if valid(TARGET):
        print(f"OK PPT Master {VERSION}: {TARGET}")
        return
    if TARGET.exists():
        raise SystemExit(f"引擎目录存在但不完整，请检查后处理: {TARGET}")
    ENGINE_ROOT.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="engine-", dir=ENGINE_ROOT) as tmp:
        tmp_path = Path(tmp)
        archive = tmp_path / "engine.zip"
        print(f"下载 PPT Master {VERSION} ...")
        with urllib.request.urlopen(URL, timeout=120) as response, archive.open("wb") as fh:
            shutil.copyfileobj(response, fh)
        extracted = tmp_path / "extracted"
        with zipfile.ZipFile(archive) as zf:
            for member in zf.infolist():
                member_path = Path(member.filename)
                if member_path.is_absolute() or ".." in member_path.parts:
                    raise SystemExit("引擎压缩包包含不安全路径")
            zf.extractall(extracted)
        source = extracted / "ppt-master" / "skills" / "ppt-master"
        if not valid(source):
            raise SystemExit("下载的引擎结构与预期不符")
        shutil.move(str(source), str(TARGET))
    print(f"OK PPT Master {VERSION}: {TARGET}")
    print("引擎仅安装在技能目录内；.engine 不进入 Git 仓库。")


if __name__ == "__main__":
    main()
