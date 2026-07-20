# -*- coding: utf-8 -*-
"""
export_preview.py — 把 PPTX 每页导出为 PNG，供逐页视觉自检
依赖本机已安装 PowerPoint（COM 自动化）。

用法:
  python scripts/export_preview.py 输出.pptx [预览目录] [宽] [高]
默认: 预览目录=./_preview, 1280x720
"""
import os
import subprocess
import sys
import tempfile

PS = r"""
param($pptx, $outdir, $w, $h)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptx, $true, $false, $false)
New-Item -ItemType Directory -Force -Path $outdir | Out-Null
foreach ($slide in $pres.Slides) {
  $n = $slide.SlideIndex.ToString("00")
  $slide.Export((Join-Path $outdir ("slide_" + $n + ".png")), "PNG", $w, $h)
}
$pres.Close()
$ppt.Quit()
"""

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    pptx = os.path.abspath(sys.argv[1])
    outdir = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.abspath("_preview")
    w = sys.argv[3] if len(sys.argv) > 3 else "1280"
    h = sys.argv[4] if len(sys.argv) > 4 else "720"
    if not os.path.exists(pptx):
        print(f"文件不存在: {pptx}")
        sys.exit(2)
    fd, ps1 = tempfile.mkstemp(suffix=".ps1")
    with os.fdopen(fd, "w", encoding="ascii") as f:
        f.write(PS)
    try:
        subprocess.run(["powershell", "-NoProfile", "-File", ps1, pptx, outdir, w, h],
                       check=True)
    finally:
        os.unlink(ps1)
    print(f"OK 已导出到 {outdir}（逐页检查：字号/密度/引导/空白/占位符）")

if __name__ == "__main__":
    main()
