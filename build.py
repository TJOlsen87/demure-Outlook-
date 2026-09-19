"""Build an installable NVDA add-on using only the Python standard library."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import ast
import re

ROOT = Path(__file__).resolve().parent

def main():
    addon = ROOT / "addon"
    manifest = (addon / "manifest.ini").read_text(encoding="utf-8-sig")
    version = re.search(r"^version\s*=\s*([0-9.]+)\s*$", manifest, re.MULTILINE).group(1)
    ast.parse((addon / "globalPlugins/demureOutlook.py").read_text(encoding="utf-8-sig"))
    output = ROOT / "dist" / f"DemureOutlook-{version}.nvda-addon"
    output.parent.mkdir(exist_ok=True)
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for path in sorted(addon.rglob("*")):
            if path.is_file() and "__pycache__" not in path.parts and path.suffix not in (".pyc", ".pyo"):
                archive.write(path, path.relative_to(addon).as_posix())
        for name in ("README.md", "CHANGELOG.md", "LICENSE"):
            if (ROOT / name).exists():
                archive.write(ROOT / name, name)
    with ZipFile(output) as archive:
        assert archive.testzip() is None
    print(output)

if __name__ == "__main__":
    main()
