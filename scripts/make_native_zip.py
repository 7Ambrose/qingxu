import hashlib
import pathlib
import shutil
import zipfile
import time

ROOT = pathlib.Path(__file__).resolve().parent.parent
NAME = "Qingxu-Native-v0.1.0"
STAGE = ROOT / "dist" / NAME
SRC_EXE = pathlib.Path(r"D:\Codex\计划！！！！！\outputs\native-win-x64\青序.exe")


def main() -> None:
    if STAGE.exists():
        shutil.rmtree(STAGE)
    STAGE.mkdir(parents=True)
    shutil.copy2(SRC_EXE, STAGE / "青序.exe")
    for name in ("计划模板.xlsx", "README.md", "创建桌面快捷方式（原生版）.ps1"):
        source = ROOT / "templates" / name if name.endswith(".xlsx") else ROOT / name
        shutil.copy2(source, STAGE / name)
    lines = []
    for item in sorted(STAGE.iterdir()):
        digest = hashlib.sha256(item.read_bytes()).hexdigest()
        lines.append(f"{digest}  {item.name}")
        print(f"  {item.name:44s} {item.stat().st_size / 1048576:8.2f} MiB", flush=True)
    (STAGE / "SHA256SUMS.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")

    target = ROOT / "dist" / f"{NAME}-win-x64.zip"
    if target.exists():
        target.unlink()
    started = time.time()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for item in sorted(STAGE.iterdir()):
            archive.write(item, f"{NAME}/{item.name}")
    size = target.stat().st_size
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    print(f"zip {size / 1048576:.1f} MiB in {time.time() - started:.0f}s", flush=True)
    print(f"sha256 {digest}", flush=True)
    (ROOT / "dist" / "release.txt").write_text(f"{target.name}\n{size}\n{digest}\n", encoding="utf-8")


if __name__ == "__main__":
    main()
