from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


RELEASE_FILES = ("HITFapiao.exe", "模板.xls", "README.md")


def create_release_zip(release_dir: str | Path, output_path: str | Path) -> Path:
    release = Path(release_dir).resolve()
    output = Path(output_path).resolve()
    missing = [name for name in RELEASE_FILES if not (release / name).is_file()]
    if missing:
        raise FileNotFoundError(f"Release 缺少文件：{', '.join(missing)}")

    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.tmp")
    with ZipFile(temporary, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for name in RELEASE_FILES:
            archive.write(release / name, arcname=name)
    temporary.replace(output)
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="生成保留中文文件名的 Windows Release ZIP")
    parser.add_argument("--release-dir", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    print(create_release_zip(args.release_dir, args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
