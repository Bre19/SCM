from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.revamp import build_workbook, run_pipeline  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Data Cleansing otomatis dari sembilan file vendor dan PO."
    )
    parser.add_argument(
        "--raw-dir",
        type=Path,
        default=PROJECT_ROOT / "data" / "raw",
        help="Folder yang berisi sembilan input wajib.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=PROJECT_ROOT / "output" / "revamp",
        help="Folder hasil workbook dan audit.",
    )
    parser.add_argument(
        "--config-dir",
        type=Path,
        default=PROJECT_ROOT / "config",
        help="Folder konfigurasi klasifikasi.",
    )
    parser.add_argument(
        "--settings",
        type=Path,
        default=PROJECT_ROOT / "config" / "revamp_settings.json",
        help="Konfigurasi kategori dan asumsi pipeline.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    raw_dir = args.raw_dir.resolve()
    output_dir = args.output_dir.resolve()
    config_dir = args.config_dir.resolve()
    settings = args.settings.resolve()

    output_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = output_dir / "Data Cleansing - Otomatis.xlsx"
    with tempfile.TemporaryDirectory(prefix="vendor-cleansing-") as temporary:
        artifacts = run_pipeline(
            raw_dir=raw_dir,
            config_dir=config_dir,
            output_dir=Path(temporary),
            settings_file=settings,
        )
        build_workbook(artifacts["bundle"], workbook_path)

    print(
        json.dumps(
            {
                "status": "ok",
                "workbook": str(workbook_path),
                "output_dir": str(output_dir),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
