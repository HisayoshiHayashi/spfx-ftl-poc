"""Replace existing scenario JPGs from ID-named PNGs using macOS sips."""

import argparse
import concurrent.futures
import json
import pathlib
import subprocess
import tempfile


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("images", type=pathlib.Path)
    args = parser.parse_args()
    root = pathlib.Path(__file__).resolve().parents[1] / "src/webparts/ftlApp"
    scenarios = json.loads((root / "data/scenarios.json").read_text(encoding="utf-8"))
    ids = [item["id"] for item in scenarios]
    for scenario_id in ids:
        if not scenario_id.startswith("MS") or not scenario_id[2:].isdigit():
            raise ValueError(f"Invalid scenario ID: {scenario_id}")
        source = args.images / f"{scenario_id}_image.png"
        if not source.is_file():
            raise FileNotFoundError(source)
    destination = root / "assets/scenarios"
    with tempfile.TemporaryDirectory(prefix="ftl-scenario-images-") as staging:
        staging_path = pathlib.Path(staging)

        def convert(scenario_id):
            source = args.images / f"{scenario_id}_image.png"
            output = staging_path / f"{scenario_id}.jpg"
            subprocess.run(
                ["sips", "-s", "format", "jpeg", "-s", "formatOptions", "80",
                 str(source), "--out", str(output)],
                check=True, capture_output=True, text=True
            )
            if output.stat().st_size == 0:
                raise ValueError(f"Empty converted image: {scenario_id}")
            return output

        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            outputs = list(executor.map(convert, ids))
        source_size = sum((args.images / f"{scenario_id}_image.png").stat().st_size for scenario_id in ids)
        output_size = sum(output.stat().st_size for output in outputs)
        for output in outputs:
            output.replace(destination / output.name)
        print(f"Replaced {len(outputs)} images. PNG {source_size / 1024**2:.1f} MiB -> JPEG {output_size / 1024**2:.1f} MiB")
        print("Original images, scenario data and ID mappings were not changed.")


if __name__ == "__main__":
    main()
