from __future__ import annotations

import argparse
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp"}
FUEL_HINTS = ("주유", "유류", "영수", "fuel", "gas")
SIGNATURE_HINTS = ("파이리", "charmander", "서명", "signature")
NATURE_HINTS = ("자연", "풍경", "활동", "activity", "landscape", "nature")


def media_type(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".png":
        return "image/png"
    if ext in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if ext == ".bmp":
        return "image/bmp"
    raise ValueError(f"Unsupported image type: {path}")


def image_size(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        header = f.read(32)

    if header.startswith(b"\x89PNG\r\n\x1a\n"):
        return struct.unpack(">II", header[16:24])

    if header.startswith(b"BM"):
        width, height = struct.unpack("<ii", header[18:26])
        return abs(width), abs(height)

    if header.startswith(b"\xff\xd8"):
        with path.open("rb") as f:
            f.read(2)
            while True:
                marker_start = f.read(1)
                if not marker_start:
                    break
                if marker_start != b"\xff":
                    continue
                marker = f.read(1)
                while marker == b"\xff":
                    marker = f.read(1)
                if marker in {b"\xc0", b"\xc1", b"\xc2", b"\xc3", b"\xc5", b"\xc6", b"\xc7", b"\xc9", b"\xca", b"\xcb", b"\xcd", b"\xce", b"\xcf"}:
                    f.read(3)
                    height, width = struct.unpack(">HH", f.read(4))
                    return width, height
                length_bytes = f.read(2)
                if len(length_bytes) != 2:
                    break
                length = struct.unpack(">H", length_bytes)[0]
                f.seek(length - 2, 1)

    raise ValueError(f"Could not read image dimensions: {path}")


def has_hint(path: Path, hints: tuple[str, ...]) -> bool:
    name = path.name.lower()
    return any(hint.lower() in name for hint in hints)


def default_source_folder() -> Path:
    cwd = Path.cwd()
    direct = cwd / "yoonskills" / "한글문서테스트원본"
    if direct.is_dir():
        return direct
    matches = [p for p in cwd.rglob("한글문서테스트원본") if p.is_dir()]
    if matches:
        return matches[0]
    return cwd


def resolve_files(values: list[str] | None, source_folder: Path) -> list[Path] | None:
    if not values:
        return None
    resolved: list[Path] = []
    for value in values:
        path = Path(value)
        if not path.is_absolute():
            path = source_folder / path
        if not path.is_file():
            raise FileNotFoundError(f"Image not found: {path}")
        resolved.append(path)
    return resolved


def detect_images(source_folder: Path, signature_values: list[str] | None, activity_values: list[str] | None, activity_count: int) -> tuple[list[Path], list[Path]]:
    explicit_signature = resolve_files(signature_values, source_folder)
    explicit_activity = resolve_files(activity_values, source_folder)
    if explicit_signature is not None and len(explicit_signature) != 2:
        raise ValueError("--signature-images requires exactly 2 images")
    if explicit_activity is not None and len(explicit_activity) < activity_count:
        raise ValueError(f"--activity-images requires at least {activity_count} images")

    files = sorted(
        [p for p in source_folder.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTS],
        key=lambda p: p.name,
    )
    if not files:
        raise FileNotFoundError(f"No supported image files found in {source_folder}")

    sizes = {p: image_size(p) for p in files}
    non_fuel = [p for p in files if not has_hint(p, FUEL_HINTS)]

    if explicit_signature is None:
        hinted = [p for p in non_fuel if has_hint(p, SIGNATURE_HINTS)]
        portraits = [p for p in non_fuel if sizes[p][1] > sizes[p][0]]
        signature = (hinted if len(hinted) >= 2 else portraits)[:2]
    else:
        signature = explicit_signature

    if len(signature) != 2:
        raise ValueError("Could not auto-detect 2 signature/파이리 images. Pass --signature-images explicitly.")

    signature_set = {p.resolve() for p in signature}
    candidates = [p for p in non_fuel if p.resolve() not in signature_set]
    if explicit_activity is None:
        hinted = [p for p in candidates if has_hint(p, NATURE_HINTS)]
        landscapes = [p for p in candidates if sizes[p][0] >= sizes[p][1]]
        activity = (hinted if len(hinted) >= activity_count else landscapes)[:activity_count]
    else:
        activity = explicit_activity[:activity_count]

    if len(activity) != activity_count:
        raise ValueError(f"Could not auto-detect {activity_count} activity/nature images. Pass --activity-images explicitly.")

    return signature, activity


def find_input_hwpx(source_folder: Path, value: str | None) -> Path:
    if value:
        path = Path(value)
        if not path.is_absolute():
            path = source_folder / path
        if not path.is_file():
            raise FileNotFoundError(f"HWPX not found: {path}")
        return path
    candidates = sorted(source_folder.glob("*.hwpx"), key=lambda p: p.name)
    if len(candidates) != 1:
        raise FileNotFoundError(f"Expected exactly one .hwpx in {source_folder}; found {len(candidates)}. Pass --input-hwpx.")
    return candidates[0]


def item_xml(image_id: str, href: str, mtype: str) -> str:
    return f'<opf:item id="{image_id}" href="{href}" media-type="{mtype}" isEmbeded="1"/>'


def item_pattern(image_id: str) -> re.Pattern[str]:
    return re.compile(rf'<opf:item\b(?=[^>]*\bid="{re.escape(image_id)}")[^>]*/>')


def replace_manifest(hpf_path: Path, replacements: dict[str, str], media: dict[str, str]) -> None:
    hpf = hpf_path.read_text(encoding="utf-8")

    for image_id in ("image1", "image2"):
        new_item = item_xml(image_id, replacements[image_id], media[image_id])
        hpf, count = item_pattern(image_id).subn(new_item, hpf, count=1)
        if count != 1:
            raise ValueError(f"Could not find manifest item for {image_id}")

    for image_id in ("image4", "image5", "image6", "image7"):
        hpf = item_pattern(image_id).sub("", hpf)

    image3_match = item_pattern("image3").search(hpf)
    if not image3_match:
        raise ValueError("Could not find manifest item for image3")

    new_items = "".join(item_xml(image_id, replacements[image_id], media[image_id]) for image_id in ("image4", "image5", "image6", "image7"))
    hpf = hpf[: image3_match.end()] + new_items + hpf[image3_match.end() :]
    hpf_path.write_text(hpf, encoding="utf-8")


def replace_activity_refs(section_path: Path, count: int) -> None:
    section = section_path.read_text(encoding="utf-8")
    ids = [f"image{i}" for i in range(4, 4 + count)]
    index = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal index
        if index < len(ids):
            value = f'binaryItemIDRef="{ids[index]}"'
            index += 1
            return value
        return match.group(0)

    updated = re.sub(r'binaryItemIDRef="image3"', repl, section)
    if index != count:
        raise ValueError(f"Expected to replace {count} activity image3 references, replaced {index}")
    if 'binaryItemIDRef="image3"' not in updated:
        raise ValueError("No image3 reference remains for the fuel photo; refusing to write output")
    section_path.write_text(updated, encoding="utf-8")


def bindata_image_id(zip_name: str) -> str | None:
    match = re.match(r"^BinData/(image\d+)\.[^/]+$", zip_name)
    return match.group(1) if match else None


def package_hwpx(source_hwpx: Path, work_dir: Path, output_hwpx: Path, zip_replacements: dict[str, Path]) -> None:
    tmp = output_hwpx.with_suffix(output_hwpx.suffix + ".tmp")
    if tmp.exists():
        tmp.unlink()
    output_hwpx.parent.mkdir(parents=True, exist_ok=True)

    written: set[str] = set()
    extra_ids = {"image4", "image5", "image6", "image7"}

    def write_file(zout: zipfile.ZipFile, local: Path, arcname: str, compression: int = zipfile.ZIP_DEFLATED) -> None:
        zout.write(local, arcname, compress_type=compression)
        written.add(arcname)

    with zipfile.ZipFile(source_hwpx, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        write_file(zout, work_dir / "mimetype", "mimetype", zipfile.ZIP_STORED)
        for info in zin.infolist():
            name = info.filename
            if name == "mimetype" or name.endswith("/"):
                continue
            image_id = bindata_image_id(name)
            if image_id in {"image1", "image2"}:
                arcname = next(k for k in zip_replacements if k.startswith(f"BinData/{image_id}."))
                write_file(zout, zip_replacements[arcname], arcname)
                continue
            if image_id in extra_ids:
                continue
            local = work_dir.joinpath(*name.split("/"))
            write_file(zout, local, name)
            if image_id == "image3":
                for arcname in sorted(k for k in zip_replacements if bindata_image_id(k) in extra_ids):
                    write_file(zout, zip_replacements[arcname], arcname)

        for arcname, local in zip_replacements.items():
            if arcname not in written:
                write_file(zout, local, arcname)

    if output_hwpx.exists():
        output_hwpx.unlink()
    tmp.replace(output_hwpx)


def run_validation(source_hwpx: Path, output_hwpx: Path) -> None:
    skill_root = Path(__file__).resolve().parents[1]
    validate = skill_root / "scripts" / "validate.py"
    page_guard = skill_root / "scripts" / "page_guard.py"
    if not validate.is_file() or not page_guard.is_file():
        raise FileNotFoundError(f"HWPX validation scripts not found under {skill_root}")
    subprocess.run([sys.executable, str(validate), str(output_hwpx)], check=True)
    subprocess.run([sys.executable, str(page_guard), "--reference", str(source_hwpx), "--output", str(output_hwpx)], check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Replace fixed 강원SW미래채움 HWPX form photos while preserving the fuel photo.")
    parser.add_argument("--source-folder", default=None, help="Folder containing the source .hwpx and replacement images.")
    parser.add_argument("--input-hwpx", default=None, help="Source HWPX filename or path. Defaults to the only .hwpx in source folder.")
    parser.add_argument("--drive-root", default=r"G:\내 드라이브", help="Google Drive Desktop My Drive root.")
    parser.add_argument("--target-folder-name", default="한글테스트", help="Folder to create under drive root.")
    parser.add_argument("--output-name", default=None, help="Output HWPX filename. Defaults to <source stem>_수정.hwpx.")
    parser.add_argument("--signature-images", nargs=2, default=None, help="Two 파이리/signature image filenames or paths.")
    parser.add_argument("--activity-images", nargs="+", default=None, help="Four natural activity image filenames or paths.")
    parser.add_argument("--activity-count", type=int, default=4, help="Number of education activity photos to replace.")
    parser.add_argument("--validate", action="store_true", help="Run the base hwpx validate.py and page_guard.py after writing.")
    args = parser.parse_args()

    source_folder = Path(args.source_folder).resolve() if args.source_folder else default_source_folder().resolve()
    source_hwpx = find_input_hwpx(source_folder, args.input_hwpx).resolve()
    signatures, activities = detect_images(source_folder, args.signature_images, args.activity_images, args.activity_count)

    output_name = args.output_name or f"{source_hwpx.stem}_수정.hwpx"
    output_hwpx = Path(args.drive_root) / args.target_folder_name / output_name

    with tempfile.TemporaryDirectory(prefix="gangwon_hwpx_") as temp_name:
        work_dir = Path(temp_name)
        with zipfile.ZipFile(source_hwpx, "r") as zin:
            zin.extractall(work_dir)

        bindata_dir = work_dir / "BinData"
        replacements: dict[str, str] = {}
        media: dict[str, str] = {}
        zip_replacements: dict[str, Path] = {}

        image_sources = {
            "image1": signatures[0],
            "image2": signatures[1],
            "image4": activities[0],
            "image5": activities[1],
            "image6": activities[2],
            "image7": activities[3],
        }

        for image_id, src in image_sources.items():
            href = f"BinData/{image_id}{src.suffix.lower()}"
            dest = bindata_dir / Path(href).name
            for old in bindata_dir.glob(f"{image_id}.*"):
                old.unlink()
            shutil.copyfile(src, dest)
            replacements[image_id] = href
            media[image_id] = media_type(src)
            zip_replacements[href] = dest

        replace_manifest(work_dir / "Contents" / "content.hpf", replacements, media)
        replace_activity_refs(work_dir / "Contents" / "section1.xml", args.activity_count)
        package_hwpx(source_hwpx, work_dir, output_hwpx, zip_replacements)

    print("Source:", source_hwpx)
    print("Signature images:", ", ".join(p.name for p in signatures))
    print("Activity images:", ", ".join(p.name for p in activities))
    print("Output:", output_hwpx)

    if args.validate:
        run_validation(source_hwpx, output_hwpx)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
