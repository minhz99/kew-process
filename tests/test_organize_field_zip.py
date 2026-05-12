import io
import os
import shutil
import tempfile
import zipfile

import pandas as pd

from modules.kew.organize_field_zip import process_field_zip_bytes


def _make_minimal_zip(tmp_root: str) -> bytes:
    """ZIP: Excel + S0001/ + BMP ở gốc."""
    excel_path = os.path.join(tmp_root, "bang.xlsx")
    df = pd.DataFrame(
        [
            {
                "STT": 1,
                "Tên thiết bị": "Bom test",
                "File": "0001",
                "IMG": 1,
                "IMG end": 2,
            }
        ]
    )
    df.to_excel(excel_path, index=False, engine="openpyxl")

    sdir = os.path.join(tmp_root, "S0001")
    os.makedirs(sdir, exist_ok=True)
    with open(os.path.join(sdir, "note.txt"), "w", encoding="utf-8") as f:
        f.write("record")

    for n in (1, 2):
        name = f"PS-SD{n:03d}.BMP"
        with open(os.path.join(tmp_root, name), "wb") as f:
            f.write(b"BMP")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(excel_path, "bang.xlsx")
        zf.write(os.path.join(sdir, "note.txt"), "S0001/note.txt")
        for n in (1, 2):
            zf.write(os.path.join(tmp_root, f"PS-SD{n:03d}.BMP"), f"PS-SD{n:03d}.BMP")
    buf.seek(0)
    return buf.read()


def test_process_field_zip_happy_path():
    src = tempfile.mkdtemp(prefix="zip_in_")
    try:
        zbytes = _make_minimal_zip(src)
    finally:
        shutil.rmtree(src, ignore_errors=True)

    work = tempfile.mkdtemp(prefix="kew_org_test_")
    try:
        out_zip, warnings, fatal = process_field_zip_bytes(zbytes, work)
        assert not fatal
        assert os.path.isfile(out_zip)
        with zipfile.ZipFile(out_zip, "r") as zf:
            names = zf.namelist()
        assert any(n.endswith("Project_Output/Bom test/note.txt") for n in names)
        assert any(n.endswith("Project_Output/Bom test/PS-SD001.BMP") for n in names)
        assert any(n.endswith("Project_Output/Bom test/PS-SD002.BMP") for n in names)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def test_overlap_error():
    src = tempfile.mkdtemp()
    try:
        excel = os.path.join(src, "b.xlsx")
        pd.DataFrame(
            [
                {"Tên thiết bị": "A", "File": "1", "IMG": 1, "IMG end": 3},
                {"Tên thiết bị": "B", "File": "2", "IMG": 2, "IMG end": 4},
            ]
        ).to_excel(excel, index=False, engine="openpyxl")
        os.makedirs(os.path.join(src, "S0001"))
        os.makedirs(os.path.join(src, "S0002"))
        for sid in ("S0001", "S0002"):
            for n in range(1, 5):
                open(os.path.join(src, sid, f"x{n}.txt"), "w").close()
        for n in range(1, 5):
            open(os.path.join(src, f"PS-SD{n:03d}.BMP"), "wb").write(b"x")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.write(excel, "b.xlsx")
            for n in range(1, 5):
                zf.write(os.path.join(src, f"PS-SD{n:03d}.BMP"), f"PS-SD{n:03d}.BMP")
            for sid in ("S0001", "S0002"):
                for n in range(1, 5):
                    zf.write(os.path.join(src, sid, f"x{n}.txt"), f"{sid}/x{n}.txt")
        zbytes = buf.getvalue()
    finally:
        shutil.rmtree(src, ignore_errors=True)

    work = tempfile.mkdtemp()
    try:
        _, _, fatal = process_field_zip_bytes(zbytes, work)
        assert fatal
        assert any("Trùng dải ảnh" in e for e in fatal)
    finally:
        shutil.rmtree(work, ignore_errors=True)
