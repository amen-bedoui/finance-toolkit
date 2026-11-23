import json
import os
import shutil
import subprocess
import sys

import pytest
from openpyxl import load_workbook

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
needs_lo = pytest.mark.skipif(SOFFICE is None, reason="LibreOffice requis pour recalculer le classeur")


def build_and_recalc(tmp_path, data=None, blank=False, name="wb.xlsx"):
    out = tmp_path / name
    cmd = [sys.executable, os.path.join(ROOT, "src", "build_workbook.py"), "--out", str(out)]
    if data:
        cmd += ["--data", os.path.join(ROOT, "data", "examples", data)]
    if blank:
        cmd += ["--blank"]
    subprocess.run(cmd, check=True, capture_output=True)
    outdir = tmp_path / "calc"
    outdir.mkdir(exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx", "--outdir", str(outdir), str(out)], check=True, capture_output=True)
    return str(outdir / name)


class Book:
    def __init__(self, path):
        self.wb = load_workbook(path, data_only=True)
        self.ix = {r[0]: (r[1], r[2]) for r in self.wb["_Index"].iter_rows(min_row=2, values_only=True) if r[0]}

    def v(self, key, col="D"):
        sh, row = self.ix[key]
        return self.wb[sh][f"{col}{row}"].value

    def errors(self):
        bad = []
        for ws in self.wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("#"):
                        bad.append(f"{ws.title}!{c.coordinate}={c.value}")
        return bad


@pytest.fixture(scope="session")
def wave(tmp_path_factory):
    if SOFFICE is None:
        pytest.skip("LibreOffice requis")
    return Book(build_and_recalc(tmp_path_factory.mktemp("wave"), "wavestone_public.json"))


@pytest.fixture(scope="session")
def fict(tmp_path_factory):
    if SOFFICE is None:
        pytest.skip("LibreOffice requis")
    return Book(build_and_recalc(tmp_path_factory.mktemp("fict"), "societe_fictive.json"))


@pytest.fixture(scope="session")
def lvmh(tmp_path_factory):
    if SOFFICE is None:
        pytest.skip("LibreOffice requis")
    return Book(build_and_recalc(tmp_path_factory.mktemp("lvmh"), "lvmh_public.json"))


@pytest.fixture(scope="session")
def nvda(tmp_path_factory):
    if SOFFICE is None:
        pytest.skip("LibreOffice requis")
    return Book(build_and_recalc(tmp_path_factory.mktemp("nvda"), "nvidia_public.json"))
