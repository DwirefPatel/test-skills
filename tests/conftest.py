import copy
import json
import shutil
from pathlib import Path

import pytest

from brokerage_1099 import package_from_dict

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "examples" / "sample_package.json"

needs_soffice = pytest.mark.skipif(
    not (shutil.which("soffice") or shutil.which("libreoffice")),
    reason="LibreOffice not installed",
)


@pytest.fixture
def sample_dict():
    return copy.deepcopy(json.loads(SAMPLE.read_text()))


@pytest.fixture
def sample_pkg(sample_dict):
    return package_from_dict(sample_dict)
