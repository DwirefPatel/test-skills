"""Brokerage consolidated 1099 -> preparer Excel workpaper.

Public API (stable for use from a larger build):

    from brokerage_1099 import (
        Package, Lot, BoxTotal, IncomeLine, FormTotal, OtherItem, StateObligation, Payer,
        load_package, package_from_dict, process_package, process_pdf, validate, analyze,
    )
"""

from .analysis import analyze
from .io import load_lots_csv, load_package, merge_overrides, package_from_dict, \
    package_to_dict, save_package
from .models import (
    BoxTotal,
    FormTotal,
    IncomeLine,
    Lot,
    OtherItem,
    Package,
    Payer,
    StateObligation,
)
from .pipeline import RunResult, process_package, process_pdf
from .validation import validate

__version__ = "0.1.0"

__all__ = [
    "BoxTotal", "FormTotal", "IncomeLine", "Lot", "OtherItem", "Package", "Payer",
    "StateObligation", "RunResult", "analyze", "load_lots_csv", "load_package",
    "merge_overrides", "package_from_dict", "package_to_dict", "process_package",
    "process_pdf", "save_package", "validate", "__version__",
]
