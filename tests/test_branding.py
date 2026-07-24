"""Branding smoke tests: product constants and logo asset resolution."""

from __future__ import annotations

from src import branding


def test_product_constants():
    assert branding.PRODUCT_NAME == "Tongkat Gaib Excel"
    assert branding.PRODUCT_BYLINE == "by crimSun"
    assert branding.LOGO_FILENAME == "crimsun_logo.png"


def test_logo_path_resolves_to_existing_png():
    path = branding.logo_path()
    assert path is not None, "canonical logo asset should resolve in a source checkout"
    assert path.exists()
    assert path.name == "crimsun_logo.png"
    # The asset must be a real PNG (Tk PhotoImage cannot load JPEG bytes).
    with path.open("rb") as handle:
        signature = handle.read(8)
    assert signature == b"\x89PNG\r\n\x1a\n", "logo must be a valid PNG for Tk"
