"""Branding smoke tests: product constants."""

from __future__ import annotations

from src import branding


def test_product_constants():
    assert branding.PRODUCT_NAME == "Tongkat Gaib Excel"


def test_contact_credentials():
    assert branding.AUTHOR_NAME == "Clairine Safa Naqayya"
    assert branding.AUTHOR_EMAIL == "cnaqayya@gmail.com"
    assert branding.AUTHOR_PHONE == "+628118078807"
    credit = branding.contact_credit_text()
    assert "by: Clairine Safa Naqayya" in credit
    assert "email: cnaqayya@gmail.com" in credit
    assert "number: +628118078807" in credit
