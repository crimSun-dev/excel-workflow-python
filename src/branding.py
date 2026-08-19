"""Product branding constants.

Single source of truth for the product name and operator-facing credits so the
GUI, packaging script, and docs cannot drift.
"""

from __future__ import annotations

PRODUCT_NAME = "Tongkat Gaib Excel"

AUTHOR_NAME = "Clairine Safa Naqayya"
AUTHOR_EMAIL = "cnaqayya@gmail.com"
AUTHOR_PHONE = "+628118078807"


def contact_credit_text() -> str:
    """Labeled name / email / phone block shown in the GUI header."""
    return (
        f"by: {AUTHOR_NAME}\n"
        f"email: {AUTHOR_EMAIL}\n"
        f"number: {AUTHOR_PHONE}"
    )
