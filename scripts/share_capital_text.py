"""Indian-style number formatting + share-capital description builder.

Turns numeric inputs (shares + nominal value) into the formal phrase used in
the Director's Report Share Capital table:

    Rs. 1,00,000/- (Rupees One Lakh divided into 10,000 Equity Shares
    of Rs. 10/- each)
"""

_UNITS = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen",
    "Sixteen", "Seventeen", "Eighteen", "Nineteen",
]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _two_digits(n):
    if n < 20:
        return _UNITS[n]
    return _TENS[n // 10] + (" " + _UNITS[n % 10] if n % 10 else "")


def _three_digits(n):
    if n < 100:
        return _two_digits(n)
    out = _UNITS[n // 100] + " Hundred"
    if n % 100:
        out += " " + _two_digits(n % 100)
    return out


def indian_words(n):
    """Non-negative integer → Indian-English words.
    100000 → 'One Lakh', 22500000 → 'Two Crore Twenty Five Lakh'."""
    n = int(n)
    if n == 0:
        return "Zero"
    if n < 0:
        return "Minus " + indian_words(-n)

    crore = n // 10_000_000
    n %= 10_000_000
    lakh = n // 100_000
    n %= 100_000
    thousand = n // 1000
    rest = n % 1000

    parts = []
    if crore:
        parts.append(_two_digits(crore) + " Crore")
    if lakh:
        parts.append(_two_digits(lakh) + " Lakh")
    if thousand:
        parts.append(_two_digits(thousand) + " Thousand")
    if rest:
        parts.append(_three_digits(rest))
    return " ".join(parts)


def indian_comma(n):
    """Indian-style grouping. 1234567 → '12,34,567'."""
    s = str(int(n))
    neg = s.startswith("-")
    if neg:
        s = s[1:]
    if len(s) <= 3:
        return ("-" if neg else "") + s
    last3 = s[-3:]
    rest = s[:-3]
    pairs = []
    while len(rest) > 2:
        pairs.append(rest[-2:])
        rest = rest[:-2]
    if rest:
        pairs.append(rest)
    return ("-" if neg else "") + ",".join(reversed(pairs)) + "," + last3


def _format_amount(amt):
    """100000 → '1,00,000';  100000.5 → '1,00,000.50'."""
    if amt == int(amt):
        return indian_comma(int(amt))
    integer = int(amt)
    paise = int(round((amt - integer) * 100))
    return indian_comma(integer) + f".{paise:02d}"


def _format_nominal(n):
    """Nominal-per-share can be 10 or 0.50."""
    if n == int(n):
        return str(int(n))
    return f"{n:.2f}"


def _amount_words(amt):
    """Words for an amount that may have paise."""
    integer = int(amt)
    paise = int(round((amt - integer) * 100))
    out = indian_words(integer)
    if paise:
        out += f" and {indian_words(paise)} Paise"
    return out


def build_description(eq_shares=0, eq_nominal=0, pf_shares=0, pf_nominal=0):
    """Assemble the Rs. X/- (Rupees ... divided into ...) sentence.
    Returns '' if neither equity nor preference is filled."""
    eq_shares = float(eq_shares or 0)
    eq_nominal = float(eq_nominal or 0)
    pf_shares = float(pf_shares or 0)
    pf_nominal = float(pf_nominal or 0)

    has_eq = eq_shares > 0 and eq_nominal > 0
    has_pf = pf_shares > 0 and pf_nominal > 0
    if not has_eq and not has_pf:
        return ""

    segments = []
    total = 0.0
    if has_eq:
        total += eq_shares * eq_nominal
        segments.append(
            f"{indian_comma(eq_shares)} Equity Shares of Rs. {_format_nominal(eq_nominal)}/- each"
        )
    if has_pf:
        total += pf_shares * pf_nominal
        segments.append(
            f"{indian_comma(pf_shares)} Preference Shares of Rs. {_format_nominal(pf_nominal)}/- each"
        )

    return (
        f"Rs. {_format_amount(total)}/- "
        f"(Rupees {_amount_words(total)} divided into {' and '.join(segments)})"
    )


if __name__ == "__main__":
    samples = [
        (10_000, 10, 0, 0),
        (22_50_000, 10, 0, 0),
        (1_05_000, 0.5, 0, 0),
        (10_000, 10, 5_000, 100),
    ]
    for args in samples:
        print(args, "→", build_description(*args))
