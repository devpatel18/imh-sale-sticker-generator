"""Regression test: every label must fit inside its box, whatever the input.

Run:  python test_fit.py
Checks the real sample file plus adversarial synthetic cases, verifying with
the same width/height model the generator uses (estimated wrap simulation):
  - total stacked line height <= USABLE_H
  - no single word wider than USABLE_W at the chosen size
  - the price line fits on one line
"""

from label_generator import (
    parse_excel, fit_label_sizes, fit_note_size, format_price,
    _est_width, _longest_word_w, _wrap_count, _line_h,
    USABLE_W, USABLE_H, EXP_PT,
)

NOTE = "No Return, No Exchange on Sale Items — All Sales Final"

ADVERSARIAL = [
    # (brand, name, price, qty) — worst cases we could think of
    ("", "Supercalifragilisticexpialidocious", 9.99, "1 kg"),       # unbreakable word
    ("WWWWWWWW", "WWWW MMMM WWWW MMMM WWWW MMMM", 99.99, "26.04 oz"),  # widest glyphs
    ("International Farmers Market", "Special Reserve Extra Long Grain Aged Basmati Rice Collection", 199.99, "Family Pack"),
    ("Royal Chef's", "Classic/Jeera/Masala/Methi/Whole Wheat Extra Crispy Khari Premium", 3.99, "400 gm"),
    ("X", "Y", 0.01, ""),                                            # tiny
    ("", "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM", 99999.99, "999.99 oz"),  # absurd everything
    ("Brand", "Name " * 30, 5.0, "1 lb"),                            # 30 words
    ("Garvi Gujarat", "Methi Thepla", 2.99, "200 gm"),               # normal
    ("Aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", "Bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb Ccccc", 12.99, "10 lb"),  # long brand
]


def check(label, small_lines, ctx):
    brand_pt, name_pt, price_pt = fit_label_sizes(label, small_lines)
    unit_pt = max(18, round(price_pt * 32 / 72))

    total_h = (_line_h(brand_pt) if label["brand"] else 0) \
        + _wrap_count(label["name"], name_pt) * _line_h(name_pt) \
        + _line_h(price_pt) + small_lines * _line_h(EXP_PT)
    assert total_h <= USABLE_H + 1e-9, \
        f"{ctx}: height {total_h:.2f}\" > {USABLE_H}\" ({label})"

    if label["brand"]:
        assert _est_width(label["brand"], brand_pt) <= USABLE_W + 1e-9, \
            f"{ctx}: brand too wide ({label})"
    if label["name"]:
        assert _longest_word_w(label["name"], name_pt) <= USABLE_W + 1e-9, \
            f"{ctx}: name word too wide ({label})"

    price_w = _est_width("$ ", unit_pt) \
        + _est_width(format_price(label["price"]) + ("/" if label["qty"] else ""), price_pt) \
        + _est_width(label["qty"], unit_pt)
    assert price_w <= USABLE_W + 1e-9, f"{ctx}: price line too wide ({label})"


def main():
    n = 0

    # Note must always be a single line.
    assert _est_width(NOTE, fit_note_size(NOTE)) <= USABLE_W

    for brand, name, price, qty in ADVERSARIAL:
        label = {"brand": brand, "name": name.strip(), "price": price, "qty": qty}
        for small_lines in (0, 1, 2):
            check(label, small_lines, "adversarial")
            n += 1

    try:
        sample = parse_excel("sample_sale_list.xlsx")
    except FileNotFoundError:
        sample = []
    for label in sample:
        for small_lines in (0, 1, 2):
            check(label, small_lines, "sample")
            n += 1

    print(f"OK — {n} label/option combinations all fit "
          f"({len(ADVERSARIAL)} adversarial + {len(sample)} sample labels)")


if __name__ == "__main__":
    main()
