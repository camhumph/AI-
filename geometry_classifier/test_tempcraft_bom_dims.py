"""Regression: Tempcraft BOM Stock Weight must never become Length.

Mirrors Module6121 PickThreeFinishedSizeDims / BomDimsLookLikeStockWeight.
"""

from __future__ import annotations


def extract_decimal_numbers(s: str) -> list[float]:
    nums: list[float] = []
    token = ""
    has_dot = False
    for ch in s + " ":
        if ch.isdigit() or ch == ".":
            token += ch
            if ch == ".":
                has_dot = True
        else:
            if token and has_dot:
                try:
                    nums.append(float(token))
                except ValueError:
                    pass
            token = ""
            has_dot = False
    return nums


def pick_three_finished_size_dims(nums: list[float]) -> tuple[float, float, float] | None:
    cand = [v for v in nums if 0.05 < v < 80.0]
    if len(cand) >= 3:
        return cand[0], cand[1], cand[2]
    if len(nums) >= 3:
        return nums[0], nums[1], nums[2]
    return None


def sort_three(a: float, b: float, c: float) -> tuple[float, float, float]:
    arr = sorted([a, b, c], reverse=True)
    return arr[0], arr[1], arr[2]  # L, W, T


def looks_like_stock_weight(t: float, w: float, l: float) -> bool:
    if t <= 0 or w <= 0 or l <= 0:
        return False
    if l >= 50 and l > (w * 2.5) and l > (t * 8):
        return True
    if (t * w * l) > 8000 and l > 40:
        return True
    return False


def parse_tempcraft_line(raw: str) -> tuple[float, float, float] | None:
    nums = extract_decimal_numbers(raw)
    picked = pick_three_finished_size_dims(nums)
    if not picked:
        return None
    l, w, t = sort_three(*picked)
    if looks_like_stock_weight(t, w, l):
        return None
    return t, w, l


def test_tcp_smed_ignores_stock_weight():
    line = (
        "101 Top SMED Material 1 4140 Holder Pre-Hardened "
        "1.375 15.875 18.000 G1C-5005010 Lbs 117.87"
    )
    # Old bug: PickThreeLargest -> T=15.875 W=18 L=117.87
    dims = parse_tempcraft_line(line)
    assert dims is not None
    t, w, l = dims
    assert abs(t - 1.375) < 1e-6
    assert abs(w - 15.875) < 1e-6
    assert abs(l - 18.0) < 1e-6
    assert l < 50


def test_holder_block_dims():
    line = (
        "103 Top Holder Block Material 1 4140 Holder Pre-Hardened "
        "6.875 7.000 13.875 G1C-5005010 Lbs 200.32"
    )
    dims = parse_tempcraft_line(line)
    assert dims is not None
    t, w, l = dims
    assert abs(t - 6.875) < 1e-6
    assert abs(w - 7.0) < 1e-6
    assert abs(l - 13.875) < 1e-6


def test_pot_block_dims():
    line = (
        "105 Top Pot Block Material 1 4140 Holder Pre-Hardened "
        "5.500 5.500 6.875 G1C-5005010 Lbs 62.39"
    )
    dims = parse_tempcraft_line(line)
    assert dims is not None
    t, w, l = dims
    assert abs(t - 5.5) < 1e-6
    assert abs(w - 5.5) < 1e-6
    assert abs(l - 6.875) < 1e-6


def test_stock_weight_triplet_rejected():
    # If only weight-like numbers remain, reject.
    assert looks_like_stock_weight(15.875, 18.0, 117.87) is True
    assert looks_like_stock_weight(1.375, 15.875, 18.0) is False


if __name__ == "__main__":
    test_tcp_smed_ignores_stock_weight()
    test_holder_block_dims()
    test_pot_block_dims()
    test_stock_weight_triplet_rejected()
    print("OK: Tempcraft BOM dim regression tests passed")
