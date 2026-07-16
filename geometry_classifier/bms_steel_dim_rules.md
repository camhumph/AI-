# BMS / Pot-Block Steel Sheet Dimension Rules

Learned from finished J000 Steel Order / Machining Sheet jobs and shop
dimensioned DXF labels (TOP / RIGHT / FRONT views).

## Where Width, Length, and Thickness come from

CMS L/W/T follow the **oriented DXF view frame**, not model XYZ and not
blind `L ≥ W ≥ T` sorting. Holder/pot blocks often have **Thickness as the
largest** size.

| View | Horizontal (X) | Vertical (Y) |
|---|---|---|
| **TOP** (CMS_TOP) | **Width** | **Length** |
| **RIGHT** | **Thickness** | **Length** |
| **FRONT / BOTTOM** | **Width** | (stack / up) |

| Steel sheet column | Meaning | CAD source after view-frame |
|---|---|---|
| **C** (col 3) | **Thickness / Height** | Extent along RIGHT view X |
| **E** (col 5) | **Width** | Extent along TOP view X |
| **G** (col 7) | **Length** | Extent along TOP view Y |
| H (col 8) | Steel type | `#2 4140` |

Macro: `CaptureCmsViewFrameFromModel` → `AssignLengthWidthThicknessFromAxes` →
`ApplyCmsViewDimsToAllParts`. Prefer CAD; BOM is backup only.

QuoteWorksheet stock sizes: Thickness gets +0.25" stock allowance;
Width/Length round up to nickel. Steel Order keeps **finished** sizes.

Plate order on steel sheet (rows starting at 19):
1. TCP
2. ID Holder
3. OD Holder
4. ID Pot
5. OD Pot
6. BCP

## Correct example (vs Tempcraft BOM)

| Plate | Thickness | Width | Length |
|---|---|---|---|
| TCP / BCP | 1.375 | 15.875 | 18.375 |
| ID Holder | 6.875 | 13.875 | 7.000 |
| OD Holder | 5.970 | 13.875 | 7.000 |
| ID Pot | 6.875 | 5.500 | 5.500 |
| OD Pot | 5.970 | 5.500 | 5.500 |

## Tempcraft / Howmet BOM trap (Stock Weight ≠ Length)

Tempcraft Base BOM columns are:

| Lth (in.) | Wth/O.D. (in.) | Hgt/I.D. (in.) | Stock Weight |
|---|---|---|---|
| … | … | … | **lbs — not a size** |

**Never** take Stock Weight as a plate dimension. PDF parse must use the first
three finished-size decimals (Lth/Wth/Hgt), not `PickThreeLargest`.

Tempcraft Lth/Wth/Hgt are **not** CMS T/W/L. When BOM is used as backup,
`MapTempcraftBomDimsToCmsSteel` remaps by plate role:

| Plate | Tempcraft order → CMS |
|---|---|
| TCP / BCP | smallest→T; remaining larger→L, smaller→W |
| Holders | Lth→T, Wth→L, Hgt→W (e.g. `6.875 7.000 13.875` → T6.875 W13.875 L7) |
| Pot blocks | Lth→W, Wth→L, Hgt→T (e.g. `5.500 5.500 6.875` → T6.875 W5.5 L5.5) |

Prefer CAD finished bbox after the view frame is locked; BOM is backup only
after Stock-Weight rejection.
