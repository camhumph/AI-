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

## Both pot rows are often named the SAME on the BOM

Real BOMs do not qualify the pot blocks. From a live job:

| # | Description | Lth | Wth | Hgt |
|---|---|---|---|---|
| 103 | Top Holder Block Material | 6.875 | 8.000 | 13.875 |
| 104 | Bottom Holder Block Material | 5.970 | 8.000 | 13.875 |
| 105 | **Pot Block Material** | 5.500 | 5.500 | 6.875 |
| 106 | **Pot Block Material** | 5.500 | 5.500 | 5.970 |

Rows 105/106 carry no Top/Bottom/ID/OD token, so `IsLikelyIdSideName` and
`IsLikelyOdSideName` both fail and `StandardPlateName` returns the generic
`POT BLOCK` — which is not one of the six roles. It then survives
`CanonicalHolderQuoteName` unchanged and lands in
`MapTempcraftBomDimsToCmsSteel`'s `Case Else`, which **blind-sorts** the
dimensions. That produced a phantom seventh steel row reading
`Pot Block Material  T=5.5  W=5.5  L=6.875` — the same block as ID Pot
(`T=6.875 W=5.5 L=5.5`) with its axes scrambled — carrying 67.4 hours and
$126.66 of steel that does not exist, alongside the ID/OD Pot rows that were
already correct.

**Ground thickness tells them apart**, and this BOM is a second independent
confirmation of the pairing rule below: holder 6.875 pairs with pot Hgt 6.875,
holder 5.970 with pot Hgt 5.970. `ResolveUnqualifiedPotBomRows` runs before
`BuildExportRowsFromBom` and renames both rows, so every downstream consumer
sees a qualified name and the right axis mapping.

Watch the column each side is read from — in Tempcraft file order the stack
dimension is **Lth for a holder** and **Hgt for a pot**. `POTBLOCK`/`POT` are
also listed in the pot branch of `MapTempcraftBomDimsToCmsSteel` as a backstop:
getting the side wrong costs a label, getting the axes wrong costs a mis-cut
block.

## Naming pots from B-rep features (not just bounding boxes)

Bounding boxes and center positions split each pair **independently** —
`AssignPairTopBottom` re-derives its own "dominant separation axis" per pair. So
the holders can be split along Y while the pots, offset laterally as pots in a
multi-cavity base usually are, get split along X, and nothing notices the two
answers disagree about which end is the top. The pot pair then comes back
swapped and ID Pot is cut to OD Pot's thickness.

Three measurements settle it. Macro: `MeasureBmsPotFeaturesAlongStackAxis` →
`RefineBmsRolesFromFeatureEvidence`, evidence in
`BMS_Pot_Feature_Evidence.csv`.

1. **Ground thickness pairs a pot to its holder.** A pot is set into its holder
   and the two are ground together, so a pot and its own holder measure the
   **same** along the stack axis — 6.875 for the ID pair, 5.970 for the OD pair
   in the reference job above. A 0.9" gap between pairings, far outside any
   tolerance. This is independent of position, which is exactly why it is
   trusted first.
2. **Nesting confirms it.** The pot sits inside the holder's opening: centers
   agree in both lateral axes and the holder's missing end-face area is at least
   the pot's cross-section. When thickness and nesting **disagree**, nothing is
   changed and the run logs a `BMS ROLE WARNING` — an ambiguous pairing is a
   thing to check against the print, not to guess at.
3. **Facing openings are the parting line.** The two molding halves present
   their cavities to each other: measured against one shared axis, the lower
   block opens toward +stack and the upper toward −stack. No other pair of
   blocks in a pot base does this. If the parting line does **not** land between
   the two pots, the pots are not the two pots — logged as a warning.

Two traps this had to work around:

- **The part's own thickness axis is the wrong axis for a pot.** The `Hs*`
  pocket columns split on `HoleSigThicknessAxis`, the part's *smallest* box
  extent. For a 1.375 × 15.875 × 18.375 TCP that is the stack direction; for a
  5.500 × 5.500 × 6.875 ID Pot it is 5.500 — two axes tied, neither of them the
  stack — so `PocketAreaUp`/`Dn` describe openings in the pot's **side**, not its
  molding face. Pot evidence must be measured against an axis handed in from the
  assembly.
- **Openings must be judged on asymmetry, not size.** An opening is measured as
  end-face area missing from the bounding-box cross-section, so any block that is
  not a rectangular prism reports an opening it does not have: a Ø5.500 round pot
  in a 5.500 × 5.500 box loses 21% of the box area at **both** ends purely to the
  corners. On a size test alone every round pot has two "open" faces. A molding
  cavity is on one end, so the honest signal is the difference between a block's
  two ends — the corner artifact, identical at both, cancels out.

Pot blocks must also never be dropped from the feature walk. The
`HOLE_SIG_MIN_FOOTPRINT_FRAC` gate is written for mold plates; the reference
job's 5.500 × 5.500 ID Pot is 10.4% of a 15.875 × 18.375 base, inside the 10%
gate by four tenths of one percent. Anything `IsPotBlockGeometry` calls a pot is
now walked regardless of footprint, and pots are promoted ahead of the
`HOLE_SIG_MAX_PARTS` cap.
