# PCS Company mold bases — reference for the classifier

Injected as LLM prompt context and used to derive the deterministic rules in
`qwen_classify_xt_csv.py`. Everything here is about **PCS Company** (Fraser, MI)
standard mold bases, which is what most of the shop's non-BMS work sits on.

Sourced from PCS Company's published mold base documentation. Where a figure is
inferred rather than published it says so.

---

## 1. The six series

PCS standard mold bases come in **A, B, T, AX, 5X and 6X**.

| Series | Plates | What distinguishes it |
|---|---|---|
| **A** | 2-plate | A-clamping plate, A plate, B plate, support plate, rails, ejector set, bottom clamp. The default. |
| **B** | 2-plate | Same as A but **no support plate** — B plate mounts straight onto the rails. |
| **T** | 3-plate | Adds **X-1** (runner stripper) and **X-2** (cavity) floating plates. **Two parting lines.** |
| **AX** | stripper | Stripper plate between AX and BX plates, for parts stripped off the core. |
| **5X** | 5-plate stripper | Stripper series without support plate. |
| **6X** | 6-plate stripper | Stripper series with support plate. |

**32 standard sizes**, from **7-7/8 × 7-7/8"** up to **23-3/4 × 35-1/2"**.
Plate thicknesses run **7/8"** to **5-7/8"**.

### Why the T-series matters for classification

A T-Series is an A-Series with an extra plate between the injection clamp plate
and the A plate — the **X-1 / runner stripper plate**. Both X-1 and **X-2 (the
cavity plate)** stay with the **stationary** half.

It runs **two parting lines**:

1. **First** parting line opens between X-1 and X-2 — this breaks the part off
   the gate before the mold proper opens.
2. **Then** the main parting line opens and the part is free to eject.

Consequences the classifier must respect:

- "Top of stack" is ambiguous on a T-series. There are two parting lines, so
  stack-extreme logic keyed on a single parting line can invert the naming.
- X-2 **is** the cavity plate. It belongs on the A-plate quote row, not on a row
  of its own. (`StdSlotForName` maps `X2 Plate → A`.)
- X-1 is a genuinely separate plate with its own row.

---

## 2. Item-number grammar — this is directly parseable

### A and B series

```
<NOMINAL SIZE><SERIES>-<A-PLATE THK CODE>-<B-PLATE THK CODE>
```

Example: **`1016A-13-37`**
= 9-7/8 × 16", A-Series, A plate 1-3/8" thick, B plate 3-7/8" thick.

### X (stripper) series

```
<NOMINAL SIZE>X<5 or 6>-<AX-PLATE THK CODE>
```

`X` = stripper plate series, then the plate-count series (5 or 6), then the
AX-plate thickness.

### The thickness code — the useful bit

Standard PCS plate thicknesses are **a whole number plus either 3/8 or 7/8**.
The code is the whole number followed by `3` (for 3/8) or `7` (for 7/8):

| Code | Thickness | Code | Thickness |
|---|---|---|---|
| `07` / `7` | 7/8" | `33` | 3-3/8" |
| `13` | 1-3/8" | `37` | 3-7/8" |
| `17` | 1-7/8" | `43` | 4-3/8" |
| `23` | 2-3/8" | `47` | 4-7/8" |
| `27` | 2-7/8" | `53` | 5-3/8" |
| | | `57` | 5-7/8" |

**Rule:** last digit `3` → +3/8", last digit `7` → +7/8". Leading digits are the
whole inches.

### ⚠ Nominal size is NOT actual size

`1016` means nominal 10 × 16 but the plate is **9-7/8 × 16"**. The nominal width
is rounded UP to the next whole inch; the length is actual.

This matters when matching a BOM row to a part number: a 9.875" wide plate is a
"10" series base. Matching on 10.000 will fail.

---

## 3. Plate stackups, top to bottom

Written in the order a classifier walking the stack axis will meet them.

### A-series (2-plate)
```
Top clamp plate  (A-clamping)
A plate          (cavity side)
B plate          (core side)
Support plate
Rails            (2, left and right)
Ejector plate    (thinner, the retainer)
Bottom ejector plate  (thicker backing)
Bottom clamp plate
```

### B-series
As A-series **minus the support plate**.

### T-series (3-plate)
```
Top clamp plate
X-1 plate        (runner stripper)   <-- first parting line below this
X-2 plate        (cavity)
B plate          (core)
Support plate
Rails
Ejector plate
Bottom ejector plate
Bottom clamp plate
```

### AX / 5X / 6X (stripper)
```
Top clamp plate
AX plate
Stripper plate   <-- strips the part off the core
BX plate
[Support plate]  <-- 6X only
Rails
Ejector plate
Bottom ejector plate
Bottom clamp plate
```

### Hot runner (an option on any series)
```
Top clamp plate
Manifold backing plate
Manifold plate
A plate
...
```
A "manifold backing plate" is a **backing plate**, not the manifold plate. Both
`StandardPlateNameStd` and `componentKind.ts` test the compound name first.

---

## 4. Naming: what CMS calls things

The shop's canonical names, which every downstream row and file must use:

| Use this | Never this |
|---|---|
| `A Plate` | `cavity_plate`, `Cavity Plate` |
| `B Plate` | `core_plate`, `Core Plate` |
| `Bottom Ejector Plate` | `Ejector Retainer Plate` (deprecated alias) |
| `"X" Plate` / `"Y" Plate` | `AX Plate` / `BX Plate` in a quote row |
| `X1 Plate` / `X2 Plate` | `X-1`, `X 1` |
| `Rails` (qty 2) | `Spacer Block`, `Parallel`, `Riser` |

Customer BOMs use all the wrong-column spellings freely, so the classifier
accepts them as INPUT and normalises. It must never EMIT them.

> ⚠ **Known inconsistency in this repo.** The YOLO vision datasets
> (`datasets/major8_full_views/data.yaml`, `datasets/right_view_major8/data.yaml`)
> use class names `a_cavity_plate` and `b_core_plate`, contradicting the rule
> above. The vision labels and the tabular role vocabulary are therefore not
> interchangeable. Anything joining the two must map explicitly.

---

## 5. Materials

| Plate | Typical grade |
|---|---|
| A plate, B plate | P20 (pre-hard, ~30 HRC) |
| X-1, X-2, stripper, manifold | P20 |
| Clamp plates, support, rails, ejector set | A36 / 1018 hot-rolled |
| Backing plates | A36 |
| Cavity/core inserts (not base plates) | H13, 420SS, S136 |

`DefaultStandardGradeForSlot` encodes this: slots `A`, `B`, `X`, `Y`,
`MANIFOLD`, `STRIPPER` get the P20-class grade; everything else A36.

DME grade codes appear on some shop sheets and other vendors' BOMs:

| Code | Grade |
|---|---|
| `#1` | A36 |
| `#2` | 4140 |
| `#3` | P20 |
| `#5` | H13 |
| `#7` | 420SS |

`NormalizeSteelType` handles these, plus 4130 → 4140 (the shop treats them
interchangeably for quoting).

---

## 6. Classification signals, ranked

What actually separates plates when the CAD name is unhelpful:

1. **Full footprint vs partial.** Clamp / A / B / support / stripper / X-1 / X-2
   all span the base. Rails, ejector plates and pins do not. This is the single
   strongest signal and it is why `IsStandardStructuralRoleKey` only whitelists
   full-footprint roles for stack-extreme detection.
2. **Position on the stack axis.** Sort full-footprint plates along the stack
   axis and the order is fixed by the series.
3. **Thickness.** Clamp plates are the thinnest full plates (7/8–1-3/8").
   A and B plates are the thickest (2-3/8–5-7/8"). Ejector plates are thin
   (5/8–1").
4. **Count of full plates.** 2-plate A-series has 5 (clamp, A, B, support,
   bottom clamp). Adding two more full plates between the top clamp and the A
   plate says T-series. A full plate between A and B says stripper series.
5. **Rails come in pairs** with identical dimensions and mirrored centres.
6. **Leader pin / bushing plane** tells you which side is stationary — leader
   pins are fixed in the A half and pass through bushings in the B half.

---

## 7. Traps

- **A "clamp plate" with no TOP/BOTTOM qualifier.** One customer (C18503)
  calls the bottom one just "Clamp Plate". That is one customer's habit, not a
  rule. Geometry names the outermost plate correctly from stack position, so do
  not guess from the name. (Deliberately not handled in
  `HardStandardRoleFromCadName`.)
- **`"X" Plate` written with literal quotes.** `CleanFileName` turns `"` into
  `_`, so the STL lands as `J8420__X_ Plate.STL`. The classifier needs a
  boundary-aware test; `CAVITYPLATE` ends in `...Y` + `PLATE` and a loose
  substring test on `YPLATE` misclassified every cavity plate as a BX plate.
- **`NormalizeKey` does not strip double quotes**, so `"X" Plate` normalises to
  `"X"PLATE`, not `XPLATE`. Whitelists must list both.
- **Insulation is never quoted.** Pyropel / HT200 / "insulation" / "insulator"
  are excluded at every capture path.
- **Stack molds** (two complete cavity sets back-to-back) have two parting lines
  like a T-series but a genuinely ambiguous "top". Not yet handled — a stack mold
  will need a human to confirm the naming.

---

## Sources

- PCS Company, *Mold Bases & Plates* catalogue and mold base product
  documentation — series list, 32 sizes, thickness range, item-number grammar,
  thickness codes, T-series description and two-parting-line operation.
- Shop practice for CMS naming, materials by slot and the DME grade codes:
  `Module6121.bas` (`StandardPlateNameStd`, `StdSlotForName`,
  `DefaultStandardGradeForSlot`, `NormalizeSteelType`) and
  `bms_steel_dim_rules.md`.
