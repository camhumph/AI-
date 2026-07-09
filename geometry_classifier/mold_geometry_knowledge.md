# CMS Mold Base Geometry Knowledge

Use this knowledge for interpreting `XT_Export_CAD_Dimensions.csv`.

## General Rule

Do not blindly trust CAD component names for final classification. Customer CAD names can be copied, stale, swapped, or misleading. Use geometry first:

- Thickness, width, length
- Bounding-box volume
- Center X/Y/Z
- Stack order
- Full-footprint plate count
- Rails and ejector-side relationship
- Leader pins, bushings, support pillars, return pins
- Parting-line relationship

Important exception: exact shop-standard tokens in an imported STEP can be strong evidence when the geometry agrees. Examples:

- `A-PLATE`, `A_PLATE` -> A Plate
- `B-PLATE`, `B_PLATE` -> B Plate
- `SC-RETAINER-PLATE` -> SC Retainer Plate
- `SC-BACKUP-PLATE` -> SC Backup Plate
- `EJ-RET-PLATE` -> Ejector Plate in CMS naming (the thinner ejector-stack plate)
- `EJ-BACKUP-PLATE` -> Bottom Ejector Plate in CMS naming (the thicker/lower ejector-stack plate; never call this or the thinner plate "Ejector Retainer Plate")
- `RAIL-BOTTOM`, `RAIL-TOP`, `RAIL` -> Rails
- `LDR-PIN` -> Leader Pin
- `LBB` -> Leader Pin Bushing
- `PLC75`, `LATCH-LOCK`, `SAFETY-STRAP` (also seen misspelled `SAFTEY-STRAP`) -> sequenced/latch-lock standard base clues

Treat those exact tokens differently from vague copied names like "plate", "block", "base", or stale assembly names.

## Stack Orientation: Bottom-Up Anchoring

Decide the bottom of the mold stack from the rails and ejector-stack plates
first. Rails and the ejector assembly are the primary, most reliable
orientation signal. Leader pins and bushings only decide stack orientation
when rails/ejector plates are missing or ambiguous.

Do not let leader-pin direction flip a stack orientation, or an A/B plate
identity, that is already clear from the rail/ejector stack or from a strong
shop-name token (`A-PLATE`, `B-PLATE`). On plate-sequenced or Stripper-Core
(SC) bases, leader pins can seat in the B-plate area and run upward toward the
A-side (reversed/seated leader pins) -- this must never force an incorrect
A/B flip.

## Quote Row Mapping

Ejector-stack and pin-plate rows must never be mapped or merged into the
A Plate row in the quoting workbook. Keep A Plate, B Plate, ejector_plate,
bottom_ejector_plate, and pin_plate as distinct quote rows.

## Standard DME / PCS Mold Base Stack

A common standard mold base is top-to-bottom:

1. Top Clamp Plate
2. Cavity Plate / A Plate
3. Core Plate / B Plate
4. Support Plate
5. Rails / risers beside the ejector box
6. Pin Plate / Ejector Retainer
7. Ejector Plate
8. Bottom Clamp Plate

Some jobs may omit a top clamp, include a stripper plate, include a manifold plate, or have extra plates. Use stack order and hardware position to decide.

## Sequenced / Latch-Lock Standard Bases

Some standard bases are not the simple A/B/support layout. If the component list contains latch-lock parts such as `PLC75`, `LATCH-LOCK`, or `SAFETY-STRAP`, classify it as a sequenced or latch-lock standard base.

These bases may have:

1. A Plate
2. B Plate
3. SC Retainer Plate
4. SC Backup Plate
5. Bottom Clamp Plate
6. Rails
7. Ejector Retainer Plate / ejector backing plate
8. Ejector Plate

Do not force `B Plate` to become a stripper/support plate just because extra full-footprint SC plates appear below it. In a latch-lock/SC base, the named `B-PLATE` is a strong anchor when leader pins or B-side bushings line up with it.

The latch locks help identify secondary parting/opening lines. Leader pins/bushings identify guide direction; latch locks identify which plates open together or in sequence.

## Full-Footprint Plates

Full plates have nearly the same width and length as the mold base footprint.

Typical roles:

- Highest full-footprint plate: Top Clamp Plate
- Next full-footprint plate toward parting line: Cavity Plate
- Next full-footprint plate toward ejector side: Core Plate
- Below core side: Support Plate
- Lowest full-footprint plate: Bottom Clamp Plate

For a 5-full-plate standard stack:

1. Top Clamp Plate
2. Cavity Plate
3. Core Plate
4. Support Plate
5. Bottom Clamp Plate

For a 5-full-plate sequenced SC stack with no top clamp:

1. A Plate
2. B Plate
3. SC Retainer Plate
4. SC Backup Plate
5. Bottom Clamp Plate

## Parting Line

The main parting line is usually the interface between A Plate and B Plate.

- A side is toward A-side bushings and the stationary side.
- B side is toward rails/ejector stack and usually the leader pins.
- In latch-lock/SC bases, there can also be secondary opening/parting lines around SC Retainer / SC Backup.

## Rails

Rails are long narrow side blocks near the ejector side.

Geometry clues:

- Long dimension close to base length
- Width much smaller than full base width
- Usually two rails
- Located left/right of the ejector stack
- Do not classify rails as ejector plates

## Ejector Stack

Ejector stack plates are long plates inside/between the rails.

Typical geometry:

- Long dimension close to base length
- Width smaller than full base width
- Usually narrower than top/core/support/full plates
- Pin Plate is above the Ejector Plate in the stack

CMS ejector-plate naming rule: in the ejector assembly, the thinner plate is
always the **Ejector Plate**. The thicker/lower plate is the **Bottom Ejector
Plate**. Do not name the thinner plate "Ejector Retainer Plate" -- that name
is deprecated; only "Ejector Plate" and "Bottom Ejector Plate" should be used
in CMS output.

Small cavity/core inserts are not ejector plates even if their area is large enough. They are usually not long across the base.

## Leader Pins

Leader pins are round long pins.

Geometry clues:

- Two equal small dimensions form diameter
- Long axis is much longer than diameter
- Often 4 pins on the main mold base
- Can be multiple sets when ejector guidance also exists
- Larger than return/ejector pins

Leader pins help identify B-side orientation, but do not use them alone when a latch-lock/SC base has multiple pin/bushing sets. Match leader-pin bottoms and bushing stacks to the plate centers. In T001015-style bases, the primary leader pins align with the B Plate while longer/shorter bushings appear in both A and B regions.

## Leader / Shoulder Bushings

Bushings are short round cylinders.

Geometry clues:

- Two equal large dimensions form diameter
- Short axis is the bushing height
- Typically near leader-pin locations
- Often 4 bushings

Bushings help identify A/cavity side orientation.

## Guided Ejector Bushings

Guided ejector bushings are short round bushings near the ejector-side stack.

Use position relative to ejector plates and rails when possible.

## Return Pins / Ejector Pins

Return pins and ejector pins are smaller long round pins.

Geometry clues:

- Smaller diameter than leader pins
- Long axis much longer than diameter
- Usually associated with ejector side

## Support Pillars

Support pillars are large round posts between support/bottom regions.

Geometry clues:

- Large diameter
- Long post length
- Often repeated many times
- Do not classify as leader pins

## Recommended AI Behavior

The AI should explain uncertain classifications, but the quote should be filled from deterministic geometry rules when confidence is high.

For uncertain rows, output:

- Index
- Suggested role
- Confidence
- Geometry reason
- What would confirm it

## Known Example: T001015 Moldbase

This is a standard base, but not a simple DME A/B/support stack. It is a sequenced/latch-lock SC-style base.

Correct classification:

- `T001015_A-PLATE` -> A Plate
- `T001015_B-PLATE` -> B Plate
- `T001015_SC-RETAINER-PLATE` -> SC Retainer Plate
- `T001015_SC-BACKUP-PLATE` -> SC Backup Plate
- `T001015_CLAMP-PLATE` -> Bottom Clamp Plate
- `T001015_RAIL-BOTTOM` / `T001015_RAIL-TOP` -> Rails
- `T001015_EJ-BACKUP-PLATE` -> Bottom Ejector Plate (thicker/lower ejector-stack plate)
- `T001015_EJ-RET-PLATE` -> Ejector Plate (thinner ejector-stack plate)
- `T0010115_LATCH-LOCK_ASM-*` / `T001015_PLC75-*` -> latch-lock hardware; confirms this is a plate-sequenced/latch-lock standard base with secondary opening/parting lines, not a plain A/B/support stack
- `PROGRESSIVE-SAFETY-STRAP_ASM-*` (component itself spelled `PROGRESSIVE-SAFTEY-STRAP`) -> Progressive Components safety strap, same latch-lock/sequenced-base clue

Why B Plate is B Plate:

- `T001015_B-PLATE` is an exact shop token.
- Primary leader pins are centered at roughly the same stack position as B Plate.
- `SC-RETAINER` and `SC-BACKUP` are extra sequenced plates below B, not replacements for B.

## BMS / Pot-Block Steel Sheet Dimensions (W × L × H)

BMS jobs do **not** use the AI classifier. The macro fills the J000 Steel Order /
Machining Sheet from CAD bbox dims via `SortThreeDimensions` (always **L ≥ W ≥ T**).

| Steel sheet column | Meaning | CAD source |
|---|---|---|
| **C** | **Thickness / Height** | Smallest bbox dim (`Thickness`) |
| **E** | **Width** | Middle bbox dim (`Width`) |
| **G** | **Length** | Largest bbox dim (`Length`) |
| H | Steel type | `#2 4140` |

Never swap Width and Length. Never put thickness into the Length column.
QuoteWorksheet stock sizes add +0.25" to thickness; steel sheet keeps finished sizes.
Training writes `geometry_classifier/data/training/bms_steel_dim_rules.md` when BMS jobs are scanned.
