# Vendor Knowledge Sources For CMS Geometry AI

These are source-backed rules and references for teaching the geometry AI about DME/PCS-style mold bases and components.

## Official / Primary Sources

- DME homepage and digital catalog area: https://www.dme.net/
- DME Inch Shouldered Leader Pin: https://store.dme.net/inch-shouldered-leader-pin
- DME Leader Pins FAQ: https://www.dme.net/leader-pins-faq/
- PCS Mold Components: https://www.pcs-company.com/mold-components
- PCS Leader Pins: https://www.pcs-company.com/leader-pins
- PCS Straight Leader Pins - Standard: https://www.pcs-company.com/straight-leader-pins-standard
- PCS Bushings: https://www.pcs-company.com/bushings
- PCS General Mold Components: https://www.pcs-company.com/general-mold-components

## Source-Backed Knowledge

PCS states that leader pins provide initial alignment of cavity and core halves. PCS standard straight leader pins are made from 4150 steel and mate with PCS standard bushings.

DME describes shouldered leader pins as precision steel components, hardened and finish ground to close tolerances.

DME's leader-pin FAQ says collars on metric leader pins align the top clamp plate to the A plate. That matters because leader pins and bushing/collar geometry can help identify the stationary/A-side region of a standard mold base.

PCS describes bushings as available in shoulder, straight, guided ejection, die bushing, and ball guided ejection styles. PCS also groups mold components such as leader pins, bushings, support pillars, ejector items, fasteners, locating rings, springs, and sprue bushings as standard mold components.

## Geometry Interpretation Rules

Exact shop-standard tokens on imported STEP/CAD component names are STRONG
ANCHOR EVIDENCE, not weak notes. When a component name carries a deliberate
shop token, trust it over generic bounding-box geometry and use geometry only
to confirm/sanity-check it:

- `A-PLATE`, `A_PLATE` -> A Plate
- `B-PLATE`, `B_PLATE` -> B Plate
- `SC-RETAINER-PLATE` -> SC Retainer Plate
- `SC-BACKUP-PLATE` -> SC Backup Plate
- `CLAMP-PLATE` -> Bottom Clamp Plate
- `EJ-RET-PLATE` -> Ejector Plate (the thinner ejector-stack plate)
- `EJ-BACKUP-PLATE` -> Bottom Ejector Plate (the thicker/lower ejector-stack plate)
- `RAIL`, `RAIL-TOP`, `RAIL-BOTTOM` -> Rails
- `LDR-PIN` -> Leader Pin
- `LBB` -> Leader Pin Bushing
- `PLC75`, `LATCH-LOCK`, `SAFETY-STRAP` (any spelling, including the common
  `SAFTEY-STRAP` misspelling seen in real shop CAD) -> plate-sequenced /
  latch-lock standard base clue

Only fall back to pure bounding-box geometry when a component name is generic,
stale, copied, or missing entirely (e.g. "plate", "block", raw McMaster/DME/PCS
catalog part numbers with no shop prefix).

Always use `A Plate` and `B Plate` in CMS output naming. Never use
`cavity_plate` or `core_plate`.

Leader pins:

- Round long geometry
- Diameter from two equal/similar dimensions
- Axis length much longer than diameter
- Usually repeated symmetrically
- Used to align cavity/core mold halves
- Main leader pins help define the mold alignment system

Shoulder / leader bushings:

- Short round geometry
- Diameter from two equal/similar large dimensions
- Short axis is bushing height
- Usually near matching leader-pin locations
- Their side is commonly the A/cavity/stationary receiver side in standard mold layouts, but confirm with stack and ejector-side geometry

Guided ejector bushings:

- Short round bushing geometry
- Located near ejector side / ejector stack
- Use center position relative to pin plate, ejector plate, rails, and return pins

Support pillars:

- Larger round posts
- Repeated in support/ejector area
- Bigger diameter than normal return pins/ejector pins
- Should not be confused with leader pins

Return/ejector pins:

- Smaller long round pins
- Usually associated with ejector side
- Diameter smaller than leader pins/support pillars

Rails:

- Long rectangular side blocks
- Long dimension close to mold base length
- Width less than full mold width
- Usually a left/right pair
- Ejector stack is between or adjacent to them

Ejector stack naming (CMS convention):

- The thinner plate in the ejector stack is always the **Ejector Plate**.
- The thicker/lower plate in the ejector stack is the **Bottom Ejector Plate**.
- Do not name the thinner plate "Ejector Retainer Plate". That name is
  deprecated in CMS output; use Ejector Plate / Bottom Ejector Plate only.
- Ejector-stack and pin-plate rows must never be mapped or merged into the
  A Plate row in the quoting workbook.

Stack orientation / bottom-up anchoring:

- Decide the bottom of the mold stack from the rails and ejector-stack plates
  first. Rails and the ejector assembly are the primary, most reliable
  orientation signal.
- Leader pins and bushings only decide stack orientation when rails/ejector
  plates are missing or ambiguous.
- Do not let leader-pin direction flip a stack orientation that is already
  clear from the rail/ejector stack.
- On plate-sequenced or Stripper-Core (SC) bases, leader pins can seat in the
  B-plate area and run upward toward the A-side (reversed/seated leader pins).
  This must never be allowed to force an incorrect A/B flip.

Latch locks / plate-sequenced standard bases:

- Components with tokens like `PLC`, `LATCH-LOCK`, `SAFETY-STRAP` (or
  Progressive Components naming such as `PLC75`, `T0010115_LATCH-LOCK_ASM`)
  identify a plate-sequenced / latch-lock standard base, not a plain
  A/B/support stack.
- Leader pins tell us guide direction. Latch locks tell us which plates open
  together or in sequence, and define secondary parting/opening lines (for
  example around a stripper/core split).
- Use latch attachment positions to map secondary opening splits; do not
  ignore them as generic hardware.
- These bases may also include SC Retainer Plate and SC Backup Plate, seen
  between B Plate and Bottom Clamp Plate.

Parting line:

- Interface between cavity/A plate and core/B plate
- Determine from full-footprint stack order plus A/B side evidence
- Leader pins/bushings help orient sides, but the full plate stack and ejector side still matter
- Plate-sequenced/latch-lock bases can have secondary opening/parting lines
  identified by latch-lock attachment positions, in addition to the main
  A Plate / B Plate parting line

## What The AI Should Learn

For each job, the AI should first define:

- Stack axis
- Full-footprint plate order
- Base footprint
- Main leader-pin/bushing pattern
- Ejector-side hardware pattern
- Rails
- Parting line
- Candidate A/cavity and B/core sides

Then it should classify rows into quote roles.

## What The AI Should Not Do

- Do not assume a generic/stale/copied CAD component name is correct just because it exists; but DO trust exact shop-standard tokens (A-PLATE, B-PLATE, LDR-PIN, PLC75, etc.) as strong anchors.
- Do not classify row-by-row without seeing the whole mold.
- Do not quote cavity/core inserts as standard base plates.
- Do not classify rails as ejector plates.
- Do not classify support pillars as leader pins.
- Do not classify short bushings as long leader pins.
- Do not let leader-pin direction flip an A/B/stack orientation already established by rails/ejector stack or by strong shop-name tokens.
- Do not name the thinner ejector-stack plate "Ejector Retainer Plate" -- it is the Ejector Plate. The thicker/lower plate is the Bottom Ejector Plate.
- Do not treat PLC/latch-lock/safety-strap hardware as generic "hardware_other" -- flag it as a plate-sequenced/latch-lock base with secondary parting lines.
- Do not merge or map ejector-stack/pin-plate rows into the A Plate row for quoting.

