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

Do not trust component names first. Use them only as weak notes.

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

Parting line:

- Interface between cavity/A plate and core/B plate
- Determine from full-footprint stack order plus A/B side evidence
- Leader pins/bushings help orient sides, but the full plate stack and ejector side still matter

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

- Do not assume CAD component names are correct.
- Do not classify row-by-row without seeing the whole mold.
- Do not quote cavity/core inserts as standard base plates.
- Do not classify rails as ejector plates.
- Do not classify support pillars as leader pins.
- Do not classify short bushings as long leader pins.

