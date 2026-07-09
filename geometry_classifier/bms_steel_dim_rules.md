# BMS / Pot-Block Steel Sheet Dimension Rules

Learned from finished J000 Steel Order / Machining Sheet jobs.

## Where Width, Length, and Height go

| Steel sheet column | Meaning | CAD source |
|---|---|---|
| **C** (col 3) | **Thickness / Height** | Smallest bbox dim (`parts.Thickness`) |
| **E** (col 5) | **Width** | Middle bbox dim (`parts.Width`) |
| **G** (col 7) | **Length** | Largest bbox dim (`parts.Length`) |
| H (col 8) | Steel type | `#2 4140` |

Macro `SortThreeDimensions` always sorts CAD bbox as **L ≥ W ≥ T**.
Never put thickness into Length, and never swap Width/Length.

QuoteWorksheet stock sizes: Thickness gets +0.25" stock allowance;
Width/Length round up to nickel. Steel Order keeps **finished** sizes.

Plate order on steel sheet (rows starting at 19):
1. TCP
2. ID Holder
3. OD Holder
4. ID Pot
5. OD Pot
6. BCP
