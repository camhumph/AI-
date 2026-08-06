"""
Patch gemini1.bas -- J8494 "only TOP INS / BOT INS quoted" bug.

Produces the full fixed file. Byte-exact: reads with latin-1, which round-trips
every byte 0-255 losslessly, so nothing outside the patched regions can change.

USAGE
    python patch_gemini1.py "C:\\path\\to\\gemini1.bas"

    # also fix the month-folder search gate (July-2026 vs "July 2026")
    python patch_gemini1.py "C:\\path\\to\\gemini1.bas" --fix-month-folders

Writes:
    gemini1.bas                      <- patched in place
    gemini1.bas.bak_<timestamp>      <- your original, untouched

Refuses to write anything unless every edit matches exactly once, so a partial
or double-applied patch is impossible.
"""

import shutil
import sys
import time
from pathlib import Path

# latin-1 maps bytes 0..255 to codepoints 0..255 and back with no loss, so a
# VBA file with cp1252 characters (degree signs, diameter symbols) survives.
ENCODING = "latin-1"


# ---------------------------------------------------------------------------
# EDIT 1 - teach NormalizeSteelType about 4130.
#
# Mapped to "4140" on purpose: 4130 pre-hardened holder steel quotes out of the
# same #2 4140 block the template already expects, so no downstream code has to
# learn a new grade name.
# ---------------------------------------------------------------------------
EDIT_1_OLD = '''    If InStr(s, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
'''

EDIT_1_NEW = '''    If InStr(s, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
    ' 4130 pre-hardened holder steel quotes out of the same #2 4140 block.
    ' Without this branch NormalizeSteelType fell through to Trim(matText) and
    ' returned "4130 Holder Pre-Hardened", which the AddExportRow material gate
    ' compared unequal to "4140" -- deleting every plate on J8494.
    If InStr(s, "4130") > 0 Then NormalizeSteelType = "4140": Exit Function
'''


# ---------------------------------------------------------------------------
# EDIT 2 - the material gate that was deleting the plates.
# ---------------------------------------------------------------------------
EDIT_2_OLD = '''    If ONLY_INCLUDE_4140_BOM_ITEMS Then
        If NormalizeSteelType(b.material) <> "4140" Then
            If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
                LogLine "Non-4140 item skipped: " & b.Description & " (" & b.material & ")"
                Exit Sub
            End If
        End If
    End If
'''

EDIT_2_NEW = '''    ' Material gate.
    '
    ' This used to demand NormalizeSteelType(material) = "4140" exactly, which
    ' silently deleted the entire steel quote on J8494: every main plate on that
    ' job is "4130 Holder Pre-Hardened", NormalizeSteelType had no 4130 branch,
    ' so it fell through to Trim(matText), compared unequal to "4140", and
    ' TOP SMED / BOTTOM SMED / TOP+BOTTOM HOLDER BLOCK / TOP+BOTTOM POT BLOCK
    ' were all dropped before matching ever ran. Only TOP INS and BOT INS came
    ' out, because IsInsertQuoteName exempts them.
    '
    ' The intent was to keep non-steel junk rows out of the steel quote, not to
    ' insist on one alloy. So test "is this a real mold steel?" instead.
    If ONLY_INCLUDE_4140_BOM_ITEMS Then
        If IsQuotableSteelMaterial(b.material) = False Then
            If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
                LogLine "Non-steel item skipped: " & b.Description & " (" & b.material & ")"
                Exit Sub
            End If
        End If
    End If
'''


# ---------------------------------------------------------------------------
# EDIT 3 - new function, inserted immediately before IsHardwareName.
# ---------------------------------------------------------------------------
EDIT_3_ANCHOR = '''Private Function IsHardwareName(ByVal d As String) As Boolean
'''

EDIT_3_NEW = '''' True when the BOM material is a real mold / holder steel and the row belongs
' in the steel quote.
'
' Replaces the old "material must normalize to exactly 4140" test. Fails OPEN --
' a blank material, or an error, returns True -- because dropping a real plate is
' far more expensive than carrying one junk row a human can delete. The old test
' failed closed and cost J8494 its entire steel quote.
Private Function IsQuotableSteelMaterial(ByVal matText As String) As Boolean
On Error GoTo eh

    IsQuotableSteelMaterial = False

    Dim s As String
    s = UCase(Trim(matText))

    ' Blank material is normalized to DEFAULT_STEEL_TYPE elsewhere in the macro,
    ' so treat it as quotable instead of silently deleting the row.
    If s = "" Then
        IsQuotableSteelMaterial = True
        Exit Function
    End If

    ' --- Pre-hardened mold base / holder steels -------------------------
    If InStr(s, "4140") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "4130") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "P20") > 0 Or InStr(s, "P-20") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "1030") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "1020") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "1045") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "A36") > 0 Or InStr(s, "A-36") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "HOT ROLLED") > 0 Then IsQuotableSteelMaterial = True: Exit Function

    ' --- Tool steels ---------------------------------------------------
    If InStr(s, "H13") > 0 Or InStr(s, "H-13") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "D2") > 0 Or InStr(s, "D-2") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "A2") > 0 Or InStr(s, "A-2") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "S7") > 0 Or InStr(s, "S-7") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "O1") > 0 Or InStr(s, "O-1") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "420") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "STAINLESS") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "DRILL ROD") > 0 Then IsQuotableSteelMaterial = True: Exit Function

    ' --- Generic catch-alls --------------------------------------------
    ' Covers descriptive spellings that carry no alloy number at all.
    If InStr(s, "STEEL") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "PRE-HARD") > 0 Or InStr(s, "PREHARD") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "HOLDER") > 0 Then IsQuotableSteelMaterial = True: Exit Function

    Exit Function

eh:
    ' Never let an error in a material string delete a plate.
    IsQuotableSteelMaterial = True
End Function

'''


# ---------------------------------------------------------------------------
# OPTIONAL EDIT 4 - month-folder search gate.
#
# ShouldSkipJobSearchTopFolder builds Format(Date, "mmmm yyyy") -> "JULY 2026"
# with a SPACE, then requires that literal inside the folder name. The share
# folder is "000000007. July-2026" with a HYPHEN, so InStr returns 0 and the
# whole month is skipped -- every job in it is unfindable.
#
# Comparing month and year separately is immune to whatever separator someone
# types next month.
# ---------------------------------------------------------------------------
EDIT_4_OLD = '''        monthText = UCase(Format(DateAdd("m", -monthOffset, Date), "mmmm yyyy"))

        If InStr(n, monthText) > 0 Then
            ShouldSkipJobSearchTopFolder = False
            Exit Function
        End If
'''

EDIT_4_NEW = '''        ' Match month and year INDEPENDENTLY.
        '
        ' This used to build Format(..., "mmmm yyyy") -> "JULY 2026" and require
        ' that exact literal, which failed on the real share folder
        ' "000000007. July-2026" because of the hyphen. The whole month folder
        ' was skipped and every job inside it reported "not found".
        Dim monthName As String
        Dim yearText As String
        monthName = UCase(Format(DateAdd("m", -monthOffset, Date), "mmmm"))
        yearText = UCase(Format(DateAdd("m", -monthOffset, Date), "yyyy"))
        monthText = monthName & " " & yearText   ' kept for logging parity

        If InStr(n, monthName) > 0 And InStr(n, yearText) > 0 Then
            ShouldSkipJobSearchTopFolder = False
            Exit Function
        End If
'''


# ---------------------------------------------------------------------------
# EDIT 5 - base top view: *Bottom -> *Top.
#
# CMS_BASE_TOP_VIEW_NAME was "*Bottom" (ViewId 6). That looks at the UNDERSIDE
# of the top plate, which mirrors the whole drawing left-to-right -- which is
# exactly what the J8494 DXF shows: "C15 PROPERTY OF GEA / HWC D/N 6019-EW"
# engraving coming out reversed.
#
# The rule wanted is: find the top plate, take ITS top view, then rotate so the
# pots sit in front of the holders. The pot-in-front half already works
# (DefineStandardFrontFromHolderAndPotCom + EnsurePotBlocksCloserThanHolders +
# EnforcePotBlocksCloserAfterFrontPersist, and the J8494 log confirms
# "Final *Front verification OK: pot blocks are closer to front than holders").
# It was the top-view half that was inverted.
#
# ViewId 5 = *Top, 6 = *Bottom (swStandardViews_e).
#
# If after this the face is correct but spun on the sheet, that is what
# CMS_TOP_ROTATE_Z_STEPS is for: 0 = none, 1 = +90, 2 = 180, -1 = -90.
# ---------------------------------------------------------------------------
EDIT_5_OLD = '''Private Const CMS_BASE_TOP_VIEW_NAME As String = "*Bottom"
Private Const CMS_BASE_TOP_VIEW_ID As Long = 6
'''

EDIT_5_NEW = '''' Standard TOP view of the top plate.
'
' This was "*Bottom" / ViewId 6, which views the top plate from underneath and
' mirrors every drawing and engraving DXF left-to-right. The J8494 prints came
' out with the GEA / HWC D/N text reversed for exactly this reason.
'
' ViewId 5 = *Top, 6 = *Bottom (swStandardViews_e).
Private Const CMS_BASE_TOP_VIEW_NAME As String = "*Top"
Private Const CMS_BASE_TOP_VIEW_ID As Long = 5
'''


def apply(text: str, old: str, new: str, label: str) -> str:
    n = text.count(old)
    if n != 1:
        raise SystemExit(
            f"REFUSING TO WRITE: {label} matched {n} times, expected exactly 1.\n"
            f"The file may already be patched, or may be a different version.\n"
            f"Nothing was changed."
        )
    print(f"  ok  {label}")
    return text.replace(old, new, 1)


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}

    if not args:
        raise SystemExit(__doc__)

    path = Path(args[0])
    if not path.is_file():
        raise SystemExit(f"Not a file: {path}")

    original = path.read_text(encoding=ENCODING)
    print(f"Read {path}  ({len(original):,} chars, {original.count(chr(10)):,} lines)")

    if "IsQuotableSteelMaterial" in original:
        raise SystemExit(
            "REFUSING TO WRITE: this file already contains IsQuotableSteelMaterial.\n"
            "It looks patched. Nothing was changed."
        )

    text = original
    print("Applying edits:")
    text = apply(text, EDIT_1_OLD, EDIT_1_NEW, "1/3 NormalizeSteelType 4130 branch")
    text = apply(text, EDIT_2_OLD, EDIT_2_NEW, "2/3 AddExportRow material gate")
    text = apply(text, EDIT_3_ANCHOR, EDIT_3_NEW + EDIT_3_ANCHOR, "3/4 IsQuotableSteelMaterial")
    text = apply(text, EDIT_5_OLD, EDIT_5_NEW, "4/4 base top view *Bottom -> *Top")

    if "--fix-month-folders" in flags:
        text = apply(text, EDIT_4_OLD, EDIT_4_NEW, "5/5 month-folder search gate")
    else:
        print("  -   month-folder gate NOT patched (pass --fix-month-folders to include)")

    backup = path.with_suffix(path.suffix + f".bak_{time.strftime('%Y%m%d_%H%M%S')}")
    shutil.copy2(path, backup)
    path.write_text(text, encoding=ENCODING)

    print()
    print(f"Backup:  {backup}")
    print(f"Patched: {path}")
    print(f"         {len(text):,} chars, {text.count(chr(10)):,} lines "
          f"({len(text) - len(original):+,} chars)")
    print()
    print("Next: open the macro in the SolidWorks VBA editor, remove the old")
    print("module contents, and paste this file in. Then re-run J8494 and check:")
    print()
    print("  1. BOM match report shows 10 EXPORT rows, not 4 (TCP, BCP,")
    print("     ID/OD HOLDER, ID/OD POT BLOCK are back).")
    print("  2. The engraving DXF text reads FORWARD, not mirrored.")
    print("  3. The log still says 'Final *Front verification OK: pot blocks")
    print("     are closer to front than holders'.")
    print()
    print("If the face is right but spun on the sheet, set CMS_TOP_ROTATE_Z_STEPS")
    print("(line ~97): 0 = none, 1 = +90, 2 = 180, -1 = -90.")


if __name__ == "__main__":
    main()
