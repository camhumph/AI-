' ============================================================================
' gemini1.bas — J8494 FIX
'
' SYMPTOM
'   J8494 (BMS-851100029) quoted only TOP INS and BOT INS. All six main plates
'   were missing from the steel quote even though the BOM parsed perfectly:
'
'     RAWBOM  TOP SMED             4130 Holder Pre-Hardened  18.375 x 15.875 x 1.375  -> TCP
'     RAWBOM  BOTTOM SMED          4130 Holder Pre-Hardened  18.375 x 15.875 x 1.375  -> BCP
'     RAWBOM  TOP HOLDER BLOCK     4130 Holder Pre-Hardened  14.25 x 6.75 x 6.13      -> ID HOLDER
'     RAWBOM  BOTTOM HOLDER BLOCK  4130 Holder Pre-Hardened  14.25 x 7.46 x 6.75      -> OD HOLDER
'     RAWBOM  TOP POT BLOCK        4130 Holder Pre-Hardened  6.13 x 6 x 5             -> ID POT BLOCK
'     RAWBOM  BOTTOM POT BLOCK     4130 Holder Pre-Hardened  7.46 x 6.5 x 5           -> OD POT BLOCK
'
'   Names right, dims right, quote names right -- then they vanished.
'
' CAUSE
'   AddExportRow, lines 13323-13330:
'
'       If ONLY_INCLUDE_4140_BOM_ITEMS Then
'           If NormalizeSteelType(b.material) <> "4140" Then
'               If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
'                   LogLine "Non-4140 item skipped: " ...
'                   Exit Sub
'
'   ONLY_INCLUDE_4140_BOM_ITEMS = True (line 64). The material is 4130, and
'   NormalizeSteelType (line 20424) has no 4130 branch -- it falls through to
'   `Trim(matText)` and returns "4130 Holder Pre-Hardened", which compares
'   unequal to "4140". Every plate was dropped.
'
'   TOP INS / BOT INS survived only because IsInsertQuoteName exempts them.
'   Pyropel survived on the isPyropel exemption. "Mold Base (j-blk, Cam, Ht200)"
'   survived because its material really is 4140. That is the whole quote.
'
'   NOTE: this is the same constant that sits at False in Module6121.bas. There
'   it is a loaded gun that has never fired. Here it is True, and this is it
'   firing. 4130 pre-hardened is standard holder-block steel -- it is not an
'   exotic material, so any shop running this macro on a 4130 job loses its
'   entire steel quote silently.
'
' ============================================================================
' HOW TO APPLY  (3 edits, all in gemini1.bas)
' ============================================================================
'
' EDIT 1 -- line 20431, in NormalizeSteelType.
'   Add the 4130 branch immediately AFTER the existing 4140 line.
'   Mapping 4130 to "4140" (rather than to its own "4130") is deliberate: it
'   routes pre-hardened holder steel into the #2 4140 block the quote template
'   already expects, so nothing downstream needs to learn a new grade.
'
'   FIND:
'       If InStr(s, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
'
'   REPLACE WITH:
'       If InStr(s, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
'       ' 4130 pre-hardened holder steel quotes out of the same #2 4140 block.
'       If InStr(s, "4130") > 0 Then NormalizeSteelType = "4140": Exit Function
'
'
' EDIT 2 -- lines 13323-13330, in AddExportRow.
'   Replace the whole material gate.
'
'   FIND:
'       If ONLY_INCLUDE_4140_BOM_ITEMS Then
'           If NormalizeSteelType(b.material) <> "4140" Then
'               If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
'                   LogLine "Non-4140 item skipped: " & b.Description & " (" & b.material & ")"
'                   Exit Sub
'               End If
'           End If
'       End If
'
'   REPLACE WITH:
'       ' Keep non-steel junk out of the steel quote -- do NOT insist on one
'       ' alloy. Demanding exactly "4140" deleted every 4130 plate on J8494.
'       If ONLY_INCLUDE_4140_BOM_ITEMS Then
'           If IsQuotableSteelMaterial(b.material) = False Then
'               If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
'                   LogLine "Non-steel item skipped: " & b.Description & " (" & b.material & ")"
'                   Exit Sub
'               End If
'           End If
'       End If
'
'
' EDIT 3 -- paste the new function below anywhere at module level.
'   Next to NormalizeSteelType (around line 20440) keeps it tidy.
'
' ============================================================================


' True when the BOM material is a real mold / holder steel and the row belongs
' in the steel quote.
'
' Replaces the old "material must normalize to exactly 4140" test. Fails OPEN --
' a blank material, or an error, returns True -- because dropping a real plate is
' far more expensive than carrying one junk row that a human can delete. The old
' test failed closed and cost J8494 its entire quote.
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
    ' "4130 Holder Pre-Hardened" already matched on 4130 above; these cover
    ' descriptive spellings that carry no alloy number at all.
    If InStr(s, "STEEL") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "PRE-HARD") > 0 Or InStr(s, "PREHARD") > 0 Then IsQuotableSteelMaterial = True: Exit Function
    If InStr(s, "HOLDER") > 0 Then IsQuotableSteelMaterial = True: Exit Function

    Exit Function

eh:
    ' Never let an error in a material string delete a plate.
    IsQuotableSteelMaterial = True
End Function


' ============================================================================
' WHAT J8494 SHOULD LOOK LIKE AFTER THE FIX
'
'   Section  QuoteName       Material                  Status
'   EXPORT   TCP             4130 Holder Pre-Hardened  (matched or NO CAD MATCH)
'   EXPORT   BCP             4130 Holder Pre-Hardened
'   EXPORT   ID HOLDER       4130 Holder Pre-Hardened
'   EXPORT   OD HOLDER       4130 Holder Pre-Hardened
'   EXPORT   ID POT BLOCK    4130 Holder Pre-Hardened
'   EXPORT   OD POT BLOCK    4130 Holder Pre-Hardened
'   EXPORT   TOP INS / BOT INS / Mold Base / Pyropel   (unchanged)
'
' Ten EXPORT rows instead of four. PULLCORE and PACKAGE sections were already
' correct and are untouched by this patch.
'
' If any of the six still says NO CAD MATCH, that is a separate problem in
' FindBestCadMatchForBom / the dimension tolerances -- not this gate. The gate
' was deleting the rows before matching ever got a chance to run.
' ============================================================================
