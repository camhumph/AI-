CMS Ron Quotes run package - BOM / job-token fix
Date: 2026-06-30

Run:
  1. Put these files together in one folder.
  2. Double-click CMS_Launcher.vbs.

Important:
  - The launcher now runs Module6121.swb first so the newest source fixes are used.
  - Module6121.swp is included only as a fallback if SolidWorks cannot run the .swb source macro.
  - Module6121.bas and Module6121.swb contain the newest source fixes.

Fixes included:
  - Ron-only output folder:
      \\Mycloudex2ultra\mexico\Cameron's stuff\RON'S QUOTES
  - Ron-only proposal folder:
      \\Mycloudex2ultra\mexico\Cameron's stuff\RON'S QUOTES\Quote-Proposals-2026
  - PowerShell SolidWorks macro runner with COM message filter to reduce Server Busy dialogs.
  - No OK prompts in the launcher/macro path.
  - Corrected top/base rotation setting.
  - Gemini1 pot-block orientation flow copied into the new source:
      BOM match first, matched holder/pot/TCP orientation second, base export third.
      This keeps the pot blocks in front of the holders and stops the base package from exporting from the old upside-down view.
  - Gemini1 fallback top-view constants copied exactly:
      CMS_BASE_TOP_VIEW_NAME = *Bottom
      CMS_BASE_TOP_VIEW_ID = 6
      CMS_TOP_ROTATE_Z_STEPS = 0
  - Gemini1 named TCP/BCP component orientation path added before the geometry fallback.
  - Gemini1 show/unsuppress cleanup now runs after orientation before base export.
  - Base DXF now follows Gemini1's native-view rule:
      the DXF is created from the saved .sldasm/.sldprt that preserves CMS_TOP,
      not from the X_T conversion path that can lose/rebuild the view.
      No manual DXF-only 90-degree rotation is used.
  - Launcher/proposal naming no longer assumes BMS:
      BMS is used only when the email/attachment actually says BMS.
      Otherwise the job folder/proposal subject uses the email company token,
      or the attachment/file token such as 215MOLDBASE.
  - Gemini1 naming/matching rules copied for pot-block jobs:
      top/id aliases map to TCP, ID HOLDER, ID POT BLOCK;
      bottom/bot/od aliases map to BCP, OD HOLDER, OD POT BLOCK.
      The deep holder/pot aliases now include IDTE/IDLE/ODTE/ODLE,
      top/bottom carrier, top/bottom mold base, TCP/BCP pot,
      and top/bottom/ID/OD SMED and clamp variants.
  - TCP/BCP same-size BOM pairs now use Gemini's mass preference:
      TCP = lighter matching CAD part, BCP = heavier matching CAD part.
  - Name matches no longer override clearly wrong dimensions.
  - Pot-block orientation is now only used for jobs detected as POT / HOLDER BLOCK.
  - Standard mold bases use plain SolidWorks *Top as CMS_TOP and skip the holder/pot front-orientation routine.
  - Standard mold steel naming is geometry-first:
      full-footprint plates are found by shared mold footprint,
      the stack axis is found from part center positions,
      top-to-bottom order is inferred from the layout,
      ejector plates, top/bottom names, and clamp thickness are used only to choose stack direction,
      CAD/BOM names are hints only after the geometry chooses the plate.
  - Unnamed standard full plates are assigned standard roles by stack count/order:
      Top Clamp Plate, Cavity Plate, Core Plate, Stripper Plate, Die Plate,
      Die Backup Plate, Support Plate, Bottom Clamp Plate as applicable.
  - Standard mold CAD aliases added:
      CAVITY PLT -> Cavity Plate
      CORE PLT -> Core Plate
      DIE PLATE -> Die Plate
      DIE BACK UP / DIE BACKUP -> Die Backup Plate
      STRIPPER PLT -> Stripper Plate
  - Standard mold rail detection now rejects near-square long CAD parts
    so rods/pins/support pillars are not mislabeled as Rails.
  - Automatic visual mold inspection added:
      after the ISO JPEG and CAD dimension CSV are written, the macro runs
      cms_visual_inspect.ps1 and creates Visual_Mold_Inspection.txt/csv.
      It estimates visible plate/layer count from the JPEG, combines that with
      CAD dimensions, guesses mold type and PCS six-series family
      (A, B, T, AX, 5 Plate Stripper, 6 Plate Stripper), guesses parting-line
      area, lists likely outside/stack part names, and warns when JPEG/CAD disagree.
  - Standard molds with no BOM now also scan CAD component names for purchased PCS-style hardware
    and include the discovered parts even when the price CSV has no match yet.
  - Gmail/customer job parser accepts alphanumeric jobs like ITWMedical-125.
  - C18605-style BOM columns are supported:
      Mat'l Spec or Mfg. Part No., Lth., Wth./O.D., Hgt./I.D.
  - BOM parser keeps reading after a bad row instead of stopping the entire BOM.
  - DME purchased families added/captured:
      B6, P6, MUD, 6143
    Prices are 0 until the CSV is filled/updated by lookup.
  - BOM material rows are included only when their size matches a CAD part.
  - 4140, cold/hot rolled 1020, A-2, and O-1 BOM material rows are recognized.
  - Multiple rails/retainers/material pieces no longer overwrite the same quote row.
  - Pot/holder jobs still use the pot path; normal bases with names like Optic Holder Mount are not misclassified as pot jobs.