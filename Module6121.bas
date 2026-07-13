Attribute VB_Name = "Module6121"
Option Explicit

' ============================================================
' CMS POT-BLOCK ENGINE  (jobs WITH a BOM)
' ------------------------------------------------------------
' Type one or more C numbers (e.g. C18454). For each one this:
'   1. Finds the BMS-...-C##### job folder by the C number.
'   2. Copies it local, extracts any ZIP, opens the CAD.
'   3. Orients to CMS_TOP (gemini1 holder/pot/ins/TCP -> *Top/*Front) and saves
'      the WHOLE BASE into the job folder:
'        base\<job> .sldasm / .easm / .igs / .x_t
'        <job> .stl               (gemini1 merged one-file STL, oriented from corrected *Front)
'        <job> .dxf               (4 views, Pyropel hidden)
'        <job> ISO.jpg / BACK ISO.jpg  (Pyropel hidden)
'   4. Scans CAD parts and writes XT_Export_CAD_Dimensions.csv + BOM match report.
'   5. Fills Quote (#2 4140) and J000 steel sheet from CAD/BOM sizes.
'
' NO prints folder, NO individual part X_T/DXF, NO J Block,
' NO Pyropel on JPEG/DXF, NO dimensioned DXF.
' ============================================================

' ============================================================
' USER SETTINGS
' ============================================================
Private Const JOB_ROOT_BASE As String = "\\Mycloudex2ultra\mexico\Cameron's stuff\RON'S QUOTES"  ' Ron-only month folders live here
Private Const EXTRACT_FOLDER_NAME As String = "_EXTRACTED_ZIP"
Private Const LOCAL_WORKSPACE_ROOT As String = "C:\CMS_Local_Workspace"

' --- Network-aware publishing (private local vs public company share) ---
' On the company Netgear Wi-Fi -> publish to the PUBLIC share so the office and
' Elgin can see it. Anywhere else -> keep everything in a PRIVATE local folder.
Private Const COMPANY_WIFI_SSID As String = "NETGEAR"        ' company Wi-Fi name (partial match, case-insensitive)
Private Const PUBLIC_DATA_ROOT As String = "\\Mycloudex2ultra\mexico\Cameron's stuff\Matching software"
Private Const PRIVATE_DATA_ROOT As String = "C:\CMS_Local_Workspace\Matching"
Private Const FORCE_LOCAL_PUBLISH As Boolean = False         ' True = always use the PRIVATE local folder
Private Const PUBLISH_OUTPUTS As Boolean = True              ' copy signature/sheets/images to the matching folder for Elgin
Private Const DELETE_EXTRACTED_ZIP_AFTER_FLATTEN As Boolean = True
Private Const SYNC_COMPLETED_JOB_TO_NETWORK As Boolean = False  ' completed local package copied back to the source job folder (off while testing)
Private Const WRITE_PCS_NAMING_ANALYSIS As Boolean = False
' Always write the Qwen-parity stack / leader-pin position analysis (lightweight CSV).
Private Const WRITE_STACK_LEADERPIN_ANALYSIS As Boolean = True
' Lateral center-plane tolerance (inches) for matching a leader pin to its bushing.
Private Const LEADER_PIN_BUSHING_PLANE_TOL As Double = 0.55

Private Const RUN_SOLIDWORKS_INVISIBLE As Boolean = True
Private Const DISABLE_MAIN_VIEWPORT_GRAPHICS As Boolean = True

' Deliverable exports: always write STL + DXF + ISO JPGs.
' FAST_QUOTE_MODE skips the slowest steps on large assemblies:
'   - ResolveAllLightWeight / Unsuppress-all before export (except STL, which always shows all)
'   - EASM + IGS (heavy neutrals)
'   - visual inspection / PCS naming analysis
' STL always exports a combined one-file mesh (merge when PartCount allows).
Private Const FAST_QUOTE_MODE As Boolean = True
Private Const CREATE_ISO_JPEGS As Boolean = True
Private Const RUN_VISUAL_MOLD_INSPECTION As Boolean = False
Private Const CREATE_DIM_DXF As Boolean = False   ' DIM DXF removed per request
Private Const EXPORT_PER_PLATE_STLS As Boolean = False
' Heavy neutrals are slow on 200+ part STEP imports; off in fast mode.
Private Const EXPORT_HEAVY_NEUTRALS As Boolean = False
Private Const EXPORT_BASE_DXF As Boolean = True
' Above this count, do NOT assembly->temp-part->Combine.
' Export assembly STL directly as one file instead. Much faster for BMS/PCS jobs.
Private Const STL_MERGE_MAX_PARTS As Long = 50
Private Const MAX_SANE_MOLD_DIM_IN As Double = 120#    ' mold parts rarely exceed 10 ft on one axis

' --- Deliverable safety ---
Private Const CREATE_FULL_ASSEMBLY_STL As Boolean = True

' For BMS / holder-pot jobs, STL should contain ONLY the quoted base components:
' TCP, BCP, ID Holder, OD Holder, ID Pot, OD Pot.
' It excludes purchased hardware, pins, bushings, straps, insulation, etc.
Private Const BMS_STL_EXPORT_QUOTED_BASE_ONLY As Boolean = True

' If the six quoted BMS components cannot be identified, do NOT silently export
' the full assembly STL. This prevents accidentally sending a 100+ part STL.
Private Const BMS_STL_SKIP_IF_KEEP_LIST_INCOMPLETE As Boolean = True

' Export STL as one merged STL file containing only quoted/steel components.
Private Const STL_EXPORT_STEEL_COMPONENTS_ONLY As Boolean = True

' Clean accidental component STL shards from the job folder after the merged STL is made.
Private Const CLEAN_EXTRA_STL_SHARDS_IN_JOB_FOLDER As Boolean = True

' If True, BMS ISO/DXF hides everything except TCP/BCP/holders/pots.
' If the keep-list is incomplete, macro falls back to full visible assembly
' instead of making a bad 2-part picture/DXF.
Private Const BMS_ISO_DXF_HIDE_NON_BASE_WHEN_COMPLETE As Boolean = True

' Require at least these BMS components before isolating for ISO/DXF.
' 6 = TCP, BCP, ID Holder, OD Holder, ID Pot, OD Pot.
Private Const BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF As Long = 6

' Force SolidWorks STL export to binary when possible so post-rotation can read it.
Private Const FORCE_BINARY_STL_EXPORT As Boolean = True

' SolidWorks STL ASCII/Binary preference. Common SW value for STL output as binary.
' If your SW version uses a different enum, this harmlessly no-ops under On Error Resume Next.
Private Const swSTLBinaryFormat As Long = 69

Private Const CMS_TOP_VIEW_NAME As String = "CMS_TOP"
Private Const CMS_FRONT_VIEW_NAME As String = "CMS_FRONT"
Private Const CMS_BASE_TOP_VIEW_NAME As String = "*Bottom"
Private Const CMS_BASE_TOP_VIEW_ID As Long = 6
Private Const CMS_TOP_ROTATE_Z_STEPS As Long = 0
Private Const AUTO_SELECT_TCP_TOP_ORIENTATION As Boolean = True
Private Const TCP_TOP_ORIENTATION_KEYS As String = "TCP|TOP CLAMPING|TOP CLAMPING PLATE|TOP SMED|TOP SMED PLATE|ID SMED"
Private Const BCP_BOTTOM_ORIENTATION_KEYS As String = "BCP|BOTTOM CLAMPING|BOTTOM CLAMPING PLATE|BOT CLAMPING|BOT CLAMPING PLATE|BOTTOM SMED|BOT SMED|OD SMED"
Private Const PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE As Boolean = True
Private Const PROMPT_FOR_TOP_ORIENTATION As Boolean = False
Private Const SUPPRESS_USER_PROMPTS As Boolean = True
' Front-orientation tuning (ported from gemini1).
Private Const POT_BLOCKS_MUST_BE_FRONT_OF_HOLDERS As Boolean = True
Private Const POT_FRONT_REQUIRE_EVERY_POT_AHEAD_OF_EVERY_HOLDER As Boolean = False
Private Const POT_FRONT_DEPTH_MIN_DELTA_IN As Double = 0.03
Private Const HOLDER_LONG_SIDE_VISIBLE_RATIO As Double = 0.8
Private Const AUTO_DEFINE_FRONT_FROM_HOLDER_POT_COM As Boolean = True
Private Const DISABLE_STABILIZE_DELAYS As Boolean = True

Private Const SW_DRAWING_TEMPLATE_PATH As String = "C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Drawing.drwdot"
Private Const E_SHEET_WIDTH_IN As Double = 44#
Private Const E_SHEET_HEIGHT_IN As Double = 34#
Private Const DXF_MARGIN_IN As Double = 1#
Private Const DXF_MAX_SCALE As Double = 1#
Private Const DXF_PROJECTED_VIEW_GAP_IN As Double = 2.25
Private Const MULTIVIEW_FIT_SAFETY As Double = 0.9
Private Const FREEZE_DXF_DRAWING_GRAPHICS As Boolean = True

Private Const DIM_DECIMALS As Long = 3
Private Const INCHES_PER_METER As Double = 39.3700787401575
Private Const PI_VALUE As Double = 3.14159265358979

' ============================================================
' SOLIDWORKS CONSTANTS
' ============================================================
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3
Private Const swOpenDocOptions_Silent As Long = 1
Private Const swOpenDocOptions_ReadOnly As Long = 2
Private Const swSaveAsCurrentVersion As Long = 0
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swSaveAsOptions_Copy As Long = 2
' swUserPreferenceToggle_e.swSTLComponentsIntoOneFile (gemini1 = 248 for SW2023).
Private Const swSTLComponentsIntoOneFile As Long = 248
' swUserPreferenceIntegerValue_e.swSaveAssemblyAsPartOptions
Private Const swSaveAssemblyAsPartOptions As Long = 201
Private Const swSaveAsmAsPart_AllComponents As Long = 1
' swBodyOperationType_e.SWBODYADD = Combine -> Add (union of bodies).
Private Const SWBODYADD As Long = 15903
' After merge-STL export, rotate mesh into corrected *Front/*Top frame (gemini1).
Private Const MATCH_STUDIO_STL_MATCH_MAIN_BASE_ORIENTATION As Boolean = True
Private Const POST_ROTATE_STL_TO_CORRECTED_FRONT As Boolean = True
Private Const swSolidBody As Long = 0
Private Const swComponentHidden As Long = 0
Private Const swComponentVisible As Long = 1

' Binary STL layout for post-rotate (ported from gemini1).
Private Type BinaryStlHeader
    HeaderText As String * 80
    TriangleCount As Long
End Type
Private Type BinaryStlTriangle
    nx As Single
    ny As Single
    nz As Single
    x1 As Single
    y1 As Single
    z1 As Single
    x2 As Single
    y2 As Single
    z2 As Single
    x3 As Single
    y3 As Single
    z3 As Single
    AttributeByteCount As Integer
End Type

' ============================================================
' GLOBALS
' ============================================================
Private swApp As Object
Private swModel As Object

Private RunLogPath As String
Private StartupLogPath As String
Private CurrentJobFolder As String
Private CurrentJobNumber As String
Private NetworkJobFolder As String
Private LocalJobFolder As String
Private JobBaseName As String

Private MacroStartTime As Date
Private StepStartTime As Date
Private JobStartTime As Date
Private CurrentStepName As String

Private MainCadOpenedByMacro As Boolean
Private MainCadTitleForClose As String
Private MainViewportGraphicsDisabled As Boolean
Private LastJobFailReason As String

Private DxfFreezeDoc As Object
Private CurrentDxfForce1to1 As Boolean

' Final corrected *Front orientation matrix for STL post-rotate (gemini1).
Private FinalStlCoordFrameReady As Boolean
Private FinalStlCoordM(0 To 8) As Double

' Set by DetectBaseTypeIsStandard / ProcessOneJob / RunActiveAssembly.
' Standard molds have no Pyropel isolation step — full-assembly DXF/ISO only.
Private gJobIsStandardBase As Boolean

' ============================================================
' POT-BLOCK ENGINE ADDITIONS  (scan + BOM read/match + Excel fill)
' ============================================================

' --- BOM reading ---
Private Const READ_PDF_BOM_WITH_PDFTOTEXT As Boolean = True
Private Const PDFTOTEXT_EXE As String = "C:\Users\lenovo\Downloads\New folder (9)\poppler-26.02.0\Library\bin\pdftotext.exe"
Private Const TURBO_READ_ONLY_BOM_SHEET As Boolean = True
Private Const TURBO_BOM_SHEET_NAME As String = "BOM"
Private Const BOM_HEADER_SEARCH_MAX_ROWS As Long = 150
Private Const STOP_BOM_READ_AFTER_BLANK_ROWS As Long = 20
Private Const ONLY_INCLUDE_4140_BOM_ITEMS As Boolean = False
Private Const DEFAULT_STEEL_TYPE As String = "4140"

' --- matching / dims ---
Private Const DIM_OK_TOL As Double = 0.03
Private Const DIM_REVIEW_TOL As Double = 0.125
Private Const SAME_SIZE_PAIR_TOL As Double = 0.125
Private Const DIM_MAX_MATCH_TOTAL_DIFF As Double = 5#
Private Const MIN_STEEL_VOLUME_CUIN As Double = 1#
Private Const CUIN_PER_CUBIC_METER As Double = 61023.7440947323
Private Const HIDE_QUARTER_INCH_THICKNESS As Boolean = False
Private Const QUARTER_INCH_THICKNESS As Double = 0.25
' Steel stock allowance: add this to the finished thickness on the QUOTE sheet,
' then round up to the nearest 0.0001". The STEEL ORDER sheet keeps finished dims.
Private Const STEEL_THICKNESS_ALLOWANCE As Double = 0.25
Private Const QUARTER_INCH_TOLERANCE As Double = 0.01

' --- Excel fill toggles ---
Private Const FILL_QUOTE_WORKBOOK As Boolean = True
Private Const FILL_J000_STEEL_SHEET As Boolean = True
Private Const DOWNLOADS_FOLDER As String = "C:\Users\lenovo\Downloads"
' Trusted location where the PDFs and Excel/xlsm templates now live.
Private Const TRUSTED_FOLDER As String = "C:\Users\lenovo\Documents\Trust"
' Gmail SMTP for the proposal email-back.
' SECURITY: credentials live in the webapp Settings page JSON file (never
' gmail_app_password.txt). Revoke any password ever committed to GitHub.
Private Const GMAIL_ADDRESS As String = "cms1engineering@gmail.com"
Private Const EMAIL_CREDENTIALS_FILE As String = "C:\CMS_Local_Workspace\cms_data\email_credentials.json"
' Proposal email behavior: "AUTO" sends with no prompt (old behavior),
' "PROMPT" asks before sending, "OFF" only writes the preview file.
' Price purchased parts in the macro / CSV only — do not auto-email proposals.
Private Const PROPOSAL_EMAIL_MODE As String = "OFF"
Private Const QUOTE_SHEET_NAME As String = "QuoteWorksheet"
Private Const POTBLOCK_STEEL_TYPE As String = "#2 4140"
' Quote worksheet shows STOCK sizes = finished size rounded UP to the next 1/4".
' The J000 steel order/machining sheet shows the FINISHED sizes as-is.
Private Const QUOTE_ROUND_UP_TO_QUARTER As Boolean = True
Private Const xlCalculationManual As Long = -4135

' Pot-block plate name keys (CAD bounding-box matching for the Excel fill)
Private Const KEYS_TCP As String = "TCP|TOP CLAMPING|TOP CLAMPING PLATE|TOP CLAMP PLATE|TOP CLAMP|TOP SMED|TOP SMED PLATE|ID SMED|ID SMED PLATE|ID CLAMPING|ID CLAMPING PLATE|ID CLAMP PLATE"
Private Const KEYS_BCP As String = "BCP|BOTTOM CLAMPING|BOTTOM CLAMPING PLATE|BOT CLAMPING|BOT CLAMPING PLATE|BOTTOM CLAMP PLATE|BOTTOM CLAMP|BOT CLAMP|BOTTOM SMED|BOTTOM SMED PLATE|BOT SMED|BOT SMED PLATE|OD SMED|OD SMED PLATE|OD CLAMPING|OD CLAMPING PLATE|OD CLAMP PLATE"
Private Const ID_HOLDER_KEYS As String = "ID HOLDER|TOP HOLDER|TOP HOLDER BLOCK|ID HOLDER BLOCK|IDTE HOLDER|IDLE HOLDER|TOP CARRIER|TOP CARRIER BLOCK|ID CARRIER|ID CARRIER BLOCK|ID MOLD BASE|ID MOLDBASE|IDTE MOLD BASE|IDLE MOLD BASE|TOP MOLD BASE|TOP MOLDBASE|TOP BASE|ID BASE"
Private Const OD_HOLDER_KEYS As String = "OD HOLDER|BOTTOM HOLDER|BOT HOLDER|BOTTOM HOLDER BLOCK|BOT HOLDER BLOCK|OD HOLDER BLOCK|ODTE HOLDER|ODLE HOLDER|BOTTOM CARRIER|BOT CARRIER|BOTTOM CARRIER BLOCK|BOT CARRIER BLOCK|OD CARRIER|OD CARRIER BLOCK|OD MOLD BASE|OD MOLDBASE|ODTE MOLD BASE|ODLE MOLD BASE|BOTTOM MOLD BASE|BOT MOLD BASE|BOTTOM MOLDBASE|BOT MOLDBASE|BOTTOM BASE|BOT BASE|OD BASE"
Private Const KEYS_ID_POT As String = "ID POT BLOCK|ID POT|IDTE POT|IDLE POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT|ID INSERT POT|TOP INSERT POT"
Private Const KEYS_OD_POT As String = "OD POT BLOCK|OD POT|ODTE POT|ODLE POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT|OD INSERT POT|BOTTOM INSERT POT|BOT INSERT POT"

' --- Types ---
Private Type PartInfo
    componentName As String
    cleanName As String
    filePath As String
    configName As String
    bodyName As String
    Quantity As Long
    Length As Double
    Width As Double
    Thickness As Double
    ' Raw assembly-axis box extents before CMS-view L/W/T assignment.
    BoxDx As Double
    BoxDy As Double
    BoxDz As Double
    BBoxVolume As Double
    massValue As Double
    hasMassCenter As Boolean
    MassCenterX As Double
    MassCenterY As Double
    MassCenterZ As Double
    hasAsmCenter As Boolean
    AsmCenterX As Double
    AsmCenterY As Double
    AsmCenterZ As Double
    UsedForBomMatch As Boolean
    isBodyOnly As Boolean
End Type

' CMS DXF view-frame axes (after CMS_TOP / *Front / *Right are locked).
' From the shop dimensioned DXF (labels on TOP / RIGHT / BOTTOM views):
'   TOP view:    horizontal (X) = Width,  vertical (Y) = Length, into-screen = Thickness
'   RIGHT view:  horizontal (X) = Thickness, vertical (Y) = Length
'   FRONT/BOTTOM:horizontal (X) = Width
' Model XYZ alone is not L/W/T — the frame rotates with the views.
Private gCmsViewFrameReady As Boolean
Private gCmsLenAxisX As Double, gCmsLenAxisY As Double, gCmsLenAxisZ As Double
Private gCmsWidAxisX As Double, gCmsWidAxisY As Double, gCmsWidAxisZ As Double
Private gCmsThkAxisX As Double, gCmsThkAxisY As Double, gCmsThkAxisZ As Double

Private Type BomInfo
    Description As String
    quoteName As String
    Quantity As Long
    material As String
    BomLength As Double
    BomWidth As Double
    BomThickness As Double
    hasDims As Boolean
    ' True when BomThickness/Width/Length hold Tempcraft Lth/Wth/Hgt file order
    ' (not yet mapped to CMS T/W/L). False when already CMS-oriented or sorted.
    BomIsTempcraftOrder As Boolean
End Type

Private Type ExportInfo
    quoteName As String
    Quantity As Long
    material As String
    CadPartIndex As Long
    HasCad As Boolean
    Thickness As Double
    Width As Double
    Length As Double
    BomThickness As Double
    BomWidth As Double
    BomLength As Double
    HasBomDims As Boolean
    Status As String
End Type

' --- Globals ---
Private swAssy As Object
Private parts() As PartInfo
Private PartCount As Long
Private BomRows() As BomInfo
Private BomCount As Long
Private ExportRows() As ExportInfo
Private ExportCount As Long

' Pot-block plates identified directly from CAD geometry (for .x_t imports
' that have no plate names, or BOMs that carry no sizes). 0 = not found.
Private gIdxTCP As Long
Private gIdxBCP As Long
Private gIdxIDH As Long
Private gIdxODH As Long
Private gIdxIDP As Long
Private gIdxODP As Long
Private Const PLATE_MIN_THICKNESS As Double = 0.5    ' below this = insulation/shim
Private Const PLATE_MIN_FOOTPRINT As Double = 20#    ' W*L below this = hardware
Private Const POT_MAX_ASPECT As Double = 1.7         ' L/W <= this => pot candidate (blocky)
Private Const POT_MIN_THICKNESS As Double = 3#       ' pots are thick blocks, not mold plates
Private Const POT_MAX_FOOTPRINT_FRAC As Double = 0.55 ' pots are clearly smaller than the mold footprint
Private Const POT_MIN_CUBE_RATIO As Double = 0.35    ' min(T,W,L)/max(T,W,L) — pots are chunky, not flat
Private Const CLAMP_THIN_RATIO As Double = 0.25      ' T <= ratio*L => clamp/smed plate
Private Const ASSIGN_ID_AS_TOP As Boolean = True     ' higher Z (or larger) = ID/top; flip if reversed

' --- Standard (non-pot) mold base ---
Private Const BASE_TYPE_MODE As String = "AUTO"      ' AUTO | POT | STANDARD
Private Const STD_FOOTPRINT_TOL As Double = 0.18     ' within this fraction of base footprint = full plate
Private Const STD_MIN_PLATE_THICKNESS As Double = 0.4
Private Const STD_RAIL_MIN_LENGTH_FRAC As Double = 0.6
Private Const STD_RAIL_MAX_WIDTH_FRAC As Double = 0.65
Private Const STD_RAIL_MIN_THICK As Double = 1#
Private Const STD_EJECTOR_MIN_FOOT_FRAC As Double = 0.15
Private Const STD_A_B_GRADE As String = "P20"        ' A & B plates default to P20 (#3 block)
Private Const STD_TRUST_CAD_NAMES_FOR_STANDARD_STACK As Boolean = False

' For PCS / standard mold bases, quote only the primary steel stack:
' A Plate, B Plate, 2 Rails, Ejector Plate, and Ejector Retainer/Backup Plate.
Private Const STD_QUOTE_PRIMARY_PCS_STACK_ONLY As Boolean = True
Private Const STD_QUOTE_INCLUDE_CLAMP_PLATES As Boolean = False
Private Const STD_QUOTE_RAIL_QTY As Long = 2
Private Const STD_QUOTE_KEEP_ONE_A_PLATE As Boolean = True
Private Const STD_QUOTE_KEEP_ONE_B_PLATE As Boolean = True
Private Const STD_QUOTE_KEEP_ONE_EJECTOR_PLATE As Boolean = True
Private Const STD_QUOTE_KEEP_ONE_EJECTOR_BACKUP As Boolean = True

' BMS TCP/BCP mass/volume sanity check.
Private Const BMS_TCP_EXPECT_LIGHTER_THAN_BCP As Boolean = True
Private Const BMS_TCP_BCP_FORCE_LIGHTER_TCP As Boolean = False
Private Const BMS_TCP_BCP_MASS_DIFF_FRAC As Double = 0.01

Private Const PULLCORE_RATE As Double = 88#          ' pullcore/key quote = total cubic inches x this
Private Const PULLCORE_QUOTE_START_ROW As Long = 218 ' Quote sheet row where the pull-core category begins
Private Const PULLCORE_PRICE_FILE As String = "Pullcore Prices.csv"

' --- Purchased components (DME / McMaster / Jaco hardware) ---
Private Const FILL_PURCHASED_COMPONENTS As Boolean = True
' Live web price lookup for purchased components (DME store / Bing). Needs internet.
' Pricing is handled by the Python tool (cms_price_lookup.py), which renders the
' DME page and writes prices into the CSV. The macro just reads the CSV, so its
' own web lookup is OFF. (Flip to True only if you want the VBA fallback back.)
Private Const ENABLE_ONLINE_PRICE_LOOKUP As Boolean = False
Private Const ENABLE_PYTHON_PRICE_LOOKUP As Boolean = False
Private Const ENABLE_ASSISTED_PRICE_PROMPT As Boolean = False
Private Const PYTHON_EXE As String = "python"
Private Const PURCHASED_PRICE_FILE As String = "Purchased Components Prices.csv"
Private Const PURCHASED_QUOTE_START_ROW As Long = 240 ' Quote row where the purchased-hardware category begins

Private PcName() As String
Private PcQty() As Long
Private PcT() As Double
Private PcW() As Double
Private PcL() As Double
Private PcMat() As String
Private PcVol() As Double
Private PcCount As Long

' Purchased hardware captured from the BOM and priced from the CSV
Private PpDesc() As String
Private PpQty() As Long
Private PpComp() As String
Private PpVendor() As String
Private PpPartNo() As String
Private PpPrice() As Double
Private PpW() As Double          ' BOM width / O.D. (used as Dia.)
Private PpL() As Double          ' BOM length
Private PpT() As Double          ' BOM thickness / height (used as Width)
Private PpDet() As String        ' BOM Det No. (e.g. 107) for the email
Private PpCount As Long
' The loaded purchased price list (from PURCHASED_PRICE_FILE)
Private PlComp() As String
Private PlVendor() As String
Private PlPartNo() As String
Private PlDescr() As String
Private PlUnit() As String
Private PlPrice() As Double
Private PlCount As Long
Private gPriceListPath As String   ' resolved CSV path, so captured prices persist

' Customer info passed in from the launcher via the handoff file
Private CustomerJobNumber As String
Private CustomerPrefix As String
Private CustomerDisplayName As String
Private AssignedQuoteNumber As String
Private SimilarToJob As String
Private ShipDateText As String
Private gRootJobPath As String          ' resolved month folder (from handoff or computed)
Private gExactJobFolderName As String    ' exact folder name from launcher handoff (avoids fuzzy match)
Private gHandoffAttachDir As String      ' email attachment folder when network job folder is missing
Private gHandoffCadPath As String        ' preferred CAD path from batch/single handoff
Private gDiagBomPath As String           ' BOM file the macro used (for the end-of-run popup)
Private gEmailStatus As String           ' result of the proposal email step
Private gLastJobDiag As String           ' summary of BOM/components/email for the popup
Private gProcessingHandoff As Boolean     ' True while launcher-supplied quote/job info must be preserved

Private stdName() As String
Private StdT() As Double
Private StdW() As Double
Private StdL() As Double
Private StdQty() As Long
Private StdGrade() As String
Private StdQuoteRow() As Long
Private StdCadIndex() As Long   ' exact CAD part index used for this standard steel/quote row
Private StdCount As Long
Private gStdRoleByPart() As String
Private gStdStackAxis As Integer
Private gStdTopIsFirst As Boolean
Private gStdDmeStackFamily As String
Private gStdPartingLineAxis As Integer
Private gStdPartingLinePos As Double
Private gStdCavityCadIndex As Long
Private gStdCoreCadIndex As Long
' Qwen-parity leader-pin / stack analysis (offline geometry path + AI bridge).
' Per-part set: "PRIMARY" (matched to shoulder/LBB bushings on B plate),
' "SECONDARY" (matched to guided-ejector bushings — must not decide A/B),
' or "" when unmatched.
Private gStdLeaderPinSetByPart() As String
Private gStdLeaderPinFromTop As Boolean      ' True = pins enter from top/A side
Private gStdLeaderPinFromKnown As Boolean    ' True when pin direction was measured
Private gStdLeaderPinReversed As Boolean     ' Seated in B area running toward A
Private gStdSequencedLatchLock As Boolean    ' PLC / latch-lock / safety-strap base
Private gStdStackRules As String             ' Pipe-separated rules_for_this_job text
Private gStdPartingLineText As String

' ============================================================
' AI BRIDGE (CMS AI Quoting local web app / geometry classifier)
'
' After WritePartDimensionCsv the macro POSTs the CSV path to the LOCAL
' AI service (127.0.0.1 only - never the network) which runs the Python
' geometry classifier (shop-token anchors, latch-lock/SC detection,
' bottom-up rail/ejector stack anchoring) and returns one role per CAD
' part index. Those roles feed ClassifyStandardBasePlates.
'
' HARD RULE: the AI is used for STANDARD (non-BMS) bases ONLY. BMS /
' pot-block jobs keep the proven BOM-driven flow untouched - see the
' isStd guard in RunAiBridgeClassification.
' ============================================================
Private Const AI_BRIDGE_ENABLED As Boolean = True
Private Const AI_BRIDGE_URL As String = "http://127.0.0.1:8000"
Private Const AI_BRIDGE_TIMEOUT_MS As Long = 30000
' Fallback folder scanned for <job>_part_names.csv when the local web app
' is not running (matches the web app's CMS_VBA_BRIDGE_DIR).
Private Const AI_BRIDGE_FILE_DIR As String = "C:\CMS_Local_Workspace\AI_Bridge"
' Minimum HIGH-confidence AI plate roles before the AI stack is trusted
' over the macro's own geometry pass.
Private Const AI_BRIDGE_MIN_PLATES As Long = 3

Private gAiRoleByPart() As String       ' AI role key per CAD part index ("a_plate", ...)
Private gAiConfByPart() As String       ' "HIGH" / "MEDIUM" / "LOW"
Private gAiRoleCount As Long            ' rows parsed from the bridge
Private gAiBridgeUsed As Boolean        ' True when AI roles drove the standard stack
Private gAiSequencedLatchLock As Boolean ' True when the AI flagged a latch-lock/SC base

' Handoff file written by CMS_Launcher.vbs and read at startup
Private Type HandoffInfo
    CNum     As String
    QuoteNum As String
    CustJob  As String
    SimilarTo As String
    ShipDate As String
    RootPath As String
    JobFolder As String
    CustomerPrefix As String
    CustomerName As String
    AttachDir As String
    CadPath As String
End Type

Private Const HANDOFF_FILE As String = "C:\CMS_Local_Workspace\cms_handoff.txt"
Private Const TRAINING_XT_HANDOFF As String = "C:\CMS_Local_Workspace\cms_training_xt.txt"

Private Const MACRO_STATUS_FILE As String = "C:\CMS_Local_Workspace\cms_macro_status.txt"
Private Const MACRO_STARTED_FILE As String = "C:\CMS_Local_Workspace\cms_macro_started.txt"
Private Const MACRO_DONE_FILE As String = "C:\CMS_Local_Workspace\cms_macro_done.txt"
Private Const MACRO_ERROR_FILE As String = "C:\CMS_Local_Workspace\cms_macro_error.txt"

Private Type TrainingXtHandoff
    JobFolder As String
    JobId As String
    OutputCsv As String
    DoneFile As String
End Type

' ============================================================
' MAIN
' ============================================================
Sub main()
On Error GoTo ErrHandler

    ' If the launcher created a handoff file, run the unattended launcher path.
    ' This lets SolidWorks start the macro through either main() or RunFromLauncher()
    ' without showing an InputBox.
    Dim fsoLaunch As Object
    Set fsoLaunch = CreateObject("Scripting.FileSystemObject")

    ' Live quote handoff must win over stale training handoff.
    If fsoLaunch.FileExists(HANDOFF_FILE) Then
        On Error Resume Next
        If fsoLaunch.FileExists(TRAINING_XT_HANDOFF) Then fsoLaunch.DeleteFile TRAINING_XT_HANDOFF, True
        On Error GoTo ErrHandler

        RunFromLauncher
        Exit Sub
    End If

    If fsoLaunch.FileExists(TRAINING_XT_HANDOFF) Then
        RunTrainingXtExport
        Exit Sub
    End If

    Set swApp = Application.SldWorks

    If ActiveCadIsOpen() Then
        Dim runActive As VbMsgBoxResult
        runActive = MsgBox("Use the CAD model that is already open in SolidWorks?" & vbCrLf & vbCrLf & _
                           "Yes = run quoting from the open XT/assembly." & vbCrLf & _
                           "No = enter C-number(s) like normal." & vbCrLf & _
                           "Cancel = stop.", _
                           vbYesNoCancel + vbQuestion, "CMS Base Export")
        If runActive = vbYes Then
            RunActiveAssembly
            Exit Sub
        ElseIf runActive = vbCancel Then
            GoTo NormalEnd
        End If
    End If

    If RUN_SOLIDWORKS_INVISIBLE Then
        On Error Resume Next
        swApp.Visible = False
        On Error GoTo ErrHandler
    End If

    MacroStartTime = Now
    StartupLogPath = DOWNLOADS_FOLDER & "\CMS_Base_Export_Log.txt"
    RunLogPath = StartupLogPath
    gRootJobPath = CurrentMonthJobFolder()

    LogLine "========================================"
    LogLine "BASE EXPORT MACRO STARTED"
    LogLine "Root path: " & gRootJobPath
    LogLine "========================================"

    Dim jobInput As String
    jobInput = Trim(InputBox("Enter one or more C numbers." & vbCrLf & _
                             "Examples:" & vbCrLf & _
                             "C18454" & vbCrLf & _
                             "C18454, C18455, C18456", _
                             "CMS Base Export"))

    If jobInput = "" Then GoTo NormalEnd

    Dim jobs As Collection
    Set jobs = ParseJobInputList(jobInput)

    If jobs Is Nothing Or jobs.Count = 0 Then
        LogLine "No valid C numbers entered."
        If Not SUPPRESS_USER_PROMPTS Then MsgBox "No valid C numbers entered.", vbExclamation
        GoTo NormalEnd
    End If

    Dim completed As Collection
    Dim failed As Collection
    Set completed = New Collection
    Set failed = New Collection

    Dim i As Long
    Dim jobText As String
    Dim ok As Boolean

    For i = 1 To jobs.Count
        jobText = UCase(Trim(CStr(jobs(i))))
        If jobText <> "" Then
            RunLogPath = StartupLogPath
            LogLine "BATCH ITEM " & i & "/" & jobs.Count & ": " & jobText
            ok = ProcessOneJob(jobText)
            If ok Then
                completed.Add jobText
            Else
                failed.Add jobText & IIf(LastJobFailReason <> "", "  ->  " & LastJobFailReason, "")
            End If
            DoEvents
        End If
    Next i

    Dim summary As String
    summary = BuildBatchSummary(completed, failed)

    CloseAllDocumentsSafely
    On Error Resume Next
    If Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo ErrHandler

    LogLine summary
    If Not SUPPRESS_USER_PROMPTS Then MsgBox summary, IIf(failed.Count > 0, vbExclamation, vbInformation)

NormalEnd:
    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing And Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo 0
    Exit Sub

ErrHandler:
    LogLine "FATAL BATCH ERROR. Step: " & CurrentStepName & "  Err " & Err.Number & ": " & Err.Description
    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing And Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo 0
    If Not SUPPRESS_USER_PROMPTS Then MsgBox "Macro error at step: " & CurrentStepName & vbCrLf & Err.Description & vbCrLf & RunLogPath, vbCritical
End Sub

' ============================================================
' ACTIVE CAD / XT ENTRY POINT
' Use this when the XT, STEP, part, or assembly is already open in SolidWorks.
' It does not need a launcher handoff or C-number, and it does not move/close
' the user's CAD files.
' ============================================================
Sub RunActiveAssembly()
On Error GoTo ErrHandler
    Set swApp = Application.SldWorks
    WriteMacroLaunchStatus "STARTED", "RunActiveAssembly entered"
    On Error Resume Next
    swApp.Visible = True
    On Error GoTo ErrHandler
    MacroStartTime = Now
    StartupLogPath = DOWNLOADS_FOLDER & "\CMS_Base_Export_Log.txt"
    RunLogPath = StartupLogPath

    If Not ActiveCadIsOpen() Then
        LogLine "RunActiveAssembly stopped: no active CAD document is open."
        MsgBox "Open the XT, STEP, part, or assembly in SolidWorks first, then run RunActiveAssembly.", vbExclamation, "CMS Base Export"
        Exit Sub
    End If

    Set swModel = swApp.ActiveDoc
    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
    MainViewportGraphicsDisabled = False

    Dim modelPath As String
    Dim modelTitle As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    modelPath = ""
    On Error Resume Next
    modelPath = swModel.GetPathName
    modelTitle = swModel.GetTitle
    On Error GoTo ErrHandler
    If modelTitle = "" Then modelTitle = "ActiveCad"

    ' Do not quote from a previously exported base assembly.
    ' Those files live in \base\ and can have moved/broken component references,
    ' changed standard views, and already-processed orientation.
    If modelPath <> "" Then
        If InStr(1, UCase(modelPath), "\BASE\", vbTextCompare) > 0 Then
            LogErrorText "Active CAD is a generated base output assembly. Open the original customer XT/STEP instead."
            LogLine "  Active CAD: " & modelTitle
            LogLine "  Active path: " & modelPath

            If Not SUPPRESS_USER_PROMPTS Then
                MsgBox "This is a generated output assembly in the \base\ folder." & vbCrLf & vbCrLf & _
                       "Do not run the quote macro on:" & vbCrLf & _
                       modelPath & vbCrLf & vbCrLf & _
                       "Open the original customer XT/STEP file instead.", _
                       vbCritical, "Wrong CAD source"
            End If

            WriteMacroLaunchStatus "ERROR", "Refused generated \\base\\ CAD: " & modelPath
            Exit Sub
        End If
    End If

    JobBaseName = CleanFileName(GetFileBaseName(modelTitle))
    If JobBaseName = "" Then JobBaseName = CleanFileName(modelTitle)
    If JobBaseName = "" Then JobBaseName = "ActiveCad"
    ' Prefer C-number from launcher handoff (e.g. C18603) over CAD file name.
    If Not gProcessingHandoff Or CurrentJobNumber = "" Then
        CurrentJobNumber = JobBaseName
    End If
    If Not gProcessingHandoff Then
        CustomerJobNumber = ""
        CustomerPrefix = ""
        CustomerDisplayName = ""
        AssignedQuoteNumber = ""
        SimilarToJob = ""
        ShipDateText = ""
        gExactJobFolderName = ""
    End If
    If Not gProcessingHandoff Then
        NetworkJobFolder = ""
        LocalJobFolder = ""
    End If
    gDiagBomPath = ""
    gEmailStatus = ""

    If modelPath <> "" And NetworkJobFolder = "" Then
        NetworkJobFolder = fso.GetParentFolderName(modelPath)
    End If
    If NetworkJobFolder = "" Then NetworkJobFolder = LOCAL_WORKSPACE_ROOT
    ' Use C-number folder when quoting from launcher (stable path for outputs).
    If gProcessingHandoff And CurrentJobNumber <> "" Then
        CurrentJobFolder = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber)
        EnsureFolderDeep CurrentJobFolder
    Else
        CurrentJobFolder = NetworkJobFolder & "\CMS_ACTIVE_QUOTE_" & JobBaseName & "_" & Format(Now, "yyyymmdd_hhnnss")
        EnsureFolderDeep CurrentJobFolder
    End If
    RunLogPath = CurrentJobFolder & "\CMS_Base_Export_Log.txt"

    LogLine "========================================"
    LogLine "BASE EXPORT MACRO STARTED (active CAD quote)"
    LogLine "Active CAD: " & modelTitle
    If modelPath <> "" Then LogLine "Active path: " & modelPath
    LogLine "Output folder: " & CurrentJobFolder
    LogLine "========================================"

    If gProcessingHandoff Then
        If CustomerJobNumber <> "" Then
            If InStr(UCase(modelTitle & " " & modelPath), UCase(CustomerJobNumber)) = 0 Then

                LogErrorText "Active CAD does not match handoff customer job. Stopping to prevent wrong BOM/CAD quote."
                LogLine "  Active CAD: " & modelTitle
                LogLine "  Active path: " & modelPath
                LogLine "  Handoff customer job: " & CustomerJobNumber

                If Not SUPPRESS_USER_PROMPTS Then
                    MsgBox "The open CAD does not match the launcher/job handoff." & vbCrLf & vbCrLf & _
                           "Open CAD: " & modelTitle & vbCrLf & _
                           "Expected customer job: " & CustomerJobNumber & vbCrLf & vbCrLf & _
                           "Close the wrong CAD or clear the handoff file before running.", _
                           vbCritical, "CMS Base Export - CAD/BOM mismatch"
                End If

                Exit Sub
            End If
        End If
    End If

    JobStartTime = Now
    FinalStlCoordFrameReady = False
    ResetCmsViewFrame
    Dim stlCoordI As Long
    For stlCoordI = 0 To 8
        FinalStlCoordM(stlCoordI) = 0#
    Next stlCoordI
    MainViewportGraphicsDisabled = False
    DoEvents

    LogStart "Scan active CAD/XT parts"
    PartCount = 0
    ReDim parts(1 To 1)
    Set swAssy = Nothing
    ScanActiveSolidWorksDocument
    SortPartsByVolumeDescending
    ClassifyPotBlockPlatesFromCad
    LogLine "CAD PartCount=" & PartCount
    WritePartDimensionCsv CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"
    WriteAllCadComponentsDebugCsv CurrentJobFolder & "\CAD_All_Components_Debug_PRE_ORIENT.csv"
    WriteJobFileInventoryCsv CurrentJobFolder, CurrentJobFolder & "\Job_File_Inventory.csv"
    LogDone "Scan active CAD/XT parts"
    DoEvents

    ' BMS / pot-block jobs MUST read the customer BOM (same as ProcessOneJob).
    ' Without this, open-CAD-first quotes skip BOM and fill wrong sizes.
    BomCount = 0
    ReDim BomRows(1 To 1)
    ExportCount = 0
    ReDim ExportRows(1 To 1)
    LoadPurchasedPriceList

    LogStart "Find + read BOM (active CAD path)"
    Dim bomSearchRoots As Collection
    Set bomSearchRoots = New Collection
    bomSearchRoots.Add CurrentJobFolder
    If NetworkJobFolder <> "" Then bomSearchRoots.Add NetworkJobFolder
    If gHandoffAttachDir <> "" Then bomSearchRoots.Add gHandoffAttachDir
    Dim bomPath As String, bi As Long, rootPath As String
    bomPath = ""
    For bi = 1 To bomSearchRoots.Count
        rootPath = CStr(bomSearchRoots(bi))
        If rootPath <> "" Then
            bomPath = FindCustomerBomFile(rootPath)
            If bomPath <> "" Then Exit For
        End If
    Next bi
    gDiagBomPath = bomPath
    If bomPath <> "" Then
        LogLine "BOM selected (active): " & bomPath
        If LCase(GetFileExtension(bomPath)) = "pdf" Then
            If READ_PDF_BOM_WITH_PDFTOTEXT Then ReadCustomerBomPdfUsingPdfToText bomPath
        Else
            ReadCustomerBom bomPath
        End If
    Else
        LogLine "No BOM file found near active CAD (continuing with CAD-only fill)."
    End If
    LogLine "BomCount=" & BomCount
    LogDone "Find + read BOM (active CAD path)"

    BuildExportRowsFromBom
    WriteExportCheckCsv CurrentJobFolder & "\XT_Export_BOM_Match_Report.csv"

    Dim isStd As Boolean
    isStd = DetectBaseTypeIsStandard()
    gJobIsStandardBase = isStd
    LogLine "Base type: " & IIf(isStd, "STANDARD MOLD BASE", "POT / HOLDER BLOCK (BOM-driven)")
    DoEvents

    ' AI bridge: classify through the LOCAL AI service (standard bases only;
    ' BMS/pot-block jobs are guarded inside and keep the BOM-driven flow).
    LogStart "AI bridge classification"
    RunAiBridgeClassification CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv", isStd
    LogDone "AI bridge classification"

    ' gemini1 orientation BEFORE exports (active-CAD path previously skipped this).
    If isStd Then
        LogStart "Set STANDARD mold base orientation (active CAD)"
        SetStandardBaseOrientation swModel
        LogDone "Set STANDARD mold base orientation (active CAD)"
    Else
        LogStart "Set BMS pot-block TCP/top orientation (active CAD)"
        EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
        LogDone "Set BMS pot-block TCP/top orientation (active CAD)"
    End If
    ' STL matrix only here — L/W/T wait until after DXF locks the same views.
    CaptureFinalStandardViewsForStlCoordinateSystem swModel
    If FAST_QUOTE_MODE Then

        LogLine "FAST QUOTE: skipped ResolveAllLightWeight / Unsuppress-all heavy prep (active CAD, all base types)."

        On Error Resume Next
        PrepareAssemblyVisibilityFast swModel
        On Error GoTo ErrHandler

    Else

        UnsuppressAllAssemblyComponents swModel
        ShowAllAssemblyComponents swModel

    End If
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 50

    If isStd Then
        LogStart "Classify STANDARD mold base from active CAD"
        ClassifyStandardBasePlates
        CaptureStandardPurchasedFromCadIfNeeded
        LogDone "Classify STANDARD mold base from active CAD"

        LogStart "Set STANDARD top from classified A/B stack"
        If SetStandardBaseTopFromClassifiedStack(swModel) Then
            CaptureFinalStandardViewsForStlCoordinateSystem swModel
        End If
        LogDone "Set STANDARD top from classified A/B stack"
    End If

    BuildPullcoreList
    If WRITE_PCS_NAMING_ANALYSIS And Not FAST_QUOTE_MODE Then
        WritePcsNamingAnalysis CurrentJobFolder & "\PCS_Naming_Analysis.csv", isStd
    End If
    DoEvents

    If isStd Then
        LogStart "Refine STANDARD *Front from rails/latch after classify"
        If DefineStandardFrontFromRailsAndFootprint(swModel) Then
            swModel.ShowNamedView2 "*Top", 5
            StabilizeActiveView swModel, 50
            On Error Resume Next
            swModel.DeleteNamedView CMS_TOP_VIEW_NAME
            Err.Clear
            swModel.NameView CMS_TOP_VIEW_NAME
            On Error GoTo ErrHandler
            CaptureFinalStandardViewsForStlCoordinateSystem swModel
        End If
        LogDone "Refine STANDARD *Front from rails/latch after classify"
    End If

    ComputePullcoreQuote
    ComputePurchasedQuote

    If isStd Then
        LogStart "Apply primary PCS standard steel filter before STL"
        ApplyPrimaryPcsStandardQuoteFilter
        WriteStandardQuoteRowsDebugCsv CurrentJobFolder & "\Standard_Quote_Rows_Debug_BEFORE_STL.csv"
        LogDone "Apply primary PCS standard steel filter before STL"
    End If

    ' Export FIRST so DXF EnsureNativeDxfSourceUsesCmsTop locks the view frame.
    LogStart "Export base package (active CAD)"
    ExportBasePackage CurrentJobFolder & "\base"
    LogDone "Export base package (active CAD)"

    ' Assign Width/Length/Thickness only AFTER DXF views match CMS_TOP / *Front / *Right.
    LogStart "Assign CMS view-frame dims after DXF"
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 50
    CaptureCmsViewFrameFromModel swModel
    ApplyCmsViewDimsToAllParts
    WritePartDimensionCsv CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"

    If isStd Then
        RefreshStandardPlateDimsFromCurrentPartRoles
        ApplyPrimaryPcsStandardQuoteFilter
        WriteStandardQuoteRowsDebugCsv CurrentJobFolder & "\Standard_Quote_Rows_Debug.csv"
    End If

    WriteAllCadComponentsDebugCsv CurrentJobFolder & "\CAD_All_Components_Debug_FINAL.csv"

    LogDone "Assign CMS view-frame dims after DXF"

    If FILL_QUOTE_WORKBOOK Then
        LogStart "Fill Quote workbook from active CAD"
        If isStd Then FillStandardBaseQuote Else FillQuoteWorkbookFromBoundingBox
        LogDone "Fill Quote workbook from active CAD"
    End If

    If FILL_J000_STEEL_SHEET Then
        LogStart "Fill J000 steel sheet from active CAD"
        If isStd Then FillStandardBaseSteel Else FillJ000SteelSheet
        LogDone "Fill J000 steel sheet from active CAD"
    End If

    AiBridgeNotifyJobComplete IIf(isStd, "standard", "bms")

    LogLine "DONE ACTIVE CAD QUOTE. Output folder: "
    WriteMacroLaunchStatus "DONE", "RunActiveAssembly completed" & CurrentJobFolder
    LogLine "TOTAL ACTIVE RUN TIME: " & DateDiff("s", JobStartTime, Now) & "s   (log: " & RunLogPath & ")"
    If Not SUPPRESS_USER_PROMPTS Then
        MsgBox "Active CAD quote finished." & vbCrLf & _
               "Base type: " & IIf(isStd, "STANDARD", "BMS/POT (BOM)") & vbCrLf & _
               "BOM: " & IIf(bomPath = "", "(none)", bomPath) & vbCrLf & vbCrLf & _
               CurrentJobFolder, vbInformation, "CMS Base Export"
    End If

CleanExit:
    On Error Resume Next
    RestoreMainViewportGraphics
    Set swModel = swApp.ActiveDoc
    Exit Sub

ErrHandler:
    LogLine "RunActiveAssembly error. Step: " & CurrentStepName & "  Err " & Err.Number & ": " & Err.Description
    WriteMacroLaunchStatus "ERROR", "RunActiveAssembly error: " & Err.Description
    On Error Resume Next
    RestoreMainViewportGraphics
    MsgBox "Active assembly run failed at step: " & CurrentStepName & vbCrLf & Err.Description & vbCrLf & RunLogPath, vbCritical, "CMS Base Export"
End Sub

Private Function ActiveCadIsOpen() As Boolean
On Error GoTo nope
    ActiveCadIsOpen = False
    If swApp Is Nothing Then Set swApp = Application.SldWorks
    If swApp Is Nothing Then Exit Function
    Dim m As Object
    Set m = swApp.ActiveDoc
    If m Is Nothing Then Exit Function
    ActiveCadIsOpen = (m.GetType = swDocASSEMBLY Or m.GetType = swDocPART)
    Exit Function
nope:
    ActiveCadIsOpen = False
End Function

Private Function ActiveAssemblyIsOpen() As Boolean
    ActiveAssemblyIsOpen = ActiveCadIsOpen()
End Function

' ============================================================
' TRAINING XT EXPORT (dimensions only)
' Called by the webapp training scan when a job folder has CAD
' but no XT_Export_CAD_Dimensions.csv yet.
' Reads cms_training_xt.txt, opens CAD, scans geometry, writes
' the XT CSV into the training folder, then exits.
' ============================================================
Sub RunTrainingXtExport()
On Error GoTo ErrHandler
    Dim h As TrainingXtHandoff
    h = ReadTrainingXtHandoff()
    If h.JobFolder = "" Then
        WriteTrainingXtDone h.DoneFile, "ERROR", "", "Training handoff missing JobFolder"
        Exit Sub
    End If

    Set swApp = Application.SldWorks
    MacroStartTime = Now
    CurrentJobFolder = h.JobFolder
    CurrentJobNumber = h.JobId
    If CurrentJobNumber = "" Then CurrentJobNumber = GetFolderLeafName(CurrentJobFolder)
    If CurrentJobNumber = "" Then CurrentJobNumber = "TRAINING"
    RunLogPath = CurrentJobFolder & "\CMS_Training_XT_Log.txt"
    StartupLogPath = RunLogPath
    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
    MainViewportGraphicsDisabled = False
    Set swModel = Nothing

    LogLine "========================================"
    LogLine "TRAINING XT EXPORT STARTED"
    LogLine "Job folder: " & CurrentJobFolder
    LogLine "========================================"

    Dim extractFolder As String
    extractFolder = CurrentJobFolder & "\" & EXTRACT_FOLDER_NAME
    LogStart "Extract ZIP files"
    EnsureFolderDeep extractFolder
    ExtractAllZipFilesInJobFolder CurrentJobFolder, extractFolder
    FlattenExtractedZipContentsIntoJobFolder CurrentJobFolder, extractFolder
    If DELETE_EXTRACTED_ZIP_AFTER_FLATTEN Then DeleteFolderSafe extractFolder
    LogDone "Extract ZIP files"

    LogStart "Find CAD file"
    Dim cadCandidates As Collection
    Set cadCandidates = FindAllCadModelsRanked(CurrentJobFolder)
    AppendCadCandidates cadCandidates, FindAllCadModelsRanked(extractFolder)
    If cadCandidates.Count = 0 Then
        WriteTrainingXtDone h.DoneFile, "ERROR", "", "No CAD file found in training folder"
        GoTo CleanExit
    End If
    LogDone "Find CAD file"

    LogStart "Open CAD"
    Dim ci As Long
    Dim cadPath As String
    For ci = 1 To cadCandidates.Count
        cadPath = CStr(cadCandidates(ci))
        LogLine "Trying CAD candidate " & ci & "/" & cadCandidates.Count & ": " & cadPath
        Set swModel = OpenCadFile(cadPath)
        If Not swModel Is Nothing Then
            LogLine "CAD opened: " & cadPath
            Exit For
        End If
    Next ci
    If swModel Is Nothing Then
        WriteTrainingXtDone h.DoneFile, "ERROR", "", "Open CAD failed"
        GoTo CleanExit
    End If
    MainCadOpenedByMacro = True
    MainCadTitleForClose = swModel.GetTitle
    LogDone "Open CAD"

    Dim errs As Long
    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    EnsureSwHidden

    LogStart "Scan CAD parts (training XT export)"
    PartCount = 0
    ReDim parts(1 To 1)
    Set swAssy = Nothing
    ScanActiveSolidWorksDocument
    SortPartsByVolumeDescending
    ClassifyPotBlockPlatesFromCad
    LogLine "CAD PartCount=" & PartCount

    Dim outCsv As String
    If h.OutputCsv <> "" Then
        outCsv = h.OutputCsv
    Else
        outCsv = CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"
    End If
    WritePartDimensionCsv outCsv
    LogDone "Scan CAD parts (training XT export)"

    WriteTrainingXtDone h.DoneFile, "OK", outCsv, "Exported " & PartCount & " components"
    LogLine "TRAINING XT EXPORT DONE: " & outCsv

CleanExit:
    On Error Resume Next
    CloseCurrentJobCadIfNeeded
    RestoreMainViewportGraphics
    Exit Sub

ErrHandler:
    LogLine "RunTrainingXtExport error: " & Err.Description
    WriteTrainingXtDone h.DoneFile, "ERROR", "", Err.Description
    Resume CleanExit
End Sub

Private Function ReadTrainingXtHandoff() As TrainingXtHandoff
On Error GoTo eh
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(TRAINING_XT_HANDOFF) Then Exit Function
    Dim f As Integer, line As String, k As String, v As String, p As Long
    f = FreeFile
    Open TRAINING_XT_HANDOFF For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        p = InStr(line, "=")
        If p > 0 Then
            k = Trim(Left(line, p - 1))
            v = Trim(Mid(line, p + 1))
            Select Case UCase(k)
                Case "JOBFOLDER": ReadTrainingXtHandoff.JobFolder = v
                Case "JOBID": ReadTrainingXtHandoff.JobId = v
                Case "OUTPUTCSV": ReadTrainingXtHandoff.OutputCsv = v
                Case "DONEFILE": ReadTrainingXtHandoff.DoneFile = v
            End Select
        End If
    Loop
    Close #f
    If ReadTrainingXtHandoff.DoneFile = "" Then
        ReadTrainingXtHandoff.DoneFile = LOCAL_WORKSPACE_ROOT & "\cms_training_xt_done.txt"
    End If
    Exit Function
eh:
    LogLine "ReadTrainingXtHandoff error: " & Err.Description
    On Error Resume Next: Close #f
End Function

Private Sub WriteTrainingXtDone(ByVal donePath As String, ByVal status As String, ByVal xtCsv As String, ByVal message As String)
On Error Resume Next
    Dim p As String
    If donePath = "" Then donePath = LOCAL_WORKSPACE_ROOT & "\cms_training_xt_done.txt"
    p = donePath
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parent As String
    parent = fso.GetParentFolderName(p)
    If parent <> "" Then EnsureFolderDeep parent
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, "Status=" & status
    Print #f, "XtCsv=" & xtCsv
    Print #f, "PartCount=" & PartCount
    Print #f, "Message=" & message
    Close #f
End Sub

' ============================================================
' LAUNCHER ENTRY POINT
' Called by CMS_Launcher.vbs via swApp.RunMacro.
' Reads C-number and assigned quote # from the handoff file
' written by the launcher - no input box needed.
' ============================================================
Sub RunFromLauncher()
On Error GoTo ErrHandler

    Set swApp = Application.SldWorks
    WriteMacroLaunchStatus "STARTED", "RunFromLauncher entered"

    ' Live quote handoff must win over stale training handoff.
    Dim fsoTrain As Object
    Set fsoTrain = CreateObject("Scripting.FileSystemObject")

    If fsoTrain.FileExists(HANDOFF_FILE) Then
        On Error Resume Next
        If fsoTrain.FileExists(TRAINING_XT_HANDOFF) Then fsoTrain.DeleteFile TRAINING_XT_HANDOFF, True
        On Error GoTo ErrHandler
    ElseIf fsoTrain.FileExists(TRAINING_XT_HANDOFF) Then
        RunTrainingXtExport
        Exit Sub
    End If

    On Error Resume Next
    swApp.Visible = True
    swApp.UserControl = True
    On Error GoTo ErrHandler

    MacroStartTime = Now
    StartupLogPath = DOWNLOADS_FOLDER & "\CMS_Base_Export_Log.txt"
    RunLogPath = StartupLogPath

    Dim batch() As HandoffInfo
    Dim batchCount As Long
    Dim firstHandoff As HandoffInfo

    batchCount = ReadHandoffBatchFile(batch)

    If batchCount < 1 Then
        LogLine "Handoff file missing or empty - launcher automation requires at least one C-number."
        WriteMacroLaunchStatus "ERROR", "Handoff file missing CNum / BatchCount jobs"
        GoTo NormalEnd
    End If

    firstHandoff = batch(1)

    If firstHandoff.RootPath <> "" Then
        gRootJobPath = firstHandoff.RootPath
    Else
        gRootJobPath = CurrentMonthJobFolder()
    End If

    LogLine "========================================"
    LogLine "BASE EXPORT MACRO STARTED (from Launcher)"
    LogLine "Root path: " & gRootJobPath
    LogLine "Batch count: " & batchCount
    LogLine "========================================"

    Dim bi As Long
    For bi = 1 To batchCount
        LogLine "Batch handoff " & bi & "/" & batchCount & _
                ": CNum=" & batch(bi).CNum & _
                " QuoteNum=" & batch(bi).QuoteNum & _
                " CustJob=" & batch(bi).CustJob & _
                " JobFolder=" & batch(bi).JobFolder
    Next bi

    WriteMacroLaunchStatus "STARTED", "RunFromLauncher batch count=" & batchCount

    If batchCount = 1 Then

        Dim handoff As HandoffInfo
        handoff = batch(1)

        LogLine "Job from launcher: " & UCase$(Trim$(handoff.CNum))
        If handoff.QuoteNum <> "" Then LogLine "Assigned quote #:  " & handoff.QuoteNum
        If handoff.CustJob <> "" Then LogLine "Customer job #:    " & handoff.CustJob
        If handoff.SimilarTo <> "" Then LogLine "Similar to:        " & handoff.SimilarTo
        If handoff.ShipDate <> "" Then LogLine "Ship date:         " & handoff.ShipDate
        If handoff.CadPath <> "" Then LogLine "CadPath from handoff: " & handoff.CadPath

        If handoff.CadPath <> "" And IsGeneratedBaseCadPath(handoff.CadPath) Then
            LogLine "WARNING: handoff CadPath is a generated \base\ assembly — ignoring it and searching for original XT/STEP."
            LogLine "  Bad CadPath: " & handoff.CadPath
            handoff.CadPath = ""
        End If

        If ActiveCadIsOpen() Then
            Dim activePath As String
            activePath = ""
            On Error Resume Next
            If Not swApp.ActiveDoc Is Nothing Then activePath = CStr(swApp.ActiveDoc.GetPathName)
            On Error GoTo ErrHandler

            If IsGeneratedBaseCadPath(activePath) Then
                LogLine "WARNING: active CAD is a generated \base\ assembly — closing it and falling back to ProcessOneJob."
                LogLine "  Active path: " & activePath
                CloseAllDocumentsSafely
            Else
                LogLine "CAD already open in SolidWorks — quoting from active document"
                RunActiveAssemblyWithHandoff handoff
                GoTo NormalEnd
            End If
        End If

        If handoff.CadPath <> "" Then
            If fsoTrain.FileExists(handoff.CadPath) Then
                LogLine "Opening CadPath from handoff before ProcessOneJob: " & handoff.CadPath

                Set swModel = OpenCadFile(handoff.CadPath)

                If Not swModel Is Nothing Then
                    MainCadOpenedByMacro = True
                    MainCadTitleForClose = swModel.GetTitle

                    Dim errsOpen As Long
                    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errsOpen

                    LogLine "CAD opened from handoff — quoting from active document"
                    RunActiveAssemblyWithHandoff handoff
                    GoTo NormalEnd
                End If

                LogLine "OpenCadFile failed for handoff CadPath — falling back to ProcessOneJob"
            End If
        End If

    Else

        If ActiveCadIsOpen() Then
            LogLine "Batch has multiple jobs; closing active CAD so each job opens its own source."
            CloseAllDocumentsSafely
        End If

    End If

    Dim completed As Collection
    Dim failed As Collection

    Set completed = New Collection
    Set failed = New Collection

    Dim ok As Boolean
    Dim jobText As String

    For bi = 1 To batchCount

        jobText = UCase$(Trim$(batch(bi).CNum))

        If jobText <> "" Then

            WriteMacroLaunchStatus "STARTED", "Batch item " & bi & "/" & batchCount & " started: " & jobText

            LogLine "========================================"
            LogLine "BATCH QUOTE " & bi & "/" & batchCount & ": " & jobText
            LogLine "========================================"

            If batch(bi).RootPath <> "" Then
                gRootJobPath = batch(bi).RootPath
            ElseIf gRootJobPath = "" Then
                gRootJobPath = CurrentMonthJobFolder()
            End If

            ok = ProcessOneJobWithHandoff(jobText, batch(bi))

            If ok Then
                completed.Add jobText
                WriteMacroLaunchStatus "STARTED", "Batch item completed: " & jobText
            Else
                failed.Add jobText & IIf(LastJobFailReason <> "", "  ->  " & LastJobFailReason, "")
                WriteMacroLaunchStatus "ERROR", "Batch item failed: " & jobText & " " & LastJobFailReason
            End If

            CloseAllDocumentsSafely
            DoEvents

        End If

    Next bi

    CloseAllDocumentsSafely

    On Error Resume Next
    swApp.Visible = True
    On Error GoTo ErrHandler

    Dim summary As String
    summary = BuildBatchSummary(completed, failed)

    LogLine summary

    If failed.Count > 0 Then
        WriteMacroLaunchStatus "ERROR", "Batch completed with " & failed.Count & " failure(s)"
    Else
        WriteMacroLaunchStatus "DONE", "Batch completed successfully: " & completed.Count & " job(s)"
    End If

    If Not SUPPRESS_USER_PROMPTS Then
        MsgBox summary, IIf(failed.Count > 0, vbExclamation, vbInformation)
    End If

NormalEnd:
    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing Then swApp.Visible = True
    Exit Sub

ErrHandler:
    LogLine "RunFromLauncher error: " & Err.Description
    WriteMacroLaunchStatus "ERROR", "RunFromLauncher error: " & Err.Description

    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely

    If Not SUPPRESS_USER_PROMPTS Then
        MsgBox "Macro error: " & Err.Description & vbCrLf & RunLogPath, vbCritical
    End If
End Sub

Private Sub RunActiveAssemblyWithHandoff(ByRef h As HandoffInfo)
On Error GoTo ErrHandler
    AssignedQuoteNumber = h.QuoteNum
    CustomerJobNumber = h.CustJob
    CustomerPrefix = h.CustomerPrefix
    CustomerDisplayName = h.CustomerName
    SimilarToJob = h.SimilarTo
    ShipDateText = h.ShipDate
    gExactJobFolderName = h.JobFolder
    gHandoffAttachDir = h.AttachDir
    gProcessingHandoff = True

    If h.CNum <> "" Then
        CurrentJobNumber = UCase(Trim(h.CNum))
    End If

    ' Prefer the staged job folder / attach dir as output root when available.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If h.RootPath <> "" And h.JobFolder <> "" Then
        Dim cand As String
        cand = h.RootPath & "\" & h.JobFolder
        If fso.FolderExists(cand) Then
            NetworkJobFolder = cand
        End If
    End If
    If NetworkJobFolder = "" And h.AttachDir <> "" Then
        If fso.FolderExists(h.AttachDir) Then NetworkJobFolder = h.AttachDir
    End If

    RunActiveAssembly

    gProcessingHandoff = False
    gHandoffAttachDir = ""
    Exit Sub
ErrHandler:
    gProcessingHandoff = False
    gHandoffAttachDir = ""
    LogLine "RunActiveAssemblyWithHandoff error: " & Err.Description
End Sub

Private Function ReadHandoffFile() As HandoffInfo
On Error GoTo eh
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(HANDOFF_FILE) Then Exit Function
    Dim f As Integer, line As String, k As String, v As String, p As Long
    f = FreeFile
    Open HANDOFF_FILE For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        p = InStr(line, "=")
        If p > 0 Then
            k = Trim(Left(line, p - 1))
            v = Trim(Mid(line, p + 1))
            Select Case UCase(k)
                Case "CNUM":      ReadHandoffFile.CNum     = v
                Case "QUOTENUM":  ReadHandoffFile.QuoteNum = v
                Case "CUSTJOB":   ReadHandoffFile.CustJob  = v
                Case "SIMILARTO": ReadHandoffFile.SimilarTo = v
                Case "SHIPDATE":  ReadHandoffFile.ShipDate  = v
                Case "ROOTPATH":  ReadHandoffFile.RootPath  = v
                Case "JOBFOLDER": ReadHandoffFile.JobFolder = v
                Case "CUSTOMERPREFIX": ReadHandoffFile.CustomerPrefix = v
                Case "CUSTOMERNAME": ReadHandoffFile.CustomerName = v
                Case "ATTACHDIR": ReadHandoffFile.AttachDir = v
                Case "CADPATH": ReadHandoffFile.CadPath = v
            End Select
        End If
    Loop
    Close #f
    Exit Function
eh:
    LogLine "ReadHandoffFile error: " & Err.Description
    On Error Resume Next: Close #f
End Function

Private Function DictGetText(ByVal dict As Object, ByVal keyName As String) As String
On Error Resume Next
    DictGetText = ""
    If dict Is Nothing Then Exit Function
    If dict.Exists(UCase$(Trim$(keyName))) Then
        DictGetText = Trim$(CStr(dict(UCase$(Trim$(keyName)))))
    End If
End Function

Private Function BatchField(ByVal dict As Object, ByVal idx As Long, ByVal fieldName As String) As String
    Dim k1 As String
    Dim k2 As String
    Dim k3 As String

    k1 = "JOB" & CStr(idx) & "." & UCase$(fieldName)
    k2 = "JOB" & CStr(idx) & "_" & UCase$(fieldName)
    k3 = CStr(idx) & "." & UCase$(fieldName)

    BatchField = DictGetText(dict, k1)
    If BatchField <> "" Then Exit Function

    BatchField = DictGetText(dict, k2)
    If BatchField <> "" Then Exit Function

    BatchField = DictGetText(dict, k3)
End Function

Private Sub ApplyHandoffField(ByRef h As HandoffInfo, ByVal keyName As String, ByVal valueText As String)
    keyName = UCase$(Trim$(keyName))
    valueText = Trim$(valueText)

    Select Case keyName
        Case "CNUM": h.CNum = valueText
        Case "QUOTENUM": h.QuoteNum = valueText
        Case "CUSTJOB": h.CustJob = valueText
        Case "SIMILARTO": h.SimilarTo = valueText
        Case "SHIPDATE": h.ShipDate = valueText
        Case "ROOTPATH": h.RootPath = valueText
        Case "JOBFOLDER": h.JobFolder = valueText
        Case "CUSTOMERPREFIX": h.CustomerPrefix = valueText
        Case "CUSTOMERNAME": h.CustomerName = valueText
        Case "ATTACHDIR": h.AttachDir = valueText
        Case "CADPATH": h.CadPath = valueText
    End Select
End Sub

Private Function HandoffFromDictUnprefixed(ByVal dict As Object) As HandoffInfo
    ApplyHandoffField HandoffFromDictUnprefixed, "CNUM", DictGetText(dict, "CNUM")
    ApplyHandoffField HandoffFromDictUnprefixed, "QUOTENUM", DictGetText(dict, "QUOTENUM")
    ApplyHandoffField HandoffFromDictUnprefixed, "CUSTJOB", DictGetText(dict, "CUSTJOB")
    ApplyHandoffField HandoffFromDictUnprefixed, "SIMILARTO", DictGetText(dict, "SIMILARTO")
    ApplyHandoffField HandoffFromDictUnprefixed, "SHIPDATE", DictGetText(dict, "SHIPDATE")
    ApplyHandoffField HandoffFromDictUnprefixed, "ROOTPATH", DictGetText(dict, "ROOTPATH")
    ApplyHandoffField HandoffFromDictUnprefixed, "JOBFOLDER", DictGetText(dict, "JOBFOLDER")
    ApplyHandoffField HandoffFromDictUnprefixed, "CUSTOMERPREFIX", DictGetText(dict, "CUSTOMERPREFIX")
    ApplyHandoffField HandoffFromDictUnprefixed, "CUSTOMERNAME", DictGetText(dict, "CUSTOMERNAME")
    ApplyHandoffField HandoffFromDictUnprefixed, "ATTACHDIR", DictGetText(dict, "ATTACHDIR")
    ApplyHandoffField HandoffFromDictUnprefixed, "CADPATH", DictGetText(dict, "CADPATH")
End Function

Private Function HandoffFromDictIndexed(ByVal dict As Object, ByVal idx As Long, ByRef defaults As HandoffInfo) As HandoffInfo
    HandoffFromDictIndexed = defaults

    Dim v As String

    v = BatchField(dict, idx, "CNUM")
    If v <> "" Then HandoffFromDictIndexed.CNum = v

    v = BatchField(dict, idx, "QUOTENUM")
    If v <> "" Then HandoffFromDictIndexed.QuoteNum = v

    v = BatchField(dict, idx, "CUSTJOB")
    If v <> "" Then HandoffFromDictIndexed.CustJob = v

    v = BatchField(dict, idx, "SIMILARTO")
    If v <> "" Then HandoffFromDictIndexed.SimilarTo = v

    v = BatchField(dict, idx, "SHIPDATE")
    If v <> "" Then HandoffFromDictIndexed.ShipDate = v

    v = BatchField(dict, idx, "ROOTPATH")
    If v <> "" Then HandoffFromDictIndexed.RootPath = v

    v = BatchField(dict, idx, "JOBFOLDER")
    If v <> "" Then HandoffFromDictIndexed.JobFolder = v

    v = BatchField(dict, idx, "CUSTOMERPREFIX")
    If v <> "" Then HandoffFromDictIndexed.CustomerPrefix = v

    v = BatchField(dict, idx, "CUSTOMERNAME")
    If v <> "" Then HandoffFromDictIndexed.CustomerName = v

    v = BatchField(dict, idx, "ATTACHDIR")
    If v <> "" Then HandoffFromDictIndexed.AttachDir = v

    v = BatchField(dict, idx, "CADPATH")
    If v <> "" Then HandoffFromDictIndexed.CadPath = v
End Function

Private Sub AddHandoffToArray(ByRef arr() As HandoffInfo, ByRef n As Long, ByRef h As HandoffInfo)
    If Trim$(h.CNum) = "" Then Exit Sub

    n = n + 1

    If n = 1 Then
        ReDim arr(1 To 1)
    Else
        ReDim Preserve arr(1 To n)
    End If

    arr(n) = h
End Sub

Private Function ReadHandoffBatchFile(ByRef jobs() As HandoffInfo) As Long
On Error GoTo ErrHandler

    ReadHandoffBatchFile = 0

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(HANDOFF_FILE) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim f As Integer
    Dim line As String
    Dim p As Long
    Dim k As String
    Dim v As String

    f = FreeFile
    Open HANDOFF_FILE For Input As #f

    Do While Not EOF(f)
        Line Input #f, line

        line = Trim$(line)

        If line <> "" Then
            If Left$(line, 1) <> "#" Then
                p = InStr(line, "=")

                If p > 0 Then
                    k = UCase$(Trim$(Left$(line, p - 1)))
                    v = Trim$(Mid$(line, p + 1))
                    dict(k) = v
                End If
            End If
        End If
    Loop

    Close #f

    Dim defaults As HandoffInfo
    defaults = HandoffFromDictUnprefixed(dict)

    Dim batchCount As Long
    batchCount = CLng(Val(DictGetText(dict, "BATCHCOUNT")))

    If batchCount <= 0 Then batchCount = CLng(Val(DictGetText(dict, "JOBCOUNT")))

    Dim n As Long
    n = 0

    Dim i As Long
    Dim h As HandoffInfo

    If batchCount > 0 Then

        For i = 1 To batchCount
            h = HandoffFromDictIndexed(dict, i, defaults)

            If Trim$(h.CNum) <> "" Then
                If Trim$(h.QuoteNum) = "" Then h.QuoteNum = h.CNum
                AddHandoffToArray jobs, n, h
            End If
        Next i

        ReadHandoffBatchFile = n
        Exit Function

    End If

    Dim cList As Collection
    Set cList = ParseJobInputList(defaults.CNum)

    If cList Is Nothing Or cList.Count = 0 Then
        ReadHandoffBatchFile = 0
        Exit Function
    End If

    If cList.Count = 1 Then

        h = defaults
        h.CNum = CStr(cList(1))
        If Trim$(h.QuoteNum) = "" Then h.QuoteNum = h.CNum
        AddHandoffToArray jobs, n, h

    Else

        For i = 1 To cList.Count
            h = defaults
            h.CNum = CStr(cList(i))
            h.QuoteNum = h.CNum

            h.CustJob = ""
            h.JobFolder = ""
            h.CadPath = ""
            h.AttachDir = ""

            AddHandoffToArray jobs, n, h
        Next i

    End If

    ReadHandoffBatchFile = n
    Exit Function

ErrHandler:
    LogLine "ReadHandoffBatchFile error: " & Err.Description
    On Error Resume Next
    Close #f
    ReadHandoffBatchFile = 0
End Function


' Current month job folder, e.g. \\Mycloudex2ultra\mexico\Cameron's stuff\RON'S QUOTES\000000006.June 2026
Private Function CurrentMonthJobFolder() As String
    CurrentMonthJobFolder = JOB_ROOT_BASE & "\" & MonthFolderName(Now)
End Function

Private Function MonthFolderName(ByVal d As Date) As String
    Dim m As Long, y As Long
    m = Month(d): y = Year(d)
    MonthFolderName = Right$("00000000" & m, 9) & "." & MonthName(m) & " " & y
End Function

' Wrapper that injects launcher info into the job globals before processing
Private Function ProcessOneJobWithHandoff(ByVal jobText As String, ByRef h As HandoffInfo) As Boolean
    AssignedQuoteNumber = h.QuoteNum
    CustomerJobNumber   = h.CustJob
    CustomerPrefix      = h.CustomerPrefix
    CustomerDisplayName = h.CustomerName
    SimilarToJob        = h.SimilarTo
    ShipDateText        = h.ShipDate
    gExactJobFolderName = h.JobFolder
    gHandoffAttachDir = h.AttachDir
    gHandoffCadPath = h.CadPath
    If IsGeneratedBaseCadPath(gHandoffCadPath) Then
        LogLine "ProcessOneJobWithHandoff: clearing generated \base\ CadPath: " & gHandoffCadPath
        gHandoffCadPath = ""
        h.CadPath = ""
    End If
    gProcessingHandoff = True
    ProcessOneJobWithHandoff = ProcessOneJob(jobText)
    gProcessingHandoff = False
    gHandoffAttachDir = ""
    gHandoffCadPath = ""
End Function

Private Function ProcessOneJob(ByVal jobSearchText As String) As Boolean
On Error GoTo ErrHandler

    ProcessOneJob = False
    LastJobFailReason = ""
    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
    MainViewportGraphicsDisabled = False
    Set swModel = Nothing

    CurrentJobNumber = UCase(Trim(jobSearchText))
    WriteMacroLaunchStatus "STARTED", "ProcessOneJob started for " & CurrentJobNumber
    JobStartTime = Now
    FinalStlCoordFrameReady = False
    ResetCmsViewFrame
    Dim stlCoordI As Long
    For stlCoordI = 0 To 8
        FinalStlCoordM(stlCoordI) = 0#
    Next stlCoordI
    If Not gProcessingHandoff Then
        CustomerJobNumber = ""
        CustomerPrefix = ""
        CustomerDisplayName = ""
        AssignedQuoteNumber = ""
        SimilarToJob = ""
        ShipDateText = ""
        gExactJobFolderName = ""
        gHandoffAttachDir = ""
        gHandoffCadPath = ""
    End If
    CurrentJobFolder = ""
    NetworkJobFolder = ""
    LocalJobFolder = ""
    JobBaseName = ""

    LogStart "Find job folder"
    ' Prefer the EXACT folder the launcher just created (handoff JobFolder).
    ' This avoids matching a stale folder that happens to share the C-number
    ' (e.g. an old BMS-868000000-C18601 test folder).
    NetworkJobFolder = ""
    Dim fsoJ As Object
    Set fsoJ = CreateObject("Scripting.FileSystemObject")
    If gExactJobFolderName <> "" Then
        Dim cand As String
        cand = gRootJobPath & "\" & gExactJobFolderName
        If fsoJ.FolderExists(cand) Then
            NetworkJobFolder = cand
            LogLine "Using exact job folder from handoff: " & gExactJobFolderName
        Else
            LogLine "Handoff job folder not found, falling back to search: " & cand
        End If
    End If
    If NetworkJobFolder = "" Then NetworkJobFolder = FindJobFolderByText(gRootJobPath, CurrentJobNumber)
    If NetworkJobFolder = "" And gHandoffAttachDir <> "" Then
        If fsoJ.FolderExists(gHandoffAttachDir) Then
            NetworkJobFolder = gHandoffAttachDir
            LogLine "Using AttachDir fallback from handoff: " & gHandoffAttachDir
        End If
    End If
    ' Prefer launcher/webapp-staged local workspace when network folder missing.
    If NetworkJobFolder = "" Then
        Dim localOnly As String
        localOnly = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber)
        If fsoJ.FolderExists(localOnly) Then
            NetworkJobFolder = localOnly
            LogLine "Using staged local workspace: " & localOnly
        End If
    End If
    LogLine "Job folder result: " & NetworkJobFolder
    If NetworkJobFolder = "" Then
        LogErrorText "Could not find job folder for: " & CurrentJobNumber
        GoTo CleanExit
    End If
    LogDone "Find job folder"

    JobBaseName = GetFolderLeafName(NetworkJobFolder)
    If JobBaseName = "" Then JobBaseName = CurrentJobNumber

    LogStart "Prepare local job workspace"
    Dim stagedLocal As String
    stagedLocal = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber)
    If gHandoffCadPath <> "" Then
        If IsGeneratedBaseCadPath(gHandoffCadPath) Then
            LogLine "WARNING: clearing generated \\base\\ CadPath before staging: " & gHandoffCadPath
            gHandoffCadPath = ""
        End If
    End If
    ' If launcher already pulled files into C:\CMS_Local_Workspace\C##### and set
    ' CadPath to the local XT, keep that folder (do not wipe while SolidWorks has it open).
    If fsoJ.FolderExists(stagedLocal) Then
        If StrComp(UCase(NetworkJobFolder), UCase(stagedLocal), vbTextCompare) = 0 Then
            LocalJobFolder = stagedLocal
            LogLine "Job source is already local workspace — skip wipe/recopy."
        ElseIf gHandoffCadPath <> "" Then
            If InStr(1, UCase(gHandoffCadPath), UCase(stagedLocal & "\"), vbTextCompare) = 1 Then
                If fsoJ.FileExists(gHandoffCadPath) Then
                    LocalJobFolder = stagedLocal
                    LogLine "Launcher-staged local XT present — skip wipe/recopy: " & gHandoffCadPath
                End If
            End If
        End If
    End If
    If LocalJobFolder = "" Then
        PrepareLocalJobWorkspace NetworkJobFolder, CurrentJobNumber, LocalJobFolder
    End If
    If LocalJobFolder = "" Then
        LogErrorText "Could not create local workspace for: " & CurrentJobNumber
        GoTo CleanExit
    End If
    CurrentJobFolder = LocalJobFolder
    RunLogPath = CurrentJobFolder & "\CMS_Base_Export_Log.txt"
    LogLine "Local job folder: " & CurrentJobFolder
    LogDone "Prepare local job workspace"

    Dim extractFolder As String
    extractFolder = CurrentJobFolder & "\" & EXTRACT_FOLDER_NAME

    LogStart "Extract ZIP files"
    EnsureFolderDeep extractFolder
    ExtractAllZipFilesInJobFolder CurrentJobFolder, extractFolder
    FlattenExtractedZipContentsIntoJobFolder CurrentJobFolder, extractFolder
    If DELETE_EXTRACTED_ZIP_AFTER_FLATTEN Then DeleteFolderSafe extractFolder
    LogDone "Extract ZIP files"

    LogStart "Find CAD file"
    Dim cadCandidates As Collection
    Set cadCandidates = New Collection

    If gHandoffCadPath <> "" Then
        Dim fsoCad As Object
        Set fsoCad = CreateObject("Scripting.FileSystemObject")
        If IsGeneratedBaseCadPath(gHandoffCadPath) Then
            LogLine "WARNING: ignoring handoff CadPath under \base\: " & gHandoffCadPath
            gHandoffCadPath = ""
        ElseIf IsForeignJobCadName(gHandoffCadPath) Then
            LogLine "WARNING: ignoring foreign-job handoff CadPath: " & gHandoffCadPath
            gHandoffCadPath = ""
        ElseIf fsoCad.FileExists(gHandoffCadPath) Then
            cadCandidates.Add gHandoffCadPath
            LogLine "Using CadPath from handoff first: " & gHandoffCadPath
        End If
    End If

    AppendCadCandidates cadCandidates, FindAllCadModelsRanked(CurrentJobFolder)
    AppendCadCandidates cadCandidates, FindAllCadModelsRanked(extractFolder)
    If cadCandidates.Count = 0 Then
        LogErrorText "No CAD file found."
        GoTo CleanExit
    End If
    LogDone "Find CAD file"

    LogStart "Open CAD"
    Dim ci As Long
    Dim cadPath As String
    For ci = 1 To cadCandidates.Count
        cadPath = CStr(cadCandidates(ci))
        LogLine "Trying CAD candidate " & ci & "/" & cadCandidates.Count & ": " & cadPath
        Set swModel = OpenCadFile(cadPath)
        If Not swModel Is Nothing Then
            LogLine "CAD opened: " & cadPath
            Exit For
        End If
    Next ci
    If swModel Is Nothing Then
        LogErrorText "Open CAD failed (tried " & cadCandidates.Count & " file(s))."
        GoTo CleanExit
    End If
    MainCadOpenedByMacro = True
    MainCadTitleForClose = swModel.GetTitle
    LogDone "Open CAD"

    Dim errs As Long
    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    EnsureSwHidden
    DisableMainViewportGraphics

    ' --- Scan parts FIRST so the plate classification (TCP/BCP, holders, pots) is
    '     available, THEN orient. Orientation reads gIdxTCP/BCP and the holder/pot
    '     indices, so it must run after ClassifyPotBlockPlatesFromCad. ---
    LogStart "Scan CAD parts"
    PartCount = 0
    ReDim parts(1 To 1)
    Set swAssy = Nothing
    ScanActiveSolidWorksDocument
    SortPartsByVolumeDescending
    ClassifyPotBlockPlatesFromCad
    LogLine "CAD PartCount=" & PartCount
    WritePartDimensionCsv CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"
    WriteAllCadComponentsDebugCsv CurrentJobFolder & "\CAD_All_Components_Debug_PRE_ORIENT.csv"
    WriteJobFileInventoryCsv CurrentJobFolder, CurrentJobFolder & "\Job_File_Inventory.csv"
    LogDone "Scan CAD parts"

    LogStart "Find + read BOM"
    BomCount = 0
    ReDim BomRows(1 To 1)
    LoadPurchasedPriceList
    Dim bomPath As String
    bomPath = FindCustomerBomFile(CurrentJobFolder)
    gDiagBomPath = bomPath
    If bomPath <> "" Then
        LogLine "BOM selected: " & bomPath
        If LCase(GetFileExtension(bomPath)) = "pdf" Then
            If READ_PDF_BOM_WITH_PDFTOTEXT Then ReadCustomerBomPdfUsingPdfToText bomPath
        Else
            ReadCustomerBom bomPath
        End If
    Else
        LogLine "No BOM file found (continuing with CAD-only Excel fill)."
    End If
    LogLine "BomCount=" & BomCount
    LogDone "Find + read BOM"

    ' Re-activate the base (scan may have loaded component part docs) before export.
    Dim reErrs As Long
    swApp.ActivateDoc3 swModel.GetTitle, False, 0, reErrs
    Set swModel = swApp.ActiveDoc

    LogStart "Match BOM to CAD"
    ExportCount = 0
    ReDim ExportRows(1 To 1)
    BuildExportRowsFromBom
    WriteExportCheckCsv CurrentJobFolder & "\XT_Export_BOM_Match_Report.csv"
    LogLine "ExportCount=" & ExportCount
    LogDone "Match BOM to CAD"

    Dim isStd As Boolean
    isStd = DetectBaseTypeIsStandard()
    gJobIsStandardBase = isStd
    LogLine "Base type: " & IIf(isStd, "STANDARD MOLD BASE", "POT / HOLDER BLOCK")

    ' AI bridge: classify through the LOCAL AI service (standard bases only;
    ' BMS/pot-block jobs are guarded inside and keep the BOM-driven flow).
    LogStart "AI bridge classification"
    RunAiBridgeClassification CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv", isStd
    LogDone "AI bridge classification"

    If isStd Then
        LogStart "Set STANDARD mold base orientation"
        SetStandardBaseOrientation swModel
        LogDone "Set STANDARD mold base orientation"
    Else
        LogStart "Set BMS pot-block TCP/top orientation from matched holder/pot/TCP"
        EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
        LogDone "Set BMS pot-block TCP/top orientation from matched holder/pot/TCP"
    End If

    ' gemini1: capture corrected *Front for STL. Do NOT assign L/W/T yet —
    ' wait until DXF has locked CMS_TOP / *Front / *Right (same frame as Width).
    CaptureFinalStandardViewsForStlCoordinateSystem swModel

    If isStd Then
        LogStart "Classify STANDARD mold base plates"
        ClassifyStandardBasePlates
        CaptureStandardPurchasedFromCadIfNeeded
        LogDone "Classify STANDARD mold base plates"

        LogStart "Set STANDARD top from classified A/B stack"
        If SetStandardBaseTopFromClassifiedStack(swModel) Then
            CaptureFinalStandardViewsForStlCoordinateSystem swModel
        End If
        LogDone "Set STANDARD top from classified A/B stack"
    End If

    BuildPullcoreList

    If WRITE_PCS_NAMING_ANALYSIS And Not FAST_QUOTE_MODE Then
        WritePcsNamingAnalysis CurrentJobFolder & "\PCS_Naming_Analysis.csv", isStd
    End If

    ' After standard stack/rails/latch roles are known, refine *Front and
    ' re-capture the STL coordinate frame (dims still deferred until after DXF).
    If isStd Then
        LogStart "Refine STANDARD *Front from rails/latch after classify"
        If DefineStandardFrontFromRailsAndFootprint(swModel) Then
            swModel.ShowNamedView2 "*Top", 5
            StabilizeActiveView swModel, 50
            On Error Resume Next
            swModel.DeleteNamedView CMS_TOP_VIEW_NAME
            Err.Clear
            swModel.NameView CMS_TOP_VIEW_NAME
            On Error GoTo ErrHandler
            CaptureFinalStandardViewsForStlCoordinateSystem swModel
        End If
        LogDone "Refine STANDARD *Front from rails/latch after classify"
    End If

    ComputePullcoreQuote
    ComputePurchasedQuote

    If isStd Then
        LogStart "Apply primary PCS standard steel filter before STL"
        ApplyPrimaryPcsStandardQuoteFilter
        WriteStandardQuoteRowsDebugCsv CurrentJobFolder & "\Standard_Quote_Rows_Debug_BEFORE_STL.csv"
        LogDone "Apply primary PCS standard steel filter before STL"
    End If

    ' Fast mode must skip heavy SolidWorks prep for BOTH standard and BMS jobs.
    If FAST_QUOTE_MODE Then

        LogLine "FAST QUOTE: skipped ResolveAllLightWeight / Unsuppress-all heavy prep (all base types)."

        On Error Resume Next
        PrepareAssemblyVisibilityFast swModel
        On Error GoTo ErrHandler

    Else

        On Error Resume Next
        swModel.ResolveAllLightWeightComponents True
        On Error GoTo ErrHandler

        UnsuppressAllAssemblyComponents swModel
        ShowAllAssemblyComponents swModel

    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 50

    LogStart "Export base package"
    ExportBasePackage CurrentJobFolder & "\base"
    LogDone "Export base package"

    ' Width/Length/Thickness from the same views DXF just used.
    LogStart "Assign CMS view-frame dims after DXF"
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 50
    CaptureCmsViewFrameFromModel swModel
    ApplyCmsViewDimsToAllParts
    WritePartDimensionCsv CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"

    If isStd Then
        RefreshStandardPlateDimsFromCurrentPartRoles
        ApplyPrimaryPcsStandardQuoteFilter
        WriteStandardQuoteRowsDebugCsv CurrentJobFolder & "\Standard_Quote_Rows_Debug.csv"
    End If

    WriteAllCadComponentsDebugCsv CurrentJobFolder & "\CAD_All_Components_Debug_FINAL.csv"

    LogDone "Assign CMS view-frame dims after DXF"

    If FILL_QUOTE_WORKBOOK Then
        LogStart "Fill Quote workbook"
        If isStd Then FillStandardBaseQuote Else FillQuoteWorkbookFromBoundingBox
        LogDone "Fill Quote workbook"
    End If
    If FILL_J000_STEEL_SHEET Then
        LogStart "Fill J000 steel sheet"
        If isStd Then FillStandardBaseSteel Else FillJ000SteelSheet
        LogDone "Fill J000 steel sheet"
    End If

    If RUN_VISUAL_MOLD_INSPECTION And Not FAST_QUOTE_MODE Then
        LogStart "Visual mold inspection"
        RunVisualMoldInspection
        LogDone "Visual mold inspection"
    End If

    OrganizeJobFiles

    PublishJobOutputs

    ' Move the loose .sldprt part files into the base subfolder (close docs first
    ' so SolidWorks releases the file locks).
    CloseAllDocumentsSafely
    MoveLooseSolidWorksPartsToBaseFolder
    SyncCompletedJobToNetworkFolder

    AiBridgeNotifyJobComplete IIf(isStd, "standard", "bms")

    gLastJobDiag = "BOM file:  " & IIf(gDiagBomPath = "", "(NONE FOUND)", gDiagBomPath) & vbCrLf & _
                   "BOM rows read:  " & BomCount & vbCrLf & _
                   "Purchased parts matched:  " & PpCount & vbCrLf & _
                   "Email:  " & IIf(gEmailStatus = "", "(step not reached)", gEmailStatus)

    LogLine "DONE JOB " & CurrentJobNumber & ". Output folder: " & CurrentJobFolder
    LogLine "TOTAL JOB TIME: " & DateDiff("s", JobStartTime, Now) & "s   (log: " & RunLogPath & ")"
    ProcessOneJob = True
    WriteMacroLaunchStatus "DONE", "ProcessOneJob completed for " & CurrentJobNumber

CleanExit:
    On Error Resume Next
    RestoreMainViewportGraphics
    If Not swApp Is Nothing Then
        If Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    End If
    CloseCurrentJobCadIfNeeded
    CloseAllDocumentsSafely
    Set swModel = Nothing
    Exit Function

ErrHandler:
    LogLine "FATAL JOB ERROR. Job: " & CurrentJobNumber & "  Step: " & CurrentStepName & "  Err " & Err.Number & ": " & Err.Description
    LastJobFailReason = "Step '" & CurrentStepName & "' - Err " & Err.Number & ": " & Err.Description
    WriteMacroLaunchStatus "ERROR", LastJobFailReason
    ProcessOneJob = False
    Resume CleanExit
End Function

Private Function ParseJobInputList(ByVal inputText As String) As Collection
On Error GoTo ErrHandler
    Dim result As New Collection
    inputText = Replace(inputText, vbCr, " ")
    inputText = Replace(inputText, vbLf, " ")
    inputText = Replace(inputText, vbTab, " ")
    inputText = Replace(inputText, ",", " ")
    inputText = Replace(inputText, ";", " ")
    Do While InStr(inputText, "  ") > 0
        inputText = Replace(inputText, "  ", " ")
    Loop
    inputText = Trim(inputText)
    If inputText = "" Then
        Set ParseJobInputList = result
        Exit Function
    End If
    Dim arr() As String
    arr = Split(inputText, " ")
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim token As String
    For i = LBound(arr) To UBound(arr)
        token = UCase(Trim(CStr(arr(i))))
        If token <> "" Then
            If Not dict.Exists(token) Then
                dict(token) = True
                result.Add token
            End If
        End If
    Next i
    Set ParseJobInputList = result
    Exit Function
ErrHandler:
    Set ParseJobInputList = New Collection
End Function

Private Function BuildBatchSummary(ByVal completed As Collection, ByVal failed As Collection) As String
On Error GoTo ErrHandler
    Dim s As String
    Dim i As Long
    s = "Base export complete." & vbCrLf & vbCrLf
    s = s & "Completed: " & completed.Count & vbCrLf
    For i = 1 To completed.Count
        s = s & "  - " & CStr(completed(i)) & vbCrLf
    Next i
    s = s & vbCrLf & "Failed: " & failed.Count & vbCrLf
    For i = 1 To failed.Count
        s = s & "  - " & CStr(failed(i)) & vbCrLf
    Next i
    If gLastJobDiag <> "" Then
        s = s & vbCrLf & "----- BOM / Components / Email -----" & vbCrLf & gLastJobDiag & vbCrLf
    End If
    BuildBatchSummary = s
    Exit Function
ErrHandler:
    BuildBatchSummary = "Base export complete."
End Function

Private Sub EnsureSwHidden()
On Error Resume Next
    If RUN_SOLIDWORKS_INVISIBLE Then
        If Not swApp Is Nothing Then
            If swApp.Visible Then swApp.Visible = False
        End If
    End If
End Sub

Private Sub CloseCurrentJobCadIfNeeded()
On Error Resume Next
    If MainCadOpenedByMacro Then
        If MainCadTitleForClose <> "" Then swApp.CloseDoc MainCadTitleForClose
    End If
    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
End Sub

Private Sub CloseAllDocumentsSafely()
On Error Resume Next
    If swApp Is Nothing Then Exit Sub
    swApp.CloseAllDocuments True
    Dim swDoc As Object
    Dim nextDoc As Object
    Set swDoc = swApp.GetFirstDocument
    Do While Not swDoc Is Nothing
        Set nextDoc = swDoc.GetNext
        On Error Resume Next
        swApp.CloseDoc swDoc.GetTitle
        On Error Resume Next
        Set swDoc = nextDoc
    Loop
End Sub

Private Sub DeleteFolderSafe(ByVal folderPath As String)
On Error Resume Next
    If folderPath = "" Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
        LogLine "Deleted folder: " & folderPath
    End If
End Sub

' ============================================================
' LOCAL STAGING
' ============================================================
Private Sub PrepareLocalJobWorkspace(ByVal sourceNetworkFolder As String, _
                                     ByVal jobNumber As String, _
                                     ByRef localFolderOut As String)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolderDeep LOCAL_WORKSPACE_ROOT
    localFolderOut = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(jobNumber)
    If fso.FolderExists(localFolderOut) Then
        On Error Resume Next
        fso.DeleteFolder localFolderOut, True
        On Error GoTo ErrHandler
    End If
    EnsureFolderDeep localFolderOut
    If fso.FolderExists(localFolderOut) = False Then
        localFolderOut = ""
        Exit Sub
    End If
    If CopyFolderWithRobocopy(sourceNetworkFolder, localFolderOut) Then
        LogLine "Local copy complete (robocopy)."
    Else
        LogLine "Robocopy failed. Using VBA copy fallback."
        CopyFolderContents sourceNetworkFolder, localFolderOut
    End If
    Exit Sub
ErrHandler:
    LogLine "PrepareLocalJobWorkspace error: " & Err.Description
    On Error Resume Next
    If fso Is Nothing Then
        localFolderOut = ""
    ElseIf fso.FolderExists(localFolderOut) = False Then
        localFolderOut = ""
    End If
End Sub

Private Function CopyFolderWithRobocopy(ByVal src As String, ByVal dst As String) As Boolean
On Error GoTo ErrHandler
    CopyFolderWithRobocopy = False
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim cleanSrc As String
    cleanSrc = src
    Do While Right(cleanSrc, 1) = "\"
        cleanSrc = Left(cleanSrc, Len(cleanSrc) - 1)
    Loop
    Dim excludeExtract As String
    excludeExtract = cleanSrc & "\" & EXTRACT_FOLDER_NAME
    Dim cmd As String
    cmd = "cmd /c robocopy " & Chr(34) & cleanSrc & Chr(34) & " " & Chr(34) & dst & Chr(34) & _
          " /MIR /XD " & Chr(34) & excludeExtract & Chr(34) & _
          " /R:1 /W:1 /MT:16 /NFL /NDL /NJH /NJS /NP"
    Dim rc As Long
    rc = sh.Run(cmd, 0, True)
    LogLine "robocopy exit code: " & rc
    If rc < 8 Then CopyFolderWithRobocopy = True
    Exit Function
ErrHandler:
    LogLine "CopyFolderWithRobocopy error: " & Err.Description
    CopyFolderWithRobocopy = False
End Function

' ============================================================
' ZIP
' ============================================================
Private Sub ExtractAllZipFilesInJobFolder(ByVal jobFolder As String, ByVal extractRoot As String)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(jobFolder) Then Exit Sub
    EnsureFolderDeep extractRoot
    Dim zips As Collection
    Set zips = New Collection
    SearchZipFilesRecursive fso.GetFolder(jobFolder), zips
    LogLine "ZIP count=" & zips.Count
    Dim i As Long
    For i = 1 To zips.Count
        LogLine "Extracting ZIP " & i & "/" & zips.Count & ": " & CStr(zips(i))
        UnzipOneZipRobust CStr(zips(i)), extractRoot
    Next i
    Exit Sub
ErrHandler:
    LogLine "ExtractAllZipFilesInJobFolder error: " & Err.Description
End Sub

Private Sub SearchZipFilesRecursive(ByVal folder As Object, ByRef zips As Collection)
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderName As String
    folderName = UCase(folder.Name)
    If folderName = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub
    Dim file As Object
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.path)) = "zip" Then zips.Add file.path
    Next file
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        SearchZipFilesRecursive subFolder, zips
    Next subFolder
End Sub

Private Function UnzipOneZipRobust(ByVal zipPath As String, ByVal extractRoot As String) As Boolean
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(zipPath) Then
        UnzipOneZipRobust = False
        Exit Function
    End If
    Dim zipBaseName As String
    zipBaseName = CleanFileName(GetFileBaseName(zipPath))
    Dim finalDest As String
    finalDest = extractRoot & "\" & zipBaseName
    EnsureFolderDeep finalDest
    Dim tempRoot As String
    tempRoot = Environ$("TEMP") & "\CMS_ZIP_" & Format(Now, "yyyymmdd_hhnnss") & "_" & zipBaseName
    Dim tempOut As String
    tempOut = tempRoot & "\OUT"
    EnsureFolderDeep tempRoot
    EnsureFolderDeep tempOut
    Dim localZip As String
    localZip = tempRoot & "\" & zipBaseName & ".zip"
    FileCopy zipPath, localZip
    Dim ok As Boolean
    ok = ExtractZipUsingPowerShell(localZip, tempOut)
    If Not ok Then ok = ExtractZipUsingShell(localZip, tempOut)
    If Not ok Then
        UnzipOneZipRobust = False
        Exit Function
    End If
    CopyFolderContents tempOut, finalDest
    On Error Resume Next
    fso.DeleteFolder tempRoot, True
    On Error GoTo 0
    UnzipOneZipRobust = True
    Exit Function
ErrHandler:
    LogLine "UnzipOneZipRobust error: " & Err.Description
    UnzipOneZipRobust = False
End Function

Private Function ExtractZipUsingPowerShell(ByVal zipPath As String, ByVal destFolder As String) As Boolean
On Error GoTo ErrHandler
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & _
          "Expand-Archive -LiteralPath " & PowerShellQuote(zipPath) & _
          " -DestinationPath " & PowerShellQuote(destFolder) & " -Force" & Chr(34)
    ExtractZipUsingPowerShell = (sh.Run(cmd, 0, True) = 0)
    Exit Function
ErrHandler:
    ExtractZipUsingPowerShell = False
End Function

Private Function ExtractZipUsingShell(ByVal zipPath As String, ByVal destFolder As String) As Boolean
On Error GoTo ErrHandler
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    Dim z As Object
    Dim d As Object
    Set z = shellApp.NameSpace(zipPath)
    Set d = shellApp.NameSpace(destFolder)
    If z Is Nothing Or d Is Nothing Then
        ExtractZipUsingShell = False
        Exit Function
    End If
    d.CopyHere z.Items, 16 + 4
    WaitMilliseconds 8000
    ExtractZipUsingShell = True
    Exit Function
ErrHandler:
    ExtractZipUsingShell = False
End Function

Private Sub FlattenExtractedZipContentsIntoJobFolder(ByVal jobFolder As String, ByVal extractRoot As String)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If jobFolder = "" Or extractRoot = "" Then Exit Sub
    If fso.FolderExists(jobFolder) = False Then Exit Sub
    If fso.FolderExists(extractRoot) = False Then Exit Sub
    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(extractRoot)
    Dim subFolder As Object
    Dim file As Object
    For Each file In rootFolder.Files
        fso.CopyFile file.path, jobFolder & "\" & file.Name, True
    Next file
    For Each subFolder In rootFolder.SubFolders
        CopyExtractedFolderContentsToMain subFolder.path, jobFolder
    Next subFolder
    Exit Sub
ErrHandler:
    LogLine "FlattenExtractedZipContentsIntoJobFolder error: " & Err.Description
End Sub

Private Sub CopyExtractedFolderContentsToMain(ByVal sourceFolder As String, ByVal mainJobFolder As String)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(sourceFolder) = False Then Exit Sub
    If fso.FolderExists(mainJobFolder) = False Then Exit Sub
    Dim src As Object
    Set src = fso.GetFolder(sourceFolder)
    Dim file As Object
    Dim subFolder As Object
    For Each file In src.Files
        fso.CopyFile file.path, mainJobFolder & "\" & file.Name, True
    Next file
    For Each subFolder In src.SubFolders
        EnsureFolderDeep mainJobFolder & "\" & subFolder.Name
        CopyFolderContents subFolder.path, mainJobFolder & "\" & subFolder.Name
    Next subFolder
    Exit Sub
ErrHandler:
    LogLine "CopyExtractedFolderContentsToMain error: " & Err.Description
End Sub

Private Sub CopyFolderContents(ByVal sourceFolder As String, ByVal destFolder As String)
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sourceFolder) Then Exit Sub
    EnsureFolderDeep destFolder
    Dim src As Object
    Set src = fso.GetFolder(sourceFolder)
    Dim file As Object
    For Each file In src.Files
        fso.CopyFile file.path, destFolder & "\" & file.Name, True
    Next file
    Dim subFolder As Object
    For Each subFolder In src.SubFolders
        EnsureFolderDeep destFolder & "\" & subFolder.Name
        CopyFolderContents subFolder.path, destFolder & "\" & subFolder.Name
    Next subFolder
End Sub

Private Sub SyncCompletedJobToNetworkFolder()
On Error GoTo eh
    If Not SYNC_COMPLETED_JOB_TO_NETWORK Then Exit Sub
    If CurrentJobFolder = "" Or NetworkJobFolder = "" Then Exit Sub
    If LCase$(CurrentJobFolder) = LCase$(NetworkJobFolder) Then Exit Sub
    EnsureFolderDeep NetworkJobFolder
    CopyFolderContents CurrentJobFolder, NetworkJobFolder
    LogLine "Synced completed job package to network folder: " & NetworkJobFolder
    Exit Sub
eh:
    LogLine "SyncCompletedJobToNetworkFolder error: " & Err.Description
End Sub

' ============================================================
' FIND JOB FOLDER (by C number)
' ============================================================
Private Function FindJobFolderByText(ByVal rootPath As String, ByVal searchText As String) As String
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootPath) Then
        LogLine "Root path does not exist: " & rootPath
        Exit Function
    End If
    Dim root As Object
    Set root = fso.GetFolder(rootPath)
    Dim wantUpper As String
    wantUpper = UCase(Trim(searchText))
    Dim bestPath As String
    Dim bestScore As Long
    bestPath = ""
    bestScore = -1
    Dim subFolder As Object
    Dim nameUpper As String
    Dim score As Long
    For Each subFolder In root.SubFolders
        nameUpper = UCase(subFolder.Name)
        If nameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextTop
        score = -1
        If nameUpper = wantUpper Then
            score = 1000
        ElseIf InStr(nameUpper, wantUpper) > 0 Then
            score = 500 - Abs(Len(nameUpper) - Len(wantUpper))
        End If
        ' Prefer a folder that ALSO contains the current customer job number,
        ' so a stale folder sharing the C-number doesn't win.
        If score > -1 And CustomerJobNumber <> "" Then
            If InStr(nameUpper, UCase(CustomerJobNumber)) > 0 Then score = score + 60
        End If
        If score > bestScore Then
            bestScore = score
            bestPath = subFolder.path
        End If
NextTop:
    Next subFolder
    If bestPath <> "" Then
        FindJobFolderByText = bestPath
        Exit Function
    End If
    For Each subFolder In root.SubFolders
        nameUpper = UCase(subFolder.Name)
        If nameUpper <> UCase(EXTRACT_FOLDER_NAME) Then
            SearchJobFolderRecursive subFolder, wantUpper, 1, bestPath, bestScore
        End If
    Next subFolder
    If bestPath = "" Then
        If InStr(UCase(root.Name), wantUpper) > 0 Then bestPath = rootPath
    End If
    FindJobFolderByText = bestPath
    Exit Function
ErrHandler:
    LogLine "FindJobFolderByText error: " & Err.Description
    FindJobFolderByText = ""
End Function

Private Sub SearchJobFolderRecursive(ByVal folder As Object, ByVal wantUpper As String, _
                                     ByVal depth As Long, ByRef bestPath As String, ByRef bestScore As Long)
On Error Resume Next
    If folder Is Nothing Then Exit Sub
    If depth > 3 Then Exit Sub
    Dim subFolder As Object
    Dim nameUpper As String
    Dim score As Long
    For Each subFolder In folder.SubFolders
        nameUpper = UCase(subFolder.Name)
        If nameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextSub
        score = -1
        If nameUpper = wantUpper Then
            score = 1000 - depth
        ElseIf InStr(nameUpper, wantUpper) > 0 Then
            score = 500 - (depth * 20) - Abs(Len(nameUpper) - Len(wantUpper))
        End If
        If score > bestScore Then
            bestScore = score
            bestPath = subFolder.path
        End If
        SearchJobFolderRecursive subFolder, wantUpper, depth + 1, bestPath, bestScore
NextSub:
    Next subFolder
End Sub

' ============================================================
' FIND / OPEN CAD
' ============================================================
Private Function FindAllCadModelsRanked(ByVal folderPath As String) As Collection
On Error GoTo ErrHandler
    Dim result As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Set FindAllCadModelsRanked = result
        Exit Function
    End If
    Dim paths As Collection
    Dim scores As Collection
    Set paths = New Collection
    Set scores = New Collection
    CollectCadModelsInFolder fso.GetFolder(folderPath), paths, scores
    Dim used() As Boolean
    If paths.Count > 0 Then
        ReDim used(1 To paths.Count)
        Dim k As Long, i As Long, bestI As Long
        Dim bestS As Long
        For k = 1 To paths.Count
            bestI = 0: bestS = -2147483647
            For i = 1 To paths.Count
                If used(i) = False Then
                    If CLng(scores(i)) > bestS Then bestS = CLng(scores(i)): bestI = i
                End If
            Next i
            If bestI > 0 Then
                used(bestI) = True
                result.Add CStr(paths(bestI))
            End If
        Next k
    End If
    Set FindAllCadModelsRanked = result
    Exit Function
ErrHandler:
    LogLine "FindAllCadModelsRanked error: " & Err.Description
    Set FindAllCadModelsRanked = New Collection
End Function

Private Sub CollectCadModelsInFolder(ByVal folder As Object, ByVal paths As Collection, ByVal scores As Collection)
On Error Resume Next

    Dim folderName As String
    folderName = UCase(folder.Name)

    ' Never use generated output folders as source CAD.
    If folderName = "BASE" Then Exit Sub
    If InStr(folderName, " BASE") > 0 Then Exit Sub
    If InStr(folderName, " PRINT") > 0 Then Exit Sub
    If folderName = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Dim ext As String
    Dim score As Long
    For Each file In folder.Files

        If Left(file.Name, 2) = "~$" Then GoTo NextCadFile

        ext = LCase(fso.GetExtensionName(file.path))
        score = CadFilePriority(ext, file.Name)

        If score > 0 Then
            paths.Add file.path
            scores.Add score
        End If

NextCadFile:
    Next file
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        CollectCadModelsInFolder subFolder, paths, scores
    Next subFolder
End Sub

Private Sub AppendCadCandidates(ByVal target As Collection, ByVal extra As Collection)
On Error Resume Next
    If target Is Nothing Or extra Is Nothing Then Exit Sub
    Dim i As Long
    Dim p As String
    Dim j As Long
    Dim dup As Boolean
    For i = 1 To extra.Count
        p = CStr(extra(i))
        dup = False
        For j = 1 To target.Count
            If LCase(CStr(target(j))) = LCase(p) Then dup = True: Exit For
        Next j
        If dup = False Then target.Add p
    Next i
End Sub

Private Function CadFilePriority(ByVal ext As String, ByVal fileName As String) As Long
    Dim nameUpper As String
    nameUpper = UCase(fileName)
    If Left(fileName, 2) = "~$" Then CadFilePriority = 0: Exit Function
    If IsForeignJobCadName(nameUpper) Then CadFilePriority = 0: Exit Function
    Dim bonus As Long
    bonus = 0
    ' Strongly prefer the assembly that belongs to THIS C-number / job folder.
    ' Example: 863700126-C18614.sldasm beats 863700102_RFQ_MB_ASM_....sldasm
    If CurrentJobNumber <> "" Then
        If InStr(nameUpper, UCase(CurrentJobNumber)) > 0 Then bonus = bonus + 500
    End If
    If CustomerJobNumber <> "" Then
        If InStr(nameUpper, UCase(CustomerJobNumber)) > 0 Then bonus = bonus + 500
    End If
    If gExactJobFolderName <> "" Then
        If InStr(nameUpper, UCase(CleanFileName(gExactJobFolderName))) > 0 Then bonus = bonus + 200
    End If
    If JobBaseName <> "" Then
        If InStr(nameUpper, UCase(JobBaseName)) > 0 Then bonus = bonus + 100
    End If
    ' Prefer mold-base naming, but do NOT boost every *.sldasm via "ASM" substring
    ' (that incorrectly preferred RFQ_MB_ASM over the C-number assembly).
    If InStr(nameUpper, "MOLDBASE") > 0 Or InStr(nameUpper, "MOLD_BASE") > 0 Or InStr(nameUpper, "MOLD BASE") > 0 Then
        bonus = bonus + 30
    ElseIf InStr(nameUpper, "BASE") > 0 And InStr(nameUpper, "DATABASE") = 0 Then
        bonus = bonus + 10
    End If
    ' Deprioritize obvious leftovers / other jobs
    If InStr(nameUpper, "RFQ") > 0 And bonus < 400 Then bonus = bonus - 40
    If InStr(nameUpper, "_EXTRACT") > 0 Or InStr(nameUpper, "OLD_") > 0 Then bonus = bonus - 80
    If bonus = 0 And (InStr(nameUpper, "MOLD_BASE") > 0 Or InStr(nameUpper, "OUTSOURCE") > 0) Then
        bonus = bonus - 20
    End If
    Select Case ext
        ' Prefer customer XT/STEP over exported SLDASM when quoting from staged local files.
        Case "x_t", "x_b": CadFilePriority = 120 + bonus
        Case "step", "stp": CadFilePriority = 110 + bonus
        Case "sldasm": CadFilePriority = 100 + bonus
        Case "easm": CadFilePriority = 90 + bonus
        Case "asm": CadFilePriority = 85 + bonus
        Case "igs", "iges": CadFilePriority = 80 + bonus
        Case "sldprt": CadFilePriority = 50 + bonus
        Case "prt": CadFilePriority = 45 + bonus
        Case Else: CadFilePriority = 0
    End Select
    If CadFilePriority < 0 Then CadFilePriority = 0
End Function

' Reject CAD whose name carries a different C##### or BMS job id than this quote.
Private Function IsForeignJobCadName(ByVal pathOrName As String) As Boolean
    IsForeignJobCadName = False
    Dim u As String
    u = UCase$(Trim$(pathOrName))
    If u = "" Then Exit Function

    Dim wantC As String, wantJob As String, digits As String
    Dim i As Long, ch As String, p As Long, tok As String
    Dim foundOtherC As Boolean, foundWantC As Boolean
    Dim foundOtherJob As Boolean, foundWantJob As Boolean
    Dim atC As Boolean

    wantC = UCase$(Trim$(CurrentJobNumber))
    wantJob = ""
    digits = ""
    For i = 1 To Len(CustomerJobNumber)
        ch = Mid$(CustomerJobNumber, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i
    wantJob = digits

    foundOtherC = False: foundWantC = False
    foundOtherJob = False: foundWantJob = False

    p = 1
    Do While p <= Len(u)
        atC = False
        If Mid$(u, p, 1) = "C" And p < Len(u) Then
            If Mid$(u, p + 1, 1) >= "0" And Mid$(u, p + 1, 1) <= "9" Then
                If p = 1 Or Not ((Mid$(u, p - 1, 1) >= "A" And Mid$(u, p - 1, 1) <= "Z") Or (Mid$(u, p - 1, 1) >= "0" And Mid$(u, p - 1, 1) <= "9")) Then
                    digits = ""
                    i = p + 1
                    Do While i <= Len(u)
                        ch = Mid$(u, i, 1)
                        If ch >= "0" And ch <= "9" Then
                            digits = digits & ch
                        Else
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                    If Len(digits) >= 4 And Len(digits) <= 6 Then
                        tok = "C" & digits
                        If wantC <> "" And tok = wantC Then
                            foundWantC = True
                        ElseIf wantC <> "" Then
                            foundOtherC = True
                        End If
                        p = i
                        atC = True
                    End If
                End If
            End If
        End If
        If Not atC Then p = p + 1
    Loop

    digits = ""
    For i = 1 To Len(u) + 1
        If i <= Len(u) Then ch = Mid$(u, i, 1) Else ch = ""
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            If Len(digits) >= 8 Then
                If wantJob <> "" And digits = wantJob Then
                    foundWantJob = True
                ElseIf wantJob <> "" Then
                    foundOtherJob = True
                End If
            End If
            digits = ""
        End If
    Next i

    If foundOtherC And Not foundWantC Then IsForeignJobCadName = True: Exit Function
    If foundOtherJob And Not foundWantJob Then IsForeignJobCadName = True: Exit Function
End Function

Private Function IsGeneratedBaseCadPath(ByVal cadPath As String) As Boolean
    IsGeneratedBaseCadPath = False
    Dim u As String
    u = UCase$(Trim$(cadPath))
    If u = "" Then Exit Function
    If InStr(u, "\BASE\") > 0 Then IsGeneratedBaseCadPath = True: Exit Function
    If InStr(u, "/BASE/") > 0 Then IsGeneratedBaseCadPath = True: Exit Function
End Function

Private Function OpenCadFile(ByVal cadPath As String) As Object
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(cadPath) Then
        Set OpenCadFile = Nothing
        Exit Function
    End If
    If IsGeneratedBaseCadPath(cadPath) Then
        LogLine "OpenCadFile refused generated \base\ output: " & cadPath
        Set OpenCadFile = Nothing
        Exit Function
    End If
    Dim ext As String
    ext = LCase(fso.GetExtensionName(cadPath))
    Dim errs As Long
    Dim warns As Long
    Dim importErrors As Long
    Dim m As Object
    Set m = Nothing
    If ext = "sldasm" Then
        Set m = swApp.OpenDoc6(cadPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    ElseIf ext = "sldprt" Then
        Set m = swApp.OpenDoc6(cadPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    ElseIf ext = "slddrw" Then
        Set m = Nothing
    Else
        Set m = swApp.LoadFile4(cadPath, "r", Nothing, importErrors)
        If m Is Nothing Then Set m = swApp.LoadFile4(cadPath, "", Nothing, importErrors)
        If m Is Nothing Then Set m = swApp.OpenDoc6(cadPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
        If m Is Nothing Then Set m = swApp.OpenDoc6(cadPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    End If
    Set OpenCadFile = m
    Exit Function
ErrHandler:
    LogLine "OpenCadFile error: " & Err.Description
    Set OpenCadFile = Nothing
End Function

' ============================================================
' GRAPHICS / ORIENTATION
' ============================================================
Private Sub DisableMainViewportGraphics()
On Error Resume Next
    If DISABLE_MAIN_VIEWPORT_GRAPHICS = False Then Exit Sub
    If swModel Is Nothing Then Exit Sub
    Dim swView As Object
    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = False
        MainViewportGraphicsDisabled = True
    End If
End Sub

Private Sub RestoreMainViewportGraphics()
On Error Resume Next
    If swModel Is Nothing Then Exit Sub
    Dim swView As Object
    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    MainViewportGraphicsDisabled = False
End Sub

' Orient so the component at topIdx ends up UP and botIdx DOWN, using their
' assembly centers. Returns False if either index is invalid or lacks a center.
Private Function OrientFromPairIndices(ByVal model As Object, ByVal topIdx As Long, _
                                       ByVal botIdx As Long, ByVal label As String) As Boolean
    OrientFromPairIndices = False
    If topIdx <= 0 Or botIdx <= 0 Then Exit Function
    If topIdx > PartCount Or botIdx > PartCount Then Exit Function
    If topIdx = botIdx Then Exit Function
    If Not parts(topIdx).hasAsmCenter Or Not parts(botIdx).hasAsmCenter Then Exit Function
    OrientFromPairIndices = OrientTcpTopFromCenters(model, _
        parts(topIdx).AsmCenterX, parts(topIdx).AsmCenterY, parts(topIdx).AsmCenterZ, _
        parts(botIdx).AsmCenterX, parts(botIdx).AsmCenterY, parts(botIdx).AsmCenterZ, _
        label & " (top=" & parts(topIdx).componentName & ", bottom=" & parts(botIdx).componentName & ")")
End Function

Private Function TryShowTcpTopViewFromComponentCenters(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    TryShowTcpTopViewFromComponentCenters = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    If TryOrientTcpUpByViewProjection(model) Then
        TryShowTcpTopViewFromComponentCenters = True
        Exit Function
    End If

    LogLine "TCP-top auto orientation: component-name method unavailable."
    Exit Function

ErrHandler:
    LogLine "TryShowTcpTopViewFromComponentCenters error: " & Err.Description
    TryShowTcpTopViewFromComponentCenters = False
End Function

Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    TryOrientTcpUpByViewProjection = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim tcpComp As Object
    Dim bcpComp As Object

    Set tcpComp = FindComponentByKeys(model, TCP_TOP_ORIENTATION_KEYS)
    Set bcpComp = FindComponentByKeys(model, BCP_BOTTOM_ORIENTATION_KEYS)

    If tcpComp Is Nothing Or bcpComp Is Nothing Then
        LogLine "World-axis orientation skipped: TCP or BCP component not found."
        Exit Function
    End If

    Dim tcpX As Double, tcpY As Double, tcpZ As Double
    Dim bcpX As Double, bcpY As Double, bcpZ As Double

    If TryGetComponentCenterInches(tcpComp, tcpX, tcpY, tcpZ) = False Then
        LogLine "World-axis orientation skipped: could not read TCP center."
        Exit Function
    End If

    If TryGetComponentCenterInches(bcpComp, bcpX, bcpY, bcpZ) = False Then
        LogLine "World-axis orientation skipped: could not read BCP center."
        Exit Function
    End If

    TryOrientTcpUpByViewProjection = OrientTcpTopFromCenters(model, _
        tcpX, tcpY, tcpZ, bcpX, bcpY, bcpZ, _
        "World-axis raw components")

    Exit Function

ErrHandler:
    LogLine "TryOrientTcpUpByViewProjection error: " & Err.Description
    TryOrientTcpUpByViewProjection = False
End Function
Private Sub SetCmsTopOrientation(ByVal model As Object, Optional ByVal persistAsStandardTop As Boolean = False)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    If PROMPT_FOR_TOP_ORIENTATION Then

        MsgBox "Rotate model so you are looking top-down at the TCP / TOP SMED plate, then click OK.", vbInformation

    Else

        Dim autoOriented As Boolean
        autoOriented = False

        If AUTO_SELECT_TCP_TOP_ORIENTATION Then
            autoOriented = TryShowTcpTopViewFromComponentCenters(model)
        End If

        If autoOriented = False Then
            model.ShowNamedView2 CMS_BASE_TOP_VIEW_NAME, CMS_BASE_TOP_VIEW_ID

            LogLine "TCP-top auto orientation unavailable. Fallback view used: " & _
                    CMS_BASE_TOP_VIEW_NAME

            RotateViewZSteps model, CMS_TOP_ROTATE_Z_STEPS
        End If

    End If

    model.DeleteNamedView CMS_TOP_VIEW_NAME
    model.NameView CMS_TOP_VIEW_NAME

    Dim persisted As Boolean
    persisted = False

    If persistAsStandardTop Then

        persisted = PersistCurrentViewAsStandardTop(model)

        If persisted Then
            model.ShowNamedView2 "*Top", 5
            model.DeleteNamedView CMS_TOP_VIEW_NAME
            model.NameView CMS_TOP_VIEW_NAME
            LogLine "CMS_TOP rebuilt from persisted SolidWorks *Top view."
        Else
            model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
            LogLine "WARNING: Could not persist standard views; CMS_TOP named view was still saved."
        End If

    End If

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    StabilizeActiveView model, 300
End Sub

Private Function FindCadIndexFromExportQuote(ByVal quoteName As String) As Long
On Error GoTo ErrHandler

    FindCadIndexFromExportQuote = 0

    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "" Then Exit Function

    FindCadIndexFromExportQuote = FindCadIndexFromExportQuoteExactKey(k)
    If FindCadIndexFromExportQuote > 0 Then Exit Function

    Select Case k
        Case "IDPOTBLOCK"
            FindCadIndexFromExportQuote = FindCadIndexFromExportQuoteExactKey("IDPOT")
        Case "IDPOT"
            FindCadIndexFromExportQuote = FindCadIndexFromExportQuoteExactKey("IDPOTBLOCK")
        Case "ODPOTBLOCK"
            FindCadIndexFromExportQuote = FindCadIndexFromExportQuoteExactKey("ODPOT")
        Case "ODPOT"
            FindCadIndexFromExportQuote = FindCadIndexFromExportQuoteExactKey("ODPOTBLOCK")
    End Select

    Exit Function

ErrHandler:
    FindCadIndexFromExportQuote = 0
End Function

Private Function FindCadIndexFromExportQuoteExactKey(ByVal key As String) As Long
On Error GoTo ErrHandler

    FindCadIndexFromExportQuoteExactKey = 0

    If key = "" Then Exit Function

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = key Then
            If ExportRows(i).HasCad Then
                FindCadIndexFromExportQuoteExactKey = ExportRows(i).CadPartIndex
                Exit Function
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    FindCadIndexFromExportQuoteExactKey = 0
End Function

Private Function TryOrientFromMatchedQuotePair(ByVal model As Object, _
                                               ByVal topQuoteName As String, _
                                               ByVal bottomQuoteName As String, _
                                               ByVal sourceLabel As String) As Boolean
On Error GoTo ErrHandler

    TryOrientFromMatchedQuotePair = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim topIdx As Long
    Dim botIdx As Long

    topIdx = FindCadIndexFromExportQuote(topQuoteName)
    botIdx = FindCadIndexFromExportQuote(bottomQuoteName)

    If topIdx <= 0 Or botIdx <= 0 Then
        LogLine sourceLabel & " orientation skipped: missing matched quote pair. " & _
                "TopQuote='" & topQuoteName & "' idx=" & CStr(topIdx) & _
                ", BottomQuote='" & bottomQuoteName & "' idx=" & CStr(botIdx)
        Exit Function
    End If

    If topIdx > PartCount Or botIdx > PartCount Then Exit Function

    If parts(topIdx).hasAsmCenter = False Or parts(botIdx).hasAsmCenter = False Then
        LogLine sourceLabel & " orientation skipped: assembly centers unavailable. " & _
                "TopQuote='" & topQuoteName & "', BottomQuote='" & bottomQuoteName & "'"
        Exit Function
    End If

    LogLine sourceLabel & " orientation source:"
    LogLine "  TOP SIDE quote '" & topQuoteName & "' -> CAD '" & _
            parts(topIdx).componentName & "' center X/Y/Z=" & _
            FormatNumberForCsv(parts(topIdx).AsmCenterX) & "/" & _
            FormatNumberForCsv(parts(topIdx).AsmCenterY) & "/" & _
            FormatNumberForCsv(parts(topIdx).AsmCenterZ)

    LogLine "  BOTTOM SIDE quote '" & bottomQuoteName & "' -> CAD '" & _
            parts(botIdx).componentName & "' center X/Y/Z=" & _
            FormatNumberForCsv(parts(botIdx).AsmCenterX) & "/" & _
            FormatNumberForCsv(parts(botIdx).AsmCenterY) & "/" & _
            FormatNumberForCsv(parts(botIdx).AsmCenterZ)

    TryOrientFromMatchedQuotePair = OrientTcpTopFromCenters(model, _
        parts(topIdx).AsmCenterX, parts(topIdx).AsmCenterY, parts(topIdx).AsmCenterZ, _
        parts(botIdx).AsmCenterX, parts(botIdx).AsmCenterY, parts(botIdx).AsmCenterZ, _
        sourceLabel)

    Exit Function

ErrHandler:
    LogLine "TryOrientFromMatchedQuotePair error (" & sourceLabel & "): " & Err.Description
    TryOrientFromMatchedQuotePair = False
End Function

Private Sub EnsureCmsTopOrientationFromMatchedTcpBcp(ByVal model As Object, _
                                                     Optional ByVal persistAsStandardTop As Boolean = False)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    Dim oriented As Boolean
    oriented = False

    ' GEMINI1 ORDER:
    ' More reliable than TCP/BCP when imported parts have generic names:
    ' try top-side/bottom-side holder/pot/insert pairs first.

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "ID HOLDER", _
                    "OD HOLDER", _
                    "Matched TOP/BOTTOM HOLDER")
    End If

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "ID POT BLOCK", _
                    "OD POT BLOCK", _
                    "Matched TOP/BOTTOM POT")
    End If

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "TOP INS", _
                    "BOT INS", _
                    "Matched TOP/BOTTOM INS")
    End If

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "TCP", _
                    "BCP", _
                    "Matched TCP/BCP")
    End If

    If oriented = False Then

        LogLine "Gemini1 matched-pair orientation failed. Trying geometry fallback before *Bottom fallback."

        ' Geometry fallback should prefer holder/pot pairs first.
        ' For generic XT imports, the thin clamp pair can be labeled TCP/BCP backward,
        ' but ID/OD holders and pots usually define the true top-side better.
        If oriented = False Then
            oriented = OrientFromPairIndices(model, _
                        gIdxIDH, _
                        gIdxODH, _
                        "Geometry fallback TOP/BOTTOM HOLDER")
        End If

        If oriented = False Then
            oriented = OrientFromPairIndices(model, _
                        gIdxIDP, _
                        gIdxODP, _
                        "Geometry fallback TOP/BOTTOM POT")
        End If

        If oriented = False Then
            oriented = OrientFromPairIndices(model, _
                        gIdxTCP, _
                        gIdxBCP, _
                        "Geometry fallback TCP/BCP")
        End If

    End If

    If oriented = False Then
        LogLine "Matched and geometry orientation both failed."
        LogLine "Falling back to SetCmsTopOrientation."
        SetCmsTopOrientation model, persistAsStandardTop
        Exit Sub
    End If

    ' First: save the currently matched top-side orientation as SolidWorks *Top.
    If persistAsStandardTop Then

        If PersistCurrentViewAsStandardTop(model) Then
            LogLine "Matched top-side orientation persisted as SolidWorks *Top."
        Else
            LogLine "WARNING: Matched top-side orientation could not be persisted as standard top."
        End If

    End If

    ' Second: define the correct SolidWorks *Front from holder long side + pot/holder COM.
    If AUTO_DEFINE_FRONT_FROM_HOLDER_POT_COM Then
        If DefineStandardFrontFromHolderAndPotCom(model) Then
            LogLine "Standard *Front defined from holder long side and pot/holder center of mass."
        Else
            LogLine "WARNING: Could not define standard *Front from holder/pot logic."
        End If
    End If

    ' Third: after front is corrected, rebuild CMS_TOP from the final SolidWorks *Top.
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 100

    On Error Resume Next
    model.DeleteNamedView CMS_TOP_VIEW_NAME
    Err.Clear
    model.NameView CMS_TOP_VIEW_NAME
    On Error GoTo ErrHandler

    LogLine "CMS_TOP named view saved from final corrected SolidWorks *Top."

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    StabilizeActiveView model, 100

    Exit Sub

ErrHandler:
    LogLine "EnsureCmsTopOrientationFromMatchedTcpBcp error: " & Err.Description
End Sub

Private Sub SetStandardBaseOrientation(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub

    ' Top first (stack axis already decided by ClassifyStandardBasePlates when
    ' available; otherwise keep SolidWorks *Top as CMS_TOP).
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 50

    On Error Resume Next
    model.DeleteNamedView CMS_TOP_VIEW_NAME
    Err.Clear
    model.NameView CMS_TOP_VIEW_NAME
    On Error GoTo ErrHandler
    LogLine "STANDARD mold orientation: SolidWorks *Top saved as CMS_TOP."

    ' Front for standard bases (no holders/pots): long side of footprint visible,
    ' rails spanning left-right, latch-locks / operator hardware toward front.
    If DefineStandardFrontFromRailsAndFootprint(model) Then
        LogLine "STANDARD *Front defined from rails / footprint / latch-lock cues."
        ' Rebuild CMS_TOP from the corrected *Top after front persist.
        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 50
        On Error Resume Next
        model.DeleteNamedView CMS_TOP_VIEW_NAME
        Err.Clear
        model.NameView CMS_TOP_VIEW_NAME
        On Error GoTo ErrHandler
    Else
        LogLine "STANDARD *Front: could not refine from rails/footprint; keeping SolidWorks *Front."
    End If
    Exit Sub

ErrHandler:
    LogLine "SetStandardBaseOrientation error: " & Err.Description
End Sub

Private Function IsStandardStructuralRoleKey(ByVal roleKey As String) As Boolean
    Select Case NormalizeKey(roleKey)
        Case "TOPCLAMPPLATE", "BOTTOMCLAMPPLATE", _
             "APLATE", "BPLATE", "CAVITYPLATE", "COREPLATE", _
             "SUPPORTPLATE", "STRIPPERPLATE", "MANIFOLDPLATE", _
             "SCRETAINERPLATE", "SCBACKUPPLATE", _
             "DIEPLATE", "DIEBACKUPPLATE"
            IsStandardStructuralRoleKey = True
    End Select
End Function

Private Function FindStandardStackExtremeIndex(ByVal wantTop As Boolean) As Long
On Error GoTo ErrHandler

    FindStandardStackExtremeIndex = 0

    If PartCount < 1 Then Exit Function
    If gStdStackAxis < 1 Or gStdStackAxis > 3 Then Exit Function
    If Not StdRoleArrayReady() Then Exit Function

    Dim i As Long
    Dim roleKey As String
    Dim v As Double
    Dim bestVal As Double
    Dim haveBest As Boolean
    Dim takeHigh As Boolean

    If wantTop Then
        takeHigh = gStdTopIsFirst
    Else
        takeHigh = Not gStdTopIsFirst
    End If

    For i = 1 To PartCount
        roleKey = NormalizeKey(StdCadRole(i))

        If IsStandardStructuralRoleKey(roleKey) Then
            v = PartAxisCenter(i, gStdStackAxis)

            If Not haveBest Then
                bestVal = v
                FindStandardStackExtremeIndex = i
                haveBest = True
            Else
                If takeHigh Then
                    If v > bestVal Then
                        bestVal = v
                        FindStandardStackExtremeIndex = i
                    End If
                Else
                    If v < bestVal Then
                        bestVal = v
                        FindStandardStackExtremeIndex = i
                    End If
                End If
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    FindStandardStackExtremeIndex = 0
End Function

Private Function SetStandardBaseTopFromClassifiedStack(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    SetStandardBaseTopFromClassifiedStack = False

    If model Is Nothing Then Exit Function
    If PartCount < 2 Then Exit Function

    If gStdStackAxis < 1 Or gStdStackAxis > 3 Then
        LogLine "STANDARD stack top orientation skipped: gStdStackAxis not ready."
        Exit Function
    End If

    Dim topIdx As Long
    Dim botIdx As Long

    topIdx = FindStandardStackExtremeIndex(True)
    botIdx = FindStandardStackExtremeIndex(False)

    If topIdx <= 0 Or botIdx <= 0 Or topIdx = botIdx Then
        LogLine "STANDARD stack top orientation skipped: could not find top/bottom structural plate pair."
        Exit Function
    End If

    LogLine "STANDARD stack top orientation source:"
    LogLine "  TOP idx=" & topIdx & " role=" & StdCadRole(topIdx) & _
            " comp='" & parts(topIdx).componentName & "'" & _
            " center X/Y/Z=" & FormatNumberForCsv(parts(topIdx).AsmCenterX) & "/" & _
                              FormatNumberForCsv(parts(topIdx).AsmCenterY) & "/" & _
                              FormatNumberForCsv(parts(topIdx).AsmCenterZ)

    LogLine "  BOTTOM idx=" & botIdx & " role=" & StdCadRole(botIdx) & _
            " comp='" & parts(botIdx).componentName & "'" & _
            " center X/Y/Z=" & FormatNumberForCsv(parts(botIdx).AsmCenterX) & "/" & _
                              FormatNumberForCsv(parts(botIdx).AsmCenterY) & "/" & _
                              FormatNumberForCsv(parts(botIdx).AsmCenterZ)

    If OrientFromPairIndices(model, topIdx, botIdx, "STANDARD classified stack") = False Then
        LogLine "STANDARD stack top orientation failed."
        Exit Function
    End If

    If PersistCurrentViewAsStandardTop(model) Then
        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 50

        On Error Resume Next
        model.DeleteNamedView CMS_TOP_VIEW_NAME
        Err.Clear
        model.NameView CMS_TOP_VIEW_NAME
        On Error GoTo ErrHandler

        model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
        StabilizeActiveView model, 50

        LogLine "STANDARD stack top orientation persisted as SolidWorks *Top and CMS_TOP."
        SetStandardBaseTopFromClassifiedStack = True
    Else
        LogLine "STANDARD stack top orientation: could not persist *Top."
    End If

    Exit Function

ErrHandler:
    LogLine "SetStandardBaseTopFromClassifiedStack error: " & Err.Description
    SetStandardBaseTopFromClassifiedStack = False
End Function

Private Function FindNextStdRolePart(ByVal wantedKey As String, ByRef usedPart() As Boolean) As Long
On Error GoTo ErrHandler

    FindNextStdRolePart = 0

    If wantedKey = "" Then Exit Function
    If PartCount < 1 Then Exit Function
    If Not StdRoleArrayReady() Then Exit Function

    Dim p As Long
    Dim roleKey As String

    For p = 1 To PartCount
        If p <= UBound(usedPart) Then
            If Not usedPart(p) Then
                roleKey = NormalizeKey(StdCadRole(p))

                If roleKey = wantedKey Then
                    FindNextStdRolePart = p
                    Exit Function
                End If
            End If
        End If
    Next p

    Exit Function

ErrHandler:
    FindNextStdRolePart = 0
End Function

Private Sub RefreshStandardPlateDimsFromCurrentPartRoles()
On Error GoTo ErrHandler

    If StdCount < 1 Then Exit Sub
    If PartCount < 1 Then Exit Sub
    If Not StdRoleArrayReady() Then Exit Sub

    Dim usedPart() As Boolean
    ReDim usedPart(1 To PartCount)

    Dim i As Long
    Dim p As Long
    Dim wantedKey As String
    Dim roleKey As String
    Dim foundIdx As Long

    Dim railCount As Long
    Dim railIdx As Long

    railCount = 0
    railIdx = 0

    For p = 1 To PartCount
        roleKey = NormalizeKey(StdCadRole(p))

        If roleKey = "RAILS" Then
            railCount = railCount + 1
            If railIdx = 0 Then railIdx = p
        End If
    Next p

    ' Mark exact indexed parts as used first.
    For i = 1 To StdCount
        If StdCadIndex(i) >= 1 And StdCadIndex(i) <= PartCount Then
            If NormalizeKey(stdName(i)) <> "RAILS" Then
                usedPart(StdCadIndex(i)) = True
            End If
        End If
    Next i

    For i = 1 To StdCount

        wantedKey = NormalizeKey(stdName(i))
        foundIdx = 0

        If wantedKey = "RAILS" Then

            If railIdx > 0 Then
                StdQty(i) = railCount
                StdT(i) = parts(railIdx).Thickness
                StdW(i) = parts(railIdx).Width
                StdL(i) = parts(railIdx).Length

                LogLine "STANDARD dims refreshed after CMS view frame: Rails qty=" & railCount & _
                        " T=" & StdT(i) & " W=" & StdW(i) & " L=" & StdL(i)
            End If

        Else

            If StdCadIndex(i) >= 1 And StdCadIndex(i) <= PartCount Then
                foundIdx = StdCadIndex(i)
            End If

            If foundIdx = 0 Then
                foundIdx = FindNextStdRolePart(wantedKey, usedPart)

                If foundIdx > 0 Then
                    usedPart(foundIdx) = True
                    StdCadIndex(i) = foundIdx
                End If
            End If

            If foundIdx > 0 Then
                StdT(i) = parts(foundIdx).Thickness
                StdW(i) = parts(foundIdx).Width
                StdL(i) = parts(foundIdx).Length

                LogLine "STANDARD dims refreshed after CMS view frame: " & stdName(i) & _
                        " <- exact idx " & foundIdx & _
                        " comp='" & parts(foundIdx).componentName & "'" & _
                        " T=" & StdT(i) & _
                        " W=" & StdW(i) & _
                        " L=" & StdL(i)
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "RefreshStandardPlateDimsFromCurrentPartRoles error: " & Err.Description
End Sub

Private Function StdRowVolume(ByVal rowIdx As Long) As Double
    If rowIdx < 1 Or rowIdx > StdCount Then Exit Function
    StdRowVolume = StdT(rowIdx) * StdW(rowIdx) * StdL(rowIdx)
End Function

Private Function StdRowScoreForPrimary(ByVal rowIdx As Long) As Double
    If rowIdx < 1 Or rowIdx > StdCount Then Exit Function
    StdRowScoreForPrimary = StdRowVolume(rowIdx) + (StdT(rowIdx) * 100000#)
End Function

Private Function BestStdRowByRole(ByVal roleKey As String) As Long
    Dim i As Long
    Dim k As String
    Dim bestI As Long
    Dim bestScore As Double
    Dim sc As Double

    roleKey = NormalizeKey(roleKey)
    bestI = 0
    bestScore = -1E+99

    For i = 1 To StdCount
        k = NormalizeKey(stdName(i))
        If k = roleKey Then
            sc = StdRowScoreForPrimary(i)
            If sc > bestScore Then
                bestScore = sc
                bestI = i
            End If
        End If
    Next i

    BestStdRowByRole = bestI
End Function

Private Function BestStdRailPartIndex() As Long
    Dim p As Long
    Dim bestIdx As Long
    Dim bestScore As Double
    Dim sc As Double

    bestIdx = 0
    bestScore = -1E+99

    If PartCount < 1 Then Exit Function
    If Not StdRoleArrayReady() Then Exit Function

    For p = 1 To PartCount
        If NormalizeKey(StdCadRole(p)) = "RAILS" Then
            sc = parts(p).Length * 100000# + parts(p).BBoxVolume
            If sc > bestScore Then
                bestScore = sc
                bestIdx = p
            End If
        End If
    Next p

    BestStdRailPartIndex = bestIdx
End Function

Private Sub AddStdRowToTemp(ByVal srcRow As Long, _
                            ByRef nCount As Long, _
                            ByRef nName() As String, _
                            ByRef nT() As Double, _
                            ByRef nW() As Double, _
                            ByRef nL() As Double, _
                            ByRef nQty() As Long, _
                            ByRef nGrade() As String, _
                            ByRef nQuoteRow() As Long, _
                            ByRef nCadIndex() As Long)
    If srcRow < 1 Or srcRow > StdCount Then Exit Sub

    nCount = nCount + 1
    nName(nCount) = stdName(srcRow)
    nT(nCount) = StdT(srcRow)
    nW(nCount) = StdW(srcRow)
    nL(nCount) = StdL(srcRow)
    nQty(nCount) = StdQty(srcRow)
    nGrade(nCount) = StdGrade(srcRow)
    nQuoteRow(nCount) = StdQuoteRow(srcRow)
    nCadIndex(nCount) = StdCadIndex(srcRow)
End Sub

Private Sub AddRailRowToTemp(ByRef nCount As Long, _
                             ByRef nName() As String, _
                             ByRef nT() As Double, _
                             ByRef nW() As Double, _
                             ByRef nL() As Double, _
                             ByRef nQty() As Long, _
                             ByRef nGrade() As String, _
                             ByRef nQuoteRow() As Long, _
                             ByRef nCadIndex() As Long)
    Dim railIdx As Long
    railIdx = BestStdRailPartIndex()
    If railIdx <= 0 Then Exit Sub

    nCount = nCount + 1
    nName(nCount) = "Rails"
    nT(nCount) = parts(railIdx).Thickness
    nW(nCount) = parts(railIdx).Width
    nL(nCount) = parts(railIdx).Length
    nQty(nCount) = STD_QUOTE_RAIL_QTY
    nGrade(nCount) = "A36"
    nQuoteRow(nCount) = StdQuoteRowFor("RAILS", "A36")
    nCadIndex(nCount) = railIdx

    LogLine "STANDARD primary-stack filter: Rails kept from idx " & railIdx & _
            " qty=" & STD_QUOTE_RAIL_QTY & _
            " comp='" & parts(railIdx).componentName & "'" & _
            " T=" & nT(nCount) & " W=" & nW(nCount) & " L=" & nL(nCount)
End Sub

Private Sub ApplyPrimaryPcsStandardQuoteFilter()
On Error GoTo ErrHandler

    If Not STD_QUOTE_PRIMARY_PCS_STACK_ONLY Then Exit Sub
    If StdCount < 1 Then Exit Sub

    Dim nName() As String
    Dim nT() As Double
    Dim nW() As Double
    Dim nL() As Double
    Dim nQty() As Long
    Dim nGrade() As String
    Dim nQuoteRow() As Long
    Dim nCadIndex() As Long
    Dim nCount As Long

    ReDim nName(1 To 80)
    ReDim nT(1 To 80)
    ReDim nW(1 To 80)
    ReDim nL(1 To 80)
    ReDim nQty(1 To 80)
    ReDim nGrade(1 To 80)
    ReDim nQuoteRow(1 To 80)
    ReDim nCadIndex(1 To 80)

    Dim rA As Long
    Dim rB As Long
    Dim rEj As Long
    Dim rEjBack As Long
    Dim rTopClamp As Long
    Dim rBottomClamp As Long

    rA = BestStdRowByRole("APLATE")
    rB = BestStdRowByRole("BPLATE")
    rEj = BestStdRowByRole("EJECTORPLATE")
    rEjBack = BestStdRowByRole("BOTTOMEJECTORPLATE")

    If STD_QUOTE_INCLUDE_CLAMP_PLATES Then
        rTopClamp = BestStdRowByRole("TOPCLAMPPLATE")
        rBottomClamp = BestStdRowByRole("BOTTOMCLAMPPLATE")
    End If

    If rB > 0 Then AddStdRowToTemp rB, nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex
    If rA > 0 Then AddStdRowToTemp rA, nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex

    If STD_QUOTE_INCLUDE_CLAMP_PLATES Then
        If rTopClamp > 0 Then AddStdRowToTemp rTopClamp, nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex
        If rBottomClamp > 0 Then AddStdRowToTemp rBottomClamp, nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex
    Else
        If rTopClamp > 0 Or rBottomClamp > 0 Then
            LogLine "STANDARD primary-stack filter: clamp plates intentionally skipped for PCS standard quote."
        End If
    End If

    If rEjBack > 0 Then AddStdRowToTemp rEjBack, nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex
    If rEj > 0 Then AddStdRowToTemp rEj, nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex

    AddRailRowToTemp nCount, nName, nT, nW, nL, nQty, nGrade, nQuoteRow, nCadIndex

    LogLine "STANDARD primary-stack filter: StdCount " & StdCount & " -> " & nCount

    ReDim stdName(1 To 80)
    ReDim StdT(1 To 80)
    ReDim StdW(1 To 80)
    ReDim StdL(1 To 80)
    ReDim StdQty(1 To 80)
    ReDim StdGrade(1 To 80)
    ReDim StdQuoteRow(1 To 80)
    ReDim StdCadIndex(1 To 80)

    Dim i As Long
    For i = 1 To nCount
        stdName(i) = nName(i)
        StdT(i) = nT(i)
        StdW(i) = nW(i)
        StdL(i) = nL(i)
        StdQty(i) = nQty(i)
        StdGrade(i) = nGrade(i)
        StdQuoteRow(i) = nQuoteRow(i)
        StdCadIndex(i) = nCadIndex(i)

        LogLine "  KEEP STD " & i & ": " & stdName(i) & _
                " qty=" & StdQty(i) & _
                " T=" & StdT(i) & " W=" & StdW(i) & " L=" & StdL(i) & _
                " cadIdx=" & StdCadIndex(i) & _
                IIf(StdCadIndex(i) > 0, " comp='" & parts(StdCadIndex(i)).componentName & "'", "")
    Next i

    StdCount = nCount

    Exit Sub

ErrHandler:
    LogLine "ApplyPrimaryPcsStandardQuoteFilter error: " & Err.Description
End Sub

' Standard-base front (non-BMS): shop molds face the operator with the long
' side of the base across the view and rails running left-right. Latch-lock /
' safety-strap hardware, when present, sits on the operator (front) face.
Private Function DefineStandardFrontFromRailsAndFootprint(ByVal model As Object) As Boolean
On Error GoTo ErrHandler
    DefineStandardFrontFromRailsAndFootprint = False
    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function
    If PartCount < 1 Then Exit Function

    Dim plateIdx As Long
    Dim i As Long
    Dim bestFp As Double, fp As Double
    Dim roleKey As String
    plateIdx = 0
    bestFp = 0#
    For i = 1 To PartCount
        roleKey = NormalizeKey(StdCadRole(i))
        If roleKey = "APLATE" Or roleKey = "BPLATE" Or roleKey = "TOPCLAMPPLATE" Or roleKey = "BOTTOMCLAMPPLATE" Then
            fp = parts(i).Width * parts(i).Length
            If fp > bestFp Then bestFp = fp: plateIdx = i
        End If
    Next i
    If plateIdx < 1 Then
        For i = 1 To PartCount
            fp = parts(i).Width * parts(i).Length
            If fp > bestFp Then bestFp = fp: plateIdx = i
        Next i
    End If
    If plateIdx < 1 Then
        LogLine "Standard front skipped: no footprint plate found."
        Exit Function
    End If

    Dim railIdx(1 To 20) As Long, nRail As Long
    Dim latchIdx(1 To 40) As Long, nLatch As Long
    nRail = 0: nLatch = 0
    For i = 1 To PartCount
        roleKey = NormalizeKey(StdCadRole(i))
        If roleKey = "RAILS" Or InStr(UCase(parts(i).componentName), "RAIL") > 0 Then
            If nRail < UBound(railIdx) Then nRail = nRail + 1: railIdx(nRail) = i
        End If
        If IsLatchLockName(parts(i).componentName) Then
            If nLatch < UBound(latchIdx) Then nLatch = nLatch + 1: latchIdx(nLatch) = i
        End If
    Next i

    ' Start from *Front; if the base long side is into the screen, try *Right.
    model.ShowNamedView2 "*Front", 1
    StabilizeActiveView model, 50

    Dim candidateViewName As String
    Dim oppositeViewName As String
    Dim oppositeViewId As Long
    Dim longIntoScreen As Boolean
    Dim gotLong As Boolean

    candidateViewName = "*Front"
    oppositeViewName = "*Back"
    oppositeViewId = 2

    gotLong = IsHolderLongSideIntoCurrentView(model, plateIdx, longIntoScreen)
    If gotLong And longIntoScreen Then
        LogLine "Standard front: base long side into screen from *Front -> trying *Right."
        model.ShowNamedView2 "*Right", 4
        StabilizeActiveView model, 50
        candidateViewName = "*Right"
        oppositeViewName = "*Left"
        oppositeViewId = 3
    Else
        LogLine "Standard front: keeping " & candidateViewName & " as front candidate (long side visible or untested)."
    End If

    ' Prefer a view where rails span left-right (large view-X separation).
    If nRail >= 2 Then
        If Not RailsSpanLeftRightInActiveView(model, railIdx, nRail) Then
            LogLine "Standard front: rails not spanning left-right in " & candidateViewName & " -> trying alternate face."
            If candidateViewName = "*Front" Then
                model.ShowNamedView2 "*Right", 4
                StabilizeActiveView model, 50
                candidateViewName = "*Right"
                oppositeViewName = "*Left"
                oppositeViewId = 3
            Else
                model.ShowNamedView2 "*Front", 1
                StabilizeActiveView model, 50
                candidateViewName = "*Front"
                oppositeViewName = "*Back"
                oppositeViewId = 2
            End If
        End If
    End If

    ' Latch-locks / safety straps toward the operator (front = larger view depth).
    If nLatch >= 1 Then
        Dim flipped As Boolean
        flipped = False
        If EnsureLatchLocksCloserToFrontInActiveView(model, latchIdx, nLatch, plateIdx, _
                                                     oppositeViewName, oppositeViewId, flipped) Then
            If flipped Then
                candidateViewName = oppositeViewName
                LogLine "Standard front: flipped so latch-lock/safety-strap hardware faces operator."
            End If
        End If
    End If

    If PersistCurrentViewAsStandardFront(model) Then
        model.ShowNamedView2 "*Front", 1
        StabilizeActiveView model, 50

        SaveCurrentViewAsNamed model, CMS_FRONT_VIEW_NAME
        LogLine "CMS_FRONT named view saved from corrected STANDARD *Front."

        LogLine "Standard front complete. Persisted as SolidWorks *Front (" & candidateViewName & " candidate)."
        DefineStandardFrontFromRailsAndFootprint = True
    Else
        LogLine "Standard front failed: could not persist current view as *Front."
    End If
    Exit Function
ErrHandler:
    LogLine "DefineStandardFrontFromRailsAndFootprint error: " & Err.Description
    DefineStandardFrontFromRailsAndFootprint = False
End Function

' True when the two rails are separated more in view-X than view-Y (span across front).
Private Function RailsSpanLeftRightInActiveView(ByVal model As Object, _
                                                ByRef railIdx() As Long, ByVal nRail As Long) As Boolean
On Error GoTo nope
    RailsSpanLeftRightInActiveView = False
    If model Is Nothing Or nRail < 2 Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView
    If swView Is Nothing Then Exit Function
    Dim mView As Variant
    mView = swView.Orientation3.ArrayData
    If IsEmpty(mView) Or IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    Dim i As Long
    Dim px As Double, py As Double, pz As Double
    Dim vx As Double, vy As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim got As Boolean
    minX = 1E+30: maxX = -1E+30: minY = 1E+30: maxY = -1E+30
    got = False

    For i = 1 To nRail
        If Not TryGetCadCenterPointForFrontCheck(railIdx(i), False, px, py, pz) Then GoTo nextRail
        vx = (px * CDbl(mView(0))) + (py * CDbl(mView(3))) + (pz * CDbl(mView(6)))
        vy = (px * CDbl(mView(1))) + (py * CDbl(mView(4))) + (pz * CDbl(mView(7)))
        If vx < minX Then minX = vx
        If vx > maxX Then maxX = vx
        If vy < minY Then minY = vy
        If vy > maxY Then maxY = vy
        got = True
nextRail:
    Next i
    If Not got Then Exit Function

    Dim spanX As Double, spanY As Double
    spanX = maxX - minX
    spanY = maxY - minY
    LogLine "Standard front rails span: viewX=" & FormatNumberForCsv(spanX) & " viewY=" & FormatNumberForCsv(spanY)
    ' Rails on opposite sides of the mold -> large lateral span across the front view.
    RailsSpanLeftRightInActiveView = (spanX >= spanY * 0.85 And spanX > 1#)
    Exit Function
nope:
    RailsSpanLeftRightInActiveView = False
End Function

' Flip to opposite face if latch-lock hardware is behind the mold center (not facing operator).
Private Function EnsureLatchLocksCloserToFrontInActiveView(ByVal model As Object, _
                                                           ByRef latchIdx() As Long, ByVal nLatch As Long, _
                                                           ByVal plateIdx As Long, _
                                                           ByVal oppositeViewName As String, _
                                                           ByVal oppositeViewId As Long, _
                                                           ByRef flipped As Boolean) As Boolean
On Error GoTo nope
    EnsureLatchLocksCloserToFrontInActiveView = False
    flipped = False
    If model Is Nothing Or nLatch < 1 Then Exit Function

    Dim latchAvg As Double, plateDepth As Double, d As Double
    Dim i As Long, n As Long
    Dim px As Double, py As Double, pz As Double

    latchAvg = 0#: n = 0
    For i = 1 To nLatch
        If TryGetCadCenterPointForFrontCheck(latchIdx(i), False, px, py, pz) Then
            If TryProjectPointToActiveViewDepth(model, px, py, pz, d) Then
                latchAvg = latchAvg + d
                n = n + 1
            End If
        End If
    Next i
    If n < 1 Then Exit Function
    latchAvg = latchAvg / n

    If Not TryGetCadCenterPointForFrontCheck(plateIdx, False, px, py, pz) Then Exit Function
    If Not TryProjectPointToActiveViewDepth(model, px, py, pz, plateDepth) Then Exit Function

    LogLine "Standard front latch depth: latchAvg=" & FormatNumberForCsv(latchAvg) & _
            " plate=" & FormatNumberForCsv(plateDepth) & " delta=" & FormatNumberForCsv(latchAvg - plateDepth)

    ' Larger view depth = closer to viewed/front face (same convention as pot/front).
    If latchAvg >= plateDepth Then
        EnsureLatchLocksCloserToFrontInActiveView = True
        Exit Function
    End If

    LogLine "Standard front: latch locks behind plate center -> flipping to " & oppositeViewName
    model.ShowNamedView2 oppositeViewName, oppositeViewId
    StabilizeActiveView model, 50
    flipped = True
    EnsureLatchLocksCloserToFrontInActiveView = True
    Exit Function
nope:
    EnsureLatchLocksCloserToFrontInActiveView = False
End Function
Private Sub ApplyCmsTopView(ByVal model As Object)
On Error Resume Next
    If model Is Nothing Then Exit Sub
    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
End Sub

Private Sub StabilizeActiveView(ByVal model As Object, Optional ByVal waitMs As Long = 200)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    If DISABLE_STABILIZE_DELAYS Then Exit Sub

    model.ViewZoomtofit2
    model.GraphicsRedraw2
    DoEvents

    If waitMs > 0 Then WaitMilliseconds waitMs

    model.GraphicsRedraw2
    DoEvents
End Sub

Private Sub RotateViewZSteps(ByVal model As Object, ByVal steps As Long)
On Error Resume Next
    Dim i As Long
    If model Is Nothing Then Exit Sub
    If steps < 0 Then
        For i = 1 To Abs(steps)
            model.ViewRotateminusz
        Next i
    Else
        For i = 1 To steps
            model.ViewRotateplusz
        Next i
    End If
End Sub

Private Sub SaveModelAs(ByVal model As Object, ByVal fullPath As String)
On Error GoTo ErrHandler
    Dim errs As Long
    Dim warns As Long
    LogLine "Saving: " & fullPath
    model.Extension.SaveAs3 fullPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns
    LogLine "Save done. Errors=" & errs & " Warnings=" & warns
    Exit Sub
ErrHandler:
    LogLine "SaveModelAs error: " & Err.Description
End Sub

Private Function SaveModelCopyAs(ByVal model As Object, ByVal fullPath As String) As Boolean
On Error GoTo ErrHandler
    SaveModelCopyAs = False
    If model Is Nothing Then Exit Function
    Dim errs As Long
    Dim warns As Long
    LogLine "Saving copy: " & fullPath
    model.Extension.SaveAs3 fullPath, swSaveAsCurrentVersion, _
                            (swSaveAsOptions_Silent Or swSaveAsOptions_Copy), _
                            Nothing, Nothing, errs, warns
    LogLine "Save copy done. Errors=" & errs & " Warnings=" & warns
    SaveModelCopyAs = (Dir(fullPath) <> "")
    Exit Function
ErrHandler:
    LogLine "SaveModelCopyAs error: " & Err.Description
    SaveModelCopyAs = False
End Function

' ============================================================
' BASE PACKAGE EXPORT  (whole base only)
' ============================================================
Private Sub PrepareAssemblyForFullStlExport(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    Set model = swApp.ActiveDoc

    If model Is Nothing Then Exit Sub

    model.ClearSelection2 True

    If model.GetType = swDocASSEMBLY Then

        If FAST_QUOTE_MODE Then
            LogLine "FAST STL: skipped ResolveAllLightWeight / Unsuppress-all / heavy rebuild before full STL."
            ShowAllAssemblyComponents model
        Else
            model.ResolveAllLightWeightComponents True
            UnsuppressAllAssemblyComponents model
            ShowAllAssemblyComponents model
            model.EditRebuild3
        End If

    ElseIf model.GetType = swDocPART Then

        ShowAllPartBodies model
        If Not FAST_QUOTE_MODE Then model.EditRebuild3

    End If

    ApplyCmsTopView model

    If Not FAST_QUOTE_MODE Then
        model.GraphicsRedraw2
    End If

    DoEvents
End Sub

Private Sub SaveAssemblyStlSingleFileBinary(ByVal assyModel As Object, ByVal stlPath As String)
On Error GoTo ErrHandler

    Dim priorOneFile As Boolean
    Dim oneFileSet As Boolean

    Dim priorBinary As Long
    Dim binarySet As Boolean

    oneFileSet = False
    binarySet = False

    If Not swApp Is Nothing Then

        On Error Resume Next

        priorOneFile = swApp.GetUserPreferenceToggle(swSTLComponentsIntoOneFile)
        swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, True
        oneFileSet = True

        If FORCE_BINARY_STL_EXPORT Then
            Err.Clear
            priorBinary = swApp.GetUserPreferenceIntegerValue(swSTLBinaryFormat)
            If Err.Number = 0 Then
                swApp.SetUserPreferenceIntegerValue swSTLBinaryFormat, 1
                binarySet = True
                LogLine "STL: binary STL preference forced."
            Else
                Err.Clear
                LogLine "STL: binary STL preference could not be set; continuing."
            End If
        End If

        On Error GoTo ErrHandler

    End If

    LogLine "STL: swSTLComponentsIntoOneFile=True. Exporting all visible components as one file."
    SaveStlWithMainBaseOrientation assyModel, stlPath, "FULL ASSEMBLY"

CleanExit:
    On Error Resume Next

    If oneFileSet Then swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, priorOneFile
    If binarySet Then swApp.SetUserPreferenceIntegerValue swSTLBinaryFormat, priorBinary

    Exit Sub

ErrHandler:
    LogLine "SaveAssemblyStlSingleFileBinary error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub PrepareAssemblyVisibilityFast(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    Set model = swApp.ActiveDoc

    If model Is Nothing Then Exit Sub

    swApp.UserControl = True

    Dim swView As Object
    Set swView = model.ActiveView

    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = True
    End If

    model.ClearSelection2 True

    If model.GetType = swDocASSEMBLY Then
        ' Fast: do not resolve/unsuppress all here.
        ShowAllAssemblyComponents model
    ElseIf model.GetType = swDocPART Then
        ShowAllPartBodies model
    End If

    model.GraphicsRedraw2
    DoEvents
End Sub

Private Sub PrepareModelForJpegCapture(ByVal model As Object, _
                                       Optional ByVal showEverything As Boolean = False)
On Error Resume Next

    If model Is Nothing Then Exit Sub
    If swApp Is Nothing Then Exit Sub

    ' JPG capture needs real visible graphics.
    swApp.Visible = True
    swApp.UserControl = True

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    Set model = swApp.ActiveDoc

    If model Is Nothing Then Exit Sub

    Dim swView As Object
    Set swView = model.ActiveView

    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = True
    End If

    model.FeatureManager.EnableFeatureTree = True
    model.ClearSelection2 True

    If showEverything Then
        If model.GetType = swDocASSEMBLY Then
            ShowAllAssemblyComponents model
            ShowAllBodiesInAssemblyComponents model
        ElseIf model.GetType = swDocPART Then
            ShowAllPartBodies model
        End If
    End If

    ' Shaded mode is much safer for screenshots than wireframe/hidden-line.
    Err.Clear
    model.ViewDisplayShaded
    Err.Clear

    model.ViewZoomtofit2
    model.GraphicsRedraw2
    DoEvents
    WaitMilliseconds 500

    model.ViewZoomtofit2
    model.GraphicsRedraw2
    DoEvents
    WaitMilliseconds 500
End Sub

Private Sub ShowAllBodiesInAssemblyComponents(ByVal assyModel As Object)
On Error Resume Next

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Sub
    If IsArray(vComps) = False Then Exit Sub

    Dim i As Long
    Dim comp As Object
    Dim partDoc As Object

    For i = 0 To UBound(vComps)

        Set comp = vComps(i)

        If Not comp Is Nothing Then
            If comp.IsSuppressed = False Then

                Set partDoc = comp.GetModelDoc2

                If Not partDoc Is Nothing Then
                    If partDoc.GetType = swDocPART Then
                        ShowAllPartBodies partDoc
                    End If
                End If

            End If
        End If

    Next i
End Sub

Private Sub ShowAllPartBodies(ByVal partModel As Object)
On Error Resume Next

    If partModel Is Nothing Then Exit Sub
    If partModel.GetType <> swDocPART Then Exit Sub

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then Exit Sub
    If IsArray(vBodies) = False Then Exit Sub

    Dim i As Long

    For i = 0 To UBound(vBodies)
        If Not vBodies(i) Is Nothing Then
            vBodies(i).Hide2 False
        End If
    Next i
End Sub

Private Sub ForceViewRedrawForImage(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    Dim swView As Object
    Set swView = model.ActiveView

    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True

    model.ViewZoomtofit2
    model.GraphicsRedraw2
    DoEvents
    WaitMilliseconds 250
    model.GraphicsRedraw2
    DoEvents
End Sub

Private Sub LogFileExistsAndSize(ByVal label As String, ByVal filePath As String)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If filePath = "" Then
        LogLine label & " missing path."
        Exit Sub
    End If

    Dim i As Long

    For i = 1 To 20
        If fso.FileExists(filePath) Then Exit For
        WaitMilliseconds 250
        DoEvents
    Next i

    If fso.FileExists(filePath) Then
        LogLine label & " OK: " & filePath & _
                " size=" & CStr(fso.GetFile(filePath).Size) & " bytes"
    Else
        LogLine "WARNING: " & label & " was not created: " & filePath
    End If
End Sub

Private Sub ExportBasePackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub
    EnsureFolderDeep outputFolder

    ApplyCmsTopView swModel

    Dim baseName As String
    baseName = JobBaseName
    If baseName = "" Then baseName = CurrentJobNumber

    Dim sldPath As String
    Dim easmPath As String
    Dim igsPath As String
    Dim xtPath As String
    Dim dxfPath As String
    Dim stlPath As String

    If swModel.GetType = swDocASSEMBLY Then
        sldPath = GetUniqueFilePath(outputFolder & "\" & baseName & ".sldasm")
    Else
        sldPath = GetUniqueFilePath(outputFolder & "\" & baseName & ".sldprt")
    End If
    easmPath = GetUniqueFilePath(CurrentJobFolder & "\" & baseName & ".easm")
    igsPath = GetUniqueFilePath(CurrentJobFolder & "\" & baseName & ".igs")
    ' DXF, X_T, STL, and ISO JPGs go in the MAIN job folder (not only base\).
    xtPath = GetUniqueFilePath(CurrentJobFolder & "\" & baseName & ".x_t")
    dxfPath = GetUniqueFilePath(CurrentJobFolder & "\" & baseName & ".dxf")
    stlPath = GetUniqueFilePath(CurrentJobFolder & "\" & baseName & ".stl")
    Dim stlBasePath As String
    stlBasePath = GetUniqueFilePath(outputFolder & "\" & baseName & ".stl")

    ' Native SolidWorks copy first (DXF needs the .sldasm path).
    SaveModelAs swModel, sldPath
    SaveModelAs swModel, xtPath

    ' ============================================================
    ' STL FIRST:
    ' Export ONE merged STL containing only quoted/steel components.
    ' This avoids SolidWorks component STL shards in the job folder.
    ' ============================================================
    If CREATE_FULL_ASSEMBLY_STL Then

        Dim stlCreated As Boolean
        stlCreated = False

        If swModel.GetType = swDocASSEMBLY And STL_EXPORT_STEEL_COMPONENTS_ONLY Then

            LogStart "Export merged steel-components STL"

            stlCreated = ExportSteelComponentsMergedStlOnly(swModel, stlPath)

            If stlCreated Then
                LogLine "Merged steel-components STL written: " & stlPath
                LogFileExistsAndSize "MERGED STEEL COMPONENTS STL", stlPath
            Else
                LogLine "WARNING: merged steel-components STL was not created."
            End If

            LogDone "Export merged steel-components STL"

        Else

            LogStart "Export base STL"

            PrepareAssemblyForFullStlExport swModel

            If swModel.GetType = swDocASSEMBLY Then

                If PartCount > STL_MERGE_MAX_PARTS Then
                    LogLine "FAST STL: PartCount=" & PartCount & _
                            " > STL_MERGE_MAX_PARTS=" & STL_MERGE_MAX_PARTS & _
                            ". Skipping temp-part merge; exporting assembly STL as one file."
                    SaveFullAssemblyStlFromAssembly swModel, stlPath
                Else
                    SaveAssemblyAsMergedPartStl swModel, stlPath
                End If

            Else
                SaveStlWithMainBaseOrientation swModel, stlPath, "PART"
            End If

            stlCreated = (Dir(stlPath) <> "")

            If stlCreated Then
                LogLine "Base STL written: " & stlPath
                LogFileExistsAndSize "BASE STL", stlPath
            Else
                LogLine "WARNING: Base STL was not created: " & stlPath
            End If

            LogDone "Export base STL"

        End If

        If stlCreated Then
            DeleteStlFilesExcept CurrentJobFolder, stlPath
        End If

        If stlCreated Then
            On Error Resume Next
            If LCase(stlPath) <> LCase(stlBasePath) Then
                Dim fsoStl As Object
                Set fsoStl = CreateObject("Scripting.FileSystemObject")
                If fsoStl.FileExists(stlPath) Then fsoStl.CopyFile stlPath, stlBasePath, True
            End If
            On Error GoTo ErrHandler
        End If

        PrepareAssemblyVisibilityFast swModel

    Else

        LogLine "FAST QUOTE: skipped STL."

    End If
    If CREATE_ISO_JPEGS Then
        If gJobIsStandardBase Then
            ExportFrontAndBackIsoJpegsFullAssembly CurrentJobFolder, baseName
            LogLine "ISO JPGs written to job folder (STANDARD full assembly — no Pyropel isolation)"
        Else
            ExportFrontAndBackIsoJpegsWithoutPyropel CurrentJobFolder, baseName
            LogLine "ISO JPGs written to job folder (BMS Pyropel hidden)"
        End If
    End If

    If EXPORT_BASE_DXF Then
        If gJobIsStandardBase Then
            LogLine "STANDARD BASE DXF: full assembly (no Pyropel isolation / no pot-block keep-list)"
            CreateProjectedDxfFromNativePath sldPath, dxfPath, "BASE", CMS_TOP_VIEW_NAME, "*Top", False, True
        Else
            CreateBaseDxfWithoutPyropel sldPath, dxfPath
        End If
        LogLine "DXF written: " & dxfPath
        LogFileExistsAndSize "BASE DXF", dxfPath
    End If

    If EXPORT_HEAVY_NEUTRALS And Not FAST_QUOTE_MODE Then
        If swModel.GetType = swDocASSEMBLY Then
            SaveModelAs swModel, easmPath
            LogLine "EASM written: " & easmPath
        End If
        SaveModelAs swModel, igsPath
        LogLine "IGS written: " & igsPath
    Else
        LogLine "FAST QUOTE: skipped EASM/IGS (heavy neutrals)"
    End If

    ' Per-plate STLs (TCP, BCP, ID/OD Holder, ID/OD Pot) — optional / slow.
    If EXPORT_PER_PLATE_STLS Then
        ExportPlateStlsForComparison CurrentJobFolder & "\stl"
    End If

    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    Exit Sub

ErrHandler:
    LogLine "ExportBasePackage error: " & Err.Description
End Sub

Private Sub CreateBaseDxfWithoutPyropel(ByVal nativeSourcePath As String, ByVal dxfPath As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    ' Standard molds: never isolate TCP/holder/pot — that path hides ~250 parts
    ' and hangs CreateProjectedDxf. Full assembly DXF only (no Pyropel on these jobs).
    If gJobIsStandardBase Then
        LogLine "CreateBaseDxfWithoutPyropel: STANDARD job — full assembly DXF (skip keep-list)"
        CreateProjectedDxfFromNativePath nativeSourcePath, dxfPath, "BASE", CMS_TOP_VIEW_NAME, "*Top", False, True
        Exit Sub
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        CreateProjectedDxfFromNativePath nativeSourcePath, dxfPath, "BASE", CMS_TOP_VIEW_NAME, "*Top", False, True
        Exit Sub
    End If

    Dim keepNames As Collection
    Set keepNames = BuildBaseDxfKeepComponentNames()

    If keepNames Is Nothing Or keepNames.Count < BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF Then
        LogLine "BASE DXF isolation skipped: BMS keep-list incomplete. " & _
                "Found " & IIf(keepNames Is Nothing, 0, keepNames.Count) & _
                " of " & BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF & _
                ". Falling back to full visible native DXF."

        PrepareAssemblyVisibilityFast swModel

        CreateProjectedDxfFromNativePath nativeSourcePath, dxfPath, "BASE", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, True
        Exit Sub
    End If

    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_BASE_DXF_NO_PYROPEL_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    Dim tempNativePath As String
    tempNativePath = tempFolder & "\" & CleanFileName(JobBaseName & "_BASE_NO_PYROPEL") & ".sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel

    LogLine "Creating BASE DXF without Pyropel. Selected base component count=" & keepNames.Count

    PrepareAssemblyVisibilityFast swModel

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "WARNING: Could not isolate base components for no-Pyropel DXF; falling back to full native DXF."
        CreateProjectedDxfFromNativePath nativeSourcePath, dxfPath, "BASE", CMS_TOP_VIEW_NAME, "*Top", False, True
        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    If SaveModelCopyAs(swModel, tempNativePath) Then
        CreateProjectedDxfFromNativePath tempNativePath, dxfPath, "BASE", CMS_TOP_VIEW_NAME, "*Top", False, True
    Else
        LogLine "WARNING: Could not save no-Pyropel temp assembly; falling back to full native DXF."
        CreateProjectedDxfFromNativePath nativeSourcePath, dxfPath, "BASE", CMS_TOP_VIEW_NAME, "*Top", False, True
    End If

CleanExit:
    On Error Resume Next
    If Not hiddenNames Is Nothing Then
        If hiddenNames.Count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    Else
        ShowAllAssemblyComponents swModel
    End If
    ApplyCmsTopView swModel

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If tempFolder <> "" Then
        If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
    End If
    Exit Sub

ErrHandler:
    LogLine "CreateBaseDxfWithoutPyropel error: " & Err.Description
    Resume CleanExit
End Sub

Private Function BuildBaseDxfKeepComponentNames() As Collection
On Error GoTo ErrHandler
    Dim keepNames As New Collection

    AddBaseDxfKeepComponentFromQuote keepNames, "TCP", gIdxTCP, "TCP", KEYS_TCP
    AddBaseDxfKeepComponentFromQuote keepNames, "BCP", gIdxBCP, "BCP", KEYS_BCP
    AddBaseDxfKeepComponentFromQuote keepNames, "ID HOLDER", gIdxIDH, "ID HOLDER", ID_HOLDER_KEYS
    AddBaseDxfKeepComponentFromQuote keepNames, "OD HOLDER", gIdxODH, "OD HOLDER", OD_HOLDER_KEYS
    AddBaseDxfKeepComponentFromQuote keepNames, "ID POT", gIdxIDP, "ID POT BLOCK", KEYS_ID_POT
    AddBaseDxfKeepComponentFromQuote keepNames, "OD POT", gIdxODP, "OD POT BLOCK", KEYS_OD_POT

    If keepNames.Count < BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF Then
        LogLine "BMS keep-list incomplete: found " & keepNames.Count & _
                " of " & BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF & _
                ". Will not isolate for ISO/DXF."
        Set BuildBaseDxfKeepComponentNames = New Collection
    Else
        Set BuildBaseDxfKeepComponentNames = keepNames
    End If

    Exit Function
ErrHandler:
    LogLine "BuildBaseDxfKeepComponentNames error: " & Err.Description
    Set BuildBaseDxfKeepComponentNames = New Collection
End Function

Private Sub AddBaseDxfKeepComponentFromQuote(ByVal keepNames As Collection, _
                                             ByVal label As String, _
                                             ByVal geometryIdx As Long, _
                                             ByVal quoteName As String, _
                                             ByVal fallbackKeys As String)
On Error GoTo ErrHandler
    If keepNames Is Nothing Then Exit Sub

    Dim cadIdx As Long
    cadIdx = FindBaseDxfCadIndex(geometryIdx, quoteName, fallbackKeys)

    If cadIdx > 0 And cadIdx <= PartCount Then
        AddUniqueComponentName keepNames, parts(cadIdx).componentName
        LogLine "BASE DXF include " & label & ": CAD '" & parts(cadIdx).componentName & "'"
    Else
        LogLine "WARNING: BASE DXF could not find no-Pyropel component for " & label
    End If
    Exit Sub
ErrHandler:
    LogLine "AddBaseDxfKeepComponentFromQuote error (" & label & "): " & Err.Description
End Sub

Private Function FindBaseDxfCadIndex(ByVal geometryIdx As Long, _
                                     ByVal quoteName As String, _
                                     ByVal fallbackKeys As String) As Long
    If geometryIdx > 0 And geometryIdx <= PartCount Then
        If Not IsPyropelPartIndex(geometryIdx) Then
            FindBaseDxfCadIndex = geometryIdx
            Exit Function
        End If
        LogLine "BASE DXF rejected Pyropel geometry match: " & parts(geometryIdx).componentName
    End If

    Dim cadIdx As Long
    cadIdx = FindCadIndexFromExportQuote(quoteName)
    If cadIdx > 0 And cadIdx <= PartCount Then
        If Not IsPyropelPartIndex(cadIdx) Then
            FindBaseDxfCadIndex = cadIdx
            Exit Function
        End If
        LogLine "BASE DXF rejected Pyropel BOM match: " & parts(cadIdx).componentName
    End If

    FindBaseDxfCadIndex = FindPartIndexByKeysForBaseDxf(fallbackKeys)
End Function

Private Function FindPartIndexByKeysForBaseDxf(ByVal pipeKeys As String) As Long
    Dim i As Long
    Dim bestIdx As Long
    Dim bestVol As Double
    bestIdx = 0
    bestVol = -1#
    For i = 1 To PartCount
        If Not IsPyropelPartIndex(i) Then
            If ContainsAnyPipeKey(parts(i).componentName, pipeKeys) Then
                If parts(i).BBoxVolume > bestVol Then
                    bestVol = parts(i).BBoxVolume
                    bestIdx = i
                End If
            End If
        End If
    Next i
    FindPartIndexByKeysForBaseDxf = bestIdx
End Function

Private Function IsPyropelPartIndex(ByVal idx As Long) As Boolean
    If idx < 1 Or idx > PartCount Then Exit Function
    Dim hay As String
    hay = UCase(parts(idx).componentName & " " & parts(idx).cleanName & " " & parts(idx).filePath)
    IsPyropelPartIndex = (InStr(hay, "PYROPEL") > 0)
End Function

Private Sub AddUniqueComponentName(ByVal names As Collection, ByVal componentName As String)
On Error Resume Next
    If names Is Nothing Then Exit Sub
    componentName = Trim(componentName)
    If componentName = "" Then Exit Sub
    Dim i As Long
    For i = 1 To names.Count
        If LCase(CStr(names(i))) = LCase(componentName) Then Exit Sub
    Next i
    names.Add componentName
End Sub

' Export one STL per quoted plate into a dedicated "stl" folder:
'   TCP, BCP, ID HOLDER, OD HOLDER, ID POT, OD POT.
Private Sub ExportPlateStlsForComparison(ByVal stlFolder As String)
On Error GoTo eh
    EnsureFolderDeep stlFolder
    Dim idx(1 To 6) As Long, lbl(1 To 6) As String, i As Long
    idx(1) = gIdxTCP: lbl(1) = "TCP"
    idx(2) = gIdxBCP: lbl(2) = "BCP"
    idx(3) = gIdxIDH: lbl(3) = "ID HOLDER"
    idx(4) = gIdxODH: lbl(4) = "OD HOLDER"
    idx(5) = gIdxIDP: lbl(5) = "ID POT"
    idx(6) = gIdxODP: lbl(6) = "OD POT"

    Dim made As Long
    made = 0
    For i = 1 To 6
        If idx(i) > 0 And idx(i) <= PartCount Then
            If ExportOnePlateStl(idx(i), stlFolder & "\" & lbl(i) & ".STL") Then made = made + 1
        Else
            LogLine "  plate STL skip: " & lbl(i) & " not classified."
        End If
    Next i
    LogLine "Plate STLs written: " & made & " of 6 -> " & stlFolder

    Dim e As Long
    If Not swModel Is Nothing Then
        swApp.ActivateDoc3 swModel.GetTitle, False, 0, e
        Set swModel = swApp.ActiveDoc
    End If
    Exit Sub
eh:
    LogLine "ExportPlateStlsForComparison error: " & Err.Description
End Sub

Private Function ExportOnePlateStl(ByVal idx As Long, ByVal stlPath As String) As Boolean
On Error GoTo eh
    ExportOnePlateStl = False
    Dim fp As String
    fp = parts(idx).filePath
    If fp = "" Then
        LogLine "  plate STL skip (no file path): " & parts(idx).componentName
        Exit Function
    End If
    If Dir(fp) = "" Then
        LogLine "  plate STL skip (file missing): " & fp
        Exit Function
    End If
    Dim errs As Long, warns As Long
    Dim pm As Object
    Set pm = swApp.OpenDoc6(fp, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    If pm Is Nothing Then
        LogLine "  plate STL open failed: " & fp
        Exit Function
    End If
    swApp.ActivateDoc3 pm.GetTitle, False, 0, errs
    SaveModelAs pm, GetUniqueFilePath(stlPath)
    swApp.CloseDoc pm.GetTitle
    ExportOnePlateStl = True
    LogLine "  plate STL: " & stlPath
    Exit Function
eh:
    LogLine "ExportOnePlateStl error: " & Err.Description
End Function

' ============================================================
' FRONT + BACK ISO JPGs — STANDARD molds (full assembly, no keep-list).
' Non-BMS jobs have no Pyropel; do not hide 250+ components.
' ============================================================
Private Sub ExportFrontAndBackIsoJpegsFullAssembly(ByVal outputFolder As String, ByVal baseName As String)
On Error GoTo ErrHandler
    If swModel Is Nothing Then Exit Sub
    If baseName = "" Then baseName = CurrentJobNumber
    EnsureFolderDeep outputFolder

    PrepareAssemblyVisibilityFast swModel

    On Error Resume Next
    swApp.Visible = True
    On Error GoTo ErrHandler

    RestoreMainViewportGraphics

    Dim isoPath As String
    Dim backIsoPath As String

    isoPath = GetUniqueFilePath(outputFolder & "\" & baseName & " ISO.jpg")
    backIsoPath = GetUniqueFilePath(outputFolder & "\" & baseName & " BACK ISO.jpg")

    ' FRONT ISO
    swModel.ShowNamedView2 "*Isometric", 7
    PrepareModelForJpegCapture swModel, True

    If SaveViewAsImage(swModel, isoPath) Then
        LogLine "Saved front ISO jpg (STANDARD full assembly): " & isoPath
    Else
        LogLine "WARNING: front ISO jpg failed: " & isoPath
    End If

    ' BACK ISO
    swModel.ShowNamedView2 "*Isometric", 7
    PrepareModelForJpegCapture swModel, True

    Dim swView As Object
    Set swView = swModel.ActiveView

    If Not swView Is Nothing Then
        swView.RotateAboutCenter 0#, PI_VALUE
    End If

    PrepareModelForJpegCapture swModel, True

    If SaveViewAsImage(swModel, backIsoPath) Then
        LogLine "Saved back ISO jpg (STANDARD full assembly): " & backIsoPath
    Else
        LogLine "WARNING: back ISO jpg failed: " & backIsoPath
    End If

    swModel.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    ApplyCmsTopView swModel
    EnsureSwHidden
    Exit Sub
ErrHandler:
    LogLine "ExportFrontAndBackIsoJpegsFullAssembly error: " & Err.Description
End Sub

' ============================================================
' FRONT + BACK ISO JPGs  (BMS only — Pyropel hidden via keep-list)
' ============================================================
Private Sub ExportFrontAndBackIsoJpegsWithoutPyropel(ByVal outputFolder As String, ByVal baseName As String)
On Error GoTo ErrHandler
    If swModel Is Nothing Then Exit Sub
    If baseName = "" Then baseName = CurrentJobNumber
    EnsureFolderDeep outputFolder

    ' Safety: standard jobs must never run the BMS keep-list hide path.
    If gJobIsStandardBase Then
        ExportFrontAndBackIsoJpegsFullAssembly outputFolder, baseName
        Exit Sub
    End If

    Dim hiddenNames As Collection
    Set hiddenNames = Nothing
    Dim keepNames As Collection
    Set keepNames = Nothing
    Dim captureShowEverything As Boolean
    captureShowEverything = False

    If swModel.GetType = swDocASSEMBLY Then

        PrepareAssemblyVisibilityFast swModel

        Set keepNames = BuildBaseDxfKeepComponentNames()

        If Not keepNames Is Nothing Then

            If keepNames.Count >= BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF And _
               BMS_ISO_DXF_HIDE_NON_BASE_WHEN_COMPLETE Then

                Set hiddenNames = New Collection

                If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) Then
                    LogLine "ISO JPG: BMS base components isolated. keep=" & keepNames.Count
                    captureShowEverything = False
                Else
                    LogLine "ISO JPG: could not isolate BMS base components; capturing full visible assembly."
                    Set hiddenNames = Nothing
                    PrepareAssemblyVisibilityFast swModel
                    PrepareModelForJpegCapture swModel, True
                    captureShowEverything = True
                End If

            Else

                LogLine "ISO JPG: BMS keep-list incomplete or isolation disabled; capturing full visible assembly."
                Set hiddenNames = Nothing
                PrepareAssemblyVisibilityFast swModel
                PrepareModelForJpegCapture swModel, True
                captureShowEverything = True

            End If

        Else

            LogLine "ISO JPG: BMS keep-list unavailable; capturing full visible assembly."
            PrepareAssemblyVisibilityFast swModel
            PrepareModelForJpegCapture swModel, True
            captureShowEverything = True

        End If

    End If

    On Error Resume Next
    swApp.Visible = True
    On Error GoTo ErrHandler

    RestoreMainViewportGraphics

    Dim isoPath As String
    Dim backIsoPath As String

    isoPath = GetUniqueFilePath(outputFolder & "\" & baseName & " ISO.jpg")
    backIsoPath = GetUniqueFilePath(outputFolder & "\" & baseName & " BACK ISO.jpg")

    ' FRONT ISO
    swModel.ShowNamedView2 "*Isometric", 7
    PrepareModelForJpegCapture swModel, captureShowEverything

    ' Make sure the FRONT ISO actually shows the pot side in front.
    EnsureBmsIsoShowsPotSideFront swModel

    If SaveViewAsImage(swModel, isoPath) Then
        LogLine "Saved front ISO jpg: " & isoPath
    Else
        LogLine "WARNING: front ISO jpg failed: " & isoPath
    End If

    ' BACK ISO
    swModel.ShowNamedView2 "*Isometric", 7
    PrepareModelForJpegCapture swModel, captureShowEverything

    ' Start from the verified front ISO, then flip to the opposite side.
    EnsureBmsIsoShowsPotSideFront swModel

    Dim swView As Object
    Set swView = swModel.ActiveView

    If Not swView Is Nothing Then
        swView.RotateAboutCenter 0#, PI_VALUE
    End If

    PrepareModelForJpegCapture swModel, captureShowEverything

    If SaveViewAsImage(swModel, backIsoPath) Then
        LogLine "Saved back ISO jpg: " & backIsoPath
    Else
        LogLine "WARNING: back ISO jpg failed: " & backIsoPath
    End If

CleanExit:
    On Error Resume Next
    If Not hiddenNames Is Nothing Then
        If hiddenNames.Count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    End If
    swModel.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    ApplyCmsTopView swModel
    EnsureSwHidden
    Exit Sub
ErrHandler:
    LogLine "ExportFrontAndBackIsoJpegsWithoutPyropel error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub EnsureBmsIsoShowsPotSideFront(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If gJobIsStandardBase Then Exit Sub

    Dim holderIndexes As Collection
    Dim potIndexes As Collection

    BuildFrontOrientIndexCollections holderIndexes, potIndexes

    If holderIndexes Is Nothing Or potIndexes Is Nothing Then Exit Sub
    If holderIndexes.Count = 0 Or potIndexes.Count = 0 Then Exit Sub

    Dim holderAvg As Double
    Dim holderMin As Double
    Dim holderMax As Double
    Dim potAvg As Double
    Dim potMin As Double
    Dim potMax As Double
    Dim delta As Double

    If TryGetPotHolderActiveViewFrontDelta(model, _
                                           holderIndexes, _
                                           potIndexes, _
                                           holderAvg, holderMin, holderMax, _
                                           potAvg, potMin, potMax, _
                                           delta) = False Then
        LogLine "BMS ISO pot-front validation skipped: could not calculate active-view depth."
        Exit Sub
    End If

    LogLine "BMS ISO pot-front validation: holderAvg=" & FormatNumberForCsv(holderAvg) & _
            " potAvg=" & FormatNumberForCsv(potAvg) & _
            " delta=" & FormatNumberForCsv(delta)

    ' Larger active-view depth means closer to camera/front.
    If delta <= POT_FRONT_DEPTH_MIN_DELTA_IN Then

        LogLine "BMS ISO view appears flipped: pots are behind holders. Rotating ISO 180 degrees."

        Dim swView As Object
        Set swView = model.ActiveView

        If Not swView Is Nothing Then
            swView.RotateAboutCenter 0#, PI_VALUE
            ForceViewRedrawForImage model
        End If

    Else
        LogLine "BMS ISO view OK: pots are closer to camera/front than holders."
    End If

    Exit Sub

ErrHandler:
    LogLine "EnsureBmsIsoShowsPotSideFront error: " & Err.Description
End Sub

' Compatibility wrapper — branches on standard vs BMS.
Private Sub ExportFrontAndBackIsoJpegs(ByVal outputFolder As String, ByVal baseName As String)
    If gJobIsStandardBase Then
        ExportFrontAndBackIsoJpegsFullAssembly outputFolder, baseName
    Else
        ExportFrontAndBackIsoJpegsWithoutPyropel outputFolder, baseName
    End If
End Sub

' ============================================================
' MERGED FULL-ASSEMBLY STL  (ported from gemini1)
' Assembly -> temp multibody part -> Combine(Add) -> one STL,
' then post-rotate mesh into corrected *Front/*Top frame.
' ============================================================
Private Sub SaveAssemblyAsMergedPartStl(ByVal assyModel As Object, ByVal stlPath As String)
On Error GoTo ErrHandler
    If assyModel Is Nothing Then Exit Sub
    If stlPath = "" Then Exit Sub

    If FAST_QUOTE_MODE And PartCount > STL_MERGE_MAX_PARTS Then
        LogLine "FAST STL: PartCount=" & PartCount & _
                " > STL_MERGE_MAX_PARTS=" & STL_MERGE_MAX_PARTS & _
                ". Skipping temp-part merge inside SaveAssemblyAsMergedPartStl."
        SaveFullAssemblyStlFromAssembly assyModel, stlPath
        Exit Sub
    End If

    On Error Resume Next
    assyModel.ClearSelection2 True

    If FAST_QUOTE_MODE Then
        LogLine "FAST STL merge prep: skipped ResolveAllLightWeight / Unsuppress-all."
        ShowAllAssemblyComponents assyModel
    Else
        assyModel.ResolveAllLightWeightComponents True
        UnsuppressAllAssemblyComponents assyModel
        ShowAllAssemblyComponents assyModel
        assyModel.EditRebuild3
    End If

    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_FULL_ASM_MERGE"
    EnsureFolderDeep tempFolder

    Dim partPath As String
    partPath = tempFolder & "\FULL_ASSEMBLY_MERGE_" & Format(Now, "yyyymmdd_hhnnss") & ".sldprt"
    partPath = GetUniqueFilePath(partPath)

    Dim asmAsPartSet As Boolean
    Dim priorAsmAsPart As Long
    asmAsPartSet = False
    If Not swApp Is Nothing Then
        priorAsmAsPart = swApp.GetUserPreferenceIntegerValue(swSaveAssemblyAsPartOptions)
        ' All components (not exterior-faces-only / not selected-only).
        swApp.SetUserPreferenceIntegerValue swSaveAssemblyAsPartOptions, swSaveAsmAsPart_AllComponents
        asmAsPartSet = True
    End If

    Dim errs As Long, warns As Long
    LogLine "FULL ASSEMBLY: saving assembly as temp multibody part (all components visible):"
    LogLine "  " & partPath
    LogLine "  CAD PartCount=" & CStr(PartCount)

    assyModel.ClearSelection2 True
    assyModel.Extension.SaveAs3 partPath, swSaveAsCurrentVersion, _
                                swSaveAsOptions_Silent + swSaveAsOptions_Copy, _
                                Nothing, Nothing, errs, warns

    If asmAsPartSet Then
        swApp.SetUserPreferenceIntegerValue swSaveAssemblyAsPartOptions, priorAsmAsPart
    End If

    If fso.FileExists(partPath) = False Then
        LogLine "FULL ASSEMBLY: temp part not created. Falling back to assembly STL."
        SaveFullAssemblyStlFromAssembly assyModel, stlPath
        Exit Sub
    End If

    Dim partModel As Object
    Set partModel = swApp.OpenDoc6(partPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    If partModel Is Nothing Then
        LogLine "FULL ASSEMBLY: could not open temp part. Falling back to assembly STL."
        SaveFullAssemblyStlFromAssembly assyModel, stlPath
        On Error Resume Next
        fso.DeleteFile partPath, True
        Exit Sub
    End If

    swApp.ActivateDoc3 partModel.GetTitle, False, 0, errs
    EnsureSwHidden

    ' Reject a merge that only captured one (or very few) bodies — that is the
    ' "STL picking up one singular part" failure mode. Fall back to assembly STL.
    Dim bodyCount As Long
    bodyCount = CountSolidBodiesInPart(partModel)
    LogLine "FULL ASSEMBLY: temp part solid body count=" & CStr(bodyCount) & _
            " (assembly PartCount=" & CStr(PartCount) & ")"

    Dim minBodies As Long
    minBodies = 2
    If PartCount >= 6 Then
        minBodies = PartCount \ 4
        If minBodies < 2 Then minBodies = 2
        If minBodies > 12 Then minBodies = 12
    End If

    If PartCount >= 3 And bodyCount < minBodies Then
        LogLine "FULL ASSEMBLY: merge captured too few bodies (" & CStr(bodyCount) & _
                " < " & CStr(minBodies) & "). Rejecting merge; exporting full assembly STL."
        On Error Resume Next
        swApp.CloseDoc partModel.GetTitle
        Set partModel = Nothing
        If fso.FileExists(partPath) Then fso.DeleteFile partPath, True
        swApp.ActivateDoc3 assyModel.GetTitle, False, 0, errs
        On Error GoTo ErrHandler
        SaveFullAssemblyStlFromAssembly assyModel, stlPath
        Exit Sub
    End If

    MergeAllPartBodies partModel
    SaveStlWithMainBaseOrientation partModel, stlPath, "FULL ASSEMBLY"

    On Error Resume Next
    swApp.CloseDoc partModel.GetTitle
    Set partModel = Nothing
    If fso.FileExists(partPath) Then
        fso.DeleteFile partPath, True
        LogLine "FULL ASSEMBLY: deleted temp merge part."
    End If
    swApp.ActivateDoc3 assyModel.GetTitle, False, 0, errs
    Set fso = Nothing
    Exit Sub

ErrHandler:
    LogLine "SaveAssemblyAsMergedPartStl error: " & Err.Description
    On Error Resume Next
    If Not partModel Is Nothing Then swApp.CloseDoc partModel.GetTitle
    Set fso = CreateObject("Scripting.FileSystemObject")
    If partPath <> "" Then
        If fso.FileExists(partPath) Then fso.DeleteFile partPath, True
    End If
    If Not assyModel Is Nothing Then swApp.ActivateDoc3 assyModel.GetTitle, False, 0, errs
    SaveFullAssemblyStlFromAssembly assyModel, stlPath
End Sub

Private Function CountSolidBodiesInPart(ByVal partModel As Object) As Long
On Error GoTo eh
    CountSolidBodiesInPart = 0
    If partModel Is Nothing Then Exit Function
    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function
    CountSolidBodiesInPart = UBound(vBodies) - LBound(vBodies) + 1
    Exit Function
eh:
    CountSolidBodiesInPart = 0
End Function


Private Function CollectStlFilesInFolder(ByVal folderPath As String) As Collection
On Error GoTo ErrHandler

    Dim result As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Set CollectStlFilesInFolder = result
        Exit Function
    End If

    Dim f As Object
    For Each f In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(f.path)) = "stl" Then
            result.Add f.path
        End If
    Next f

    Set CollectStlFilesInFolder = result
    Exit Function

ErrHandler:
    Set CollectStlFilesInFolder = New Collection
End Function

Private Function IsValidBinaryStlFile(ByVal stlPath As String, ByRef triCount As Long) As Boolean
On Error GoTo ErrHandler

    IsValidBinaryStlFile = False
    triCount = 0

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If stlPath = "" Then Exit Function
    If Not fso.FileExists(stlPath) Then Exit Function

    Dim f As Integer
    f = FreeFile

    Open stlPath For Binary Access Read As #f

    Dim hdr As BinaryStlHeader
    Get #f, 1, hdr

    triCount = hdr.TriangleCount

    If triCount <= 0 Then
        Close #f
        Exit Function
    End If

    Dim expectedLen As Double
    expectedLen = 84# + CDbl(triCount) * 50#

    If CDbl(LOF(f)) <> expectedLen Then
        Close #f
        triCount = 0
        Exit Function
    End If

    Close #f

    IsValidBinaryStlFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #f
    triCount = 0
    IsValidBinaryStlFile = False
End Function

Private Function MergeBinaryStlFilesToOne(ByVal stlFiles As Collection, ByVal outPath As String) As Boolean
On Error GoTo ErrHandler

    MergeBinaryStlFilesToOne = False

    If stlFiles Is Nothing Then Exit Function
    If stlFiles.Count < 1 Then Exit Function
    If outPath = "" Then Exit Function

    Dim validFiles As New Collection
    Dim totalTris As Long
    Dim i As Long
    Dim tc As Long
    Dim p As String

    totalTris = 0

    For i = 1 To stlFiles.Count
        p = CStr(stlFiles(i))
        If IsValidBinaryStlFile(p, tc) Then
            validFiles.Add p
            totalTris = totalTris + tc
        Else
            LogLine "STL merge: skipping non-binary/invalid STL shard: " & p
        End If
    Next i

    If validFiles.Count < 1 Or totalTris <= 0 Then
        LogLine "STL merge failed: no valid binary STL shards."
        Exit Function
    End If

    Dim outF As Integer
    outF = FreeFile

    Dim outHdr As BinaryStlHeader
    outHdr.HeaderText = "CMS MERGED STEEL COMPONENT STL"
    outHdr.TriangleCount = totalTris

    Open outPath For Binary Access Write As #outF
    Put #outF, 1, outHdr

    Dim inF As Integer
    Dim inHdr As BinaryStlHeader
    Dim tri As BinaryStlTriangle
    Dim t As Long

    For i = 1 To validFiles.Count

        p = CStr(validFiles(i))
        inF = FreeFile

        Open p For Binary Access Read As #inF

        Get #inF, 1, inHdr

        For t = 0 To inHdr.TriangleCount - 1
            Get #inF, 85 + t * 50, tri
            Put #outF, , tri
        Next t

        Close #inF
    Next i

    Close #outF

    LogLine "STL merge: merged " & validFiles.Count & " STL shard(s), triangles=" & totalTris & " -> " & outPath

    MergeBinaryStlFilesToOne = (Dir(outPath) <> "")
    Exit Function

ErrHandler:
    LogLine "MergeBinaryStlFilesToOne error: " & Err.Description
    On Error Resume Next
    Close #inF
    Close #outF
    MergeBinaryStlFilesToOne = False
End Function

Private Function ExportVisibleModelStlToTempAndMerge(ByVal model As Object, _
                                                     ByVal finalStlPath As String, _
                                                     ByVal label As String) As Boolean
On Error GoTo ErrHandler

    ExportVisibleModelStlToTempAndMerge = False

    If model Is Nothing Then Exit Function
    If finalStlPath = "" Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_STL_TEMP_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    Dim tempStlPath As String
    tempStlPath = tempFolder & "\CMS_STEEL_COMPONENTS.stl"

    Dim priorOneFile As Boolean
    Dim oneFileSet As Boolean
    Dim priorBinary As Long
    Dim binarySet As Boolean

    oneFileSet = False
    binarySet = False

    If Not swApp Is Nothing Then

        On Error Resume Next

        priorOneFile = swApp.GetUserPreferenceToggle(swSTLComponentsIntoOneFile)
        swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, True
        oneFileSet = True

        If FORCE_BINARY_STL_EXPORT Then
            Err.Clear
            priorBinary = swApp.GetUserPreferenceIntegerValue(swSTLBinaryFormat)
            If Err.Number = 0 Then
                swApp.SetUserPreferenceIntegerValue swSTLBinaryFormat, 1
                binarySet = True
                LogLine "STL temp export: binary STL preference forced."
            Else
                Err.Clear
                LogLine "STL temp export: binary preference could not be set."
            End If
        End If

        On Error GoTo ErrHandler

    End If

    Dim errs As Long
    Dim warns As Long

    LogLine "STL temp export: saving visible " & label & " to temp folder:"
    LogLine "  " & tempStlPath

    model.Extension.SaveAs3 tempStlPath, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent, _
                            Nothing, Nothing, errs, warns

    LogLine "STL temp export save done. Errors=" & errs & " Warnings=" & warns

    Dim stlFiles As Collection
    Set stlFiles = CollectStlFilesInFolder(tempFolder)

    LogLine "STL temp export produced " & stlFiles.Count & " STL file(s)."

    If stlFiles.Count = 0 Then
        LogLine "STL temp export failed: no STL files found in temp folder."
        GoTo CleanExit
    End If

    If stlFiles.Count = 1 Then

        fso.CopyFile CStr(stlFiles(1)), finalStlPath, True
        LogLine "STL temp export: single STL copied to final path."

    Else

        If MergeBinaryStlFilesToOne(stlFiles, finalStlPath) = False Then
            LogLine "STL temp export failed: could not merge component STL shards."
            GoTo CleanExit
        End If

    End If

    If MATCH_STUDIO_STL_MATCH_MAIN_BASE_ORIENTATION And POST_ROTATE_STL_TO_CORRECTED_FRONT Then

        If FinalStlCoordFrameReady Then
            If ReorientStlFileToMatrix(finalStlPath, FinalStlCoordM) Then
                LogLine "Merged STL post-rotated into corrected final coordinate frame."
            Else
                LogLine "WARNING: merged STL post-rotation failed. File still exists."
            End If
        Else
            LogLine "WARNING: final STL coordinate frame not ready; merged STL was not post-rotated."
        End If

    End If

    ExportVisibleModelStlToTempAndMerge = (Dir(finalStlPath) <> "")

CleanExit:
    On Error Resume Next

    If oneFileSet Then swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, priorOneFile
    If binarySet Then swApp.SetUserPreferenceIntegerValue swSTLBinaryFormat, priorBinary

    If tempFolder <> "" Then
        If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
    End If

    Exit Function

ErrHandler:
    LogLine "ExportVisibleModelStlToTempAndMerge error: " & Err.Description
    Resume CleanExit
End Function

Private Sub AddLargestRailComponentsToKeep(ByVal keepNames As Collection, ByVal qtyNeeded As Long)
On Error GoTo ErrHandler

    If keepNames Is Nothing Then Exit Sub
    If qtyNeeded < 1 Then Exit Sub
    If PartCount < 1 Then Exit Sub
    If Not StdRoleArrayReady() Then Exit Sub

    Dim railIdx() As Long
    Dim n As Long
    Dim i As Long

    ReDim railIdx(1 To PartCount)
    n = 0

    For i = 1 To PartCount
        If NormalizeKey(StdCadRole(i)) = "RAILS" Then
            n = n + 1
            railIdx(n) = i
        End If
    Next i

    If n < 1 Then Exit Sub

    Dim a As Long
    Dim b As Long
    Dim tmp As Long

    For a = 1 To n - 1
        For b = a + 1 To n
            If (parts(railIdx(b)).Length * 100000# + parts(railIdx(b)).BBoxVolume) > _
               (parts(railIdx(a)).Length * 100000# + parts(railIdx(a)).BBoxVolume) Then

                tmp = railIdx(a)
                railIdx(a) = railIdx(b)
                railIdx(b) = tmp
            End If
        Next b
    Next a

    Dim maxAdd As Long
    maxAdd = qtyNeeded
    If maxAdd > n Then maxAdd = n

    For i = 1 To maxAdd
        AddUniqueComponentName keepNames, parts(railIdx(i)).componentName
        LogLine "STL steel keep rail: idx=" & railIdx(i) & _
                " comp='" & parts(railIdx(i)).componentName & "'"
    Next i

    Exit Sub

ErrHandler:
    LogLine "AddLargestRailComponentsToKeep error: " & Err.Description
End Sub

Private Function BuildStandardSteelStlKeepComponentNames() As Collection
On Error GoTo ErrHandler

    Dim keepNames As New Collection

    If StdCount < 1 Then
        Set BuildStandardSteelStlKeepComponentNames = keepNames
        Exit Function
    End If

    Dim i As Long
    Dim ci As Long
    Dim k As String

    For i = 1 To StdCount

        k = NormalizeKey(stdName(i))

        If k = "RAILS" Then

            AddLargestRailComponentsToKeep keepNames, StdQty(i)

        Else

            ci = 0
            If i <= UBound(StdCadIndex) Then ci = StdCadIndex(i)

            If ci >= 1 And ci <= PartCount Then
                AddUniqueComponentName keepNames, parts(ci).componentName
                LogLine "STL steel keep standard: " & stdName(i) & _
                        " idx=" & ci & _
                        " comp='" & parts(ci).componentName & "'"
            Else
                LogLine "WARNING: STL standard steel keep skipped row " & i & _
                        " (" & stdName(i) & ") because StdCadIndex is missing."
            End If

        End If

    Next i

    Set BuildStandardSteelStlKeepComponentNames = keepNames
    Exit Function

ErrHandler:
    LogLine "BuildStandardSteelStlKeepComponentNames error: " & Err.Description
    Set BuildStandardSteelStlKeepComponentNames = New Collection
End Function

Private Function BuildSteelStlKeepComponentNamesForCurrentJob() As Collection
On Error GoTo ErrHandler

    If gJobIsStandardBase Then
        Set BuildSteelStlKeepComponentNamesForCurrentJob = BuildStandardSteelStlKeepComponentNames()
    Else
        Set BuildSteelStlKeepComponentNamesForCurrentJob = BuildBaseDxfKeepComponentNames()
    End If

    Exit Function

ErrHandler:
    LogLine "BuildSteelStlKeepComponentNamesForCurrentJob error: " & Err.Description
    Set BuildSteelStlKeepComponentNamesForCurrentJob = New Collection
End Function

Private Function ExportSteelComponentsMergedStlOnly(ByVal assyModel As Object, ByVal stlPath As String) As Boolean
On Error GoTo ErrHandler

    ExportSteelComponentsMergedStlOnly = False

    If assyModel Is Nothing Then Exit Function
    If stlPath = "" Then Exit Function

    If assyModel.GetType <> swDocASSEMBLY Then
        LogLine "Steel STL: source is not assembly; exporting current model."
        ExportSteelComponentsMergedStlOnly = ExportVisibleModelStlToTempAndMerge(assyModel, stlPath, "PART")
        Exit Function
    End If

    Dim keepNames As Collection
    Set keepNames = BuildSteelStlKeepComponentNamesForCurrentJob()

    If keepNames Is Nothing Or keepNames.Count = 0 Then
        LogLine "Steel STL export skipped: no steel components identified for keep-list."
        Exit Function
    End If

    If Not gJobIsStandardBase Then
        If keepNames.Count < BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF Then
            LogLine "BMS steel STL export skipped: keep-list incomplete. Found " & keepNames.Count
            If BMS_STL_SKIP_IF_KEEP_LIST_INCOMPLETE Then Exit Function
        End If
    End If

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    LogLine "Steel STL: isolating steel/quoted components only. keep=" & keepNames.Count

    PrepareAssemblyVisibilityFast assyModel

    If HideAllExceptComponentNamesOnce(assyModel, keepNames, hiddenNames) = False Then
        LogLine "Steel STL export failed: could not isolate steel components."
        GoTo CleanExit
    End If

    ApplyCmsTopView assyModel
    StabilizeActiveView assyModel, 50

    ExportSteelComponentsMergedStlOnly = ExportVisibleModelStlToTempAndMerge(assyModel, stlPath, "STEEL COMPONENTS")

    If ExportSteelComponentsMergedStlOnly Then
        LogLine "Steel components merged STL written: " & stlPath
    Else
        LogLine "WARNING: Steel components merged STL was not created."
    End If

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.Count > 0 Then
            ShowNamedComponentsOnce assyModel, hiddenNames
        Else
            ShowAllAssemblyComponents assyModel
        End If
    Else
        ShowAllAssemblyComponents assyModel
    End If

    ApplyCmsTopView assyModel
    Exit Function

ErrHandler:
    LogLine "ExportSteelComponentsMergedStlOnly error: " & Err.Description
    Resume CleanExit
End Function

Private Sub DeleteStlFilesExcept(ByVal folderPath As String, ByVal keepPath As String)
On Error GoTo ErrHandler

    If Not CLEAN_EXTRA_STL_SHARDS_IN_JOB_FOLDER Then Exit Sub
    If folderPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then Exit Sub

    Dim keepLower As String
    keepLower = LCase$(keepPath)

    Dim f As Object
    For Each f In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(f.path)) = "stl" Then
            If LCase$(f.path) <> keepLower Then
                LogLine "Deleting extra STL shard: " & f.path
                fso.DeleteFile f.path, True
            End If
        End If
    Next f

    Exit Sub

ErrHandler:
    LogLine "DeleteStlFilesExcept error: " & Err.Description
End Sub

Private Function ExportBmsQuotedBaseStlOnly(ByVal assyModel As Object, ByVal stlPath As String) As Boolean
On Error GoTo ErrHandler

    ExportBmsQuotedBaseStlOnly = False

    If assyModel Is Nothing Then Exit Function
    If stlPath = "" Then Exit Function

    If assyModel.GetType <> swDocASSEMBLY Then
        LogLine "BMS quoted-base STL: source is not an assembly; exporting current model STL."
        SaveStlWithMainBaseOrientation assyModel, stlPath, "BMS QUOTED BASE PART"
        ExportBmsQuotedBaseStlOnly = (Dir(stlPath) <> "")
        Exit Function
    End If

    Dim keepNames As Collection
    Set keepNames = BuildBaseDxfKeepComponentNames()

    If keepNames Is Nothing Or keepNames.Count < BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF Then

        LogLine "BMS quoted-base STL skipped: keep-list incomplete. Found " & _
                IIf(keepNames Is Nothing, 0, keepNames.Count) & _
                " of " & BMS_MIN_KEEP_COMPONENTS_FOR_ISO_DXF & " required quoted base components."

        If BMS_STL_SKIP_IF_KEEP_LIST_INCOMPLETE Then
            LogLine "BMS quoted-base STL: NOT falling back to full assembly because BMS_STL_SKIP_IF_KEEP_LIST_INCOMPLETE=True."
            Exit Function
        Else
            LogLine "BMS quoted-base STL: falling back to full assembly STL."
            SaveFullAssemblyStlFromAssembly assyModel, stlPath
            ExportBmsQuotedBaseStlOnly = (Dir(stlPath) <> "")
            Exit Function
        End If

    End If

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    LogLine "BMS quoted-base STL: isolating quoted components only. keep=" & keepNames.Count

    ' Start visible, then hide everything except TCP/BCP/holders/pots.
    PrepareAssemblyVisibilityFast assyModel

    If HideAllExceptComponentNamesOnce(assyModel, keepNames, hiddenNames) = False Then

        LogLine "BMS quoted-base STL failed: could not isolate quoted base components."

        If BMS_STL_SKIP_IF_KEEP_LIST_INCOMPLETE Then
            LogLine "BMS quoted-base STL: NOT falling back to full assembly."
            GoTo CleanExit
        Else
            LogLine "BMS quoted-base STL: falling back to full assembly STL."
            SaveFullAssemblyStlFromAssembly assyModel, stlPath
            ExportBmsQuotedBaseStlOnly = (Dir(stlPath) <> "")
            GoTo CleanExit
        End If

    End If

    ApplyCmsTopView assyModel
    StabilizeActiveView assyModel, 50

    ' Export visible quoted components as ONE STL file.
    ' This does NOT boolean-combine them; it writes one multi-shell STL file.
    SaveAssemblyStlSingleFileBinary assyModel, stlPath

    If Dir(stlPath) <> "" Then
        ExportBmsQuotedBaseStlOnly = True
        LogLine "BMS quoted-base STL written: " & stlPath
    Else
        LogLine "WARNING: BMS quoted-base STL was not created: " & stlPath
    End If

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.Count > 0 Then
            ShowNamedComponentsOnce assyModel, hiddenNames
        Else
            ShowAllAssemblyComponents assyModel
        End If
    Else
        ShowAllAssemblyComponents assyModel
    End If

    ApplyCmsTopView assyModel
    Exit Function

ErrHandler:
    LogLine "ExportBmsQuotedBaseStlOnly error: " & Err.Description
    Resume CleanExit
End Function

Private Sub SaveFullAssemblyStlFromAssembly(ByVal assyModel As Object, ByVal stlPath As String)
On Error GoTo ErrHandler

    If assyModel Is Nothing Then Exit Sub

    PrepareAssemblyForFullStlExport assyModel
    SaveAssemblyStlSingleFileBinary assyModel, stlPath

    Exit Sub

ErrHandler:
    LogLine "SaveFullAssemblyStlFromAssembly error: " & Err.Description
End Sub

Private Sub MergeAllPartBodies(ByVal partModel As Object)
On Error GoTo ErrHandler
    If partModel Is Nothing Then Exit Sub
    If partModel.GetType <> swDocPART Then Exit Sub

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then
        LogLine "FULL ASSEMBLY merge: no solid bodies found."
        Exit Sub
    End If

    Dim bodyCount As Long
    bodyCount = UBound(vBodies) - LBound(vBodies) + 1
    If bodyCount < 2 Then
        LogLine "FULL ASSEMBLY merge: single body already; no combine needed."
        Exit Sub
    End If

    partModel.ClearSelection2 True
    Dim i As Long
    For i = LBound(vBodies) To UBound(vBodies)
        If Not vBodies(i) Is Nothing Then vBodies(i).Select2 True, Nothing
    Next i

    Dim combineFeat As Object
    Set combineFeat = partModel.FeatureManager.InsertCombineFeature(SWBODYADD, Nothing, Nothing)
    partModel.ClearSelection2 True
    partModel.EditRebuild3

    If combineFeat Is Nothing Then
        LogLine "FULL ASSEMBLY merge: Combine(Add) returned nothing; bodies left separate (STL still one file)."
    Else
        LogLine "FULL ASSEMBLY merge: combined " & CStr(bodyCount) & " bodies into one."
    End If
    Exit Sub
ErrHandler:
    LogLine "MergeAllPartBodies error: " & Err.Description
    On Error Resume Next
    partModel.ClearSelection2 True
End Sub

Private Sub SaveStlWithMainBaseOrientation(ByVal model As Object, _
                                           ByVal stlPath As String, _
                                           Optional ByVal label As String = "", _
                                           Optional ByVal sourceComponentName As String = "")
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If stlPath = "" Then Exit Sub

    Dim orientM(0 To 8) As Double
    Dim gotOrient As Boolean
    Dim i As Long

    gotOrient = False

    If MATCH_STUDIO_STL_MATCH_MAIN_BASE_ORIENTATION And POST_ROTATE_STL_TO_CORRECTED_FRONT Then

        ' Use the one final coordinate frame captured after *Top and *Front were corrected.
        ' Do NOT use the component's original axes.
        ' Do NOT use the orientation of the isolated component state.
        If FinalStlCoordFrameReady Then

            For i = 0 To 8
                orientM(i) = FinalStlCoordM(i)
            Next i

            gotOrient = True

            LogLine "STL using FINAL corrected standard-view coordinate system:"
            LogLine "  Label=" & label
            LogLine "  Path=" & stlPath

        Else

            LogLine "WARNING: Final STL coordinate system was not captured before STL save."
            LogLine "Attempting to capture it now from main model/current model."

            If Not swModel Is Nothing Then
                gotOrient = CaptureFinalStandardViewsForStlCoordinateSystem(swModel)
            Else
                gotOrient = CaptureFinalStandardViewsForStlCoordinateSystem(model)
            End If

            If gotOrient Then
                For i = 0 To 8
                    orientM(i) = FinalStlCoordM(i)
                Next i
            Else
                LogLine "WARNING: Could not capture final STL coordinate system. STL may stay in original imported axes."
            End If

        End If

    End If

    ' SolidWorks STL export ignores named views and standard views.
    ' It writes mesh coordinates in model/original coordinate space.
    SaveModelAs model, stlPath

    ' Convert exported STL mesh from original model coordinates into your corrected
    ' Top/Front standard-view coordinate system.
    If gotOrient Then

        If ReorientStlFileToMatrix(stlPath, orientM) Then

            LogLine "STL post-rotated into FINAL corrected Top/Front coordinate system:"
            LogLine "  " & stlPath

        Else

            LogLine "WARNING: STL post-rotation failed:"
            LogLine "  " & stlPath

        End If

    End If

    On Error Resume Next
    ApplyCmsTopView model
    On Error GoTo 0

    Exit Sub

ErrHandler:
    LogLine "SaveStlWithMainBaseOrientation error (" & label & "): " & Err.Description

    On Error Resume Next
    SaveModelAs model, stlPath
    ApplyCmsTopView model
End Sub

Private Function CaptureFinalStandardViewsForStlCoordinateSystem(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    CaptureFinalStandardViewsForStlCoordinateSystem = False
    FinalStlCoordFrameReady = False

    If model Is Nothing Then Exit Function

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    EnsureSwHidden

    ' Make sure viewport orientation actually updates even when graphics are frozen/hidden.
    Dim swView As Object
    Set swView = model.ActiveView

    On Error Resume Next
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    On Error GoTo ErrHandler

    ' Use final corrected SolidWorks *Front.
    ' Because your macro has already redefined *Top and *Front, this view contains
    ' the full corrected coordinate frame.
    model.ShowNamedView2 "*Front", 1

    On Error Resume Next
    model.ViewZoomtofit2
    model.GraphicsRedraw2
    DoEvents
    WaitMilliseconds 50
    model.GraphicsRedraw2
    DoEvents
    On Error GoTo ErrHandler

    Set swView = model.ActiveView

    If swView Is Nothing Then
        LogLine "STL coordinate capture failed: ActiveView is Nothing."
        Exit Function
    End If

    Dim v As Variant
    v = swView.Orientation3.ArrayData

    If IsEmpty(v) Then
        LogLine "STL coordinate capture failed: Orientation3.ArrayData is empty."
        Exit Function
    End If

    If IsArray(v) = False Then
        LogLine "STL coordinate capture failed: Orientation3.ArrayData is not an array."
        Exit Function
    End If

    If UBound(v) < 8 Then
        LogLine "STL coordinate capture failed: Orientation matrix has fewer than 9 values."
        Exit Function
    End If

    Dim i As Long

    For i = 0 To 8
        FinalStlCoordM(i) = CDbl(v(i))
    Next i

    FinalStlCoordFrameReady = True
    CaptureFinalStandardViewsForStlCoordinateSystem = True

    LogLine "FINAL STL coordinate system captured from corrected SolidWorks *Front."
    LogLine "This is the only orientation matrix that will be used for Match Studio STLs."
    LogLine "  Matrix=[" & _
            FormatNumberForCsv(FinalStlCoordM(0)) & "," & _
            FormatNumberForCsv(FinalStlCoordM(1)) & "," & _
            FormatNumberForCsv(FinalStlCoordM(2)) & "; " & _
            FormatNumberForCsv(FinalStlCoordM(3)) & "," & _
            FormatNumberForCsv(FinalStlCoordM(4)) & "," & _
            FormatNumberForCsv(FinalStlCoordM(5)) & "; " & _
            FormatNumberForCsv(FinalStlCoordM(6)) & "," & _
            FormatNumberForCsv(FinalStlCoordM(7)) & "," & _
            FormatNumberForCsv(FinalStlCoordM(8)) & "]"

CleanExit:
    On Error Resume Next
    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    If Err.Number <> 0 Then
        Err.Clear
        model.ShowNamedView2 "*Top", 5
    End If
    Exit Function

ErrHandler:
    LogLine "CaptureFinalStandardViewsForStlCoordinateSystem error: " & Err.Description
    FinalStlCoordFrameReady = False
    CaptureFinalStandardViewsForStlCoordinateSystem = False
    Resume CleanExit
End Function

Private Function ReorientStlFileToMatrix(ByVal stlPath As String, _
                                         ByRef m() As Double) As Boolean
On Error GoTo ErrHandler

    ReorientStlFileToMatrix = False

    ' Try binary STL first.
    If ReorientBinaryStlFileToMatrix(stlPath, m) Then
        LogLine "STL reorient: binary STL rotated."
        ReorientStlFileToMatrix = True
        Exit Function
    End If

    ' If SolidWorks exported ASCII STL, rotate that too.
    If ReorientAsciiStlFileToMatrix(stlPath, m) Then
        LogLine "STL reorient: ASCII STL rotated."
        ReorientStlFileToMatrix = True
        Exit Function
    End If

    LogLine "STL reorient failed: file was not successfully processed as binary or ASCII STL."
    Exit Function

ErrHandler:
    LogLine "ReorientStlFileToMatrix error: " & Err.Description
    ReorientStlFileToMatrix = False
End Function

Private Function ReorientBinaryStlFileToMatrix(ByVal stlPath As String, ByRef m() As Double) As Boolean
On Error GoTo ErrHandler
    ReorientBinaryStlFileToMatrix = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If stlPath = "" Or fso.FileExists(stlPath) = False Then Exit Function

    Dim f As Integer
    f = FreeFile
    Open stlPath For Binary Access Read Write As #f

    Dim hdr As BinaryStlHeader
    Get #f, 1, hdr
    If hdr.TriangleCount <= 0 Then
        Close #f
        LogLine "STL reorient skipped: triangle count <= 0."
        Exit Function
    End If

    Dim expectedLen As Double
    expectedLen = 84# + CDbl(hdr.TriangleCount) * 50#
    If CDbl(LOF(f)) <> expectedLen Then
        Close #f
        LogLine "STL reorient skipped: not binary STL size (got " & LOF(f) & " expected " & expectedLen & ")."
        Exit Function
    End If

    Dim tri As BinaryStlTriangle
    If Len(tri) <> 50 Then
        Close #f
        LogLine "STL reorient skipped: BinaryStlTriangle size=" & Len(tri)
        Exit Function
    End If

    Dim i As Long, triPos As Long
    For i = 0 To hdr.TriangleCount - 1
        triPos = 85 + i * 50
        Get #f, triPos, tri
        TransformStlTriangleByMatrix tri, m
        Put #f, triPos, tri
    Next i
    Close #f
    ReorientBinaryStlFileToMatrix = True
    Exit Function
ErrHandler:
    LogLine "ReorientBinaryStlFileToMatrix error: " & Err.Description
    On Error Resume Next
    Close #f
    ReorientBinaryStlFileToMatrix = False
End Function

Private Sub TransformStlTriangleByMatrix(ByRef tri As BinaryStlTriangle, _
                                         ByRef m() As Double)
On Error Resume Next

    TransformStlVectorByMatrix tri.nx, tri.ny, tri.nz, m
    NormalizeStlVector tri.nx, tri.ny, tri.nz

    TransformStlVectorByMatrix tri.x1, tri.y1, tri.z1, m
    TransformStlVectorByMatrix tri.x2, tri.y2, tri.z2, m
    TransformStlVectorByMatrix tri.x3, tri.y3, tri.z3, m
End Sub

Private Sub TransformStlVectorByMatrix(ByRef x As Single, _
                                       ByRef y As Single, _
                                       ByRef z As Single, _
                                       ByRef m() As Double)
On Error Resume Next

    Dim ox As Double
    Dim oy As Double
    Dim oz As Double

    ox = CDbl(x)
    oy = CDbl(y)
    oz = CDbl(z)

    ' Same projection convention already used elsewhere in your macro:
    ' view X = m(0), m(3), m(6)
    ' view Y = m(1), m(4), m(7)
    ' view Z = m(2), m(5), m(8)
    x = CSng((ox * m(0)) + (oy * m(3)) + (oz * m(6)))
    y = CSng((ox * m(1)) + (oy * m(4)) + (oz * m(7)))
    z = CSng((ox * m(2)) + (oy * m(5)) + (oz * m(8)))
End Sub

Private Sub NormalizeStlVector(ByRef x As Single, _
                               ByRef y As Single, _
                               ByRef z As Single)
On Error Resume Next

    Dim L As Double

    L = Sqr(CDbl(x) * CDbl(x) + CDbl(y) * CDbl(y) + CDbl(z) * CDbl(z))

    If L <= 0.0000001 Then Exit Sub

    x = CSng(CDbl(x) / L)
    y = CSng(CDbl(y) / L)
    z = CSng(CDbl(z) / L)
End Sub

Private Function ReorientAsciiStlFileToMatrix(ByVal stlPath As String, ByRef m() As Double) As Boolean
On Error GoTo ErrHandler
    ReorientAsciiStlFileToMatrix = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If stlPath = "" Or fso.FileExists(stlPath) = False Then Exit Function

    Dim txt As String
    txt = ReadAllTextFile(stlPath)
    If Trim(txt) = "" Then Exit Function
    If InStr(1, Left$(txt, 512), "solid", vbTextCompare) = 0 Then Exit Function
    If InStr(1, txt, "vertex", vbTextCompare) = 0 Then Exit Function
    If InStr(1, txt, "facet normal", vbTextCompare) = 0 Then Exit Function

    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String
    lines = Split(txt, vbLf)
    Dim outLines() As String
    ReDim outLines(LBound(lines) To UBound(lines))

    Dim i As Long, rawLine As String, T As String, indent As String
    Dim toks() As String
    Dim x As Double, y As Double, z As Double
    Dim sx As Single, sy As Single, sz As Single
    Dim changed As Boolean
    changed = False

    For i = LBound(lines) To UBound(lines)
        rawLine = lines(i)
        T = LTrim(rawLine)
        indent = Left$(rawLine, Len(rawLine) - Len(T))
        If Left$(LCase$(T), 12) = "facet normal" Or Left$(LCase$(T), 6) = "vertex" Then
            toks = Split(Replace(Replace(T, vbTab, " "), "  ", " "), " ")
            ' facet normal nx ny nz  OR  vertex x y z
            Dim nTok As Long, k As Long, nums() As Double, nNum As Long
            nTok = UBound(toks)
            ReDim nums(1 To 3)
            nNum = 0
            For k = 0 To nTok
                If IsNumeric(toks(k)) Then
                    nNum = nNum + 1
                    If nNum <= 3 Then nums(nNum) = CDbl(toks(k))
                End If
            Next k
            If nNum >= 3 Then
                sx = CSng(nums(1)): sy = CSng(nums(2)): sz = CSng(nums(3))
                TransformStlVectorByMatrix sx, sy, sz, m
                If Left$(LCase$(T), 12) = "facet normal" Then NormalizeStlVector sx, sy, sz
                If Left$(LCase$(T), 12) = "facet normal" Then
                    outLines(i) = indent & "facet normal " & _
                        Format(sx, "0.000000E+00") & " " & Format(sy, "0.000000E+00") & " " & Format(sz, "0.000000E+00")
                Else
                    outLines(i) = indent & "vertex " & _
                        Format(sx, "0.000000E+00") & " " & Format(sy, "0.000000E+00") & " " & Format(sz, "0.000000E+00")
                End If
                changed = True
            Else
                outLines(i) = rawLine
            End If
        Else
            outLines(i) = rawLine
        End If
    Next i

    If Not changed Then Exit Function

    Dim outTxt As String
    outTxt = Join(outLines, vbLf)
    Dim f As Integer
    f = FreeFile
    Open stlPath For Output As #f
    Print #f, outTxt;
    Close #f
    ReorientAsciiStlFileToMatrix = True
    Exit Function
ErrHandler:
    LogLine "ReorientAsciiStlFileToMatrix error: " & Err.Description
    On Error Resume Next
    Close #f
    ReorientAsciiStlFileToMatrix = False
End Function

Private Function SaveViewAsImage(ByVal model As Object, ByVal imagePath As String) As Boolean
On Error GoTo ErrHandler

    SaveViewAsImage = False

    If model Is Nothing Then Exit Function
    If imagePath = "" Then Exit Function

    Dim errs As Long
    Dim warns As Long

    swApp.Visible = True
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    Set model = swApp.ActiveDoc

    If model Is Nothing Then Exit Function

    PrepareModelForJpegCapture model, False

    LogLine "Saving JPG image: " & imagePath

    model.Extension.SaveAs3 imagePath, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent, _
                            Nothing, Nothing, errs, warns

    LogLine "Image save (" & imagePath & ") errs=" & errs & " warns=" & warns

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim i As Long
    For i = 1 To 20
        If fso.FileExists(imagePath) Then Exit For
        WaitMilliseconds 250
        DoEvents
    Next i

    If fso.FileExists(imagePath) Then
        LogLine "JPG OK: " & imagePath & " size=" & CStr(fso.GetFile(imagePath).Size) & " bytes"
        SaveViewAsImage = True
    Else
        LogLine "WARNING: JPG was not created: " & imagePath
    End If

    Exit Function

ErrHandler:
    LogLine "SaveViewAsImage error: " & Err.Description
    SaveViewAsImage = False
End Function

' ============================================================
Private Sub RunVisualMoldInspection()
On Error GoTo ErrHandler
    If RUN_VISUAL_MOLD_INSPECTION = False Then Exit Sub
    If CurrentJobFolder = "" Then Exit Sub

    Dim scriptPath As String
    scriptPath = "C:\CMS_Local_Workspace\cms_visual_inspect.ps1"
    If Dir(scriptPath) = "" Then scriptPath = DOWNLOADS_FOLDER & "\New folder (17)\cms_visual_inspect.ps1"
    If Dir(scriptPath) = "" Then
        LogLine "Visual mold inspection skipped: cms_visual_inspect.ps1 not found."
        Exit Sub
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & _
          Chr(34) & scriptPath & Chr(34) & _
          " -JobFolder " & Chr(34) & CurrentJobFolder & Chr(34) & _
          " -CadCsv " & Chr(34) & CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv" & Chr(34)
    LogLine "Running visual mold inspection."
    sh.Run cmd, 0, True
    LogLine "Visual mold inspection done."
    Exit Sub
ErrHandler:
    LogLine "RunVisualMoldInspection error: " & Err.Description
End Sub

' DXF FROM SAVED X_T  (4 projected views, optional dimensions)
' ============================================================
Private Sub CreateProjectedDxfFromXtPath(ByVal xtPath As String, _
                                         ByVal dxfPath As String, _
                                         ByVal quoteName As String, _
                                         ByVal parentPrimaryView As String, _
                                         ByVal parentFallbackView As String, _
                                         ByVal addDimensions As Boolean, _
                                         ByVal allSolid As Boolean)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(xtPath) = False Then
        LogLine "DXF skipped. XT source missing: " & xtPath
        Exit Sub
    End If
    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_DXF_FROM_XT_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder
    Dim nativePath As String
    nativePath = OpenXtAndSaveNativeForDrawing(xtPath, tempFolder, "BASE")
    If nativePath = "" Then
        LogLine "DXF skipped. Could not open XT native source for: " & xtPath
        GoTo CleanExit
    End If
    CreateProjectedDxfFromNativePath nativePath, dxfPath, quoteName, _
                                     parentPrimaryView, parentFallbackView, addDimensions, allSolid
CleanExit:
    On Error Resume Next
    If tempFolder <> "" Then fso.DeleteFolder tempFolder, True
    Exit Sub
ErrHandler:
    LogLine "CreateProjectedDxfFromXtPath error: " & Err.Description
    Resume CleanExit
End Sub

Private Function OpenXtAndSaveNativeForDrawing(ByVal xtPath As String, _
                                               ByVal tempFolder As String, _
                                               ByVal baseName As String) As String
On Error GoTo ErrHandler
    OpenXtAndSaveNativeForDrawing = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(xtPath) = False Then Exit Function
    EnsureFolderDeep tempFolder
    Dim importErrors As Long
    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object
    Set mdl = swApp.LoadFile4(xtPath, "", Nothing, importErrors)
    If mdl Is Nothing Then Set mdl = swApp.OpenDoc6(xtPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    If mdl Is Nothing Then Set mdl = swApp.OpenDoc6(xtPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    If mdl Is Nothing Then
        LogLine "OpenXtAndSaveNativeForDrawing failed to open: " & xtPath
        Exit Function
    End If
    Dim nativePath As String
    If mdl.GetType = swDocASSEMBLY Then
        nativePath = tempFolder & "\" & CleanFileName(baseName) & ".sldasm"
    Else
        nativePath = tempFolder & "\" & CleanFileName(baseName) & ".sldprt"
    End If
    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs
    SetCmsTopOrientation mdl
    ApplyCmsTopView mdl
    mdl.Extension.SaveAs3 nativePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns
    If fso.FileExists(nativePath) Then OpenXtAndSaveNativeForDrawing = nativePath
    On Error Resume Next
    swApp.CloseDoc mdl.GetTitle
    Exit Function
ErrHandler:
    LogLine "OpenXtAndSaveNativeForDrawing error: " & Err.Description
    OpenXtAndSaveNativeForDrawing = ""
End Function

Private Sub CreateProjectedDxfFromNativePath(ByVal nativePath As String, _
                                             ByVal dxfPath As String, _
                                             ByVal quoteName As String, _
                                             ByVal parentPrimaryView As String, _
                                             ByVal parentFallbackView As String, _
                                             ByVal addDimensions As Boolean, _
                                             ByVal allSolid As Boolean)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim drawTitle As String
    drawTitle = ""
    Dim freezeApplied As Boolean
    freezeApplied = False

    If fso.FileExists(SW_DRAWING_TEMPLATE_PATH) = False Then
        LogLine "DXF skipped. Drawing template not found: " & SW_DRAWING_TEMPLATE_PATH
        Exit Sub
    End If
    If fso.FileExists(nativePath) = False Then
        LogLine "DXF skipped. Native source missing: " & nativePath
        Exit Sub
    End If

    EnsureNativeDxfSourceUsesCmsTop nativePath, parentPrimaryView

    Dim partL As Double, partW As Double, partT As Double
    partL = 0: partW = 0: partT = 0
    TryGetNativeModelDimsInches nativePath, partL, partW, partT
    If partL <= 0 Then partL = 42#
    If partW <= 0 Then partW = 30#
    If partT <= 0 Then partT = 10#

    ' Base sheet: auto-fit so all four projected views land on the E-size sheet.
    CurrentDxfForce1to1 = False
    Dim scaleVal As Double
    scaleVal = CalculateProjectedFourViewDxfScale(partL, partW, partT) * MULTIVIEW_FIT_SAFETY
    If scaleVal <= 0 Then scaleVal = 0.1

    Dim swDraw As Object
    Set swDraw = swApp.NewDocument(SW_DRAWING_TEMPLATE_PATH, 0, _
                                   E_SHEET_WIDTH_IN / INCHES_PER_METER, _
                                   E_SHEET_HEIGHT_IN / INCHES_PER_METER)
    If swDraw Is Nothing Then
        LogLine "DXF skipped. Could not create drawing."
        GoTo CleanExit
    End If
    drawTitle = swDraw.GetTitle
    Dim errs As Long
    swApp.ActivateDoc3 drawTitle, False, 0, errs
    EnsureSwHidden
    SetupDrawingAsESize swDraw
    If FREEZE_DXF_DRAWING_GRAPHICS Then
        FreezeDxfDrawingGraphics swDraw
        freezeApplied = True
    End If

    Dim centerX As Double, centerY As Double
    centerX = E_SHEET_WIDTH_IN / 2#
    centerY = E_SHEET_HEIGHT_IN / 2#

    Dim projectedXOffset As Double, projectedYOffset As Double
    ' Parent view is CMS_TOP: sheet X ~ Width, sheet Y ~ Length (shop DXF labels).
    ' Projected RIGHT sits to the side: its DXF X is Thickness, DXF Y is Length.
    projectedXOffset = ((partW / 2#) + DXF_PROJECTED_VIEW_GAP_IN + (partT / 2#)) * scaleVal
    projectedYOffset = ((partL / 2#) + DXF_PROJECTED_VIEW_GAP_IN + (partT / 2#)) * scaleVal

    Dim xLeft As Double, yLeft As Double, xRight As Double, yRight As Double
    Dim xTop As Double, yTop As Double, xBottom As Double, yBottom As Double
    xLeft = centerX - projectedXOffset: yLeft = centerY
    xRight = centerX + projectedXOffset: yRight = centerY
    xTop = centerX: yTop = centerY + projectedYOffset
    xBottom = centerX: yBottom = centerY - projectedYOffset
    If xLeft < DXF_MARGIN_IN Then xLeft = DXF_MARGIN_IN
    If xRight > E_SHEET_WIDTH_IN - DXF_MARGIN_IN Then xRight = E_SHEET_WIDTH_IN - DXF_MARGIN_IN
    If yTop > E_SHEET_HEIGHT_IN - DXF_MARGIN_IN Then yTop = E_SHEET_HEIGHT_IN - DXF_MARGIN_IN
    If yBottom < DXF_MARGIN_IN Then yBottom = DXF_MARGIN_IN

    Dim parentView As Object
    Dim viewLeft As Object
    Dim viewRight As Object
    Dim viewTop As Object
    Dim viewBottom As Object

    Set parentView = CreateParentDrawingView(swDraw, nativePath, parentPrimaryView, parentFallbackView, centerX, centerY, scaleVal)
    If parentView Is Nothing Then
        LogLine "DXF skipped. Could not create parent drawing view."
        GoTo CleanExit
    End If

    Set viewLeft = CreateProjectedDrawingView(swDraw, parentView, xLeft, yLeft, scaleVal, "LEFT")
    Set viewRight = CreateProjectedDrawingView(swDraw, parentView, xRight, yRight, scaleVal, "RIGHT")
    Set viewTop = CreateProjectedDrawingView(swDraw, parentView, xTop, yTop, scaleVal, "TOP")
    Set viewBottom = CreateProjectedDrawingView(swDraw, parentView, xBottom, yBottom, scaleVal, "BOTTOM")


    If allSolid Then
        ForceAllDrawingViewsSolid swDraw
    Else
        ForceAllDrawingViewsWireframe swDraw
    End If

    If addDimensions Then
        AddBaseOverallDimensions swDraw, parentView, viewRight, partL, partW, partT
    End If

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    Dim saveErrs As Long
    Dim saveWarns As Long
    LogLine "Saving DXF: " & dxfPath
    swDraw.Extension.SaveAs3 dxfPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, saveErrs, saveWarns
    LogLine "DXF save done. Errors=" & saveErrs & " Warnings=" & saveWarns

CleanExit:
    On Error Resume Next
    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If
    If drawTitle <> "" Then swApp.CloseDoc drawTitle
    Exit Sub
ErrHandler:
    LogLine "CreateProjectedDxfFromNativePath error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub EnsureNativeDxfSourceUsesCmsTop(ByVal nativePath As String, ByVal primaryViewName As String)
On Error GoTo ErrHandler

    If nativePath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(nativePath) = False Then Exit Sub

    Dim docType As Long
    docType = swDocPART
    If LCase$(Right$(nativePath, 6)) = "sldasm" Then docType = swDocASSEMBLY

    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object

    Set mdl = swApp.OpenDoc6(nativePath, docType, swOpenDocOptions_Silent, "", errs, warns)
    If mdl Is Nothing Then
        LogLine "DXF view prep skipped. Could not open native source: " & nativePath
        Exit Sub
    End If

    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs

    ' Preserve corrected FRONT before touching TOP.
    Dim frontReady As Boolean
    frontReady = False

    On Error Resume Next
    Err.Clear
    mdl.ShowNamedView2 CMS_FRONT_VIEW_NAME, -1
    If Err.Number = 0 Then
        frontReady = True
        LogLine "DXF view prep: found CMS_FRONT named view."
    Else
        Err.Clear
        mdl.ShowNamedView2 "*Front", 1
        If Err.Number = 0 Then
            SaveCurrentViewAsNamed mdl, CMS_FRONT_VIEW_NAME
            frontReady = True
            LogLine "DXF view prep: CMS_FRONT not found; saved current *Front as CMS_FRONT."
        Else
            Err.Clear
            LogLine "DXF view prep warning: could not preserve corrected front view."
        End If
    End If
    On Error GoTo ErrHandler

    Dim usedView As String
    usedView = ""

    If primaryViewName <> "" Then
        On Error Resume Next
        Err.Clear
        mdl.ShowNamedView2 primaryViewName, -1
        If Err.Number = 0 Then usedView = primaryViewName
        Err.Clear
        On Error GoTo ErrHandler
    End If

    If usedView = "" Then
        On Error Resume Next
        Err.Clear
        mdl.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
        If Err.Number = 0 Then usedView = CMS_TOP_VIEW_NAME
        Err.Clear
        On Error GoTo ErrHandler
    End If

    If usedView = "" Then
        On Error Resume Next
        Err.Clear
        mdl.ShowNamedView2 "*Top", 5
        If Err.Number = 0 Then usedView = "*Top"
        Err.Clear
        On Error GoTo ErrHandler
    End If

    If usedView = "" Then
        LogLine "DXF view prep warning: no CMS_TOP/*Top view could be applied for " & nativePath
        GoTo CleanExit
    End If

    StabilizeActiveView mdl, 100

    If PersistCurrentViewAsStandardTop(mdl) Then
        LogLine "DXF view prep: corrected top persisted as SolidWorks *Top from " & usedView
    Else
        LogLine "DXF view prep warning: could not persist corrected top as *Top."
    End If

    ' Re-apply corrected FRONT after TOP persistence.
    If frontReady Then
        On Error Resume Next
        Err.Clear
        mdl.ShowNamedView2 CMS_FRONT_VIEW_NAME, -1
        On Error GoTo ErrHandler

        StabilizeActiveView mdl, 100

        If PersistCurrentViewAsStandardFront(mdl) Then
            LogLine "DXF view prep: corrected CMS_FRONT persisted as SolidWorks *Front."
        Else
            LogLine "DXF view prep warning: could not persist CMS_FRONT as *Front."
        End If
    End If

    On Error Resume Next
    mdl.ShowNamedView2 "*Top", 5
    On Error GoTo ErrHandler

    StabilizeActiveView mdl, 100

    mdl.Extension.SaveAs3 nativePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns

    LogLine "DXF view prep: saved native source with corrected *Top and *Front. Errors=" & errs & " Warnings=" & warns

CleanExit:
    On Error Resume Next
    If Not mdl Is Nothing Then swApp.CloseDoc mdl.GetTitle
    Exit Sub

ErrHandler:
    LogLine "EnsureNativeDxfSourceUsesCmsTop error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub FreezeDxfDrawingGraphics(Optional ByVal docToFreeze As Object = Nothing)
On Error Resume Next
    Set DxfFreezeDoc = docToFreeze
    swApp.UserControl = False
    If Not DxfFreezeDoc Is Nothing Then DxfFreezeDoc.FeatureManager.EnableFeatureTree = False
    If Not swModel Is Nothing Then swModel.FeatureManager.EnableFeatureTree = False
    DoEvents
End Sub

Private Sub UnfreezeDxfDrawingGraphics()
On Error Resume Next
    If Not DxfFreezeDoc Is Nothing Then DxfFreezeDoc.FeatureManager.EnableFeatureTree = True
    If Not swModel Is Nothing Then swModel.FeatureManager.EnableFeatureTree = True
    swApp.UserControl = True
    Set DxfFreezeDoc = Nothing
    DoEvents
End Sub

Private Function CalculateProjectedFourViewDxfScale(ByVal partL As Double, ByVal partW As Double, ByVal partT As Double) As Double
    Dim scaleVal As Double
    scaleVal = DXF_MAX_SCALE
    Dim usableW As Double
    Dim usableH As Double
    usableW = E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)
    usableH = E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)
    Dim layoutW As Double
    Dim layoutH As Double
    ' TOP footprint is W (horizontal) x L (vertical); side projections add Thickness.
    layoutW = partW + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
    layoutH = partL + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
    If layoutW > 0 Then
        If usableW / layoutW < scaleVal Then scaleVal = usableW / layoutW
    End If
    If layoutH > 0 Then
        If usableH / layoutH < scaleVal Then scaleVal = usableH / layoutH
    End If
    If scaleVal <= 0 Then scaleVal = 1#
    If scaleVal > DXF_MAX_SCALE Then scaleVal = DXF_MAX_SCALE
    CalculateProjectedFourViewDxfScale = scaleVal
End Function

Private Function CreateParentDrawingView(ByVal swDraw As Object, ByVal nativePath As String, _
                                         ByVal primaryViewName As String, ByVal fallbackViewName As String, _
                                         ByVal xIn As Double, ByVal yIn As Double, ByVal scaleVal As Double) As Object
On Error GoTo ErrHandler
    Set CreateParentDrawingView = Nothing
    If swDraw Is Nothing Then Exit Function
    Dim swView As Object
    Set swView = Nothing
    On Error Resume Next
    Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, primaryViewName, xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    If swView Is Nothing Then Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, fallbackViewName, xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    If swView Is Nothing Then Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, "*Top", xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    On Error GoTo ErrHandler
    If swView Is Nothing Then Exit Function
    SetDrawingViewScale swView, scaleVal
    Set CreateParentDrawingView = swView
    Exit Function
ErrHandler:
    LogLine "CreateParentDrawingView error: " & Err.Description
    Set CreateParentDrawingView = Nothing
End Function

Private Function CreateProjectedDrawingView(ByVal swDraw As Object, ByVal parentView As Object, _
                                            ByVal xIn As Double, ByVal yIn As Double, _
                                            ByVal scaleVal As Double, ByVal labelText As String) As Object
On Error GoTo ErrHandler
    Set CreateProjectedDrawingView = Nothing
    If swDraw Is Nothing Then Exit Function
    If parentView Is Nothing Then Exit Function
    swDraw.ClearSelection2 True
    If SelectDrawingView(swDraw, parentView) = False Then
        LogLine "Could not select parent view for " & labelText
        Exit Function
    End If
    Dim projView As Object
    Set projView = Nothing
    On Error Resume Next
    Set projView = swDraw.CreateUnfoldedViewAt3(xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#, False)
    On Error GoTo ErrHandler
    swDraw.ClearSelection2 True
    If projView Is Nothing Then
        LogLine "Could not create projected view: " & labelText
        Exit Function
    End If
    SetDrawingViewScale projView, scaleVal
    Set CreateProjectedDrawingView = projView
    Exit Function
ErrHandler:
    LogLine "CreateProjectedDrawingView error for " & labelText & ": " & Err.Description
    On Error Resume Next
    swDraw.ClearSelection2 True
    Set CreateProjectedDrawingView = Nothing
End Function

Private Function SelectDrawingView(ByVal swDraw As Object, ByVal swView As Object) As Boolean
On Error GoTo ErrHandler
    SelectDrawingView = False
    If swDraw Is Nothing Then Exit Function
    If swView Is Nothing Then Exit Function
    swDraw.ClearSelection2 True
    On Error Resume Next
    SelectDrawingView = CBool(swView.SelectEntity(False))
    On Error GoTo ErrHandler
    If SelectDrawingView Then Exit Function
    Dim viewName As String
    viewName = ""
    On Error Resume Next
    viewName = CStr(swView.GetName2)
    On Error GoTo ErrHandler
    If viewName <> "" Then
        Dim ok As Boolean
        On Error Resume Next
        ok = CBool(swDraw.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0))
        On Error GoTo ErrHandler
        SelectDrawingView = ok
        Exit Function
    End If
    SelectDrawingView = False
    Exit Function
ErrHandler:
    SelectDrawingView = False
End Function

Private Sub ForceAllDrawingViewsWireframe(ByVal swDraw As Object)
On Error Resume Next
    If swDraw Is Nothing Then Exit Sub
    Dim v As Object
    Set v = swDraw.GetFirstView
    If Not v Is Nothing Then Set v = v.GetNextView
    Do While Not v Is Nothing
        SetDrawingViewWireframe v
        Set v = v.GetNextView
    Loop
    swDraw.GraphicsRedraw2
End Sub

Private Sub ForceAllDrawingViewsSolid(ByVal swDraw As Object)
On Error Resume Next
    If swDraw Is Nothing Then Exit Sub
    Dim v As Object
    Set v = swDraw.GetFirstView
    If Not v Is Nothing Then Set v = v.GetNextView
    Do While Not v Is Nothing
        SetDrawingViewSolid v
        Set v = v.GetNextView
    Loop
    swDraw.GraphicsRedraw2
End Sub

Private Sub SetupDrawingAsESize(ByVal swDraw As Object)
On Error Resume Next
    If swDraw Is Nothing Then Exit Sub
    Dim swSheet As Object
    Set swSheet = swDraw.GetCurrentSheet
    If Not swSheet Is Nothing Then
        swSheet.SetSize 12, E_SHEET_WIDTH_IN / INCHES_PER_METER, E_SHEET_HEIGHT_IN / INCHES_PER_METER
    End If
    swDraw.GraphicsRedraw2
End Sub

Private Sub SetDrawingViewWireframe(ByVal swView As Object)
On Error Resume Next
    If swView Is Nothing Then Exit Sub
    swView.UseParentStyle = False
    swView.SetDisplayMode3 False, 0, False, True
    swView.DisplayMode = 0
End Sub

Private Sub SetDrawingViewSolid(ByVal swView As Object)
On Error Resume Next
    If swView Is Nothing Then Exit Sub
    swView.UseParentStyle = False
    swView.SetDisplayMode3 False, 2, False, True
    swView.DisplayMode = 2
End Sub

Private Sub SetDrawingViewScale(ByVal swView As Object, ByVal scaleVal As Double)
On Error Resume Next
    If swView Is Nothing Then Exit Sub
    If CurrentDxfForce1to1 Then scaleVal = 1#
    If scaleVal <= 0 Then scaleVal = 1#
    swView.UseSheetScale = False
    swView.ScaleDecimal = scaleVal
    If scaleVal = 1# Then swView.ScaleRatio = "1:1"
End Sub

' ============================================================
' OVERALL DIMENSIONS ON THE DIM DXF
' ============================================================
Private Sub AddBaseOverallDimensions(ByVal swDraw As Object, ByVal parentView As Object, _
                                     ByVal sideView As Object, _
                                     ByVal partL As Double, ByVal partW As Double, ByVal partT As Double)
On Error GoTo ErrHandler
    If swDraw Is Nothing Then Exit Sub

    ' Overall length (horizontal) and width/height (vertical) on the top view.
    AddViewOverallDimension swDraw, parentView, True
    AddViewOverallDimension swDraw, parentView, False
    ' Overall thickness on a side view.
    If Not sideView Is Nothing Then AddViewOverallDimension swDraw, sideView, True

    ' Always-visible labeled notes (render even if edge-attached dims miss).
    AddBaseDimensionNotes swDraw, parentView, partL, partW, partT
    swDraw.GraphicsRedraw2
    Exit Sub
ErrHandler:
    LogLine "AddBaseOverallDimensions error: " & Err.Description
End Sub

Private Sub AddViewOverallDimension(ByVal swDraw As Object, ByVal swView As Object, ByVal horizontal As Boolean)
On Error GoTo ErrHandler
    If swDraw Is Nothing Then Exit Sub
    If swView Is Nothing Then Exit Sub
    Dim vOut As Variant
    vOut = swView.GetOutline
    If IsEmpty(vOut) Then Exit Sub
    If IsArray(vOut) = False Then Exit Sub
    If UBound(vOut) < 3 Then Exit Sub
    Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
    Dim midx As Double, midy As Double, gap As Double
    xmin = CDbl(vOut(0)): ymin = CDbl(vOut(1)): xmax = CDbl(vOut(2)): ymax = CDbl(vOut(3))
    midx = (xmin + xmax) / 2#
    midy = (ymin + ymax) / 2#
    gap = 0.625 / INCHES_PER_METER
    swDraw.ClearSelection2 True
    Dim ok1 As Boolean, ok2 As Boolean
    Dim dispDim As Object
    If horizontal Then
        ok1 = TrySelectViewSilhouetteEdge(swDraw, xmin, midy, True, False)
        ok2 = TrySelectViewSilhouetteEdge(swDraw, xmax, midy, True, True)
        If ok1 And ok2 Then Set dispDim = swDraw.AddHorizontalDimension2(midx, ymin - gap, 0#)
    Else
        ok1 = TrySelectViewSilhouetteEdge(swDraw, midx, ymin, False, False)
        ok2 = TrySelectViewSilhouetteEdge(swDraw, midx, ymax, False, True)
        If ok1 And ok2 Then Set dispDim = swDraw.AddVerticalDimension2(xmax + gap, midy, 0#)
    End If
    swDraw.ClearSelection2 True
    Exit Sub
ErrHandler:
    On Error Resume Next
    swDraw.ClearSelection2 True
    LogLine "AddViewOverallDimension error: " & Err.Description
End Sub

Private Function TrySelectViewSilhouetteEdge(ByVal swDraw As Object, ByVal x As Double, ByVal y As Double, _
                                             ByVal edgeIsVertical As Boolean, ByVal appendToSelection As Boolean) As Boolean
On Error GoTo ErrHandler
    TrySelectViewSilhouetteEdge = False
    Dim nudges(0 To 4) As Double
    nudges(0) = 0#
    nudges(1) = 0.015 / INCHES_PER_METER
    nudges(2) = -0.015 / INCHES_PER_METER
    nudges(3) = 0.04 / INCHES_PER_METER
    nudges(4) = -0.04 / INCHES_PER_METER
    Dim i As Long
    Dim tx As Double, ty As Double, okSel As Boolean
    For i = 0 To 4
        tx = x: ty = y
        If edgeIsVertical Then tx = x + nudges(i) Else ty = y + nudges(i)
        okSel = CBool(swDraw.Extension.SelectByID2("", "EDGE", tx, ty, 0#, appendToSelection, 0, Nothing, 0))
        If okSel Then
            TrySelectViewSilhouetteEdge = True
            Exit Function
        End If
    Next i
    Exit Function
ErrHandler:
    TrySelectViewSilhouetteEdge = False
End Function

Private Sub AddBaseDimensionNotes(ByVal swDraw As Object, ByVal swView As Object, _
                                  ByVal partL As Double, ByVal partW As Double, ByVal partT As Double)
On Error GoTo ErrHandler
    If swDraw Is Nothing Then Exit Sub
    Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
    xmin = E_SHEET_WIDTH_IN / 2# / INCHES_PER_METER
    ymin = E_SHEET_HEIGHT_IN / 2# / INCHES_PER_METER
    xmax = xmin: ymax = ymin
    If Not swView Is Nothing Then
        Dim vOut As Variant
        vOut = swView.GetOutline
        If IsArray(vOut) Then
            If UBound(vOut) >= 3 Then
                xmin = CDbl(vOut(0)): ymin = CDbl(vOut(1))
                xmax = CDbl(vOut(2)): ymax = CDbl(vOut(3))
            End If
        End If
    End If
    Dim midx As Double, midy As Double, gap As Double
    midx = (xmin + xmax) / 2#
    midy = (ymin + ymax) / 2#
    gap = 0.6 / INCHES_PER_METER
    ' TOP view: horizontal = Width, vertical = Length (matches shop DXF labels).
    PlaceDrawingNote swDraw, midx - gap, ymin - gap, "OVERALL WIDTH (W) = " & FormatDim(partW)
    PlaceDrawingNote swDraw, xmin - (2.4 / INCHES_PER_METER), midy, "OVERALL LENGTH (L) = " & FormatDim(partL)
    PlaceDrawingNote swDraw, xmax + gap, ymax + gap, "THICKNESS (T) = " & FormatDim(partT)
    swDraw.ClearSelection2 True
    Exit Sub
ErrHandler:
    LogLine "AddBaseDimensionNotes error: " & Err.Description
End Sub

Private Sub PlaceDrawingNote(ByVal swDraw As Object, ByVal xMeters As Double, ByVal yMeters As Double, ByVal text As String)
On Error Resume Next
    If swDraw Is Nothing Then Exit Sub
    swDraw.ClearSelection2 True
    Dim swNote As Object
    Set swNote = swDraw.InsertNote(text)
    If Not swNote Is Nothing Then
        Dim swAnn As Object
        Set swAnn = swNote.GetAnnotation
        If Not swAnn Is Nothing Then swAnn.SetPosition xMeters, yMeters, 0#
    End If
    swDraw.ClearSelection2 True
End Sub

Private Function FormatDim(ByVal v As Double) As String
    FormatDim = Format(v, "0.000")
End Function

' ============================================================
' BOUNDING-BOX DIMENSIONS (for DXF sizing)
' ============================================================
Private Function TryGetNativeModelDimsInches(ByVal nativePath As String, _
                                             ByRef l As Double, ByRef w As Double, ByRef t As Double) As Boolean
On Error GoTo ErrHandler
    TryGetNativeModelDimsInches = False
    l = 0: w = 0: t = 0
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(nativePath) = False Then Exit Function
    Dim ext As String
    ext = LCase(fso.GetExtensionName(nativePath))
    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object
    If ext = "sldasm" Then
        Set mdl = swApp.OpenDoc6(nativePath, swDocASSEMBLY, swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, "", errs, warns)
    Else
        Set mdl = swApp.OpenDoc6(nativePath, swDocPART, swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, "", errs, warns)
    End If
    If mdl Is Nothing Then Exit Function
    Dim dx As Double, dy As Double, dz As Double
    Dim gotBox As Boolean
    gotBox = TryGetModelDocBoxDimsInches(mdl, dx, dy, dz)
    On Error Resume Next
    swApp.CloseDoc mdl.GetTitle
    On Error GoTo ErrHandler
    If gotBox = False Then Exit Function
    ' Prefer CMS Top/Right/Front axes when the view frame is locked.
    If gCmsViewFrameReady Then
        AssignLengthWidthThicknessFromAxes dx, dy, dz, l, w, t
    Else
        SortThreeDimensions dx, dy, dz, l, w, t
    End If
    l = Round(l, DIM_DECIMALS): w = Round(w, DIM_DECIMALS): t = Round(t, DIM_DECIMALS)
    TryGetNativeModelDimsInches = (l > 0 And w > 0 And t > 0)
    Exit Function
ErrHandler:
    TryGetNativeModelDimsInches = False
End Function

Private Function TryGetModelDocBoxDimsInches(ByVal mdl As Object, _
                                             ByRef dx As Double, ByRef dy As Double, ByRef dz As Double) As Boolean
On Error GoTo ErrHandler
    TryGetModelDocBoxDimsInches = False
    If mdl Is Nothing Then Exit Function
    If mdl.GetType = swDocPART Then
        TryGetModelDocBoxDimsInches = GetPartBoundingBoxInches(mdl, dx, dy, dz)
        Exit Function
    End If
    Dim vBox As Variant
    On Error Resume Next
    vBox = mdl.GetBox(False, False)
    On Error GoTo ErrHandler
    If IsValidBoxArray(vBox) = False Then Exit Function
    Dim cxTmp As Double, cyTmp As Double, czTmp As Double
    BoxCornersToInches vBox, dx, dy, dz, cxTmp, cyTmp, czTmp
    TryGetModelDocBoxDimsInches = True
    Exit Function
ErrHandler:
    TryGetModelDocBoxDimsInches = False
End Function

Private Function GetPartBoundingBoxInches(ByVal swPartModel As Object, _
                                          ByRef dxIn As Double, ByRef dyIn As Double, ByRef dzIn As Double) As Boolean
On Error GoTo ErrHandler
    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then
        GetPartBoundingBoxInches = False
        Exit Function
    End If
    Dim firstBody As Boolean
    firstBody = True
    Dim xmin As Double, ymin As Double, zmin As Double
    Dim xmax As Double, ymax As Double, zmax As Double
    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant
    For i = 0 To UBound(vBodies)
        Set swBody = vBodies(i)
        If Not swBody Is Nothing Then
            vBox = swBody.GetBodyBox
            If Not IsEmpty(vBox) Then
                If firstBody Then
                    xmin = CDbl(vBox(0)): ymin = CDbl(vBox(1)): zmin = CDbl(vBox(2))
                    xmax = CDbl(vBox(3)): ymax = CDbl(vBox(4)): zmax = CDbl(vBox(5))
                    firstBody = False
                Else
                    If CDbl(vBox(0)) < xmin Then xmin = CDbl(vBox(0))
                    If CDbl(vBox(1)) < ymin Then ymin = CDbl(vBox(1))
                    If CDbl(vBox(2)) < zmin Then zmin = CDbl(vBox(2))
                    If CDbl(vBox(3)) > xmax Then xmax = CDbl(vBox(3))
                    If CDbl(vBox(4)) > ymax Then ymax = CDbl(vBox(4))
                    If CDbl(vBox(5)) > zmax Then zmax = CDbl(vBox(5))
                End If
            End If
        End If
    Next i
    If firstBody Then
        GetPartBoundingBoxInches = False
        Exit Function
    End If
    Dim fakeBox(0 To 5) As Double, cxT As Double, cyT As Double, czT As Double
    fakeBox(0) = xmin: fakeBox(1) = ymin: fakeBox(2) = zmin
    fakeBox(3) = xmax: fakeBox(4) = ymax: fakeBox(5) = zmax
    BoxCornersToInches fakeBox, dxIn, dyIn, dzIn, cxT, cyT, czT
    GetPartBoundingBoxInches = (dxIn > 0 And dyIn > 0 And dzIn > 0)
    Exit Function
ErrHandler:
    GetPartBoundingBoxInches = False
End Function

Private Function IsValidBoxArray(ByVal vBox As Variant) As Boolean
    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function
    IsValidBoxArray = True
End Function

' SolidWorks box APIs normally return meters. Some imports / document-unit
' edge cases return inches (or mm) already — blindly *39.37 makes dims huge.
' Pick the scale that yields sane mold-base sizes (under ~10 ft on an axis).
Private Function BoxAxisToInches(ByVal rawDelta As Double) As Double
    Dim asMeters As Double, asInches As Double, asMm As Double
    Dim best As Double
    asMeters = Abs(rawDelta) * INCHES_PER_METER
    asInches = Abs(rawDelta)
    asMm = Abs(rawDelta) / 25.4
    best = asMeters
    ' Prefer already-inches when meter conversion is absurdly large for a mold plate
    If asMeters > MAX_SANE_MOLD_DIM_IN And asInches <= MAX_SANE_MOLD_DIM_IN And asInches > 0.05 Then
        best = asInches
    ElseIf asMeters > MAX_SANE_MOLD_DIM_IN And asMm <= MAX_SANE_MOLD_DIM_IN And asMm > 0.05 Then
        best = asMm
    ElseIf asMeters <= 0.05 And asInches > 0.05 And asInches <= MAX_SANE_MOLD_DIM_IN Then
        ' Meter conversion collapsed to near-zero — raw was likely already inches
        best = asInches
    End If
    BoxAxisToInches = best
End Function

Private Sub BoxCornersToInches(ByVal vBox As Variant, _
                               ByRef dxIn As Double, ByRef dyIn As Double, ByRef dzIn As Double, _
                               ByRef cxIn As Double, ByRef cyIn As Double, ByRef czIn As Double)
    Dim rx As Double, ry As Double, rz As Double
    rx = Abs(CDbl(vBox(3)) - CDbl(vBox(0)))
    ry = Abs(CDbl(vBox(4)) - CDbl(vBox(1)))
    rz = Abs(CDbl(vBox(5)) - CDbl(vBox(2)))
    ' Detect unit scale from the largest axis so all three share one unit system.
    ' NOTE: do not name a variable "Scale" — reserved in SolidWorks VBA.
    Dim unitScale As Double
    Dim rawMax As Double
    rawMax = rx
    If ry > rawMax Then rawMax = ry
    If rz > rawMax Then rawMax = rz
    Dim asMeters As Double, asInches As Double, asMm As Double
    asMeters = rawMax * INCHES_PER_METER
    asInches = rawMax
    asMm = rawMax / 25.4
    unitScale = INCHES_PER_METER
    If asMeters > MAX_SANE_MOLD_DIM_IN And asInches <= MAX_SANE_MOLD_DIM_IN And asInches > 0.05 Then
        unitScale = 1#
        LogLine "Box units: treating as inches (meter convert was " & FormatNumberForCsv(asMeters) & " in)"
    ElseIf asMeters > MAX_SANE_MOLD_DIM_IN And asMm <= MAX_SANE_MOLD_DIM_IN And asMm > 0.05 Then
        unitScale = 1# / 25.4
        LogLine "Box units: treating as mm (meter convert was " & FormatNumberForCsv(asMeters) & " in)"
    End If
    dxIn = rx * unitScale
    dyIn = ry * unitScale
    dzIn = rz * unitScale
    cxIn = ((CDbl(vBox(0)) + CDbl(vBox(3))) / 2#) * unitScale
    cyIn = ((CDbl(vBox(1)) + CDbl(vBox(4))) / 2#) * unitScale
    czIn = ((CDbl(vBox(2)) + CDbl(vBox(5))) / 2#) * unitScale
End Sub

Private Sub SortThreeDimensions(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
                                ByRef l As Double, ByRef w As Double, ByRef t As Double)
    Dim arr(1 To 3) As Double
    Dim i As Long, j As Long, tmp As Double
    arr(1) = a: arr(2) = b: arr(3) = c
    For i = 1 To 2
        For j = i + 1 To 3
            If arr(j) > arr(i) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    l = arr(1): w = arr(2): t = arr(3)
End Sub

' ============================================================
' CMS DXF VIEW-FRAME dimensions (L/W/T follow oriented views)
'
' Shop dimensioned DXF labels (verified):
'   TOP:    X/horizontal = Width,  Y/vertical = Length, into-screen = Thickness
'   RIGHT:  X/horizontal = Thickness, Y/vertical = Length
'   FRONT/BOTTOM: X/horizontal = Width
' Model XYZ alone is NOT L/W/T — the frame rotates with the views.
' ============================================================
Private Sub ResetCmsViewFrame()
    gCmsViewFrameReady = False
    gCmsLenAxisX = 0#: gCmsLenAxisY = 1#: gCmsLenAxisZ = 0#
    gCmsWidAxisX = 1#: gCmsWidAxisY = 0#: gCmsWidAxisZ = 0#
    gCmsThkAxisX = 0#: gCmsThkAxisY = 0#: gCmsThkAxisZ = 1#
End Sub

Private Sub NormalizeAxis3(ByRef ax As Double, ByRef ay As Double, ByRef az As Double)
    Dim mag As Double
    mag = Sqr(ax * ax + ay * ay + az * az)
    If mag <= 0.0000001 Then
        ax = 0#: ay = 0#: az = 0#
    Else
        ax = ax / mag: ay = ay / mag: az = az / mag
    End If
End Sub

Private Function AbsDotAxis3(ByVal ax As Double, ByVal ay As Double, ByVal az As Double, _
                             ByVal bx As Double, ByVal by As Double, ByVal bz As Double) As Double
    AbsDotAxis3 = Abs(ax * bx + ay * by + az * bz)
End Function

' Capture Length/Width/Thickness model-space axes from the oriented CMS views.
' Uses ActiveView.Orientation3 like the pot/front depth code:
'   view X = col0 [m0,m3,m6], view Y = col1 [m1,m4,m7], depth = col2 [m2,m5,m8]
Private Function CaptureCmsViewFrameFromModel(ByVal model As Object) As Boolean
On Error GoTo eh
    CaptureCmsViewFrameFromModel = False
    ResetCmsViewFrame
    If model Is Nothing Then Exit Function

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    EnsureSwHidden

    Dim swView As Object
    Dim m As Variant
    Dim lx As Double, ly As Double, lz As Double
    Dim wx As Double, wy As Double, wz As Double
    Dim tx As Double, ty As Double, tz As Double
    Dim fx As Double, fy As Double, fz As Double
    Dim fudx As Double, fudy As Double, fudz As Double
    Dim rx As Double, ry As Double, rz As Double
    Dim rudx As Double, rudy As Double, rudz As Double
    Dim gotTop As Boolean, gotFront As Boolean, gotRight As Boolean
    gotTop = False
    gotFront = False
    gotRight = False

    ' --- TOP / CMS_TOP: Width = view X, Length = view Y, Thickness = into screen ---
    On Error Resume Next
    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    If Err.Number <> 0 Then
        Err.Clear
        model.ShowNamedView2 "*Top", 5
    End If
    On Error GoTo eh
    StabilizeActiveView model, 30
    Set swView = model.ActiveView
    If Not swView Is Nothing Then
        m = swView.Orientation3.ArrayData
        If (Not IsEmpty(m)) And IsArray(m) Then
            If UBound(m) >= 8 Then
                wx = CDbl(m(0)): wy = CDbl(m(3)): wz = CDbl(m(6))   ' view X -> Width
                lx = CDbl(m(1)): ly = CDbl(m(4)): lz = CDbl(m(7))   ' view Y -> Length
                tx = CDbl(m(2)): ty = CDbl(m(5)): tz = CDbl(m(8))   ' into screen -> Thickness
                NormalizeAxis3 wx, wy, wz
                NormalizeAxis3 lx, ly, lz
                NormalizeAxis3 tx, ty, tz
                gotTop = (Abs(wx) + Abs(wy) + Abs(wz) > 0.1)
            End If
        End If
    End If

    ' --- FRONT / BOTTOM of TOP: Width = front view X ---
    model.ShowNamedView2 "*Front", 1
    StabilizeActiveView model, 30
    Set swView = model.ActiveView
    If Not swView Is Nothing Then
        m = swView.Orientation3.ArrayData
        If (Not IsEmpty(m)) And IsArray(m) Then
            If UBound(m) >= 8 Then
                fx = CDbl(m(0)): fy = CDbl(m(3)): fz = CDbl(m(6))     ' front X = Width
                fudx = CDbl(m(1)): fudy = CDbl(m(4)): fudz = CDbl(m(7)) ' front Y = up (often Thickness)
                NormalizeAxis3 fx, fy, fz
                NormalizeAxis3 fudx, fudy, fudz
                gotFront = (Abs(fx) + Abs(fy) + Abs(fz) > 0.1)
                If gotFront Then
                    If gotTop Then
                        ' Force Width toward front X when TOP Width drifted.
                        If AbsDotAxis3(fx, fy, fz, wx, wy, wz) < 0.7 Then
                            wx = fx: wy = fy: wz = fz
                        End If
                    Else
                        wx = fx: wy = fy: wz = fz
                        tx = fudx: ty = fudy: tz = fudz
                        ' Length = Width x Thickness (placeholder until RIGHT sets Length)
                        lx = wy * tz - wz * ty
                        ly = wz * tx - wx * tz
                        lz = wx * ty - wy * tx
                        NormalizeAxis3 lx, ly, lz
                        NormalizeAxis3 tx, ty, tz
                        gotTop = True
                    End If
                End If
            End If
        End If
    End If

    ' --- RIGHT of TOP: confirm Thickness = right X, Length = right Y ---
    ' Only overwrite TOP-derived L/T when RIGHT axes agree with TOP (else RIGHT
    ' from a sideways import swaps Thickness↔Length → TCP T≈18 L≈1.4).
    Dim topTx As Double, topTy As Double, topTz As Double
    Dim topLx As Double, topLy As Double, topLz As Double
    Dim topWx As Double, topWy As Double, topWz As Double
    Dim hadTopAxes As Boolean
    hadTopAxes = gotTop
    If hadTopAxes Then
        topTx = tx: topTy = ty: topTz = tz
        topLx = lx: topLy = ly: topLz = lz
        topWx = wx: topWy = wy: topWz = wz
    End If

    model.ShowNamedView2 "*Right", 4
    StabilizeActiveView model, 30
    Set swView = model.ActiveView
    If Not swView Is Nothing Then
        m = swView.Orientation3.ArrayData
        If (Not IsEmpty(m)) And IsArray(m) Then
            If UBound(m) >= 8 Then
                rx = CDbl(m(0)): ry = CDbl(m(3)): rz = CDbl(m(6))       ' right X = Thickness
                rudx = CDbl(m(1)): rudy = CDbl(m(4)): rudz = CDbl(m(7)) ' right Y = Length
                NormalizeAxis3 rx, ry, rz
                NormalizeAxis3 rudx, rudy, rudz
                If Abs(rx) + Abs(ry) + Abs(rz) > 0.1 Then
                    Dim rightAgrees As Boolean
                    rightAgrees = True
                    If hadTopAxes Then
                        ' RIGHT X should align with TOP into-screen (Thickness).
                        If AbsDotAxis3(rx, ry, rz, topTx, topTy, topTz) < 0.7 Then
                            rightAgrees = False
                            LogLine "CMS view frame: RIGHT X does not match TOP thickness — keeping TOP L/W/T axes."
                        End If
                    End If
                    If rightAgrees Or (Not hadTopAxes) Then
                        tx = rx: ty = ry: tz = rz
                        gotRight = True
                        If Abs(rudx) + Abs(rudy) + Abs(rudz) > 0.1 Then
                            lx = rudx: ly = rudy: lz = rudz
                        End If
                        ' Rebuild Width orthogonal to Length & Thickness.
                        wx = ly * tz - lz * ty
                        wy = lz * tx - lx * tz
                        wz = lx * ty - ly * tx
                        NormalizeAxis3 wx, wy, wz
                        If gotFront Then
                            If AbsDotAxis3(fx, fy, fz, wx, wy, wz) < AbsDotAxis3(fx, fy, fz, -wx, -wy, -wz) Then
                                wx = -wx: wy = -wy: wz = -wz
                            End If
                            If AbsDotAxis3(fx, fy, fz, wx, wy, wz) < 0.5 Then
                                wx = fx: wy = fy: wz = fz
                            End If
                        ElseIf hadTopAxes Then
                            ' Prefer original TOP Width direction/sign.
                            If AbsDotAxis3(topWx, topWy, topWz, wx, wy, wz) < AbsDotAxis3(topWx, topWy, topWz, -wx, -wy, -wz) Then
                                wx = -wx: wy = -wy: wz = -wz
                            End If
                        End If
                        gotTop = True
                    Else
                        ' Keep TOP axes; optionally flip thickness sign toward RIGHT X.
                        tx = topTx: ty = topTy: tz = topTz
                        lx = topLx: ly = topLy: lz = topLz
                        wx = topWx: wy = topWy: wz = topWz
                        If AbsDotAxis3(rx, ry, rz, tx, ty, tz) < AbsDotAxis3(rx, ry, rz, -tx, -ty, -tz) Then
                            tx = -tx: ty = -ty: tz = -tz
                        End If
                    End If
                End If
            End If
        End If
    End If

    On Error Resume Next
    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    If Err.Number <> 0 Then Err.Clear: model.ShowNamedView2 "*Top", 5
    On Error GoTo eh

    If Not gotTop Then
        LogLine "CMS view frame: could not read TOP/RIGHT orientations; dims stay sort-based until frame is ready."
        Exit Function
    End If

    ' Thin-plate sanity (TCP/BCP): Thickness must be the smallest extent. If the
    ' frame mapped T↔L (symptom T≈18 L≈1.4), swap Length and Thickness axes.
    If gIdxTCP > 0 Then
        If parts(gIdxTCP).BoxDx > 0 And parts(gIdxTCP).BoxDy > 0 And parts(gIdxTCP).BoxDz > 0 Then
            Dim chkL As Double, chkW As Double, chkT As Double
            gCmsLenAxisX = lx: gCmsLenAxisY = ly: gCmsLenAxisZ = lz
            gCmsWidAxisX = wx: gCmsWidAxisY = wy: gCmsWidAxisZ = wz
            gCmsThkAxisX = tx: gCmsThkAxisY = ty: gCmsThkAxisZ = tz
            gCmsViewFrameReady = True
            AssignLengthWidthThicknessFromAxes parts(gIdxTCP).BoxDx, parts(gIdxTCP).BoxDy, parts(gIdxTCP).BoxDz, chkL, chkW, chkT
            If chkT > chkL + 0.05 And chkT > chkW + 0.05 Then
                LogLine "CMS view frame: TCP T=" & FormatNumberForCsv(chkT) & _
                        " looks like Length — swapping L/T axes."
                Dim sx As Double, sy As Double, sz As Double
                sx = lx: sy = ly: sz = lz
                lx = tx: ly = ty: lz = tz
                tx = sx: ty = sy: tz = sz
            End If
            gCmsViewFrameReady = False
        End If
    End If

    gCmsLenAxisX = lx: gCmsLenAxisY = ly: gCmsLenAxisZ = lz
    gCmsWidAxisX = wx: gCmsWidAxisY = wy: gCmsWidAxisZ = wz
    gCmsThkAxisX = tx: gCmsThkAxisY = ty: gCmsThkAxisZ = tz
    gCmsViewFrameReady = True
    CaptureCmsViewFrameFromModel = True
    LogLine "CMS view frame L/W/T axes locked from DXF labels (TOP X=W, TOP Y=L, RIGHT X=T, RIGHT Y=L):" & _
            " L=[" & FormatNumberForCsv(lx) & "," & FormatNumberForCsv(ly) & "," & FormatNumberForCsv(lz) & "]" & _
            " W=[" & FormatNumberForCsv(wx) & "," & FormatNumberForCsv(wy) & "," & FormatNumberForCsv(wz) & "]" & _
            " T=[" & FormatNumberForCsv(tx) & "," & FormatNumberForCsv(ty) & "," & FormatNumberForCsv(tz) & "]" & _
            " gotRight=" & CStr(gotRight)
    Exit Function
eh:
    LogLine "CaptureCmsViewFrameFromModel error: " & Err.Description
    ResetCmsViewFrame
    CaptureCmsViewFrameFromModel = False
End Function

' Map assembly-axis box extents (dx,dy,dz along model X/Y/Z) onto CMS L/W/T axes.
Private Sub AssignLengthWidthThicknessFromAxes(ByVal dx As Double, ByVal dy As Double, ByVal dz As Double, _
                                               ByRef l As Double, ByRef w As Double, ByRef t As Double)
    If Not gCmsViewFrameReady Then
        SortThreeDimensions dx, dy, dz, l, w, t
        Exit Sub
    End If

    ' Extent along an axis ~ |axis · modelAxis| * size on that model axis, summed.
    l = Abs(gCmsLenAxisX) * dx + Abs(gCmsLenAxisY) * dy + Abs(gCmsLenAxisZ) * dz
    w = Abs(gCmsWidAxisX) * dx + Abs(gCmsWidAxisY) * dy + Abs(gCmsWidAxisZ) * dz
    t = Abs(gCmsThkAxisX) * dx + Abs(gCmsThkAxisY) * dy + Abs(gCmsThkAxisZ) * dz

    If l <= 0 Or w <= 0 Or t <= 0 Then
        SortThreeDimensions dx, dy, dz, l, w, t
    End If
End Sub

Private Sub ApplyCmsViewDimsToAllParts()
    Dim i As Long
    Dim l As Double, w As Double, t As Double
    Dim n As Long
    If PartCount < 1 Then Exit Sub
    If Not gCmsViewFrameReady Then
        LogLine "CMS view dims: frame not ready — leaving sort-based L/W/T."
        Exit Sub
    End If
    n = 0
    For i = 1 To PartCount
        If parts(i).BoxDx > 0 And parts(i).BoxDy > 0 And parts(i).BoxDz > 0 Then
            AssignLengthWidthThicknessFromAxes parts(i).BoxDx, parts(i).BoxDy, parts(i).BoxDz, l, w, t
            parts(i).Length = Round(l, DIM_DECIMALS)
            parts(i).Width = Round(w, DIM_DECIMALS)
            parts(i).Thickness = Round(t, DIM_DECIMALS)
            parts(i).BBoxVolume = l * w * t
            n = n + 1
        End If
    Next i
    LogLine "CMS view dims applied to " & n & " parts (TOP: W x L, RIGHT: T x L, FRONT: W)."
End Sub

Private Function TryGetCmsOrientedAssemblyDims(ByVal model As Object, _
                                               ByRef l As Double, ByRef w As Double, ByRef t As Double) As Boolean
On Error GoTo nope
    TryGetCmsOrientedAssemblyDims = False
    l = 0#: w = 0#: t = 0#
    If model Is Nothing Then Exit Function
    Dim dx As Double, dy As Double, dz As Double
    If TryGetModelDocBoxDimsInches(model, dx, dy, dz) = False Then Exit Function
    If Not gCmsViewFrameReady Then
        If Not CaptureCmsViewFrameFromModel(model) Then
            SortThreeDimensions dx, dy, dz, l, w, t
            TryGetCmsOrientedAssemblyDims = (l > 0 And w > 0 And t > 0)
            Exit Function
        End If
    End If
    AssignLengthWidthThicknessFromAxes dx, dy, dz, l, w, t
    TryGetCmsOrientedAssemblyDims = (l > 0 And w > 0 And t > 0)
    Exit Function
nope:
    TryGetCmsOrientedAssemblyDims = False
End Function

' ============================================================
' GENERAL HELPERS
' ============================================================
Private Function GetFolderLeafName(ByVal p As String) As String
On Error Resume Next
    Do While Len(p) > 0 And Right(p, 1) = "\"
        p = Left(p, Len(p) - 1)
    Loop
    Dim n As Long
    n = InStrRev(p, "\")
    If n > 0 Then GetFolderLeafName = Mid(p, n + 1) Else GetFolderLeafName = p
End Function

Private Function GetFileBaseName(ByVal path As String) As String
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileBaseName = fso.GetBaseName(path)
End Function

Private Function GetFileExtension(ByVal path As String) As String
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileExtension = fso.GetExtensionName(path)
End Function

Private Function CleanFileName(ByVal s As String) As String
    s = Trim(s)
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, Chr(34), "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    CleanFileName = Trim(s)
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) = False Then fso.CreateFolder folderPath
End Sub

Private Sub EnsureFolderDeep(ByVal folderPath As String)
On Error Resume Next
    If folderPath = "" Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If parent <> "" And fso.FolderExists(parent) = False Then EnsureFolderDeep parent
    If fso.FolderExists(folderPath) = False Then fso.CreateFolder folderPath
End Sub

Private Function GetUniqueFilePath(ByVal path As String) As String
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) = False Then
        GetUniqueFilePath = path
        Exit Function
    End If
    Dim folder As String, base As String, ext As String
    folder = fso.GetParentFolderName(path)
    base = fso.GetBaseName(path)
    ext = fso.GetExtensionName(path)
    Dim n As Long
    Dim candidate As String
    n = 2
    Do
        candidate = folder & "\" & base & "_" & n & "." & ext
        If fso.FileExists(candidate) = False Then
            GetUniqueFilePath = candidate
            Exit Function
        End If
        n = n + 1
    Loop While n < 1000
    GetUniqueFilePath = path
    Exit Function
ErrHandler:
    GetUniqueFilePath = path
End Function

Private Function PowerShellQuote(ByVal s As String) As String
    PowerShellQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Sub WaitMilliseconds(ByVal ms As Long)
On Error Resume Next
    If ms <= 0 Then Exit Sub
    Dim startT As Double
    Dim target As Double
    startT = Timer
    target = ms / 1000#
    Do
        DoEvents
        If Timer < startT Then Exit Do
        If (Timer - startT) >= target Then Exit Do
    Loop
End Sub

Private Function FormatNumberForCsv(ByVal v As Double) As String
    FormatNumberForCsv = Format(v, "0.000")
End Function

Private Function NormalizeText(ByVal s As String) As String
    s = UCase(Trim(s))
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeText = Trim(s)
End Function

Private Function NormalizeKey(ByVal s As String) As String
    s = UCase(Trim(s))
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, ".", "")
    NormalizeKey = s
End Function

' ============================================================
' LOGGING
' ============================================================
Private Sub WriteMacroLaunchStatus(ByVal statusText As String, Optional ByVal messageText As String = "")
On Error Resume Next

    EnsureFolderDeep LOCAL_WORKSPACE_ROOT

    Dim f As Integer

    f = FreeFile
    Open MACRO_STATUS_FILE For Output As #f
    Print #f, "Status=" & statusText
    Print #f, "Message=" & messageText
    Print #f, "Time=" & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "CurrentJobNumber=" & CurrentJobNumber
    Print #f, "CurrentJobFolder=" & CurrentJobFolder
    Print #f, "RunLogPath=" & RunLogPath
    Close #f

    Select Case UCase$(statusText)
        Case "STARTED"
            f = FreeFile
            Open MACRO_STARTED_FILE For Output As #f
            Print #f, "Started=" & Format(Now, "yyyy-mm-dd hh:nn:ss")
            Print #f, "Message=" & messageText
            Close #f

        Case "DONE", "COMPLETED"
            f = FreeFile
            Open MACRO_DONE_FILE For Output As #f
            Print #f, "Done=" & Format(Now, "yyyy-mm-dd hh:nn:ss")
            Print #f, "Message=" & messageText
            Close #f

        Case "ERROR", "FAILED"
            f = FreeFile
            Open MACRO_ERROR_FILE For Output As #f
            Print #f, "Error=" & Format(Now, "yyyy-mm-dd hh:nn:ss")
            Print #f, "Message=" & messageText
            Print #f, "Step=" & CurrentStepName
            Print #f, "RunLogPath=" & RunLogPath
            Close #f
    End Select
End Sub

Private Sub LogLine(ByVal msg As String)
On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Dim path As String
    path = RunLogPath
    If path = "" Then path = StartupLogPath
    If path = "" Then path = Environ$("USERPROFILE") & "\Desktop\CMS_Base_Export_STARTUP_Log.txt"
    Open path For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #f
End Sub

Private Sub LogStart(ByVal stepName As String)
    CurrentStepName = stepName
    StepStartTime = Now
    LogLine ">>> START: " & stepName
End Sub

Private Sub LogDone(ByVal stepName As String)
    LogLine "<<< DONE : " & stepName & " (" & DateDiff("s", StepStartTime, Now) & "s)"
End Sub

Private Sub LogErrorText(ByVal msg As String)
    LogLine "ERROR: " & msg
    LastJobFailReason = msg
End Sub

' ============================================================
' ============================================================
' POT-BLOCK ENGINE ADDED PROCEDURES
' (scan / BOM read / match / CSV reports / Excel fill)
' ============================================================
' ============================================================

' ============================================================
' SCAN CAD PARTS  (bounding box + mass + location)
' ============================================================
Private Sub ScanActiveSolidWorksDocument()
On Error GoTo ErrHandler
    If swModel Is Nothing Then Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    If swModel.GetType = swDocASSEMBLY Then
        Set swAssy = swModel
        On Error Resume Next
        swAssy.ResolveAllLightWeightComponents True
        On Error GoTo ErrHandler
        Dim vComps As Variant
        vComps = swModel.GetComponents(False)
        If IsEmpty(vComps) Then Exit Sub
        Dim i As Long
        For i = 0 To UBound(vComps)
            ProcessAssemblyComponent vComps(i)
        Next i
    ElseIf swModel.GetType = swDocPART Then
        ScanPartBodies swModel, GetFileBaseName(swModel.GetPathName)
    End If
    Exit Sub
ErrHandler:
    LogLine "ScanActiveSolidWorksDocument error: " & Err.Description
End Sub

Private Sub ProcessAssemblyComponent(ByVal swComp As Object)
On Error GoTo ErrHandler
    If swComp Is Nothing Then Exit Sub
    If swComp.IsSuppressed Then Exit Sub
    Dim swCompModel As Object
    Set swCompModel = swComp.GetModelDoc2
    If swCompModel Is Nothing And swComp.GetPathName <> "" Then
        Dim e As Long, w As Long
        Set swCompModel = swApp.OpenDoc6(swComp.GetPathName, swDocPART, _
            swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, swComp.ReferencedConfiguration, e, w)
    End If
    If swCompModel Is Nothing Then Exit Sub
    If swCompModel.GetType <> swDocPART Then Exit Sub
    Dim dx As Double, dy As Double, dz As Double
    ' MUST use assembly-space AABB (component.GetBox). Part-local bbox axes do not
    ' match CMS Top/Right/Front when holders/pots are rotated → inconsistent W/L/T.
    If TryGetComponentBoxDimsInches(swComp, dx, dy, dz) = False Then
        If GetPartBoundingBoxInches(swCompModel, dx, dy, dz) = False Then Exit Sub
    End If
    Dim massV As Double
    massV = GetModelMassOrVolumeValue(swCompModel)
    Dim cx As Double, cy As Double, cz As Double, hasC As Boolean
    hasC = TryGetComponentCenterInches(swComp, cx, cy, cz)
    AddCadPart swComp.Name2, swComp.GetPathName, swComp.ReferencedConfiguration, "", _
               dx, dy, dz, massV, hasC, cx, cy, cz, False
    Exit Sub
ErrHandler:
    LogLine "ProcessAssemblyComponent error: " & Err.Description
End Sub

Private Sub ScanPartBodies(ByVal partModel As Object, ByVal baseName As String)
On Error GoTo ErrHandler
    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Sub
    Dim i As Long, swBody As Object, vBox As Variant
    Dim dx As Double, dy As Double, dz As Double, massV As Double
    Dim cx As Double, cy As Double, cz As Double
    For i = 0 To UBound(vBodies)
        Set swBody = vBodies(i)
        If Not swBody Is Nothing Then
            vBox = swBody.GetBodyBox
            If Not IsEmpty(vBox) Then
                BoxCornersToInches vBox, dx, dy, dz, cx, cy, cz
                massV = GetBodyMassOrVolumeValue(swBody)
                AddCadPart baseName & " [" & swBody.Name & "]", partModel.GetPathName, "", swBody.Name, _
                           dx, dy, dz, massV, True, cx, cy, cz, True
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    LogLine "ScanPartBodies error: " & Err.Description
End Sub

Private Sub AddCadPart(ByVal compName As String, ByVal filePath As String, ByVal configName As String, _
                       ByVal bodyName As String, ByVal dx As Double, ByVal dy As Double, ByVal dz As Double, _
                       ByVal massV As Double, ByVal hasC As Boolean, ByVal cx As Double, ByVal cy As Double, _
                       ByVal cz As Double, ByVal bodyOnly As Boolean)
    Dim l As Double, w As Double, t As Double
    ' Final sanity: if still absurdly large, assume inches were double-converted and undo *39.37
    If dx > MAX_SANE_MOLD_DIM_IN Or dy > MAX_SANE_MOLD_DIM_IN Or dz > MAX_SANE_MOLD_DIM_IN Then
        If (dx / INCHES_PER_METER) <= MAX_SANE_MOLD_DIM_IN And (dy / INCHES_PER_METER) <= MAX_SANE_MOLD_DIM_IN Then
            LogLine "Dim sanity: undoing meter scale for " & compName & " (was " & FormatNumberForCsv(dx) & "x" & FormatNumberForCsv(dy) & "x" & FormatNumberForCsv(dz) & ")"
            dx = dx / INCHES_PER_METER
            dy = dy / INCHES_PER_METER
            dz = dz / INCHES_PER_METER
            cx = cx / INCHES_PER_METER
            cy = cy / INCHES_PER_METER
            cz = cz / INCHES_PER_METER
        End If
    End If
    ' Keep raw assembly-axis extents. L/W/T are assigned from the CMS Top/Right/Front
    ' view frame after orientation (not by sorting XYZ largest-first forever).
    AssignLengthWidthThicknessFromAxes dx, dy, dz, l, w, t
    If l * w * t <= 0 Then Exit Sub
    ' Skip hardware-scale noise and still-impossible mold sizes
    If l > MAX_SANE_MOLD_DIM_IN Then
        LogLine "Skipping part with impossible size: " & compName & " L=" & FormatNumberForCsv(l)
        Exit Sub
    End If
    PartCount = PartCount + 1
    ReDim Preserve parts(1 To PartCount)
    parts(PartCount).componentName = compName
    parts(PartCount).cleanName = NormalizeText(compName)
    parts(PartCount).filePath = filePath
    parts(PartCount).configName = configName
    parts(PartCount).bodyName = bodyName
    parts(PartCount).Quantity = 1
    parts(PartCount).BoxDx = Round(dx, DIM_DECIMALS)
    parts(PartCount).BoxDy = Round(dy, DIM_DECIMALS)
    parts(PartCount).BoxDz = Round(dz, DIM_DECIMALS)
    parts(PartCount).Length = Round(l, DIM_DECIMALS)
    parts(PartCount).Width = Round(w, DIM_DECIMALS)
    parts(PartCount).Thickness = Round(t, DIM_DECIMALS)
    parts(PartCount).BBoxVolume = l * w * t
    parts(PartCount).massValue = massV
    parts(PartCount).hasAsmCenter = hasC
    parts(PartCount).AsmCenterX = cx
    parts(PartCount).AsmCenterY = cy
    parts(PartCount).AsmCenterZ = cz
    parts(PartCount).UsedForBomMatch = False
    parts(PartCount).isBodyOnly = bodyOnly
End Sub

Private Function GetModelMassOrVolumeValue(ByVal model As Object) As Double
On Error GoTo ErrHandler
    Dim mp As Object
    Set mp = model.Extension.CreateMassProperty
    If mp Is Nothing Then Exit Function
    Dim m As Double
    m = mp.Mass
    If m > 0 Then
        GetModelMassOrVolumeValue = m * 2.20462
    Else
        GetModelMassOrVolumeValue = mp.Volume * CUIN_PER_CUBIC_METER
    End If
    Exit Function
ErrHandler:
    GetModelMassOrVolumeValue = 0#
End Function

Private Function GetBodyMassOrVolumeValue(ByVal swBody As Object) As Double
On Error GoTo ErrHandler
    Dim vProps As Variant
    vProps = swBody.GetMassProperties(1#)
    If IsArray(vProps) Then
        If UBound(vProps) >= 3 Then GetBodyMassOrVolumeValue = CDbl(vProps(3)) * CUIN_PER_CUBIC_METER
    End If
    Exit Function
ErrHandler:
    GetBodyMassOrVolumeValue = 0#
End Function

Private Function TryGetComponentCenterInches(ByVal swComp As Object, ByRef cx As Double, ByRef cy As Double, ByRef cz As Double) As Boolean
On Error GoTo ErrHandler
    Dim vBox As Variant
    Dim dxT As Double, dyT As Double, dzT As Double
    vBox = swComp.GetBox(False, False)
    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function
    BoxCornersToInches vBox, dxT, dyT, dzT, cx, cy, cz
    TryGetComponentCenterInches = True
    Exit Function
ErrHandler:
    TryGetComponentCenterInches = False
End Function

' Assembly-axis extents of a component (transformed into the assembly).
Private Function TryGetComponentBoxDimsInches(ByVal swComp As Object, _
                                              ByRef dx As Double, ByRef dy As Double, ByRef dz As Double) As Boolean
On Error GoTo ErrHandler
    TryGetComponentBoxDimsInches = False
    dx = 0#: dy = 0#: dz = 0#
    If swComp Is Nothing Then Exit Function
    Dim vBox As Variant
    vBox = swComp.GetBox(False, False)
    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function
    Dim cx As Double, cy As Double, cz As Double
    BoxCornersToInches vBox, dx, dy, dz, cx, cy, cz
    TryGetComponentBoxDimsInches = (dx > 0.001 And dy > 0.001 And dz > 0.001)
    Exit Function
ErrHandler:
    TryGetComponentBoxDimsInches = False
End Function

Private Sub SortPartsByVolumeDescending()
    If PartCount < 2 Then Exit Sub
    Dim i As Long, j As Long
    Dim tmp As PartInfo
    For i = 1 To PartCount - 1
        For j = i + 1 To PartCount
            If parts(j).BBoxVolume > parts(i).BBoxVolume Then
                tmp = parts(i): parts(i) = parts(j): parts(j) = tmp
            End If
        Next j
    Next i
End Sub

' ============================================================
' CSV REPORTS
' ============================================================
Private Sub WritePartDimensionCsv(ByVal csvPath As String)
On Error GoTo ErrHandler
    Dim p As String
    p = GetWritableCsvPath(csvPath)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, "Index,Component,Qty,Thickness,Width,Length,BBoxVolume_cuin,Mass_or_Vol,CenterX,CenterY,CenterZ"
    Dim i As Long
    For i = 1 To PartCount
        Print #f, i & "," & CsvText(parts(i).componentName) & "," & parts(i).Quantity & "," & _
            FormatNumberForCsv(parts(i).Thickness) & "," & FormatNumberForCsv(parts(i).Width) & "," & _
            FormatNumberForCsv(parts(i).Length) & "," & FormatNumberForCsv(parts(i).BBoxVolume) & "," & _
            FormatNumberForCsv(parts(i).massValue) & "," & FormatNumberForCsv(parts(i).AsmCenterX) & "," & _
            FormatNumberForCsv(parts(i).AsmCenterY) & "," & FormatNumberForCsv(parts(i).AsmCenterZ)
    Next i
    Close #f
    LogLine "Wrote CAD dimensions CSV: " & p
    Exit Sub
ErrHandler:
    LogLine "WritePartDimensionCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteExportCheckCsv(ByVal csvPath As String)
On Error GoTo ErrHandler
    Dim p As String
    p = GetWritableCsvPath(csvPath)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, "QuoteName,Qty,Material,Status,CAD_Thickness,CAD_Width,CAD_Length,BOM_Thickness,BOM_Width,BOM_Length,CAD_Component"
    Dim i As Long
    Dim cadName As String
    For i = 1 To ExportCount
        cadName = ""
        If ExportRows(i).HasCad Then cadName = parts(ExportRows(i).CadPartIndex).componentName
        Print #f, CsvText(ExportRows(i).quoteName) & "," & ExportRows(i).Quantity & "," & _
            CsvText(ExportRows(i).material) & "," & CsvText(ExportRows(i).Status) & "," & _
            FormatNumberForCsv(ExportRows(i).Thickness) & "," & FormatNumberForCsv(ExportRows(i).Width) & "," & _
            FormatNumberForCsv(ExportRows(i).Length) & "," & FormatNumberForCsv(ExportRows(i).BomThickness) & "," & _
            FormatNumberForCsv(ExportRows(i).BomWidth) & "," & FormatNumberForCsv(ExportRows(i).BomLength) & "," & _
            CsvText(cadName)
    Next i
    Close #f
    LogLine "Wrote BOM match report CSV: " & p
    Exit Sub
ErrHandler:
    LogLine "WriteExportCheckCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Function CsvText(ByVal s As String) As String
    s = Replace(s, Chr(34), "'")
    If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CsvText = Chr(34) & s & Chr(34)
    Else
        CsvText = s
    End If
End Function

Private Function GetWritableCsvPath(ByVal csvPath As String) As String
On Error GoTo ErrHandler
    Dim f As Integer
    f = FreeFile
    Open csvPath For Output As #f
    Close #f
    GetWritableCsvPath = csvPath
    Exit Function
ErrHandler:
    On Error Resume Next
    Close #f
    GetWritableCsvPath = AppendBeforeExtension(csvPath, "_" & Format(Now, "hhnnss"))
End Function

Private Function BmsRoleForCadIndex(ByVal idx As Long) As String
    If idx = gIdxTCP Then BmsRoleForCadIndex = "TCP": Exit Function
    If idx = gIdxBCP Then BmsRoleForCadIndex = "BCP": Exit Function
    If idx = gIdxIDH Then BmsRoleForCadIndex = "ID HOLDER": Exit Function
    If idx = gIdxODH Then BmsRoleForCadIndex = "OD HOLDER": Exit Function
    If idx = gIdxIDP Then BmsRoleForCadIndex = "ID POT": Exit Function
    If idx = gIdxODP Then BmsRoleForCadIndex = "OD POT": Exit Function
End Function

Private Sub WriteAllCadComponentsDebugCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    Dim p As String
    p = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open p For Output As #f

    Print #f, "Index,ComponentName,CleanName,StdRole,BmsRole,Qty,Thickness,Width,Length,BoxDx,BoxDy,BoxDz,BBoxVolume,MassOrVol,CenterX,CenterY,CenterZ,HasCenter,FilePath,ConfigName,BodyName,IsBodyOnly,UsedForBomMatch,FileExists"

    Dim i As Long
    Dim existsText As String

    For i = 1 To PartCount

        existsText = "FALSE"
        If parts(i).filePath <> "" Then
            If Dir(parts(i).filePath) <> "" Then existsText = "TRUE"
        End If

        Print #f, _
            i & "," & _
            CsvText(parts(i).componentName) & "," & _
            CsvText(parts(i).cleanName) & "," & _
            CsvText(StdCadRole(i)) & "," & _
            CsvText(BmsRoleForCadIndex(i)) & "," & _
            parts(i).Quantity & "," & _
            FormatNumberForCsv(parts(i).Thickness) & "," & _
            FormatNumberForCsv(parts(i).Width) & "," & _
            FormatNumberForCsv(parts(i).Length) & "," & _
            FormatNumberForCsv(parts(i).BoxDx) & "," & _
            FormatNumberForCsv(parts(i).BoxDy) & "," & _
            FormatNumberForCsv(parts(i).BoxDz) & "," & _
            FormatNumberForCsv(parts(i).BBoxVolume) & "," & _
            FormatNumberForCsv(parts(i).massValue) & "," & _
            FormatNumberForCsv(parts(i).AsmCenterX) & "," & _
            FormatNumberForCsv(parts(i).AsmCenterY) & "," & _
            FormatNumberForCsv(parts(i).AsmCenterZ) & "," & _
            CsvText(CStr(parts(i).hasAsmCenter)) & "," & _
            CsvText(parts(i).filePath) & "," & _
            CsvText(parts(i).configName) & "," & _
            CsvText(parts(i).bodyName) & "," & _
            CsvText(CStr(parts(i).isBodyOnly)) & "," & _
            CsvText(CStr(parts(i).UsedForBomMatch)) & "," & _
            CsvText(existsText)
    Next i

    Close #f

    LogLine "Wrote all-CAD components debug CSV: " & p
    Exit Sub

ErrHandler:
    LogLine "WriteAllCadComponentsDebugCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteJobFileInventoryCsv(ByVal rootFolder As String, ByVal csvPath As String)
On Error GoTo ErrHandler

    If rootFolder = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rootFolder) Then Exit Sub

    Dim p As String
    p = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open p For Output As #f
    Print #f, "FullPath,Folder,FileName,Extension,SizeBytes,DateModified"

    WriteJobFileInventoryFolderRows fso.GetFolder(rootFolder), f

    Close #f

    LogLine "Wrote job file inventory CSV: " & p
    Exit Sub

ErrHandler:
    LogLine "WriteJobFileInventoryCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteJobFileInventoryFolderRows(ByVal folder As Object, ByVal fileNum As Integer)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim file As Object
    For Each file In folder.Files
        Print #fileNum, _
            CsvText(file.path) & "," & _
            CsvText(folder.path) & "," & _
            CsvText(file.Name) & "," & _
            CsvText(fso.GetExtensionName(file.path)) & "," & _
            CStr(file.Size) & "," & _
            CsvText(CStr(file.DateLastModified))
    Next file

    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        WriteJobFileInventoryFolderRows subFolder, fileNum
    Next subFolder
End Sub

Private Sub WriteStandardQuoteRowsDebugCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    Dim p As String
    p = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open p For Output As #f

    Print #f, "StdRow,Name,Qty,Thickness,Width,Length,Grade,QuoteRow,CadIndex,CadComponent,StdRole,CenterX,CenterY,CenterZ"

    Dim i As Long
    Dim ci As Long
    Dim comp As String
    Dim role As String
    Dim cx As Double, cy As Double, cz As Double

    For i = 1 To StdCount
        ci = 0
        comp = ""
        role = ""
        cx = 0#: cy = 0#: cz = 0#

        If i <= UBound(StdCadIndex) Then ci = StdCadIndex(i)

        If ci >= 1 And ci <= PartCount Then
            comp = parts(ci).componentName
            role = StdCadRole(ci)
            cx = parts(ci).AsmCenterX
            cy = parts(ci).AsmCenterY
            cz = parts(ci).AsmCenterZ
        End If

        Print #f, _
            i & "," & _
            CsvText(stdName(i)) & "," & _
            StdQty(i) & "," & _
            FormatNumberForCsv(StdT(i)) & "," & _
            FormatNumberForCsv(StdW(i)) & "," & _
            FormatNumberForCsv(StdL(i)) & "," & _
            CsvText(StdGrade(i)) & "," & _
            StdQuoteRow(i) & "," & _
            ci & "," & _
            CsvText(comp) & "," & _
            CsvText(role) & "," & _
            FormatNumberForCsv(cx) & "," & _
            FormatNumberForCsv(cy) & "," & _
            FormatNumberForCsv(cz)
    Next i

    Close #f

    LogLine "Wrote standard quote rows debug CSV: " & p
    Exit Sub

ErrHandler:
    LogLine "WriteStandardQuoteRowsDebugCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub



Private Function AppendBeforeExtension(ByVal path As String, ByVal suffix As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(path, ".")
    If dotPos > 0 Then
        AppendBeforeExtension = Left(path, dotPos - 1) & suffix & Mid(path, dotPos)
    Else
        AppendBeforeExtension = path & suffix
    End If
End Function

' ============================================================
' FIND BOM FILE
' ============================================================
Private Function FindCustomerBomFile(ByVal jobFolder As String) As String
On Error GoTo ErrHandler
    FindCustomerBomFile = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(jobFolder) = False Then Exit Function
    Dim cands As Collection
    Set cands = New Collection
    SearchBomFilesRecursive fso.GetFolder(jobFolder), cands
    Dim bestPath As String, bestScore As Long
    bestPath = "": bestScore = -1
    Dim i As Long, p As String, nm As String, ext As String, score As Long
    For i = 1 To cands.Count
        p = CStr(cands(i))
        nm = UCase(fso.GetFileName(p))
        ext = LCase(fso.GetExtensionName(p))
        score = 0
        If InStr(nm, "BOM") > 0 Then score = score + 50
        If InStr(nm, "RFQ") > 0 Then score = score + 12
        If InStr(nm, "HTE") > 0 Then score = score + 12     ' customer (Howmet/Tempcraft) BOM prefix
        If InStr(nm, "BASE") > 0 And InStr(nm, "BOM") > 0 Then score = score + 20
        If ext = "xlsm" Then score = score + 10             ' prefer live Tempcraft .xlsm over PDF export
        If ext = "xlsx" Or ext = "xls" Then score = score + 8
        If ext = "pdf" Then score = score + 4               ' PDF last — layout OCR often mixes Stock Weight into dims
        If CurrentJobNumber <> "" Then
            If InStr(nm, UCase(CurrentJobNumber)) > 0 Then score = score + 5
        End If
        If CustomerJobNumber <> "" Then
            If InStr(nm, UCase(CustomerJobNumber)) > 0 Then score = score + 15
        End If
        If InStr(nm, "QUOTE") > 0 Or InStr(nm, "PROPOSAL") > 0 Then score = score - 40
        If InStr(nm, "DWG") > 0 And InStr(nm, "BOM") = 0 Then score = score - 30
        If score > bestScore Then bestScore = score: bestPath = p
    Next i
    FindCustomerBomFile = bestPath
    Exit Function
ErrHandler:
    LogLine "FindCustomerBomFile error: " & Err.Description
    FindCustomerBomFile = ""
End Function

Private Sub SearchBomFilesRecursive(ByVal folder As Object, ByVal cands As Collection)
On Error Resume Next
    If UCase(folder.Name) = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object, nm As String, ext As String
    For Each file In folder.Files
        nm = UCase(file.Name)
        ext = LCase(fso.GetExtensionName(file.path))
        If Left(file.Name, 2) <> "~$" Then
            If (ext = "pdf" Or ext = "xls" Or ext = "xlsx" Or ext = "xlsm") Then
                ' Skip our own quote/proposal/output files - only customer BOMs.
                If InStr(nm, "CUSTOM QUOTE") = 0 _
                   And InStr(nm, "QUOTE_STEEL") = 0 _
                   And InStr(nm, "STEEL_SHEET") = 0 _
                   And InStr(nm, "PURCHASED COMPONENT") = 0 _
                   And InStr(nm, "XT_EXPORT") = 0 _
                   And InStr(nm, "EXPORT_LOG") = 0 _
                   And InStr(nm, "BOM_MATCH") = 0 Then
                    cands.Add file.path
                End If
            End If
        End If
    Next file
    Dim sub1 As Object
    For Each sub1 In folder.SubFolders
        SearchBomFilesRecursive sub1, cands
    Next sub1
End Sub

' ============================================================
' READ EXCEL BOM
' ============================================================
Private Sub ReadCustomerBom(ByVal bomPath As String)
On Error GoTo ErrHandler
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    Set xlWb = xlApp.Workbooks.Open(bomPath, False, True)
    xlApp.Calculation = -4135  ' manual calc - valid only after a workbook is open
    Dim handled As Boolean
    handled = False
    If TURBO_READ_ONLY_BOM_SHEET Then
        On Error Resume Next
        Set xlWs = xlWb.Worksheets(TURBO_BOM_SHEET_NAME)
        On Error GoTo ErrHandler
        If Not xlWs Is Nothing Then
            ReadBomWorksheetFastArray xlWs
            handled = True
        End If
    End If
    If Not handled Then
        Dim ws As Object
        For Each ws In xlWb.Worksheets
            If IsLikelyBomWorksheet(ws) And ShouldSkipBomWorksheet(ws) = False Then
                ReadBomWorksheetFastArray ws
                If BomCount > 0 Then Exit For
            End If
        Next ws
    End If
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    Exit Sub
ErrHandler:
    LogLine "ReadCustomerBom error: " & Err.Description
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

Private Sub ReadBomWorksheetFastArray(ByVal xlWs As Object)
On Error GoTo ErrHandler
    Dim data As Variant
    data = xlWs.usedRange.value
    If IsEmpty(data) Then Exit Sub
    If IsArray(data) = False Then Exit Sub
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(data, 1): rHi = UBound(data, 1)
    cLo = LBound(data, 2): cHi = UBound(data, 2)
    Dim headerRow As Long, descCol As Long, qtyCol As Long, matCol As Long
    Dim thkCol As Long, widCol As Long, lenCol As Long
    Dim sheetIsTempcraft As Boolean
    headerRow = 0: descCol = 0
    sheetIsTempcraft = False
    Dim rEnd As Long
    rEnd = rLo + BOM_HEADER_SEARCH_MAX_ROWS
    If rEnd > rHi Then rEnd = rHi
    Dim r As Long
    For r = rLo To rEnd
        descCol = FindBomHeaderLikeInArrayRow(data, r, cLo, cHi)
        If descCol > 0 Then headerRow = r: Exit For
    Next r
    If headerRow = 0 Then Exit Sub
    qtyCol = FindBomQtyColumnInArrayRow(data, headerRow, cLo, cHi)
    matCol = FindBomMaterialColumnInArrayRow(data, headerRow, cLo, cHi)
    FindBomDimensionColumnsInArrayRow data, headerRow, cLo, cHi, thkCol, widCol, lenCol, sheetIsTempcraft
    ' Part number column ("Mat'l Spec or Mfg. Part No.") and manufacturer column.
    Dim partCol As Long, manufCol As Long, hc As Long, ht As String
    partCol = 0: manufCol = 0
    Dim detCol As Long
    detCol = 0
    Dim typeCol As Long
    typeCol = 0
    For hc = cLo To cHi
        ht = UCase(GetArrayValue(data, headerRow, hc))
        If partCol = 0 Then
            If InStr(ht, "PART NO") > 0 Or InStr(ht, "PART #") > 0 Or InStr(ht, "MFG. PART") > 0 _
               Or InStr(ht, "PART NUMBER") > 0 Or InStr(ht, "MAT'L SPEC") > 0 Or InStr(ht, "MFG PART") > 0 Then partCol = hc
        End If
        If manufCol = 0 Then
            If InStr(ht, "MANUFACTURER") > 0 Or InStr(ht, "VENDOR") > 0 Or InStr(ht, "MFG") > 0 Then manufCol = hc
        End If
        If typeCol = 0 Then
            If ht = "TYPE" Or InStr(ht, "TYPE") > 0 Then typeCol = hc
        End If
        If detCol = 0 Then
            If InStr(ht, "DET NO") > 0 Or InStr(ht, "DET #") > 0 Or InStr(ht, "DET.") > 0 _
               Or InStr(ht, "ITEM NO") > 0 Or ht = "DET NO." Or ht = "NO." Or ht = "DET" Then detCol = hc
        End If
    Next hc
    Dim blanks As Long
    blanks = 0
    Dim desc As String, mat As String, qty As Long
    Dim tt As Double, ww As Double, ll As Double, hasD As Boolean
    For r = headerRow + 1 To rHi
        desc = Trim(GetArrayValue(data, r, descCol))
        If desc = "" Then
            blanks = blanks + 1
            If blanks >= STOP_BOM_READ_AFTER_BLANK_ROWS Then Exit For
        Else
            blanks = 0
            qty = 1
            If qtyCol > 0 Then
                If IsNumeric(GetArrayValue(data, r, qtyCol)) Then qty = CLng(Val(GetArrayValue(data, r, qtyCol)))
            End If
            If qty < 1 Then qty = 1
            mat = ""
            If matCol > 0 Then mat = Trim(GetArrayValue(data, r, matCol))
            tt = 0#: ww = 0#: ll = 0#: hasD = False
            Dim rowIsTempcraft As Boolean
            rowIsTempcraft = sheetIsTempcraft
            If thkCol > 0 Then tt = Val(GetArrayValue(data, r, thkCol))
            If widCol > 0 Then ww = Val(GetArrayValue(data, r, widCol))
            If lenCol > 0 Then ll = Val(GetArrayValue(data, r, lenCol))
            ' Tempcraft Lth/Wth/Hgt are finished sizes in FILE order — do not
            ' SortThreeDimensions here. MapTempcraftBomDimsToCmsSteel remaps by
            ' plate role at steel-fill time (holders/pots are not L≥W≥T).
            If tt > 0 And ww > 0 And ll > 0 Then
                Dim srtL As Double, srtW As Double, srtT As Double
                SortThreeDimensions tt, ww, ll, srtL, srtW, srtT
                ' Reject absurd "length" that is really stock weight (lbs >> plate size).
                If BomDimsLookLikeStockWeight(srtT, srtW, srtL) Then
                    tt = 0#: ww = 0#: ll = 0#
                End If
                ' else keep tt/ww/ll in column order (Tempcraft Lth/Wth/Hgt when flagged)
            End If
            If Not (tt > 0 And ww > 0 And ll > 0) Then
                ' No clean dimension columns: look for a combined size cell
                ' anywhere in the row, e.g. "1.375 X 15.875 X 18".
                Dim cc2 As Long, cellTxt As String, snums() As Double, sn As Long
                Dim sa As Double, sb As Double, scc As Double, sl As Double, sW As Double, sT As Double
                For cc2 = cLo To cHi
                    cellTxt = GetArrayValue(data, r, cc2)
                    If InStr(UCase(cellTxt), "LBS") > 0 Or InStr(UCase(cellTxt), "WEIGHT") > 0 Then GoTo NextBomSizeCell
                    If InStr(cellTxt, ".") > 0 Then
                        sn = ExtractDecimalNumbers(cellTxt, snums)
                        If sn >= 3 Then
                            ' Prefer first three decimals (finished sizes), not three largest
                            ' (which pulls Stock Weight into Length).
                            If PickThreeFinishedSizeDims(snums, sn, sa, sb, scc) Then
                                SortThreeDimensions sa, sb, scc, sl, sW, sT
                                If Not BomDimsLookLikeStockWeight(sT, sW, sl) Then
                                    ' Keep Tempcraft file order (not sorted L/W/T).
                                    tt = sa: ww = sb: ll = scc
                                    rowIsTempcraft = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
NextBomSizeCell:
                Next cc2
            End If
            ' Still nothing: parse fractional dims embedded in the description,
            ' e.g. "Top clamp plate, 1.375 x 9-7/8 x 11-7/8".
            If Not (tt > 0 And ww > 0 And ll > 0) Then
                Dim ftt As Double, fww As Double, fll As Double
                If ParseInchDimsFromText(desc, ftt, fww, fll) Then tt = ftt: ww = fww: ll = fll
            End If
            ' Material not in a labeled column: scan the row for a steel token
            ' (e.g. "#7 steel" sitting in an "Addt'l Comments" column).
            Dim matKnown As Boolean
            matKnown = False
            Select Case NormalizeSteelType(mat)
                Case "A36", "4140", "P20", "420SS", "H13", "6061": matKnown = True
            End Select
            If Not matKnown Then
                Dim mc As Long, mct As String
                For mc = cLo To cHi
                    mct = GetArrayValue(data, r, mc)
                    Select Case NormalizeSteelType(mct)
                        Case "A36", "4140", "P20", "420SS", "H13", "6061", "A2", "O1": mat = mct: Exit For
                    End Select
                Next mc
            End If
            If tt > 0 And ww > 0 And ll > 0 Then hasD = True
            Dim rowPart As String, rowManuf As String, rowDet As String, rowType As String
            rowPart = "": rowManuf = "": rowDet = "": rowType = ""
            If partCol > 0 Then rowPart = Trim(GetArrayValue(data, r, partCol))
            If manufCol > 0 Then rowManuf = Trim(GetArrayValue(data, r, manufCol))
            If detCol > 0 Then rowDet = Trim(GetArrayValue(data, r, detCol))
            If typeCol > 0 Then rowType = Trim(GetArrayValue(data, r, typeCol))
            On Error Resume Next
            AddBomRow desc, qty, mat, tt, ww, ll, hasD, rowPart, rowManuf, rowDet, rowType, rowIsTempcraft
            If Err.Number <> 0 Then
                LogLine "BOM row " & r & " skipped after parse error: " & Err.Description & " | " & desc
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
    Next r
    Exit Sub
ErrHandler:
    LogLine "ReadBomWorksheetFastArray error: " & Err.Description
End Sub

Private Function GetArrayValue(ByVal data As Variant, ByVal r As Long, ByVal c As Long) As String
On Error Resume Next
    If r < LBound(data, 1) Or r > UBound(data, 1) Then Exit Function
    If c < LBound(data, 2) Or c > UBound(data, 2) Then Exit Function
    Dim v As Variant
    v = data(r, c)
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    GetArrayValue = CStr(v)
End Function

Private Function FindBomHeaderLikeInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal cLo As Long, ByVal cHi As Long) As Long
    Dim c As Long, t As String
    For c = cLo To cHi
        t = NormalizeText(GetArrayValue(data, r, c))
        If t = "DESCRIPTION" Or t = "PART NAME" Or t = "PART DESCRIPTION" Or t = "ITEM DESCRIPTION" _
           Or t = "DESC" Or t = "PART" Or t = "NAME" Or t = "COMPONENT" Then
            FindBomHeaderLikeInArrayRow = c
            Exit Function
        End If
    Next c
    FindBomHeaderLikeInArrayRow = 0
End Function

Private Function FindBomQtyColumnInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal cLo As Long, ByVal cHi As Long) As Long
    Dim c As Long, t As String
    For c = cLo To cHi
        t = NormalizeText(GetArrayValue(data, r, c))
        If t = "QTY" Or t = "QUANTITY" Or t = "QTY." Or t = "QTY REQD" Or t = "QTY REQ" _
           Or t = "NO REQD" Or t = "NOREQD" Or t = "REQD" Or t = "NO REQUIRED" _
           Or InStr(t, "REQ") > 0 Then          ' "No. Req'd" -> "NO REQD"
            FindBomQtyColumnInArrayRow = c
            Exit Function
        End If
    Next c
End Function

Private Function FindBomMaterialColumnInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal cLo As Long, ByVal cHi As Long) As Long
    Dim c As Long, t As String
    For c = cLo To cHi
        t = NormalizeText(GetArrayValue(data, r, c))
        If t = "MATERIAL" Or t = "MATL" Or t = "STEEL" Or t = "STEEL TYPE" Or t = "MATERIAL TYPE" _
           Or InStr(t, "MATL SPEC") > 0 Or InStr(t, "MAT L SPEC") > 0 Or InStr(t, "MATERIAL SPEC") > 0 Then
            FindBomMaterialColumnInArrayRow = c
            Exit Function
        End If
    Next c
End Function

Private Sub FindBomDimensionColumnsInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal cLo As Long, ByVal cHi As Long, _
                                             ByRef thkCol As Long, ByRef widCol As Long, ByRef lenCol As Long, _
                                             ByRef usedTempcraftOrder As Boolean)
    ' Tempcraft / Howmet BOMs use Lth / Wth.O.D. / Hgt.I.D. — finished sizes in
    ' FILE order. Callers store them as tt=Lth, ww=Wth, ll=Hgt and remap later
    ' via MapTempcraftBomDimsToCmsSteel (do NOT SortThreeDimensions on store).
    thkCol = 0: widCol = 0: lenCol = 0
    usedTempcraftOrder = False
    Dim c As Long, t As String
    Dim lthCol As Long, wthCol As Long, hgtCol As Long
    lthCol = 0: wthCol = 0: hgtCol = 0
    For c = cLo To cHi
        t = NormalizeText(GetArrayValue(data, r, c))
        ' Never treat Stock Weight / Volume / Area as a dimension column.
        If InStr(t, "WEIGHT") > 0 Or InStr(t, "VOLUME") > 0 Or InStr(t, "AREA") > 0 Or InStr(t, "STOCK WT") > 0 Then GoTo NextBomDimCol
        If InStr(t, "ORACLE") > 0 Or InStr(t, "UOM") > 0 Then GoTo NextBomDimCol
        If lthCol = 0 And (InStr(t, "LTH") > 0 Or t = "LTH" Or t = "LTH IN") Then lthCol = c
        If wthCol = 0 And (InStr(t, "WTH") > 0 Or InStr(t, "O D") > 0 Or t = "WTH" Or t = "WTH IN") Then wthCol = c
        If hgtCol = 0 And (InStr(t, "HGT") > 0 Or InStr(t, "I D") > 0 Or t = "HGT" Or t = "HGT IN") Then hgtCol = c
        If thkCol = 0 And (t = "THICKNESS" Or t = "THICK" Or t = "THK" Or t = "T" Or InStr(t, "HEIGHT") > 0) Then
            If InStr(t, "HGT") = 0 And InStr(t, "I D") = 0 Then thkCol = c
        End If
        If widCol = 0 And (t = "WIDTH" Or t = "WIDE" Or t = "W") Then widCol = c
        If lenCol = 0 And (t = "LENGTH" Or t = "LONG" Or t = "LEN" Or t = "L") Then lenCol = c
NextBomDimCol:
    Next c
    ' Prefer Tempcraft Lth/Wth/Hgt: map into tt/ww/ll slots as Lth/Wth/Hgt order.
    If lthCol > 0 And wthCol > 0 And hgtCol > 0 Then
        thkCol = lthCol
        widCol = wthCol
        lenCol = hgtCol
        usedTempcraftOrder = True
    End If
End Sub

Private Function IsLikelyBomWorksheet(ByVal ws As Object) As Boolean
On Error Resume Next
    IsLikelyBomWorksheet = True
End Function

Private Function ShouldSkipBomWorksheet(ByVal ws As Object) As Boolean
On Error Resume Next
    Dim nm As String
    nm = NormalizeText(ws.Name)
    If InStr(nm, "QUOTE") > 0 Or InStr(nm, "INSTRUCT") > 0 Or InStr(nm, "NOTES") > 0 Then ShouldSkipBomWorksheet = True
End Function

' ============================================================
' READ PDF BOM  (via Poppler pdftotext)
' ============================================================
Private Sub ReadCustomerBomPdfUsingPdfToText(ByVal pdfPath As String)
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim exe As String
    exe = ResolvePdfToTextExe()
    LogLine "pdftotext exe: " & exe
    Dim txtPath As String
    txtPath = Environ$("TEMP") & "\CMS_BOM_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = Chr(34) & exe & Chr(34) & " -layout " & Chr(34) & pdfPath & Chr(34) & " " & Chr(34) & txtPath & Chr(34)
    sh.Run "cmd /c " & Chr(34) & cmd & Chr(34), 0, True
    If fso.FileExists(txtPath) = False Then
        LogLine "pdftotext produced no output (exe=" & exe & "). Is Poppler installed / on PATH?"
        Exit Sub
    End If
    ParseBomTextFromPdf ReadAllTextFile(txtPath)
    On Error Resume Next
    fso.DeleteFile txtPath
    Exit Sub
ErrHandler:
    LogLine "ReadCustomerBomPdfUsingPdfToText error: " & Err.Description
End Sub

' Find pdftotext.exe: the configured path, then Trust/Downloads/workspace, then any
' poppler* folder under Downloads, then fall back to PATH.
Private Function ResolvePdfToTextExe() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(PDFTOTEXT_EXE) Then ResolvePdfToTextExe = PDFTOTEXT_EXE: Exit Function
    Dim cands(4) As String, i As Long
    cands(0) = TRUSTED_FOLDER & "\pdftotext.exe"
    cands(1) = DOWNLOADS_FOLDER & "\pdftotext.exe"
    cands(2) = LOCAL_WORKSPACE_ROOT & "\pdftotext.exe"
    cands(3) = "C:\Users\lenovo\Downloads\New folder (17)\pdftotext.exe"
    cands(4) = TRUSTED_FOLDER & "\bin\pdftotext.exe"
    For i = 0 To 4
        If fso.FileExists(cands(i)) Then ResolvePdfToTextExe = cands(i): Exit Function
    Next i
    On Error Resume Next
    Dim sub1 As Object, p As String
    If fso.FolderExists(DOWNLOADS_FOLDER) Then
        For Each sub1 In fso.GetFolder(DOWNLOADS_FOLDER).SubFolders
            If InStr(LCase(sub1.Name), "poppler") > 0 Then
                p = sub1.path & "\Library\bin\pdftotext.exe"
                If fso.FileExists(p) Then ResolvePdfToTextExe = p: Exit Function
                p = sub1.path & "\bin\pdftotext.exe"
                If fso.FileExists(p) Then ResolvePdfToTextExe = p: Exit Function
            End If
        Next sub1
    End If
    On Error GoTo 0
    ResolvePdfToTextExe = "pdftotext"     ' rely on PATH as a last resort
End Function

Private Function ReadAllTextFile(ByVal path As String) As String
On Error GoTo ErrHandler
    Dim f As Integer
    f = FreeFile
    Open path For Input As #f
    If LOF(f) > 0 Then ReadAllTextFile = Input(LOF(f), #f)
    Close #f
    Exit Function
ErrHandler:
    On Error Resume Next
    Close #f
    ReadAllTextFile = ""
End Function

Private Sub ParseBomTextFromPdf(ByVal allText As String)
On Error GoTo ErrHandler
    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)
    Dim lines() As String
    lines = Split(allText, vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        ParseBomPdfTextLine lines(i)
    Next i
    Exit Sub
ErrHandler:
    LogLine "ParseBomTextFromPdf error: " & Err.Description
End Sub

Private Sub ParseBomPdfTextLine(ByVal lineText As String)
On Error Resume Next

    Dim raw As String
    raw = Trim(lineText)

    If raw = "" Then Exit Sub

    ' Purchase / hardware rows often do NOT have 3 decimal dimensions.
    ' Capture them before the material-line parser rejects them.
    If TryCapturePurchasedPdfLine(raw) Then Exit Sub

    ' Material / steel rows with dimensions.
    TryParseTempcraftBasePdfMaterialLine raw
End Sub

Private Function TryCapturePurchasedPdfLine(ByVal raw As String) As Boolean
On Error GoTo ErrHandler

    TryCapturePurchasedPdfLine = False

    Dim u As String
    u = UCase(raw)

    If InStr(u, " PURCHASE ") = 0 And InStr(u, " PURCHASE") = 0 Then Exit Function

    Dim desc As String
    Dim qty As Long
    Dim vendor As String
    Dim partNo As String
    Dim detNo As String

    desc = ExtractPdfPurchaseDescription(raw)
    If desc = "" Then Exit Function

    qty = ExtractPdfPurchaseQty(raw)
    If qty <= 0 Then qty = 1

    vendor = ExtractPdfPurchaseVendor(raw)
    partNo = ExtractPdfPurchasePartNo(raw)
    detNo = ExtractLeadingDetailNumber(raw)

    Dim nums() As Double
    Dim nCount As Long
    Dim tt As Double
    Dim ww As Double
    Dim ll As Double

    tt = 0#
    ww = 0#
    ll = 0#

    nCount = ExtractDecimalNumbers(raw, nums)

    ' Optional dimensions for insulation / spacers.
    If nCount >= 3 Then
        PickThreeFinishedSizeDims nums, nCount, tt, ww, ll
    End If

    If FILL_PURCHASED_COMPONENTS Then
        CapturePurchased desc, qty, "", tt, ww, ll, partNo, vendor, detNo, "Purchase"
    End If

    LogLine "PDF purchase captured: det=" & detNo & _
            " desc='" & desc & "'" & _
            " qty=" & CStr(qty) & _
            " vendor='" & vendor & "'" & _
            " part='" & partNo & "'"

    TryCapturePurchasedPdfLine = True
    Exit Function

ErrHandler:
    LogLine "TryCapturePurchasedPdfLine error: " & Err.Description & " | " & raw
    TryCapturePurchasedPdfLine = False
End Function

Private Function ExtractPdfPurchaseDescription(ByVal raw As String) As String
On Error GoTo ErrHandler

    Dim s As String
    s = RemoveLeadingItemNumber(raw)

    Dim p As Long
    p = InStr(1, UCase(s), " PURCHASE", vbTextCompare)

    If p > 1 Then
        ExtractPdfPurchaseDescription = ProperCaseText(Trim(Left(s, p - 1)))
    Else
        ExtractPdfPurchaseDescription = ProperCaseText(Trim(s))
    End If

    Exit Function

ErrHandler:
    ExtractPdfPurchaseDescription = ""
End Function

Private Function ExtractPdfPurchaseQty(ByVal raw As String) As Long
On Error GoTo ErrHandler

    ExtractPdfPurchaseQty = 1

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "\bPURCHASE\s+(\d+)\b"

    Dim m As Object
    Set m = re.Execute(raw)

    If m.Count > 0 Then
        ExtractPdfPurchaseQty = CLng(Val(m(0).SubMatches(0)))
    End If

    If ExtractPdfPurchaseQty <= 0 Then ExtractPdfPurchaseQty = 1
    Exit Function

ErrHandler:
    ExtractPdfPurchaseQty = 1
End Function

Private Function ExtractPdfPurchaseVendor(ByVal raw As String) As String
    Dim u As String
    u = UCase(raw)

    If InStr(u, "MCMASTER") > 0 Or InStr(u, "MCMASTER CARR") > 0 Then
        ExtractPdfPurchaseVendor = "McMaster-Carr"
        Exit Function
    End If

    If InStr(u, "D.M.E") > 0 Or InStr(u, "DME") > 0 Then
        ExtractPdfPurchaseVendor = "DME"
        Exit Function
    End If

    If InStr(u, "JACO") > 0 Then
        ExtractPdfPurchaseVendor = "JACO"
        Exit Function
    End If

    If InStr(u, "PCS") > 0 Then
        ExtractPdfPurchaseVendor = "PCS"
        Exit Function
    End If

    If InStr(u, "PYROPEL") > 0 Then
        ExtractPdfPurchaseVendor = "Pyropel"
        Exit Function
    End If

    ExtractPdfPurchaseVendor = ""
End Function

Private Function ExtractLeadingDetailNumber(ByVal raw As String) As String
On Error Resume Next

    Dim s As String
    s = Trim(raw)

    Dim i As Long
    Dim ch As String
    Dim token As String

    token = ""

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            token = token & ch
        Else
            Exit For
        End If
    Next i

    ExtractLeadingDetailNumber = token
End Function

Private Function ExtractPdfPurchasePartNo(ByVal raw As String) As String
On Error GoTo ErrHandler

    ExtractPdfPurchasePartNo = ""

    Dim u As String
    u = UCase(raw)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[A-Z0-9][A-Z0-9\-]*"

    Dim matches As Object
    Set matches = re.Execute(u)

    Dim i As Long
    Dim tok As String

    For i = 0 To matches.Count - 1

        tok = Trim(CStr(matches(i).Value))

        If IsUsefulPurchasePartToken(tok) Then
            ExtractPdfPurchasePartNo = tok
            Exit Function
        End If

    Next i

    Exit Function

ErrHandler:
    ExtractPdfPurchasePartNo = ""
End Function

Private Function IsUsefulPurchasePartToken(ByVal tok As String) As Boolean
    IsUsefulPurchasePartToken = False

    tok = UCase(Trim(tok))
    If tok = "" Then Exit Function

    Select Case tok
        Case "PURCHASE", "MATERIAL", "DME", "D", "M", "E", "CO", "CARR", _
             "MCMASTER", "JACO", "MANUFACTURING", "MANUFACTURIN", "PCS", _
             "PYROPEL", "EA", "LBS", "LB", "SQ", "IN", "CU", "HOLDER", _
             "INSULATION", "EXTERNAL", "GUIDE", "BUSHING", "LEADER", _
             "PIN", "PINS", "SAFETY", "STRAP", "SPRING", "DOWN", "HOLD"
            Exit Function
    End Select

    If Left(tok, 3) = "G1C" Then Exit Function
    If Left(tok, 3) = "500" Then Exit Function
    If Len(tok) < 3 Then Exit Function

    Dim hasLetter As Boolean
    Dim hasDigit As Boolean
    Dim hasHyphen As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(tok)
        ch = Mid(tok, i, 1)

        If ch >= "A" And ch <= "Z" Then hasLetter = True
        If ch >= "0" And ch <= "9" Then hasDigit = True
        If ch = "-" Then hasHyphen = True
    Next i

    If hasLetter And hasDigit Then
        IsUsefulPurchasePartToken = True
        Exit Function
    End If

    If hasHyphen And hasDigit And Len(tok) >= 5 Then
        IsUsefulPurchasePartToken = True
        Exit Function
    End If
End Function

Private Function TryParseTempcraftBasePdfMaterialLine(ByVal raw As String) As Boolean
On Error GoTo ErrHandler
    TryParseTempcraftBasePdfMaterialLine = False
    Dim nums() As Double
    Dim nCount As Long
    nCount = ExtractDecimalNumbers(raw, nums)
    If nCount < 3 Then Exit Function
    Dim desc As String
    desc = ExtractTempcraftPdfDescription(raw)
    If desc = "" Then Exit Function
    If FILL_PURCHASED_COMPONENTS Then CapturePurchased desc, ExtractTempcraftPdfQty(raw), ExtractTempcraftPdfMaterial(raw)
    If IsHardwareName(desc) Then Exit Function
    Dim qty As Long
    qty = ExtractTempcraftPdfQty(raw)
    Dim mat As String
    mat = ExtractTempcraftPdfMaterial(raw)
    Dim a As Double, b As Double, c As Double, l As Double, w As Double, t As Double
    ' Tempcraft PDF layout after description/material:
    '   Lth  Wth/O.D.  Hgt/I.D.  [Oracle]  [UOM]  StockWeight
    ' NEVER PickThreeLargest — that turns Stock Weight (117.87 lbs) into Length.
    ' Keep file order (Lth/Wth/Hgt); MapTempcraftBomDimsToCmsSteel remaps by plate role.
    If Not PickThreeFinishedSizeDims(nums, nCount, a, b, c) Then Exit Function
    SortThreeDimensions a, b, c, l, w, t
    If BomDimsLookLikeStockWeight(t, w, l) Then Exit Function
    AddBomRow desc, qty, mat, a, b, c, (a > 0 And b > 0 And c > 0), , , , , True
    TryParseTempcraftBasePdfMaterialLine = True
    Exit Function
ErrHandler:
    TryParseTempcraftBasePdfMaterialLine = False
End Function

' Prefer the first three plausible finished-size decimals on a Tempcraft BOM line.
' Skips leading Det No / Qty integers (ExtractDecimalNumbers only keeps dotted values)
' and drops trailing Stock Weight when a 4th+ number is much larger than the plate.
Private Function PickThreeFinishedSizeDims(ByRef nums() As Double, ByVal n As Long, _
                                           ByRef a As Double, ByRef b As Double, ByRef c As Double) As Boolean
    PickThreeFinishedSizeDims = False
    a = 0#: b = 0#: c = 0#
    If n < 3 Then Exit Function
    Dim i As Long, picked As Long
    Dim cand(1 To 12) As Double
    picked = 0
    For i = 1 To n
        ' Finished plate sizes are almost always under ~80"; stock weight often 50–300+.
        If nums(i) > 0.05 And nums(i) < 80# Then
            picked = picked + 1
            If picked <= 12 Then cand(picked) = nums(i)
        End If
    Next i
    If picked >= 3 Then
        a = cand(1): b = cand(2): c = cand(3)
        PickThreeFinishedSizeDims = True
        Exit Function
    End If
    ' Fallback: first three decimals regardless (still better than three largest).
    a = nums(1): b = nums(2): c = nums(3)
    PickThreeFinishedSizeDims = True
End Function

' True when the "length" looks like Tempcraft Stock Weight (lbs), not inches.
' Example bug: TCP 1.375 x 15.875 x 117.87  ← 117.87 is weight, real L is 18.
Private Function BomDimsLookLikeStockWeight(ByVal t As Double, ByVal w As Double, ByVal l As Double) As Boolean
    BomDimsLookLikeStockWeight = False
    If t <= 0 Or w <= 0 Or l <= 0 Then Exit Function
    ' Length far larger than the other two and in the typical weight band.
    If l >= 50# And l > (w * 2.5) And l > (t * 8#) Then
        BomDimsLookLikeStockWeight = True
        Exit Function
    End If
    ' Density sanity: steel ~0.283 lb/in^3. If L were inches, mass ≈ T*W*L*0.283.
    ' If L is actually weight, T*W*L is huge vs any real plate.
    Dim vol As Double
    vol = t * w * l
    If vol > 8000# And l > 40# Then BomDimsLookLikeStockWeight = True
End Function

Private Function ExtractDecimalNumbers(ByVal s As String, ByRef nums() As Double) As Long
    ReDim nums(1 To 60)
    Dim cnt As Long
    cnt = 0
    Dim i As Long, ch As String, token As String, hasDot As Boolean
    token = "": hasDot = False
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then ch = Mid(s, i, 1) Else ch = " "
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            token = token & ch
            If ch = "." Then hasDot = True
        Else
            If token <> "" And hasDot And IsNumeric(token) Then
                cnt = cnt + 1
                If cnt <= 60 Then nums(cnt) = CDbl(token)
            End If
            token = "": hasDot = False
        End If
    Next i
    ExtractDecimalNumbers = cnt
End Function

Private Sub PickThreeLargest(ByRef nums() As Double, ByVal n As Long, ByRef a As Double, ByRef b As Double, ByRef c As Double)
    ' Prefer PickThreeFinishedSizeDims for BOM lines that may include Stock Weight.
    a = 0#: b = 0#: c = 0#
    Dim i As Long, v As Double
    For i = 1 To n
        v = nums(i)
        If v > a Then
            c = b: b = a: a = v
        ElseIf v > b Then
            c = b: b = v
        ElseIf v > c Then
            c = v
        End If
    Next i
End Sub

Private Function ExtractTempcraftPdfDescription(ByVal raw As String) As String
    Dim s As String
    s = RemoveLeadingItemNumber(raw)
    Dim i As Long, ch As String, cutPos As Long
    cutPos = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then cutPos = i: Exit For
    Next i
    Dim d As String
    If cutPos > 0 Then d = Left(s, cutPos - 1) Else d = s
    ExtractTempcraftPdfDescription = ProperCaseText(Trim(d))
End Function

Private Function RemoveLeadingItemNumber(ByVal s As String) As String
    Dim t As String
    t = LTrim(s)
    Dim i As Long, ch As String, p As Long
    p = 0
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            p = i
        ElseIf ch = " " Then
            If p > 0 Then Exit For
        Else
            Exit For
        End If
    Next i
    If p > 0 And p < Len(t) Then
        RemoveLeadingItemNumber = LTrim(Mid(t, p + 1))
    Else
        RemoveLeadingItemNumber = t
    End If
End Function

Private Function ExtractTempcraftPdfQty(ByVal raw As String) As Long
    Dim q As Long
    q = ExtractQtyAfterOutsource(raw)
    If q > 0 Then ExtractTempcraftPdfQty = q Else ExtractTempcraftPdfQty = 1
End Function

Private Function ExtractQtyAfterOutsource(ByVal raw As String) As Long
    Dim u As String
    u = UCase(raw)
    Dim p As Long
    p = InStr(u, "OUTSOURCE")
    If p = 0 Then Exit Function
    Dim rest As String
    rest = Mid(raw, p + Len("OUTSOURCE"))
    Dim i As Long, ch As String, token As String
    For i = 1 To Len(rest) + 1
        If i <= Len(rest) Then ch = Mid(rest, i, 1) Else ch = " "
        If ch >= "0" And ch <= "9" Then
            token = token & ch
        Else
            If token <> "" Then ExtractQtyAfterOutsource = CLng(token): Exit Function
        End If
    Next i
End Function

Private Function ExtractTempcraftPdfMaterial(ByVal raw As String) As String
    Dim u As String
    u = UCase(raw)
    If InStr(u, "4140") > 0 Then ExtractTempcraftPdfMaterial = "4140": Exit Function
    If InStr(u, "P20") > 0 Or InStr(u, "P-20") > 0 Then ExtractTempcraftPdfMaterial = "P20": Exit Function
    If InStr(u, "A36") > 0 Or InStr(u, "A-36") > 0 Or InStr(u, "1045") > 0 Then ExtractTempcraftPdfMaterial = "A36": Exit Function
    If InStr(u, "420") > 0 Or InStr(u, "S136") > 0 Or InStr(u, "STAINLESS") > 0 Then ExtractTempcraftPdfMaterial = "420SS": Exit Function
    If InStr(u, "H13") > 0 Then ExtractTempcraftPdfMaterial = "H13": Exit Function
    If InStr(u, "6061") > 0 Or InStr(u, "ALUM") > 0 Then ExtractTempcraftPdfMaterial = "6061": Exit Function
    ExtractTempcraftPdfMaterial = ""
End Function

' ============================================================
' BOM COMMON
' ============================================================
Private Sub AddBomRow(ByVal desc As String, ByVal qty As Long, ByVal mat As String, _
                      ByVal tt As Double, ByVal ww As Double, ByVal ll As Double, ByVal hasD As Boolean, _
                      Optional ByVal partNo As String = "", Optional ByVal manuf As String = "", _
                      Optional ByVal detNo As String = "", Optional ByVal purchType As String = "", _
                      Optional ByVal isTempcraftOrder As Boolean = False)
    If desc = "" Then Exit Sub
    If FILL_PURCHASED_COMPONENTS Then CapturePurchased desc, qty, mat, tt, ww, ll, partNo, manuf, detNo, purchType
    If ShouldUseBomItem(desc, mat, tt, ww, ll, hasD) = False Then Exit Sub
    BomCount = BomCount + 1
    ReDim Preserve BomRows(1 To BomCount)
    BomRows(BomCount).Description = desc
    BomRows(BomCount).quoteName = StandardPlateName(desc)
    BomRows(BomCount).Quantity = IIf(qty < 1, 1, qty)
    BomRows(BomCount).material = NormalizeSteelType(mat)
    BomRows(BomCount).BomThickness = Round(tt, DIM_DECIMALS)
    BomRows(BomCount).BomWidth = Round(ww, DIM_DECIMALS)
    BomRows(BomCount).BomLength = Round(ll, DIM_DECIMALS)
    BomRows(BomCount).hasDims = hasD
    BomRows(BomCount).BomIsTempcraftOrder = isTempcraftOrder
End Sub

Private Function ShouldUseBomItem(ByVal desc As String, ByVal mat As String, _
                                  ByVal tt As Double, ByVal ww As Double, ByVal ll As Double, ByVal hasD As Boolean) As Boolean
    ShouldUseBomItem = False
    If IsHardwareName(desc) Then Exit Function
    If ONLY_INCLUDE_4140_BOM_ITEMS Then
        If InStr(NormalizeSteelType(mat), "4140") = 0 Then Exit Function
    End If
    If HIDE_QUARTER_INCH_THICKNESS And hasD Then
        If IsQuarterInchThickness(tt) Then Exit Function
    End If
    ShouldUseBomItem = True
End Function

Private Function StandardPlateName(ByVal desc As String) As String
On Error GoTo ErrHandler

    Dim s As String
    s = NormalizeText(desc)

    If InStr(s, "EJECTOR J-BLOCK") > 0 Or InStr(s, "EJ J-BLOCK") > 0 Or _
       InStr(s, "EJ J BLOCK") > 0 Or InStr(s, "J-BLOCK") > 0 Or InStr(s, "J BLOCK") > 0 Then
        StandardPlateName = "EJECTOR J-BLOCK"
        Exit Function
    End If

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then
        Dim pcLoc As String
        Dim pcKind As String

        pcLoc = GetPullcoreLocationCode(s)
        If InStr(s, "CAM") > 0 Then
            pcKind = "PULLCORE CAM"
        ElseIf InStr(s, "KEY") > 0 Then
            pcKind = "PULLCORE KEY"
        Else
            pcKind = "PULLCORE"
        End If

        If pcLoc <> "" Then
            StandardPlateName = pcLoc & " " & pcKind
        Else
            StandardPlateName = pcKind
        End If
        Exit Function
    End If

    If InStr(s, "ID HOLDER") > 0 Or InStr(s, "IDTE HOLDER") > 0 Or InStr(s, "IDLE HOLDER") > 0 Or _
       InStr(s, "TOP HOLDER") > 0 Or InStr(s, "TOP HOLDER BLOCK") > 0 Or _
       InStr(s, "TOP CARRIER") > 0 Or InStr(s, "ID CARRIER") > 0 Or _
       InStr(s, "ID MOLD BASE") > 0 Or InStr(s, "ID MOLDBASE") > 0 Or _
       InStr(s, "TOP MOLD BASE") > 0 Or InStr(s, "TOP MOLDBASE") > 0 Then
        StandardPlateName = "ID HOLDER"
        Exit Function
    End If

    If InStr(s, "OD HOLDER") > 0 Or InStr(s, "ODTE HOLDER") > 0 Or InStr(s, "ODLE HOLDER") > 0 Or _
       InStr(s, "BOTTOM HOLDER") > 0 Or InStr(s, "BOT HOLDER") > 0 Or _
       InStr(s, "BOTTOM HOLDER BLOCK") > 0 Or InStr(s, "BOT HOLDER BLOCK") > 0 Or _
       InStr(s, "BOTTOM CARRIER") > 0 Or InStr(s, "BOT CARRIER") > 0 Or _
       InStr(s, "OD CARRIER") > 0 Or InStr(s, "OD MOLD BASE") > 0 Or _
       InStr(s, "OD MOLDBASE") > 0 Or InStr(s, "BOTTOM MOLD BASE") > 0 Or _
       InStr(s, "BOT MOLD BASE") > 0 Or InStr(s, "BOTTOM MOLDBASE") > 0 Or _
       InStr(s, "BOT MOLDBASE") > 0 Then
        StandardPlateName = "OD HOLDER"
        Exit Function
    End If

    If InStr(s, "POT BLOCK") > 0 Or InStr(s, "POT BLK") > 0 Or InStr(s, " POT ") > 0 Or Right(s, 4) = " POT" Then
        If IsLikelyOdSideName(s) Then
            StandardPlateName = "OD POT BLOCK"
            Exit Function
        End If

        If IsLikelyIdSideName(s) Then
            StandardPlateName = "ID POT BLOCK"
            Exit Function
        End If

        StandardPlateName = "POT BLOCK"
        Exit Function
    End If

    ' Tempcraft: "Top Pot Block Material" / "Bottom Pot Block Material" without "POT " token edge cases
    If InStr(s, "POT") > 0 And InStr(s, "BLOCK") > 0 Then
        If IsLikelyOdSideName(s) Then StandardPlateName = "OD POT BLOCK": Exit Function
        If IsLikelyIdSideName(s) Then StandardPlateName = "ID POT BLOCK": Exit Function
    End If

    If InStr(s, "SMED") > 0 Then
        If IsLikelyIdSideName(s) Then
            StandardPlateName = "TCP"
            Exit Function
        End If

        If IsLikelyOdSideName(s) Then
            StandardPlateName = "BCP"
            Exit Function
        End If

        StandardPlateName = "SMED PLATE"
        Exit Function
    End If

    If InStr(s, "TOP CLAMP") > 0 Or InStr(s, "TOP CLAMPING") > 0 Then
        StandardPlateName = "TCP"
        Exit Function
    End If

    If InStr(s, "BOTTOM CLAMP") > 0 Or InStr(s, "BOT CLAMP") > 0 Or _
       InStr(s, "BOTTOM CLAMPING") > 0 Or InStr(s, "BOT CLAMPING") > 0 Then
        StandardPlateName = "BCP"
        Exit Function
    End If

    If InStr(s, "TCP") > 0 Then
        StandardPlateName = "TCP"
        Exit Function
    End If

    If InStr(s, "BCP") > 0 Then
        StandardPlateName = "BCP"
        Exit Function
    End If

    If InStr(s, "TOP INS") > 0 Or InStr(s, "TOP INSULATION") > 0 Then
        StandardPlateName = "TOP INS"
        Exit Function
    End If

    If InStr(s, "BOT INS") > 0 Or InStr(s, "BOTTOM INS") > 0 Or InStr(s, "BOTTOM INSULATION") > 0 Then
        StandardPlateName = "BOT INS"
        Exit Function
    End If

    If InStr(s, "PULLCORE STOP") > 0 Or InStr(s, "PULL CORE STOP") > 0 Then
        StandardPlateName = "PULLCORE STOP"
        Exit Function
    End If

    If InStr(s, "FLIPPER") > 0 And InStr(s, "KEY") > 0 Then StandardPlateName = "Flipper Angle Plate Key": Exit Function
    If InStr(s, "FLIPPER") > 0 And InStr(s, "CAM") > 0 Then StandardPlateName = "Flipper Cam Mount": Exit Function
    If InStr(s, "FLIPPER") > 0 Then StandardPlateName = "Flipper Angle Plate": Exit Function

    StandardPlateName = ProperCaseText(Trim(desc))
    Exit Function

ErrHandler:
    StandardPlateName = Trim(desc)
End Function

Private Function IsLikelyIdSideName(ByVal s As String) As Boolean
    s = NormalizeText(s)
    If InStr(s, "ID ") > 0 Or Left(s, 2) = "ID" Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, "IDTE") > 0 Or InStr(s, "IDLE") > 0 Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, " TOP") > 0 Or Left(s, 3) = "TOP" Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, "TCP") > 0 Then IsLikelyIdSideName = True
    If InStr(s, "A SIDE") > 0 Or InStr(s, "CAVITY SIDE") > 0 Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, "STATIONARY") > 0 Or InStr(s, "FIXED") > 0 Or InStr(s, "INJECTION") > 0 Then IsLikelyIdSideName = True
End Function

Private Function IsLikelyOdSideName(ByVal s As String) As Boolean
    s = NormalizeText(s)
    If InStr(s, "OD ") > 0 Or Left(s, 2) = "OD" Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "ODTE") > 0 Or InStr(s, "ODLE") > 0 Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "BOTTOM") > 0 Or InStr(s, "BOT ") > 0 Or Left(s, 3) = "BOT" Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "BCP") > 0 Then IsLikelyOdSideName = True
    If InStr(s, "B SIDE") > 0 Or InStr(s, "CORE SIDE") > 0 Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "MOVABLE") > 0 Or InStr(s, "EJECTOR") > 0 Or InStr(s, "EJECTION") > 0 Then IsLikelyOdSideName = True
End Function

Private Function IsHardwareName(ByVal desc As String) As Boolean
    Dim s As String
    s = NormalizeText(desc)
    Dim hw As Variant
    hw = Array("SCREW", "BOLT", "SHCS", "DOWEL", "O-RING", "ORING", "WASHER", "SPRING", _
               "BUSHING", "LEADER PIN", "GUIDE PIN", "RETURN PIN", "EYE BOLT", "EYEBOLT", "LIFTING", _
               "PLUG", "FITTING", "GREASE", "BAFFLE", "THERMOCOUPLE", "HEATER", "NIPPLE", "QUICK DISCONNECT", _
               "PILLAR", "STOP DISC", "SIDE LOCK", "SIDELOCK", "STRAP", "SPRUE PULLER", "LOCATING RING", _
               "SPRUE BUSHING", "PULL DOWEL", "JIFFY", "WATER", "SOCKET")
    Dim i As Long
    For i = LBound(hw) To UBound(hw)
        If InStr(s, CStr(hw(i))) > 0 Then IsHardwareName = True: Exit Function
    Next i
End Function

Private Function NormalizeSteelType(ByVal mat As String) As String
    Dim u As String
    u = UCase(Trim(mat))
    If u = "" Then NormalizeSteelType = "": Exit Function
    If InStr(u, "PYROPEL") > 0 Then NormalizeSteelType = "Pyropel": Exit Function
    If InStr(u, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
    If InStr(u, "P20") > 0 Or InStr(u, "P-20") > 0 Then NormalizeSteelType = "P20": Exit Function
    If InStr(u, "A36") > 0 Or InStr(u, "A-36") > 0 Or InStr(u, "1045") > 0 Or InStr(u, "1030") > 0 _
       Or InStr(u, "1020") > 0 Or InStr(u, "HOT ROLLED") > 0 Or InStr(u, "COLD ROLLED") > 0 Then NormalizeSteelType = "A36": Exit Function
    If InStr(u, "420") > 0 Or InStr(u, "S136") > 0 Then NormalizeSteelType = "420SS": Exit Function
    If InStr(u, "H13") > 0 Or InStr(u, "H-13") > 0 Then NormalizeSteelType = "H13": Exit Function
    If InStr(u, "6061") > 0 Or InStr(u, "ALUM") > 0 Then NormalizeSteelType = "6061": Exit Function
    If InStr(u, "A-2") > 0 Or InStr(u, " A2") > 0 Or u = "A2" Then NormalizeSteelType = "A2": Exit Function
    If InStr(u, "O-1") > 0 Or InStr(u, "0-1") > 0 Or InStr(u, "O1") > 0 Or InStr(u, "FLAT GROUND") > 0 Then NormalizeSteelType = "O1": Exit Function
    ' DME grade codes used on the shop sheets / "other software" BOMs.
    If InStr(u, "#7") > 0 Then NormalizeSteelType = "420SS": Exit Function
    If InStr(u, "#5") > 0 Then NormalizeSteelType = "H13": Exit Function
    If InStr(u, "#3") > 0 Then NormalizeSteelType = "P20": Exit Function
    If InStr(u, "#2") > 0 Then NormalizeSteelType = "4140": Exit Function
    If InStr(u, "#1") > 0 Then NormalizeSteelType = "A36": Exit Function
    NormalizeSteelType = u
End Function

Private Function IsQuarterInchThickness(ByVal t As Double) As Boolean
    IsQuarterInchThickness = (Abs(t - QUARTER_INCH_THICKNESS) <= QUARTER_INCH_TOLERANCE)
End Function

Private Function ProperCaseText(ByVal s As String) As String
    s = Trim(s)
    If s = "" Then Exit Function
    ProperCaseText = StrConv(s, vbProperCase)
End Function

' ============================================================
' MATCH BOM -> CAD
' ============================================================
Private Sub BuildExportRowsFromBom()
    Dim i As Long, cadIdx As Long
    For i = 1 To BomCount
        cadIdx = FindBestCadMatchForBom(i)
        AddExportRow i, cadIdx
        If cadIdx > 0 Then parts(cadIdx).UsedForBomMatch = True
    Next i
End Sub

Private Sub AddExportRow(ByVal bomIdx As Long, ByVal cadIdx As Long)
    ExportCount = ExportCount + 1
    ReDim Preserve ExportRows(1 To ExportCount)
    ExportRows(ExportCount).quoteName = CanonicalHolderQuoteName(BomRows(bomIdx).quoteName)
    ExportRows(ExportCount).Quantity = BomRows(bomIdx).Quantity
    ExportRows(ExportCount).material = BomRows(bomIdx).material
    ExportRows(ExportCount).BomThickness = BomRows(bomIdx).BomThickness
    ExportRows(ExportCount).BomWidth = BomRows(bomIdx).BomWidth
    ExportRows(ExportCount).BomLength = BomRows(bomIdx).BomLength
    ExportRows(ExportCount).HasBomDims = BomRows(bomIdx).hasDims
    If cadIdx > 0 Then
        ExportRows(ExportCount).HasCad = True
        ExportRows(ExportCount).CadPartIndex = cadIdx
        ExportRows(ExportCount).Thickness = parts(cadIdx).Thickness
        ExportRows(ExportCount).Width = parts(cadIdx).Width
        ExportRows(ExportCount).Length = parts(cadIdx).Length
    Else
        ExportRows(ExportCount).HasCad = False
    End If
    ExportRows(ExportCount).Status = CompareBomToCadStatus(bomIdx, cadIdx)
End Sub

Private Function CanonicalHolderQuoteName(ByVal quoteName As String) As String
    CanonicalHolderQuoteName = quoteName

    Dim d As String
    Dim k As String
    d = NormalizeText(quoteName)
    k = NormalizeKey(quoteName)

    Select Case k
        Case "IDHOLDER", "TOPHOLDER", "TOPHOLDERBLOCK", "IDHOLDERBLOCK", _
             "IDTEHOLDER", "IDLEHOLDER", "TOPCARRIER", "IDCARRIER", _
             "IDMOLDBASE", "IDTEMOLDBASE", "IDLEMOLDBASE", "TOPMOLDBASE"
            CanonicalHolderQuoteName = "ID HOLDER"
            Exit Function

        Case "ODHOLDER", "BOTTOMHOLDER", "BOTHOLDER", "BOTTOMHOLDERBLOCK", _
             "BOTHOLDERBLOCK", "ODHOLDERBLOCK", "ODTEHOLDER", "ODLEHOLDER", _
             "BOTTOMCARRIER", "BOTCARRIER", "ODCARRIER", "ODMOLDBASE", _
             "ODTEMOLDBASE", "ODLEMOLDBASE", "BOTTOMMOLDBASE", "BOTMOLDBASE"
            CanonicalHolderQuoteName = "OD HOLDER"
            Exit Function

        Case "IDPOT", "IDPOTBLOCK", "IDTEPOT", "IDLEPOT", "TOPPOT", _
             "TOPPOTBLOCK", "TCPPOT", "TCPPOTBLOCK"
            CanonicalHolderQuoteName = "ID POT BLOCK"
            Exit Function

        Case "ODPOT", "ODPOTBLOCK", "ODTEPOT", "ODLEPOT", "BOTTOMPOT", _
             "BOTPOT", "BOTTOMPOTBLOCK", "BOTPOTBLOCK", "BCPPOT", "BCPPOTBLOCK"
            CanonicalHolderQuoteName = "OD POT BLOCK"
            Exit Function

        Case "TOPCLAMPING", "TOPCLAMPINGPLATE", "TOPCLAMPPLATE", _
             "TOPSMED", "TOPSMEDPLATE", "IDSMED", "IDSMEDPLATE", _
             "IDCLAMPING", "IDCLAMPINGPLATE"
            CanonicalHolderQuoteName = "TCP"
            Exit Function

        Case "BOTTOMCLAMPING", "BOTCLAMPING", "BOTTOMCLAMPINGPLATE", _
             "BOTCLAMPINGPLATE", "BOTTOMCLAMPPLATE", "BOTCLAMPPLATE", _
             "BOTTOMSMED", "BOTSMED", "BOTTOMSMEDPLATE", "BOTSMEDPLATE", _
             "ODSMED", "ODSMEDPLATE", "ODCLAMPING", "ODCLAMPINGPLATE"
            CanonicalHolderQuoteName = "BCP"
            Exit Function
    End Select

    If ContainsAnyPipeKey(quoteName, ID_HOLDER_KEYS) Then
        CanonicalHolderQuoteName = "ID HOLDER"
        Exit Function
    End If

    If ContainsAnyPipeKey(quoteName, OD_HOLDER_KEYS) Then
        CanonicalHolderQuoteName = "OD HOLDER"
        Exit Function
    End If

    If ContainsAnyPipeKey(quoteName, KEYS_ID_POT) Then
        CanonicalHolderQuoteName = "ID POT BLOCK"
        Exit Function
    End If

    If ContainsAnyPipeKey(quoteName, KEYS_OD_POT) Then
        CanonicalHolderQuoteName = "OD POT BLOCK"
        Exit Function
    End If

    If ContainsAnyPipeKey(quoteName, KEYS_TCP) Then
        CanonicalHolderQuoteName = "TCP"
        Exit Function
    End If

    If ContainsAnyPipeKey(quoteName, KEYS_BCP) Then
        CanonicalHolderQuoteName = "BCP"
        Exit Function
    End If
End Function

Private Function QuoteAliasKeys(ByVal quoteName As String) As String
    Dim k As String
    Dim d As String
    k = NormalizeKey(quoteName)
    d = NormalizeText(quoteName)

    Select Case k
        Case "IDHOLDER"
            QuoteAliasKeys = ID_HOLDER_KEYS
            Exit Function
        Case "ODHOLDER"
            QuoteAliasKeys = OD_HOLDER_KEYS
            Exit Function
        Case "IDPOT", "IDPOTBLOCK"
            QuoteAliasKeys = KEYS_ID_POT
            Exit Function
        Case "ODPOT", "ODPOTBLOCK"
            QuoteAliasKeys = KEYS_OD_POT
            Exit Function
        Case "TCP"
            QuoteAliasKeys = KEYS_TCP
            Exit Function
        Case "BCP"
            QuoteAliasKeys = KEYS_BCP
            Exit Function
        Case "TOPINS"
            QuoteAliasKeys = "TOP INS|TOP INSULATION|ID INS|ID INSERT|TOP INSERT|IDTE INS|IDLE INS"
            Exit Function
        Case "BOTINS"
            QuoteAliasKeys = "BOT INS|BOTTOM INS|BOTTOM INSULATION|OD INS|OD INSERT|BOTTOM INSERT|BOT INSERT|ODTE INS|ODLE INS"
            Exit Function
    End Select

    If ContainsAnyPipeKey(quoteName, ID_HOLDER_KEYS) Then QuoteAliasKeys = ID_HOLDER_KEYS: Exit Function
    If ContainsAnyPipeKey(quoteName, OD_HOLDER_KEYS) Then QuoteAliasKeys = OD_HOLDER_KEYS: Exit Function
    If ContainsAnyPipeKey(quoteName, KEYS_ID_POT) Then QuoteAliasKeys = KEYS_ID_POT: Exit Function
    If ContainsAnyPipeKey(quoteName, KEYS_OD_POT) Then QuoteAliasKeys = KEYS_OD_POT: Exit Function
    If ContainsAnyPipeKey(quoteName, KEYS_TCP) Then QuoteAliasKeys = KEYS_TCP: Exit Function
    If ContainsAnyPipeKey(quoteName, KEYS_BCP) Then QuoteAliasKeys = KEYS_BCP: Exit Function

    If InStr(d, "POT") > 0 Then
        If IsLikelyIdSideName(d) Then QuoteAliasKeys = KEYS_ID_POT: Exit Function
        If IsLikelyOdSideName(d) Then QuoteAliasKeys = KEYS_OD_POT: Exit Function
    End If

    If InStr(d, "HOLDER") > 0 Or InStr(d, "CARRIER") > 0 Or _
       InStr(d, "MOLD BASE") > 0 Or InStr(d, "MOLDBASE") > 0 Then
        If IsLikelyIdSideName(d) Then QuoteAliasKeys = ID_HOLDER_KEYS: Exit Function
        If IsLikelyOdSideName(d) Then QuoteAliasKeys = OD_HOLDER_KEYS: Exit Function
    End If

    If InStr(d, "SMED") > 0 Or InStr(d, "CLAMP") > 0 Or InStr(d, "CLAMPING") > 0 Then
        If IsLikelyIdSideName(d) Then QuoteAliasKeys = KEYS_TCP: Exit Function
        If IsLikelyOdSideName(d) Then QuoteAliasKeys = KEYS_BCP: Exit Function
    End If
End Function

Private Function IsTcpBcpQuoteName(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "TCP", "BCP", _
             "TOPCLAMPINGPLATE", "BOTTOMCLAMPINGPLATE", _
             "TOPSMEDPLATE", "BOTTOMSMEDPLATE", _
             "TOPSMED", "BOTTOMSMED", "BOTSMED"
            IsTcpBcpQuoteName = True
    End Select
End Function

Private Function ShouldUseSameSizePairMassRule(ByVal bomIdx As Long) As Boolean
On Error GoTo ErrHandler

    ShouldUseSameSizePairMassRule = False

    If bomIdx <= 0 Or bomIdx > BomCount Then Exit Function
    If BomRows(bomIdx).hasDims = False Then Exit Function
    If GetMassPreferenceForQuoteName(BomRows(bomIdx).quoteName) = "" Then Exit Function

    Dim counterpart As String
    counterpart = GetCounterpartQuoteName(BomRows(bomIdx).quoteName)
    If counterpart = "" Then Exit Function

    Dim i As Long
    For i = 1 To BomCount
        If i <> bomIdx Then
            If NormalizeKey(BomRows(i).quoteName) = NormalizeKey(counterpart) Then
                If BomRows(i).hasDims Then
                    If Abs(BomRows(i).BomLength - BomRows(bomIdx).BomLength) <= SAME_SIZE_PAIR_TOL Then
                        If Abs(BomRows(i).BomWidth - BomRows(bomIdx).BomWidth) <= SAME_SIZE_PAIR_TOL Then
                            If Abs(BomRows(i).BomThickness - BomRows(bomIdx).BomThickness) <= SAME_SIZE_PAIR_TOL Then
                                ShouldUseSameSizePairMassRule = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    ShouldUseSameSizePairMassRule = False
End Function

Private Function GetMassPreferenceForQuoteName(ByVal quoteName As String) As String
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "TOPINS" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "BOTINS" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function

    If k = "TCP" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "BCP" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function
    If k = "IDPOT" Or k = "IDPOTBLOCK" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "ODPOT" Or k = "ODPOTBLOCK" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function
    If k = "IDHOLDER" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function
    If k = "ODHOLDER" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function

    If k = "TOPCLAMPINGPLATE" Or k = "TOPSMEDPLATE" Or k = "TOPSMED" Then
        GetMassPreferenceForQuoteName = "LIGHT"
        Exit Function
    End If

    If k = "BOTTOMCLAMPINGPLATE" Or k = "BOTCLAMPINGPLATE" Or _
       k = "BOTTOMSMEDPLATE" Or k = "BOTSMEDPLATE" Or _
       k = "BOTTOMSMED" Or k = "BOTSMED" Then
        GetMassPreferenceForQuoteName = "HEAVY"
        Exit Function
    End If

    GetMassPreferenceForQuoteName = ""
End Function

Private Function GetCounterpartQuoteName(ByVal quoteName As String) As String
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "IDHOLDER" Then GetCounterpartQuoteName = "OD HOLDER": Exit Function
    If k = "ODHOLDER" Then GetCounterpartQuoteName = "ID HOLDER": Exit Function
    If k = "IDPOT" Or k = "IDPOTBLOCK" Then GetCounterpartQuoteName = "OD POT BLOCK": Exit Function
    If k = "ODPOT" Or k = "ODPOTBLOCK" Then GetCounterpartQuoteName = "ID POT BLOCK": Exit Function
    If k = "TOPINS" Then GetCounterpartQuoteName = "BOT INS": Exit Function
    If k = "BOTINS" Then GetCounterpartQuoteName = "TOP INS": Exit Function
    If k = "TCP" Then GetCounterpartQuoteName = "BCP": Exit Function
    If k = "BCP" Then GetCounterpartQuoteName = "TCP": Exit Function

    If k = "TOPCLAMPINGPLATE" Or k = "TOPSMEDPLATE" Or k = "TOPSMED" Then
        GetCounterpartQuoteName = "BCP"
        Exit Function
    End If

    If k = "BOTTOMCLAMPINGPLATE" Or k = "BOTCLAMPINGPLATE" Or _
       k = "BOTTOMSMEDPLATE" Or k = "BOTSMEDPLATE" Or _
       k = "BOTTOMSMED" Or k = "BOTSMED" Then
        GetCounterpartQuoteName = "TCP"
        Exit Function
    End If
End Function

Private Function FindBestCadMatchForBomByDimsAndMassPreference(ByVal bomIdx As Long, _
                                                               ByVal massPreference As String) As Long
On Error GoTo ErrHandler

    FindBestCadMatchForBomByDimsAndMassPreference = 0

    If bomIdx <= 0 Or bomIdx > BomCount Then Exit Function
    If BomRows(bomIdx).hasDims = False Then Exit Function
    If massPreference = "" Then Exit Function

    Dim i As Long
    Dim bestDimDiff As Double
    bestDimDiff = 1E+99

    Dim dL As Double, dW As Double, dT As Double, totalDiff As Double

    For i = 1 To PartCount
        If parts(i).UsedForBomMatch = False Then
            dL = Abs(parts(i).Length - BomRows(bomIdx).BomLength)
            dW = Abs(parts(i).Width - BomRows(bomIdx).BomWidth)
            dT = Abs(parts(i).Thickness - BomRows(bomIdx).BomThickness)
            totalDiff = dL + dW + dT

            If dL <= DIM_REVIEW_TOL * 4 And _
               dW <= DIM_REVIEW_TOL * 4 And _
               dT <= DIM_REVIEW_TOL * 4 Then
                If totalDiff < bestDimDiff Then bestDimDiff = totalDiff
            End If
        End If
    Next i

    If bestDimDiff = 1E+99 Then Exit Function

    Dim dimBand As Double
    dimBand = DIM_OK_TOL
    If dimBand < 0.05 Then dimBand = 0.05

    Dim bestIdx As Long
    Dim bestMass As Double

    bestIdx = 0
    If UCase(massPreference) = "LIGHT" Then
        bestMass = 1E+99
    Else
        bestMass = -1E+99
    End If

    For i = 1 To PartCount
        If parts(i).UsedForBomMatch = False Then
            dL = Abs(parts(i).Length - BomRows(bomIdx).BomLength)
            dW = Abs(parts(i).Width - BomRows(bomIdx).BomWidth)
            dT = Abs(parts(i).Thickness - BomRows(bomIdx).BomThickness)
            totalDiff = dL + dW + dT

            If dL <= DIM_REVIEW_TOL * 4 And _
               dW <= DIM_REVIEW_TOL * 4 And _
               dT <= DIM_REVIEW_TOL * 4 Then
                If totalDiff <= bestDimDiff + dimBand Then
                    If UCase(massPreference) = "LIGHT" Then
                        If parts(i).massValue < bestMass Then
                            bestMass = parts(i).massValue
                            bestIdx = i
                        End If
                    ElseIf UCase(massPreference) = "HEAVY" Then
                        If parts(i).massValue > bestMass Then
                            bestMass = parts(i).massValue
                            bestIdx = i
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If bestIdx > 0 Then
        LogLine "Same-size pair mass match: " & BomRows(bomIdx).quoteName & _
                " preference=" & massPreference & _
                " -> CAD '" & parts(bestIdx).componentName & "'" & _
                " mass=" & FormatNumberForCsv(parts(bestIdx).massValue)
    End If

    FindBestCadMatchForBomByDimsAndMassPreference = bestIdx
    Exit Function

ErrHandler:
    LogLine "FindBestCadMatchForBomByDimsAndMassPreference error: " & Err.Description
    FindBestCadMatchForBomByDimsAndMassPreference = 0
End Function
Private Function FindBestCadMatchForBom(ByVal bomIdx As Long) As Long
    If ShouldUseSameSizePairMassRule(bomIdx) Then
        Dim pref As String
        Dim massIdx As Long
        pref = GetMassPreferenceForQuoteName(BomRows(bomIdx).quoteName)
        If pref <> "" Then
            massIdx = FindBestCadMatchForBomByDimsAndMassPreference(bomIdx, pref)
            If massIdx > 0 Then
                FindBestCadMatchForBom = massIdx
                Exit Function
            End If
        End If
    End If

    Dim bestIdx As Long, bestScore As Double
    bestIdx = 0: bestScore = DIM_MAX_MATCH_TOTAL_DIFF + 1#
    Dim i As Long, nm As Boolean, diff As Double
    For i = 1 To PartCount
        If parts(i).UsedForBomMatch = False Then
            nm = IsNameMatch(BomRows(bomIdx).quoteName, parts(i).componentName)
            If BomRows(bomIdx).hasDims Then
                Dim dL As Double
                Dim dW As Double
                Dim dT As Double

                dL = Abs(parts(i).Length - BomRows(bomIdx).BomLength)
                dW = Abs(parts(i).Width - BomRows(bomIdx).BomWidth)
                dT = Abs(parts(i).Thickness - BomRows(bomIdx).BomThickness)
                diff = dL + dW + dT

                If dL <= DIM_REVIEW_TOL * 4 And _
                   dW <= DIM_REVIEW_TOL * 4 And _
                   dT <= DIM_REVIEW_TOL * 4 Then
                    If nm Then diff = diff - 0.75
                    If diff < bestScore Then bestScore = diff: bestIdx = i
                End If
            Else
                If nm Then
                    If -parts(i).BBoxVolume < bestScore Then bestScore = -parts(i).BBoxVolume: bestIdx = i
                End If
            End If
        End If
    Next i
    If bestIdx > 0 And BomRows(bomIdx).hasDims Then
        Dim realDiff As Double
        realDiff = Abs(parts(bestIdx).Length - BomRows(bomIdx).BomLength) + _
                   Abs(parts(bestIdx).Width - BomRows(bomIdx).BomWidth) + _
                   Abs(parts(bestIdx).Thickness - BomRows(bomIdx).BomThickness)
        If realDiff > DIM_MAX_MATCH_TOTAL_DIFF And IsNameMatch(BomRows(bomIdx).quoteName, parts(bestIdx).componentName) = False Then
            bestIdx = FindMassBasedCadMatchForBom(bomIdx)
        End If
    End If
    FindBestCadMatchForBom = bestIdx
End Function

Private Function FindMassBasedCadMatchForBom(ByVal bomIdx As Long) As Long
    Dim i As Long, bestIdx As Long, bestVol As Double
    bestIdx = 0: bestVol = -1#
    For i = 1 To PartCount
        If parts(i).UsedForBomMatch = False Then
            If IsNameMatch(BomRows(bomIdx).quoteName, parts(i).componentName) Then
                If parts(i).BBoxVolume > bestVol Then bestVol = parts(i).BBoxVolume: bestIdx = i
            End If
        End If
    Next i
    FindMassBasedCadMatchForBom = bestIdx
End Function

Private Function CompareBomToCadStatus(ByVal bomIdx As Long, ByVal cadIdx As Long) As String
    If cadIdx = 0 Then CompareBomToCadStatus = "NO CAD MATCH": Exit Function
    If BomRows(bomIdx).hasDims = False Then CompareBomToCadStatus = "OK (no BOM dims)": Exit Function
    Dim d As Double
    d = Abs(parts(cadIdx).Length - BomRows(bomIdx).BomLength) + _
        Abs(parts(cadIdx).Width - BomRows(bomIdx).BomWidth) + _
        Abs(parts(cadIdx).Thickness - BomRows(bomIdx).BomThickness)
    If d <= DIM_OK_TOL * 3# Then
        CompareBomToCadStatus = "OK"
    ElseIf d <= DIM_REVIEW_TOL * 3# Then
        CompareBomToCadStatus = "REVIEW"
    Else
        CompareBomToCadStatus = "MISMATCH"
    End If
End Function

Private Function IsNameMatch(ByVal quoteName As String, ByVal cadName As String) As Boolean
    Dim q As String, c As String
    q = NormalizeKey(quoteName)
    c = NormalizeKey(cadName)
    If q = "" Or c = "" Then Exit Function
    If InStr(c, q) > 0 Or InStr(q, c) > 0 Then IsNameMatch = True

    Dim aliases As String
    aliases = QuoteAliasKeys(quoteName)
    If aliases <> "" Then
        If ContainsAnyPipeKey(cadName, aliases) Then
            IsNameMatch = True
            Exit Function
        End If
    End If
End Function

Private Function ContainsAnyPipeKey(ByVal haystack As String, ByVal pipeKeys As String) As Boolean
    If haystack = "" Or pipeKeys = "" Then Exit Function
    Dim hayText As String, hayKey As String
    hayText = NormalizeText(haystack)
    hayKey = NormalizeKey(haystack)
    Dim arr() As String
    arr = Split(pipeKeys, "|")
    Dim i As Long, k As String
    For i = LBound(arr) To UBound(arr)
        k = NormalizeText(arr(i))
        If k <> "" Then If InStr(hayText, k) > 0 Then ContainsAnyPipeKey = True: Exit Function
    Next i
    For i = LBound(arr) To UBound(arr)
        k = NormalizeKey(arr(i))
        If k <> "" Then If InStr(hayKey, k) > 0 Then ContainsAnyPipeKey = True: Exit Function
    Next i
End Function

' ============================================================
' EXCEL FILL  (Quote sheet + J000 steel sheet)
' ============================================================
' Resolve a plate's finished dims: try CAD bounding-box (by name key) first,
' then fall back to the BOM row (by standard name). Returns True if found.
' Identify the six pot-block plates directly from CAD geometry:
'   - clamps = broad, flat, near-full-footprint plates (TCP / BCP / SMED)
'   - pots   = thick, chunky, clearly SMALLER than the mold footprint
'   - holders = thick elongated plates (not skinny rails/straps)
' Within each pair the dominant-axis separation picks ID/top vs OD/bottom.
Private Sub ClassifyPotBlockPlatesFromCad()
On Error GoTo ErrHandler

    gIdxTCP = 0
    gIdxBCP = 0
    gIdxIDH = 0
    gIdxODH = 0
    gIdxIDP = 0
    gIdxODP = 0

    If PartCount < 1 Then Exit Sub

    Dim maxFp As Double
    Dim i As Long
    Dim fp As Double

    maxFp = 0#

    For i = 1 To PartCount
        If Not IsPyropelPartIndex(i) Then
            fp = parts(i).Width * parts(i).Length
            If fp > maxFp Then maxFp = fp
        End If
    Next i

    If maxFp <= 0# Then Exit Sub

    Dim potList() As Long
    Dim otherList() As Long
    Dim nPot As Long
    Dim nOther As Long

    ReDim potList(1 To PartCount)
    ReDim otherList(1 To PartCount)

    Dim t As Double
    Dim w As Double
    Dim l As Double

    nPot = 0
    nOther = 0

    For i = 1 To PartCount

        If IsPyropelPartIndex(i) Then GoTo NextPart

        t = parts(i).Thickness
        w = parts(i).Width
        l = parts(i).Length
        fp = w * l

        If t < PLATE_MIN_THICKNESS Then GoTo NextPart
        If fp < PLATE_MIN_FOOTPRINT Then GoTo NextPart

        ' First separate true pot blocks.
        If IsPotBlockGeometry(t, w, l, maxFp) Then

            nPot = nPot + 1
            potList(nPot) = i

        Else

            ' Remaining large/thick plate-like candidates.
            ' In generic imported XT assemblies, raw pre-orientation boxes can
            ' look too thick, so do not reject by apparent thickness here.
            If IsBmsMajorPlateCandidate(i, maxFp) Then
                nOther = nOther + 1
                otherList(nOther) = i
            Else
                LogLine "BMS geometry classifier ignored: idx=" & i & _
                        " comp='" & parts(i).componentName & "'" & _
                        " T/W/L=" & FormatNumberForCsv(t) & "/" & _
                                  FormatNumberForCsv(w) & "/" & _
                                  FormatNumberForCsv(l) & _
                        " fp=" & FormatNumberForCsv(fp)
            End If

        End If

NextPart:
    Next i

    ' Sort remaining major candidates by footprint descending.
    SortIndexArrayByFootprintDesc otherList, nOther

    ' Largest footprint pair = TCP/BCP.
    If nOther >= 2 Then
        Dim clampPair(1 To 2) As Long
        clampPair(1) = otherList(1)
        clampPair(2) = otherList(2)
        AssignPairTopBottom clampPair, 2, gIdxTCP, gIdxBCP
    End If

    ' Next footprint pair = ID/OD holders.
    If nOther >= 4 Then
        Dim holderPair(1 To 2) As Long
        holderPair(1) = otherList(3)
        holderPair(2) = otherList(4)
        AssignPairTopBottom holderPair, 2, gIdxIDH, gIdxODH
    End If

    ' Pots use their own pair.
    If nPot >= 2 Then
        SortIndexArrayByVolumeDesc potList, nPot
        AssignPairTopBottom potList, nPot, gIdxIDP, gIdxODP
    End If

    ' Make the thin clamp labels agree with the holder/pot top side.
    ValidateAndCorrectBmsTcpBcpAgainstHolderPot

    ' Check whether TCP is lighter than BCP.
    ValidateAndLogBmsTcpBcpMass

    LogLine "Geometry plates: TCP=" & gIdxTCP & " BCP=" & gIdxBCP & _
            " IDholder=" & gIdxIDH & " ODholder=" & gIdxODH & _
            " IDpot=" & gIdxIDP & " ODpot=" & gIdxODP

    LogPlateIndexDetails "TCP", gIdxTCP
    LogPlateIndexDetails "BCP", gIdxBCP
    LogPlateIndexDetails "ID HOLDER", gIdxIDH
    LogPlateIndexDetails "OD HOLDER", gIdxODH
    LogPlateIndexDetails "ID POT", gIdxIDP
    LogPlateIndexDetails "OD POT", gIdxODP

    Exit Sub

ErrHandler:
    LogLine "ClassifyPotBlockPlatesFromCad error: " & Err.Description
End Sub

Private Function AxisCenterForBms(ByVal idx As Long, ByVal axis As Integer) As Double
    If idx < 1 Or idx > PartCount Then Exit Function

    Select Case axis
        Case 1
            AxisCenterForBms = parts(idx).AsmCenterX
        Case 2
            AxisCenterForBms = parts(idx).AsmCenterY
        Case Else
            AxisCenterForBms = parts(idx).AsmCenterZ
    End Select
End Function

Private Function DominantAxisBetweenParts(ByVal idxA As Long, ByVal idxB As Long) As Integer
    DominantAxisBetweenParts = 0

    If idxA < 1 Or idxB < 1 Then Exit Function
    If idxA > PartCount Or idxB > PartCount Then Exit Function

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = Abs(parts(idxA).AsmCenterX - parts(idxB).AsmCenterX)
    dy = Abs(parts(idxA).AsmCenterY - parts(idxB).AsmCenterY)
    dz = Abs(parts(idxA).AsmCenterZ - parts(idxB).AsmCenterZ)

    If dy >= dx And dy >= dz Then
        DominantAxisBetweenParts = 2
    ElseIf dz >= dx And dz >= dy Then
        DominantAxisBetweenParts = 3
    Else
        DominantAxisBetweenParts = 1
    End If
End Function

Private Function AxisDeltaBetweenParts(ByVal idxA As Long, ByVal idxB As Long, ByVal axis As Integer) As Double
    AxisDeltaBetweenParts = AxisCenterForBms(idxA, axis) - AxisCenterForBms(idxB, axis)
End Function

Private Sub SwapLongValues(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a
    a = b
    b = t
End Sub

Private Sub ValidateAndCorrectBmsTcpBcpAgainstHolderPot()
On Error GoTo ErrHandler

    If gIdxTCP <= 0 Or gIdxBCP <= 0 Then Exit Sub
    If gIdxTCP > PartCount Or gIdxBCP > PartCount Then Exit Sub

    Dim refTop As Long
    Dim refBot As Long
    Dim refLabel As String

    refTop = 0
    refBot = 0
    refLabel = ""

    If gIdxIDH > 0 And gIdxODH > 0 Then
        refTop = gIdxIDH
        refBot = gIdxODH
        refLabel = "ID/OD HOLDER"
    ElseIf gIdxIDP > 0 And gIdxODP > 0 Then
        refTop = gIdxIDP
        refBot = gIdxODP
        refLabel = "ID/OD POT"
    Else
        LogLine "BMS TCP/BCP validation skipped: no holder or pot reference pair."
        Exit Sub
    End If

    If refTop > PartCount Or refBot > PartCount Then Exit Sub

    Dim ax As Integer
    ax = DominantAxisBetweenParts(refTop, refBot)
    If ax < 1 Or ax > 3 Then Exit Sub

    Dim refDelta As Double
    Dim tcpDelta As Double

    refDelta = AxisDeltaBetweenParts(refTop, refBot, ax)
    tcpDelta = AxisDeltaBetweenParts(gIdxTCP, gIdxBCP, ax)

    If Abs(refDelta) < 0.001 Or Abs(tcpDelta) < 0.001 Then
        LogLine "BMS TCP/BCP validation skipped: tiny axis delta. refDelta=" & _
                FormatNumberForCsv(refDelta) & " tcpDelta=" & FormatNumberForCsv(tcpDelta)
        Exit Sub
    End If

    LogLine "BMS TCP/BCP validation against " & refLabel & _
            ": axis=" & CStr(ax) & _
            " refDelta=" & FormatNumberForCsv(refDelta) & _
            " tcpDelta=" & FormatNumberForCsv(tcpDelta)

    If refDelta * tcpDelta < 0# Then
        LogLine "BMS TCP/BCP validation: TCP/BCP appear flipped relative to " & refLabel & ". Swapping gIdxTCP/gIdxBCP."
        SwapLongValues gIdxTCP, gIdxBCP
    Else
        LogLine "BMS TCP/BCP validation OK: TCP is on same side as " & refLabel & " top."
    End If

    Exit Sub

ErrHandler:
    LogLine "ValidateAndCorrectBmsTcpBcpAgainstHolderPot error: " & Err.Description
End Sub

Private Function BmsHasHolderOrPotSideReference() As Boolean
    BmsHasHolderOrPotSideReference = False

    If gIdxIDH > 0 And gIdxODH > 0 Then
        BmsHasHolderOrPotSideReference = True
        Exit Function
    End If

    If gIdxIDP > 0 And gIdxODP > 0 Then
        BmsHasHolderOrPotSideReference = True
        Exit Function
    End If
End Function

Private Sub ValidateAndLogBmsTcpBcpMass()
On Error GoTo ErrHandler

    If Not BMS_TCP_EXPECT_LIGHTER_THAN_BCP Then Exit Sub

    If gIdxTCP <= 0 Or gIdxBCP <= 0 Then
        LogLine "BMS TCP/BCP mass check skipped: TCP or BCP index missing."
        Exit Sub
    End If

    If gIdxTCP > PartCount Or gIdxBCP > PartCount Then Exit Sub

    Dim mTcp As Double
    Dim mBcp As Double
    Dim diff As Double
    Dim tol As Double
    Dim maxM As Double

    mTcp = parts(gIdxTCP).massValue
    mBcp = parts(gIdxBCP).massValue

    If mTcp <= 0# Or mBcp <= 0# Then
        LogLine "BMS TCP/BCP mass check skipped: mass/volume unavailable. " & _
                "TCP mass=" & FormatNumberForCsv(mTcp) & _
                " BCP mass=" & FormatNumberForCsv(mBcp)
        Exit Sub
    End If

    maxM = mTcp
    If mBcp > maxM Then maxM = mBcp

    tol = maxM * BMS_TCP_BCP_MASS_DIFF_FRAC
    If tol < 0.001 Then tol = 0.001

    diff = mTcp - mBcp

    LogLine "BMS TCP/BCP mass check:"
    LogLine "  TCP idx=" & gIdxTCP & _
            " mass/vol=" & FormatNumberForCsv(mTcp) & _
            " comp='" & parts(gIdxTCP).componentName & "'"
    LogLine "  BCP idx=" & gIdxBCP & _
            " mass/vol=" & FormatNumberForCsv(mBcp) & _
            " comp='" & parts(gIdxBCP).componentName & "'"
    LogLine "  TCP-BCP mass/vol diff=" & FormatNumberForCsv(diff) & _
            " tolerance=" & FormatNumberForCsv(tol)

    If Abs(diff) <= tol Then
        LogLine "BMS TCP/BCP mass check: TCP and BCP are effectively same weight/volume."
        Exit Sub
    End If

    If diff < 0# Then
        LogLine "BMS TCP/BCP mass check OK: TCP is lighter than BCP."
        Exit Sub
    End If

    LogLine "WARNING: BMS TCP/BCP mass check says TCP is HEAVIER than BCP."

    If BMS_TCP_BCP_FORCE_LIGHTER_TCP Then

        If BmsHasHolderOrPotSideReference() Then
            LogLine "WARNING: Not force-swapping TCP/BCP because holder/pot side reference exists. " & _
                    "Holder/pot side is safer than mass when imported CAD has missing material/density."
        Else
            LogLine "BMS TCP/BCP mass check: force-swapping TCP/BCP because no holder/pot side reference exists."
            SwapLongValues gIdxTCP, gIdxBCP
        End If

    End If

    Exit Sub

ErrHandler:
    LogLine "ValidateAndLogBmsTcpBcpMass error: " & Err.Description
End Sub

Private Function IsBmsMajorPlateCandidate(ByVal idx As Long, ByVal maxFp As Double) As Boolean
On Error GoTo ErrHandler

    IsBmsMajorPlateCandidate = False

    If idx <= 0 Or idx > PartCount Then Exit Function
    If IsPyropelPartIndex(idx) Then Exit Function

    Dim t As Double
    Dim w As Double
    Dim l As Double
    Dim fp As Double

    t = parts(idx).Thickness
    w = parts(idx).Width
    l = parts(idx).Length
    fp = w * l

    If t < PLATE_MIN_THICKNESS Then Exit Function
    If fp < PLATE_MIN_FOOTPRINT Then Exit Function

    ' Ignore small skinny rails/straps/hardware.
    If fp < 0.15 * maxFp Then Exit Function

    ' Ignore very skinny pieces.
    If w <= 0# Or l <= 0# Then Exit Function
    If l / w > 8# Then Exit Function

    ' Do not include pots here.
    If IsPotBlockGeometry(t, w, l, maxFp) Then Exit Function

    IsBmsMajorPlateCandidate = True
    Exit Function

ErrHandler:
    IsBmsMajorPlateCandidate = False
End Function

Private Sub SortIndexArrayByFootprintDesc(ByRef idx() As Long, ByVal n As Long)
On Error GoTo ErrHandler

    If n < 2 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    Dim fpI As Double
    Dim fpJ As Double

    For i = 1 To n - 1
        For j = i + 1 To n

            fpI = parts(idx(i)).Width * parts(idx(i)).Length
            fpJ = parts(idx(j)).Width * parts(idx(j)).Length

            If fpJ > fpI Then
                tmp = idx(i)
                idx(i) = idx(j)
                idx(j) = tmp
            End If

        Next j
    Next i

    Exit Sub

ErrHandler:
    LogLine "SortIndexArrayByFootprintDesc error: " & Err.Description
End Sub

Private Sub SortIndexArrayByVolumeDesc(ByRef idx() As Long, ByVal n As Long)
On Error GoTo ErrHandler

    If n < 2 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    For i = 1 To n - 1
        For j = i + 1 To n
            If parts(idx(j)).BBoxVolume > parts(idx(i)).BBoxVolume Then
                tmp = idx(i)
                idx(i) = idx(j)
                idx(j) = tmp
            End If
        Next j
    Next i

    Exit Sub

ErrHandler:
    LogLine "SortIndexArrayByVolumeDesc error: " & Err.Description
End Sub

Private Sub LogPlateIndexDetails(ByVal label As String, ByVal idx As Long)
On Error Resume Next

    If idx <= 0 Or idx > PartCount Then
        LogLine "GEOM " & label & ": NOT FOUND"
        Exit Sub
    End If

    LogLine "GEOM " & label & ": idx=" & idx & _
            " comp='" & parts(idx).componentName & "'" & _
            " T/W/L=" & FormatNumberForCsv(parts(idx).Thickness) & "/" & _
                      FormatNumberForCsv(parts(idx).Width) & "/" & _
                      FormatNumberForCsv(parts(idx).Length) & _
            " BoxDx/Dy/Dz=" & FormatNumberForCsv(parts(idx).BoxDx) & "/" & _
                            FormatNumberForCsv(parts(idx).BoxDy) & "/" & _
                            FormatNumberForCsv(parts(idx).BoxDz) & _
            " CtrX/Y/Z=" & FormatNumberForCsv(parts(idx).AsmCenterX) & "/" & _
                         FormatNumberForCsv(parts(idx).AsmCenterY) & "/" & _
                         FormatNumberForCsv(parts(idx).AsmCenterZ)
End Sub

Private Function IsClampPlateGeometry(ByVal t As Double, _
                                      ByVal w As Double, _
                                      ByVal l As Double, _
                                      ByVal maxFp As Double) As Boolean
    IsClampPlateGeometry = False

    If t <= 0# Or w <= 0# Or l <= 0# Then Exit Function
    If maxFp <= 0# Then Exit Function

    Dim fp As Double
    fp = w * l

    ' Clamp plates are broad plates, not skinny rails/hardware.
    If fp < 0.65 * maxFp Then Exit Function

    ' They are thin relative to footprint.
    If t > CLAMP_THIN_RATIO * l Then Exit Function

    ' Avoid extremely skinny objects being called TCP/BCP.
    If w < 0.55 * l Then Exit Function

    ' Normal BMS clamps are not 0.25" insulation and not 8-12" thick holders.
    If t < 0.5 Then Exit Function
    If t > 3# Then Exit Function

    IsClampPlateGeometry = True
End Function

Private Function IsHolderBlockGeometry(ByVal t As Double, _
                                       ByVal w As Double, _
                                       ByVal l As Double, _
                                       ByVal maxFp As Double) As Boolean
    IsHolderBlockGeometry = False

    If t <= 0# Or w <= 0# Or l <= 0# Then Exit Function

    Dim fp As Double
    fp = w * l

    ' Holders are thick blocks.
    If t < 3# Then Exit Function

    ' Holders can be much smaller than TCP/BCP footprint.
    ' Example: 6.5 x 12.875 compared with 15.875 x 16.
    If maxFp > 0# Then
        If fp < 0.20 * maxFp Then Exit Function
    End If

    ' Holders are elongated, not compact pots.
    If l / w < 1.15 Then Exit Function

    ' Avoid full clamp plates.
    If maxFp > 0# Then
        If fp > 0.75 * maxFp Then Exit Function
    End If

    IsHolderBlockGeometry = True
End Function

' Pots are easy to tell from mold plates:
'   thick (>= 3"), blocky aspect, chunky (not flat), footprint << mold base.
Private Function IsPotBlockGeometry(ByVal t As Double, ByVal w As Double, ByVal l As Double, _
                                    ByVal maxFp As Double) As Boolean
    IsPotBlockGeometry = False
    If t < POT_MIN_THICKNESS Then Exit Function
    If w <= 0# Or l <= 0# Then Exit Function
    If (l / w) > POT_MAX_ASPECT Then Exit Function

    Dim fp As Double, dimMax As Double, dimMin As Double
    fp = w * l
    If maxFp > 0# Then
        If fp >= POT_MAX_FOOTPRINT_FRAC * maxFp Then Exit Function
    End If

    dimMax = t
    If w > dimMax Then dimMax = w
    If l > dimMax Then dimMax = l
    dimMin = t
    If w < dimMin Then dimMin = w
    If l < dimMin Then dimMin = l
    If dimMax <= 0# Then Exit Function
    If (dimMin / dimMax) < POT_MIN_CUBE_RATIO Then Exit Function

    IsPotBlockGeometry = True
End Function

' Assign ID/top vs OD/bottom using the dominant center-separation axis
' (same idea as OrientTcpTopFromCenters). Do NOT assume Z is always top/bottom.
Private Sub AssignPairTopBottom(ByRef lst() As Long, ByVal n As Long, ByRef topIdx As Long, ByRef botIdx As Long)
On Error GoTo ErrHandler

    topIdx = 0
    botIdx = 0

    If n < 1 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    ' Sort by volume descending so the two best candidates come first.
    For i = 1 To n - 1
        For j = i + 1 To n
            If parts(lst(j)).BBoxVolume > parts(lst(i)).BBoxVolume Then
                tmp = lst(i)
                lst(i) = lst(j)
                lst(j) = tmp
            End If
        Next j
    Next i

    Dim a As Long
    Dim b As Long

    a = lst(1)

    If n >= 2 Then
        b = lst(2)
    Else
        topIdx = a
        botIdx = 0
        Exit Sub
    End If

    If a <= 0 Or b <= 0 Then Exit Sub
    If a > PartCount Or b > PartCount Then Exit Sub

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = parts(a).AsmCenterX - parts(b).AsmCenterX
    dy = parts(a).AsmCenterY - parts(b).AsmCenterY
    dz = parts(a).AsmCenterZ - parts(b).AsmCenterZ

    Dim ax As Double
    Dim ay As Double
    Dim az As Double

    ax = Abs(dx)
    ay = Abs(dy)
    az = Abs(dz)

    Dim aIsTop As Boolean

    ' Use the dominant separation axis, same idea as OrientTcpTopFromCenters.
    ' Do NOT assume Z is top/bottom.
    If ay >= ax And ay >= az Then
        aIsTop = (dy >= 0#)
    ElseIf az >= ax And az >= ay Then
        aIsTop = (dz >= 0#)
    Else
        aIsTop = (dx >= 0#)
    End If

    If Not ASSIGN_ID_AS_TOP Then aIsTop = Not aIsTop

    If aIsTop Then
        topIdx = a
        botIdx = b
    Else
        topIdx = b
        botIdx = a
    End If

    LogLine "AssignPairTopBottom: top=" & topIdx & " bottom=" & botIdx & _
            " sep X/Y/Z=" & FormatNumberForCsv(dx) & "/" & _
            FormatNumberForCsv(dy) & "/" & FormatNumberForCsv(dz)

    Exit Sub

ErrHandler:
    LogLine "AssignPairTopBottom error: " & Err.Description
    topIdx = 0
    botIdx = 0
End Sub

Private Function GeometryIndexForStd(ByVal stdName As String) As Long
    Select Case NormalizeKey(stdName)
        Case "TCP": GeometryIndexForStd = gIdxTCP
        Case "BCP": GeometryIndexForStd = gIdxBCP
        Case "IDHOLDER": GeometryIndexForStd = gIdxIDH
        Case "ODHOLDER": GeometryIndexForStd = gIdxODH
        Case "IDPOT", "IDPOTBLOCK": GeometryIndexForStd = gIdxIDP
        Case "ODPOT", "ODPOTBLOCK": GeometryIndexForStd = gIdxODP
    End Select
End Function

Private Function GetPlateDims(ByVal stdName As String, ByVal pipeKeys As String, ByRef usedPart() As Boolean, _
                              ByRef t As Double, ByRef w As Double, ByRef l As Double, ByRef srcOut As String) As Boolean
    ' Prefer CAD finished bbox after CMS view-frame L/W/T axes are applied.
    ' BOM is backup only when CAD has no match — Tempcraft order is remapped
    ' by plate role (never blind L≥W≥T; holders/pots can have T as largest).
    GetPlateDims = False
    Dim ci As Long
    ci = FindPartIndexByKeys(pipeKeys, usedPart)
    If ci > 0 Then
        usedPart(ci) = True
        t = parts(ci).Thickness: w = parts(ci).Width: l = parts(ci).Length
        srcOut = "CAD:" & parts(ci).componentName
        GetPlateDims = True
        Exit Function
    End If
    ' Geometry classification before BOM so bad PDF Stock-Weight dims cannot win.
    Dim gi As Long
    gi = GeometryIndexForStd(stdName)
    If gi > 0 Then
        If gi <= UBound(usedPart) Then usedPart(gi) = True
        t = parts(gi).Thickness: w = parts(gi).Width: l = parts(gi).Length
        srcOut = "CAD-geom:" & parts(gi).componentName
        GetPlateDims = True
        Exit Function
    End If
    Dim bi As Long
    bi = FindBomIndexByStdName(stdName)
    If bi > 0 Then
        If BomRows(bi).hasDims Then
            t = BomRows(bi).BomThickness: w = BomRows(bi).BomWidth: l = BomRows(bi).BomLength
            If BomRows(bi).BomIsTempcraftOrder Then
                MapTempcraftBomDimsToCmsSteel stdName, t, w, l
            End If
            If BomDimsLookLikeStockWeight(t, w, l) Then
                LogLine "GetPlateDims: rejecting BOM dims that look like Stock Weight for " & stdName & _
                        " (T=" & t & " W=" & w & " L=" & l & ")"
            Else
                srcOut = "BOM:" & BomRows(bi).Description
                GetPlateDims = True
                Exit Function
            End If
        End If
    End If
End Function

' Tempcraft Base BOM finished sizes are listed as Lth, Wth/O.D., Hgt/I.D. in file
' order — NOT CMS Thickness/Width/Length. Map by plate role (from correct J000
' steel sheets + shop dimensioned DXF):
'   TCP/BCP:     smallest=T; of remaining, larger=L smaller=W
'   Holders:     Lth=T, Wth=L, Hgt=W   (e.g. 6.875 x 7.000 x 13.875 → T6.875 W13.875 L7)
'   Pot blocks:  Lth=W, Wth=L, Hgt=T   (e.g. 5.500 x 5.500 x 6.875 → T6.875 W5.5 L5.5)
Private Sub MapTempcraftBomDimsToCmsSteel(ByVal stdName As String, _
                                          ByRef t As Double, ByRef w As Double, ByRef l As Double)
    Dim a As Double, b As Double, c As Double
    a = t: b = w: c = l
    If a <= 0 Or b <= 0 Or c <= 0 Then Exit Sub
    Select Case NormalizeKey(stdName)
        Case "TCP", "BCP"
            SortThreeDimensions a, b, c, l, w, t
        Case "IDHOLDER", "ODHOLDER"
            t = a: l = b: w = c
        Case "IDPOT", "IDPOTBLOCK", "ODPOT", "ODPOTBLOCK"
            w = a: l = b: t = c
        Case Else
            ' Unknown extras: keep prior L≥W≥T behavior.
            SortThreeDimensions a, b, c, l, w, t
    End Select
End Sub

Private Function FindBomIndexByStdName(ByVal stdName As String) As Long
    Dim i As Long, k As String
    k = NormalizeKey(stdName)
    For i = 1 To BomCount
        If NormalizeKey(BomRows(i).quoteName) = k Then FindBomIndexByStdName = i: Exit Function
    Next i

    Select Case k
        Case "IDPOT"
            For i = 1 To BomCount
                If NormalizeKey(BomRows(i).quoteName) = "IDPOTBLOCK" Then FindBomIndexByStdName = i: Exit Function
            Next i
        Case "IDPOTBLOCK"
            For i = 1 To BomCount
                If NormalizeKey(BomRows(i).quoteName) = "IDPOT" Then FindBomIndexByStdName = i: Exit Function
            Next i
        Case "ODPOT"
            For i = 1 To BomCount
                If NormalizeKey(BomRows(i).quoteName) = "ODPOTBLOCK" Then FindBomIndexByStdName = i: Exit Function
            Next i
        Case "ODPOTBLOCK"
            For i = 1 To BomCount
                If NormalizeKey(BomRows(i).quoteName) = "ODPOT" Then FindBomIndexByStdName = i: Exit Function
            Next i
    End Select
End Function

Private Function CeilToQuarter(ByVal v As Double) As Double
    If v <= 0 Then Exit Function
    Dim n As Double
    n = Int(v / 0.25)
    If (v - n * 0.25) > 0.0000001 Then n = n + 1
    CeilToQuarter = Round(n * 0.25, 3)
End Function

' Round UP to the nearest 0.05" (so values end in 0 or 5): 7.125 -> 7.15, 6.227 -> 6.25.
Private Function RoundUpToNickel(ByVal v As Double) As Double
    If v <= 0 Then Exit Function
    Dim n As Double
    n = Int(v / 0.05)
    If (v / 0.05 - n) > 0.0000001 Then n = n + 1
    RoundUpToNickel = Round(n * 0.05, 2)
End Function

' Steel STOCK thickness for the quote: add STEEL_THICKNESS_ALLOWANCE (0.250") to
' the finished thickness, then round UP to the nearest 0.05".
Private Function SteelStockThickness(ByVal finished As Double) As Double
    If finished <= 0 Then Exit Function
    SteelStockThickness = RoundUpToNickel(finished + STEEL_THICKNESS_ALLOWANCE)
End Function

' Write today's date and the job ref number next to their labels on every
' sheet of a workbook (DATE -> today; REF #/JOB # -> C-number). Robust to
' which cell the value lives in: it writes to the first non-"X" cell to the
' right of the label.
Private Sub StampWorkbookDateAndRef(ByVal xlWb As Object, ByVal cnumFmt As String)
On Error Resume Next
    Dim ws As Object
    For Each ws In xlWb.Worksheets
        SetCellRightOfLabel ws, Array("DATE"), Date
        SetCellRightOfLabel ws, Array("REF #", "REF#", "REF", "JOB #", "JOB#", "JOB #"), cnumFmt
    Next ws
End Sub

Private Function SetCellRightOfLabel(ByVal ws As Object, ByVal labels As Variant, ByVal value As Variant) As Boolean
On Error Resume Next
    Dim r As Long, c As Long, t As String, li As Long, lab As String
    For r = 1 To 20
        For c = 1 To 14
            t = UCase(Trim(CStr(ws.Cells(r, c).value)))
            If t <> "" Then
                For li = LBound(labels) To UBound(labels)
                    lab = UCase(Trim(CStr(labels(li))))
                    If lab <> "" Then
                        If t = lab Or Left(t, Len(lab)) = lab Then
                            Dim cc As Long, rv As String
                            cc = c + 1
                            Do While cc <= 16
                                rv = UCase(Trim(CStr(ws.Cells(r, cc).value)))
                                If rv <> "X" Then Exit Do
                                cc = cc + 1
                            Loop
                            ws.Cells(r, cc).value = value
                            SetCellRightOfLabel = True
                            Exit Function
                        End If
                    End If
                Next li
            End If
        Next c
    Next r
End Function

Private Function FormatRefNumber() As String
    Dim n As String
    n = UCase(Trim(CurrentJobNumber))
    If Left(n, 1) = "C" And Len(n) > 1 And Mid(n, 2, 1) <> "-" Then
        FormatRefNumber = "C-" & Mid(n, 2)
    Else
        FormatRefNumber = n
    End If
End Function

Private Function FindPartIndexByKeys(ByVal pipeKeys As String, ByRef usedPart() As Boolean) As Long
    Dim i As Long, bestIdx As Long, bestVol As Double
    bestIdx = 0: bestVol = -1#
    For i = 1 To PartCount
        If usedPart(i) = False Then
            If Not IsPyropelPartIndex(i) And ContainsAnyPipeKey(parts(i).componentName, pipeKeys) Then
                If parts(i).BBoxVolume > bestVol Then bestVol = parts(i).BBoxVolume: bestIdx = i
            End If
        End If
    Next i
    FindPartIndexByKeys = bestIdx
End Function

Private Function IsStdSixName(ByVal stdName As String) As Boolean
    Select Case NormalizeKey(stdName)
        Case "TCP", "BCP", "IDHOLDER", "ODHOLDER", "IDPOT", "IDPOTBLOCK", "ODPOT", "ODPOTBLOCK": IsStdSixName = True
    End Select
End Function

' Any BOM line that is 4140 and is NOT one of the six standard pot plates
' (e.g. Pullcore Stop, Flipper Cam Cover Plate). These get listed too.
Private Function CollectExtra4140Parts(ByRef exDesc() As String, ByRef exQty() As Long, _
        ByRef ext() As Double, ByRef exW() As Double, ByRef exL() As Double) As Long
    Dim n As Long
    n = 0
    ReDim exDesc(1 To 60): ReDim exQty(1 To 60)
    ReDim ext(1 To 60): ReDim exW(1 To 60): ReDim exL(1 To 60)
    Dim i As Long
    For i = 1 To BomCount
        If BomRows(i).hasDims Then
            If InStr(NormalizeSteelType(BomRows(i).material), "4140") > 0 Then
                If Not IsStdSixName(BomRows(i).quoteName) Then
                    If Not IsHardwareName(BomRows(i).Description) Then
                        n = n + 1
                        exDesc(n) = ProperCaseText(BomRows(i).Description)
                        exQty(n) = BomRows(i).Quantity
                        ext(n) = BomRows(i).BomThickness
                        exW(n) = BomRows(i).BomWidth
                        exL(n) = BomRows(i).BomLength
                        If BomRows(i).BomIsTempcraftOrder Then
                            ' Extras are not in the six-plate map; use L≥W≥T for steel sheet.
                            MapTempcraftBomDimsToCmsSteel "", ext(n), exW(n), exL(n)
                        End If
                    End If
                End If
            End If
        End If
    Next i
    CollectExtra4140Parts = n
End Function

Private Sub FillQuoteWorkbookFromBoundingBox()
On Error GoTo ErrHandler
    Dim templatePath As String
    templatePath = FindQuoteTemplateAnywhere()
    If templatePath = "" Then
        LogLine "Quote template not found in Trust (" & TRUSTED_FOLDER & ") or Downloads; skipping Quote fill."
        Exit Sub
    End If
    Dim quotePath As String
    quotePath = CopyTemplateToJobFolder(templatePath)
    If quotePath = "" Then quotePath = templatePath
    LogLine "Quote template: " & templatePath
    LogLine "Quote workbook (job copy): " & quotePath

    Dim usedPart() As Boolean
    ReDim usedPart(1 To IIf(PartCount < 1, 1, PartCount))

    Dim keys(1 To 6) As String
    Dim rowN(1 To 6) As Long
    Dim stdN(1 To 6) As String
    keys(1) = KEYS_TCP:       rowN(1) = 22: stdN(1) = "TCP"
    keys(2) = KEYS_BCP:       rowN(2) = 23: stdN(2) = "BCP"
    keys(3) = ID_HOLDER_KEYS: rowN(3) = 31: stdN(3) = "ID HOLDER"
    keys(4) = OD_HOLDER_KEYS: rowN(4) = 32: stdN(4) = "OD HOLDER"
    keys(5) = KEYS_ID_POT:    rowN(5) = 33: stdN(5) = "ID POT"
    keys(6) = KEYS_OD_POT:    rowN(6) = 34: stdN(6) = "OD POT"

    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    Set xlWb = xlApp.Workbooks.Open(quotePath)
    xlApp.Calculation = -4135  ' manual calc - valid only after a workbook is open
    On Error Resume Next
    Set xlWs = xlWb.Worksheets(QUOTE_SHEET_NAME)
    On Error GoTo ErrHandler
    If xlWs Is Nothing Then
        xlWb.Close False: xlApp.Quit
        LogLine "QuoteWorksheet missing in: " & quotePath
        Exit Sub
    End If

    ' Clear the #2 (4140) block plate rows first (keep the column-A labels) so
    ' rows for parts this job does not have - e.g. the flipper plates - do not
    ' keep a stale quantity of 1.
    Dim cr As Long
    For cr = 22 To 34
        xlWs.Cells(cr, 3).value = ""
        xlWs.Cells(cr, 4).value = ""
        xlWs.Cells(cr, 5).value = ""
        xlWs.Cells(cr, 6).value = ""
    Next cr

    Dim i As Long
    Dim tt As Double, ww As Double, ll As Double, src As String
    Dim qt As Double, qw As Double, ql As Double
    For i = 1 To 6
        If GetPlateDims(stdN(i), keys(i), usedPart, tt, ww, ll, src) Then
            qt = tt: qw = RoundUpToNickel(ww): ql = RoundUpToNickel(ll)
            If QUOTE_ROUND_UP_TO_QUARTER Then
                qt = SteelStockThickness(tt)
            End If
            xlWs.Cells(rowN(i), 3).value = 1
            xlWs.Cells(rowN(i), 4).value = qt
            xlWs.Cells(rowN(i), 5).value = qw
            xlWs.Cells(rowN(i), 6).value = ql
            LogLine "Quote row " & rowN(i) & " (" & stdN(i) & ") <- " & src & _
                    " stock T=" & qt & " W=" & qw & " L=" & ql
        Else
            LogLine "Quote row " & rowN(i) & " (" & stdN(i) & ") : no CAD or BOM match"
        End If
    Next i

    ' Extra 4140 parts (Pullcore Stop, Flipper Cam Cover Plate, etc.) into spare rows.
    Dim exDesc() As String, exQty() As Long, ext() As Double, exW() As Double, exL() As Double
    Dim nx As Long
    nx = CollectExtra4140Parts(exDesc, exQty, ext, exW, exL)
    Dim spare As Variant
    spare = Array(26, 27, 28, 29, 24, 25, 30)
    Dim sp As Long, exr As Long, rr As Long, et As Double, ew As Double, el As Double
    sp = 0
    For exr = 1 To nx
        If sp > UBound(spare) Then
            LogLine "Quote: no spare row for extra 4140 part: " & exDesc(exr)
        Else
            rr = CLng(spare(sp))
            et = ext(exr): ew = exW(exr): el = exL(exr)
            If QUOTE_ROUND_UP_TO_QUARTER Then et = SteelStockThickness(et)
            xlWs.Cells(rr, 1).value = exDesc(exr)
            xlWs.Cells(rr, 3).value = exQty(exr)
            xlWs.Cells(rr, 4).value = et
            xlWs.Cells(rr, 5).value = ew
            xlWs.Cells(rr, 6).value = el
            LogLine "Quote extra 4140 row " & rr & " <- " & exDesc(exr) & " qty " & exQty(exr)
            sp = sp + 1
        End If
    Next exr

    If PcCount > 0 Then WritePullcoreCategoryToSheet xlWs
    WritePullcoreTotalToSummary xlWs
    If PpCount > 0 Then
        WritePurchasedToComponentsArea xlWs
        WritePurchasedCategoryToSheet xlWs
    End If
    StampWorkbookDateAndRef xlWb, FormatRefNumber
    xlWb.Save
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    LogLine "Quote workbook saved."
    Exit Sub
ErrHandler:
    LogLine "FillQuoteWorkbookFromBoundingBox error: " & Err.Description
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

Private Sub FillJ000SteelSheet()
On Error GoTo ErrHandler
    Dim templatePath As String
    templatePath = FindJ000TemplateAnywhere()
    If templatePath = "" Then
        LogLine "J000 steel-sheet template not found in Trust (" & TRUSTED_FOLDER & ") or Downloads; skipping."
        Exit Sub
    End If
    Dim jPath As String
    jPath = CopyTemplateToJobFolder(templatePath)
    If jPath = "" Then jPath = templatePath
    LogLine "J000 template: " & templatePath
    LogLine "J000 workbook (job copy): " & jPath

    Dim usedPart() As Boolean
    ReDim usedPart(1 To IIf(PartCount < 1, 1, PartCount))
    Dim names(1 To 6) As String
    Dim keys(1 To 6) As String
    Dim stdN(1 To 6) As String
    names(1) = "TCP":       keys(1) = KEYS_TCP:       stdN(1) = "TCP"
    names(2) = "ID Holder": keys(2) = ID_HOLDER_KEYS: stdN(2) = "ID HOLDER"
    names(3) = "OD Holder": keys(3) = OD_HOLDER_KEYS: stdN(3) = "OD HOLDER"
    names(4) = "ID Pot":    keys(4) = KEYS_ID_POT:    stdN(4) = "ID POT"
    names(5) = "OD Pot":    keys(5) = KEYS_OD_POT:    stdN(5) = "OD POT"
    names(6) = "BCP":       keys(6) = KEYS_BCP:       stdN(6) = "BCP"

    ' Resolve finished dims once (CAD bbox first, else BOM row).
    Dim ft(1 To 6) As Double, fw(1 To 6) As Double, fl(1 To 6) As Double
    Dim found(1 To 6) As Boolean
    Dim i As Long, tt As Double, ww As Double, ll As Double, src As String
    For i = 1 To 6
        found(i) = GetPlateDims(stdN(i), keys(i), usedPart, tt, ww, ll, src)
        If found(i) Then
            ft(i) = tt: fw(i) = ww: fl(i) = ll
            LogLine "J000 " & names(i) & " <- " & src & " (T=" & tt & " W=" & ww & " L=" & ll & ")"
        Else
            LogLine "J000 " & names(i) & " : no CAD or BOM match"
        End If
    Next i

    ' Paired plates share the same footprint (W×L); only Thickness differs.
    HarmonizeBmsSteelPairFootprints found, ft, fw, fl

    Dim exDesc() As String, exQty() As Long, ext() As Double, exW() As Double, exL() As Double
    Dim nx As Long
    nx = CollectExtra4140Parts(exDesc, exQty, ext, exW, exL)

    Dim xlApp As Object, xlWb As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    Set xlWb = xlApp.Workbooks.Open(jPath)
    xlApp.Calculation = -4135  ' manual calc - valid only after a workbook is open

    Dim sheetNames As Variant
    sheetNames = Array("Steel Order", "Machining Sheet")
    Dim sName As Variant
    Dim ws As Object
    Dim writeRow As Long
    For Each sName In sheetNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = xlWb.Worksheets(CStr(sName))
        On Error GoTo ErrHandler
        If Not ws Is Nothing Then
            writeRow = 19
            For i = 1 To 6
                If found(i) Then
                    ws.Cells(writeRow, 1).value = 1
                    ws.Cells(writeRow, 2).value = names(i)
                    ' CMS steel sheet: C=Thickness, E=Width, G=Length
                    ' from CMS view-frame axes (TOP X=W, TOP Y=L, RIGHT X=T) — not L≥W≥T sort.
                    ws.Cells(writeRow, 3).value = ft(i)   ' C = Thickness / Height
                    ws.Cells(writeRow, 5).value = fw(i)   ' E = Width
                    ws.Cells(writeRow, 7).value = fl(i)   ' G = Length
                    ws.Cells(writeRow, 8).value = POTBLOCK_STEEL_TYPE
                    writeRow = writeRow + 1
                End If
            Next i
            Dim ex As Long
            For ex = 1 To nx
                ws.Cells(writeRow, 1).value = exQty(ex)
                ws.Cells(writeRow, 2).value = exDesc(ex)
                ws.Cells(writeRow, 3).value = ext(ex)
                ws.Cells(writeRow, 5).value = exW(ex)
                ws.Cells(writeRow, 7).value = exL(ex)
                ws.Cells(writeRow, 8).value = POTBLOCK_STEEL_TYPE
                writeRow = writeRow + 1
            Next ex
            LogLine "Filled '" & CStr(sName) & "' rows 19.." & (writeRow - 1)
        End If
    Next sName

    StampWorkbookDateAndRef xlWb, FormatRefNumber
    ' REF # in I11 (right next to its label); clear the stray J12 placement.
    On Error Resume Next
    Dim refSh As Object, refNm As Variant
    For Each refNm In Array("Steel Order", "Machining Sheet")
        Set refSh = Nothing
        Set refSh = xlWb.Sheets(CStr(refNm))
        If Not refSh Is Nothing Then
            refSh.Cells(11, 9).Value = FormatRefNumber
            refSh.Cells(12, 10).ClearContents
        End If
    Next refNm
    On Error GoTo ErrHandler
    xlWb.Save
    xlWb.Close False
    xlApp.Quit
    Set xlWb = Nothing: Set xlApp = Nothing
    LogLine "J000 steel sheet saved."
    Exit Sub
ErrHandler:
    LogLine "FillJ000SteelSheet error: " & Err.Description
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

' Make paired BMS plates share one footprint so TCP/BCP, holders, and pots are consistent.
' Indexes: 1=TCP 2=IDH 3=ODH 4=IDP 5=ODP 6=BCP
Private Sub HarmonizeBmsSteelPairFootprints(ByRef found() As Boolean, _
                                            ByRef ft() As Double, ByRef fw() As Double, ByRef fl() As Double)
On Error Resume Next
    HarmonizeOneBmsFootprintPair found, ft, fw, fl, 1, 6, "TCP/BCP"
    HarmonizeOneBmsFootprintPair found, ft, fw, fl, 2, 3, "ID/OD Holder"
    HarmonizeOneBmsFootprintPair found, ft, fw, fl, 4, 5, "ID/OD Pot"
End Sub

Private Sub HarmonizeOneBmsFootprintPair(ByRef found() As Boolean, _
                                         ByRef ft() As Double, ByRef fw() As Double, ByRef fl() As Double, _
                                         ByVal a As Long, ByVal b As Long, ByVal label As String)
    If Not found(a) Or Not found(b) Then Exit Sub
    Dim wa As Double, la As Double, wb As Double, lb As Double
    wa = fw(a): la = fl(a): wb = fw(b): lb = fl(b)

    ' Detect W↔L swap between the pair (same sizes, axes flipped).
    If Abs(wa - lb) < 0.2 And Abs(la - wb) < 0.2 And Abs(wa - wb) > 0.25 Then
        LogLine "J000 " & label & ": footprint W/L swapped between pair — aligning to first plate."
        fw(b) = wa: fl(b) = la
        Exit Sub
    End If

    ' Thin clamp plates: Length should be the longer footprint side.
    If a = 1 And b = 6 Then
        If ft(a) < 3# And ft(b) < 3# Then
            If fw(a) > fl(a) Then
                Dim tmp As Double
                tmp = fw(a): fw(a) = fl(a): fl(a) = tmp
                LogLine "J000 " & label & ": TCP had W>L — swapped to L>=W for footprint."
            End If
            fw(b) = fw(a): fl(b) = fl(a)
            Exit Sub
        End If
    End If

    ' Holders/pots: share average footprint when close; if one side looks
    ' like thickness leaked into W, prefer the other's footprint.
    Dim avgW As Double, avgL As Double
    avgW = (wa + wb) / 2#
    avgL = (la + lb) / 2#
    If Abs(wa - wb) < 0.35 And Abs(la - lb) < 0.35 Then
        fw(a) = Round(avgW, DIM_DECIMALS): fl(a) = Round(avgL, DIM_DECIMALS)
        fw(b) = fw(a): fl(b) = fl(a)
        LogLine "J000 " & label & ": footprint averaged to W=" & fw(a) & " L=" & fl(a)
        Exit Sub
    End If

    ' Prefer the footprint whose sides are both larger than either thickness
    ' (true face dims), otherwise keep each.
    Dim aOk As Boolean, bOk As Boolean
    aOk = (wa > ft(a) + 0.05 And la > ft(a) + 0.05) Or (Abs(wa - la) < 0.15)
    bOk = (wb > ft(b) + 0.05 And lb > ft(b) + 0.05) Or (Abs(wb - lb) < 0.15)
    If aOk And Not bOk Then
        fw(b) = wa: fl(b) = la
        LogLine "J000 " & label & ": copied footprint from first plate onto second."
    ElseIf bOk And Not aOk Then
        fw(a) = wb: fl(a) = lb
        LogLine "J000 " & label & ": copied footprint from second plate onto first."
    End If
End Sub

Private Function CopyTemplateToJobFolder(ByVal templatePath As String) As String
On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim prefix As String
    prefix = JobBaseName
    If prefix = "" Then prefix = CurrentJobNumber
    Dim destPath As String
    destPath = GetUniqueFilePath(CurrentJobFolder & "\" & prefix & " " & fso.GetFileName(templatePath))
    fso.CopyFile templatePath, destPath, True
    CopyTemplateToJobFolder = destPath
    Exit Function
ErrHandler:
    LogLine "CopyTemplateToJobFolder error: " & Err.Description
    CopyTemplateToJobFolder = ""
End Function

Private Function FindQuoteWorkbookInJobFolder(ByVal jobFolder As String) As String
    ' Internal CMS quote: name has QUOTE + STEEL + GRIND. Prefer the exact
    ' "Quote_Steel_Grinding" over dated copies like "...-4-1-2024-rj".
    FindQuoteWorkbookInJobFolder = PickBestTemplate(jobFolder, "QUOTE_STEEL_GRINDING", "QUOTE|STEEL|GRIND")
End Function

' Try the trusted templates folder first, then Downloads.
Private Function FindQuoteTemplateAnywhere() As String
    FindQuoteTemplateAnywhere = FindQuoteWorkbookInJobFolder(TRUSTED_FOLDER)
    If FindQuoteTemplateAnywhere = "" Then FindQuoteTemplateAnywhere = FindQuoteWorkbookInJobFolder(DOWNLOADS_FOLDER)
End Function

Private Function FindJ000TemplateAnywhere() As String
    FindJ000TemplateAnywhere = FindJ000WorkbookInJobFolder(TRUSTED_FOLDER)
    If FindJ000TemplateAnywhere = "" Then FindJ000TemplateAnywhere = FindJ000WorkbookInJobFolder(DOWNLOADS_FOLDER)
End Function

Private Function FindJ000WorkbookInJobFolder(ByVal jobFolder As String) As String
    ' Steel order / machining sheet: name has STEEL + SHEET. Prefer the exact
    ' "STEEL_SHEET" over copies like "J000-STEEL_SHEET-std".
    FindJ000WorkbookInJobFolder = PickBestTemplate(jobFolder, "STEEL_SHEET", "STEEL|SHEET")
End Function

' Pick the best-matching template under a folder. exactStem (no extension) scores
' highest; otherwise every pipe-separated word in needWords must appear in the
' name. Ties break toward the shortest file name (fewest date/initials suffixes).
Private Function PickBestTemplate(ByVal folder As String, ByVal exactStem As String, ByVal needWords As String) As String
On Error GoTo eh
    PickBestTemplate = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function
    Dim best As String, bestScore As Long
    best = "": bestScore = 0
    RankTemplatesRecursive fso.GetFolder(folder), UCase(exactStem), UCase(needWords), best, bestScore
    PickBestTemplate = best
    Exit Function
eh:
    PickBestTemplate = ""
End Function

Private Sub RankTemplatesRecursive(ByVal folder As Object, ByVal exactStem As String, _
                                   ByVal needWords As String, ByRef best As String, ByRef bestScore As Long)
On Error Resume Next
    If UCase(folder.Name) = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim words() As String
    words = Split(needWords, "|")
    Dim file As Object, nm As String, stem As String, ext As String
    Dim ok As Boolean, j As Long, sc As Long
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.path))
        If (ext = "xls" Or ext = "xlsx" Or ext = "xlsm") And Left(file.Name, 2) <> "~$" Then
            nm = UCase(file.Name)
            stem = UCase(fso.GetBaseName(file.path))
            ok = True
            For j = LBound(words) To UBound(words)
                If words(j) <> "" Then
                    If InStr(nm, words(j)) = 0 Then ok = False
                End If
            Next j
            If ok Then
                If stem = exactStem Then sc = 3 Else sc = 1
                If sc > bestScore Then
                    best = file.path: bestScore = sc
                ElseIf sc = bestScore And best <> "" Then
                    If Len(file.Name) < Len(fso.GetFileName(best)) Then best = file.path
                End If
            End If
        End If
    Next file
    Dim sub1 As Object
    For Each sub1 In folder.SubFolders
        RankTemplatesRecursive sub1, exactStem, needWords, best, bestScore
    Next sub1
End Sub

' ============================================================
' ============================================================
' STANDARD (NON-POT) MOLD BASE
' Identify plates from footprint + stack position, name them by
' the stack-up rules, and fill the A-36 (#1) and P20 (#3) Quote
' blocks plus the J000 steel sheet. Works with no BOM.
' ============================================================
' ============================================================

Private Function PartAxisCenter(ByVal idx As Long, ByVal axis As Integer) As Double
    Select Case axis
        Case 1: PartAxisCenter = parts(idx).AsmCenterX
        Case 2: PartAxisCenter = parts(idx).AsmCenterY
        Case Else: PartAxisCenter = parts(idx).AsmCenterZ
    End Select
End Function

Private Sub StdSortByAxisDesc(ByRef idx() As Long, ByVal n As Long, ByVal axis As Integer)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If PartAxisCenter(idx(j), axis) > PartAxisCenter(idx(i), axis) Then t = idx(i): idx(i) = idx(j): idx(j) = t
        Next j
    Next i
End Sub

Private Sub StdReverse(ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, t As Long
    For i = 1 To n \ 2
        t = idx(i): idx(i) = idx(n - i + 1): idx(n - i + 1) = t
    Next i
End Sub

' Count full-footprint structural plates (standard stacks have many; pot-blocks ~2).
Private Function CountFullFootprintPlates() As Long
    CountFullFootprintPlates = 0
    If PartCount < 1 Then Exit Function
    Dim baseFoot As Double, fp As Double, i As Long
    baseFoot = 0#
    For i = 1 To PartCount
        fp = parts(i).Width * parts(i).Length
        If fp > baseFoot Then baseFoot = fp
    Next i
    If baseFoot <= 0# Then Exit Function
    For i = 1 To PartCount
        If parts(i).Thickness >= STD_MIN_PLATE_THICKNESS Then
            If parts(i).Width * parts(i).Length >= (1 - STD_FOOTPRINT_TOL) * baseFoot Then
                CountFullFootprintPlates = CountFullFootprintPlates + 1
            End If
        End If
    Next i
End Function

' Detect whether this job is a standard base (vs pot/holder block).
' BMS/pot-block signals must win before BOM standard-plate detection so BMS
' jobs use EnsureCmsTopOrientationFromMatchedTcpBcp (gemini1), not SetStandardBaseOrientation.
Private Function DetectBaseTypeIsStandard() As Boolean
On Error GoTo ErrHandler

    If UCase(BASE_TYPE_MODE) = "STANDARD" Then
        DetectBaseTypeIsStandard = True
        LogLine "Base type forced by BASE_TYPE_MODE=STANDARD"
        Exit Function
    End If

    If UCase(BASE_TYPE_MODE) = "POT" Then
        DetectBaseTypeIsStandard = False
        LogLine "Base type forced by BASE_TYPE_MODE=POT"
        Exit Function
    End If

    Dim hardBmsName As Boolean
    Dim bmsBomRoles As Long
    Dim stdBomRoles As Long
    Dim stdStrongCad As Long
    Dim stdCadHits As Long
    Dim nFull As Long

    Dim nFullThin As Long
    Dim nThickInner As Long
    Dim nPotLike As Long
    Dim nThinSheet As Long

    Dim bmsGeo As Boolean
    Dim stdGeo As Boolean

    Dim bmsScore As Long
    Dim stdScore As Long

    hardBmsName = HasHardBmsNameSignal()
    bmsBomRoles = CountDistinctBmsBomRoles()
    stdBomRoles = CountDistinctStandardBomRoles()
    stdStrongCad = CountPcsStrongPlateNameHits()
    stdCadHits = CountPcsStandardPlateNameHits()
    nFull = CountFullFootprintPlates()

    GetBaseTypeGeometryStats nFullThin, nThickInner, nPotLike, nThinSheet

    ' BMS holder/pot geometry:
    ' Usually only 1-2 broad thin clamp/smed plates, plus thick inner holders/pots.
    bmsGeo = False
    If nFull < 5 Then
        If nFullThin <= 3 And nThickInner >= 2 Then
            If nPotLike >= 2 Or (nPotLike >= 1 And nThinSheet >= 1) Then
                bmsGeo = True
            End If
        End If
    End If

    ' Standard geometry:
    ' 3+ full-footprint plates is a normal PCS/standard mold-base signal.
    stdGeo = False
    If nFull >= 3 Then stdGeo = True
    If stdStrongCad >= 2 Then stdGeo = True

    ' ------------------------------
    ' Score BMS
    ' ------------------------------
    If hardBmsName Then bmsScore = bmsScore + 80

    If bmsBomRoles >= 2 Then
        bmsScore = bmsScore + 80
    ElseIf bmsBomRoles = 1 Then
        bmsScore = bmsScore + 25
    End If

    If bmsGeo Then bmsScore = bmsScore + 70

    ' ------------------------------
    ' Score STANDARD / PCS
    ' ------------------------------
    If stdStrongCad >= 2 Then
        stdScore = stdScore + 90
    ElseIf stdStrongCad = 1 Then
        stdScore = stdScore + 35
    End If

    If stdCadHits >= 3 Then stdScore = stdScore + 35
    If stdBomRoles >= 2 Then stdScore = stdScore + 60

    If nFull >= 3 Then stdScore = stdScore + 45
    If nFull >= 5 Then stdScore = stdScore + 35

    If stdGeo Then stdScore = stdScore + 25

    ' Big standard stacks argue against BMS.
    If nFull >= 5 Then bmsScore = bmsScore - 30

    LogLine "Base type decision signals:"
    LogLine "  hardBmsName=" & CStr(hardBmsName)
    LogLine "  bmsBomRoles=" & CStr(bmsBomRoles)
    LogLine "  stdBomRoles=" & CStr(stdBomRoles)
    LogLine "  stdStrongCad=" & CStr(stdStrongCad)
    LogLine "  stdCadHits=" & CStr(stdCadHits)
    LogLine "  nFull=" & CStr(nFull)
    LogLine "  geometry nFullThin=" & CStr(nFullThin) & _
            " nThickInner=" & CStr(nThickInner) & _
            " nPotLike=" & CStr(nPotLike) & _
            " nThinSheet=" & CStr(nThinSheet)
    LogLine "  bmsGeo=" & CStr(bmsGeo) & " stdGeo=" & CStr(stdGeo)
    LogLine "  BMS score=" & CStr(bmsScore) & " STANDARD score=" & CStr(stdScore)

    ' ============================================================
    ' Final decision rules
    ' ============================================================

    ' Strong BMS BOM wins unless strong PCS tokens exist.
    If bmsBomRoles >= 2 And stdStrongCad < 2 Then
        DetectBaseTypeIsStandard = False
        LogLine "Base type selected: BMS/POT — BOM has multiple BMS holder/pot roles."
        Exit Function
    End If

    ' Strong BMS geometry wins unless strong PCS tokens/full standard stack exists.
    If bmsGeo And stdStrongCad < 2 And nFull < 4 Then
        DetectBaseTypeIsStandard = False
        LogLine "Base type selected: BMS/POT — geometry matches holder/pot base."
        Exit Function
    End If

    ' Strong PCS names win when there is no real BMS evidence.
    If stdStrongCad >= 2 And bmsBomRoles = 0 And Not bmsGeo Then
        DetectBaseTypeIsStandard = True
        LogLine "Base type selected: STANDARD — strong PCS A/B/ejector/SC/rail CAD names."
        Exit Function
    End If

    ' Hard BMS name wins when BMS score is close enough.
    If hardBmsName Then
        If bmsScore >= stdScore - 20 Then
            DetectBaseTypeIsStandard = False
            LogLine "Base type selected: BMS/POT — hard BMS name signal."
            Exit Function
        Else
            LogLine "WARNING: hard BMS name signal exists, but standard evidence is stronger. Selecting STANDARD."
        End If
    End If

    ' Standard full-stack geometry.
    If nFull >= 3 And Not bmsGeo Then
        DetectBaseTypeIsStandard = True
        LogLine "Base type selected: STANDARD — 3+ full-footprint plates and no BMS geometry."
        Exit Function
    End If

    ' Score comparison.
    If stdScore > bmsScore Then
        DetectBaseTypeIsStandard = True
        LogLine "Base type selected: STANDARD — score comparison."
        Exit Function
    End If

    If bmsScore > stdScore Then
        DetectBaseTypeIsStandard = False
        LogLine "Base type selected: BMS/POT — score comparison."
        Exit Function
    End If

    ' Tie/no evidence: default standard.
    DetectBaseTypeIsStandard = True
    LogLine "Base type selected: STANDARD — default tie/no strong BMS evidence."

    Exit Function

ErrHandler:
    LogLine "DetectBaseTypeIsStandard error: " & Err.Description
    DetectBaseTypeIsStandard = True
End Function

Private Function BuildBaseTypeBlob() As String
On Error Resume Next

    Dim blob As String

    blob = UCase$(CurrentJobNumber & "|" & _
                  CustomerJobNumber & "|" & _
                  CustomerPrefix & "|" & _
                  CustomerDisplayName & "|" & _
                  gExactJobFolderName & "|" & _
                  NetworkJobFolder & "|" & _
                  CurrentJobFolder & "|" & _
                  gHandoffAttachDir & "|" & _
                  JobBaseName)

    If Not swModel Is Nothing Then
        blob = blob & "|" & UCase$(swModel.GetTitle)
        blob = blob & "|" & UCase$(swModel.GetPathName)
    End If

    BuildBaseTypeBlob = blob
End Function

Private Function HasHardBmsNameSignal() As Boolean
On Error GoTo ErrHandler

    HasHardBmsNameSignal = False

    Dim blob As String
    blob = BuildBaseTypeBlob()

    If blob = "" Then Exit Function

    If InStr(blob, "BMS") > 0 Then HasHardBmsNameSignal = True: Exit Function
    If InStr(blob, "POTBLOCK") > 0 Then HasHardBmsNameSignal = True: Exit Function
    If InStr(blob, "POT-BLOCK") > 0 Then HasHardBmsNameSignal = True: Exit Function
    If InStr(blob, "POT_BLOCK") > 0 Then HasHardBmsNameSignal = True: Exit Function
    If InStr(blob, "RFQ_MB_ASM") > 0 Then HasHardBmsNameSignal = True: Exit Function
    If InStr(blob, "MB_ASM") > 0 Then HasHardBmsNameSignal = True: Exit Function

    ' SMED is normally BMS/pot-block style.
    If InStr(blob, "SMED") > 0 Then HasHardBmsNameSignal = True: Exit Function

    Exit Function

ErrHandler:
    HasHardBmsNameSignal = False
End Function

Private Function CountDistinctBmsBomRoles() As Long
On Error GoTo ErrHandler

    CountDistinctBmsBomRoles = 0

    If BomCount < 1 Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim q As String
    Dim d As String

    For i = 1 To BomCount

        q = NormalizeKey(BomRows(i).quoteName)
        d = NormalizeText(BomRows(i).Description)

        Select Case q
            Case "IDHOLDER"
                If Not dict.Exists("IDHOLDER") Then dict.Add "IDHOLDER", True

            Case "ODHOLDER"
                If Not dict.Exists("ODHOLDER") Then dict.Add "ODHOLDER", True

            Case "IDPOT", "IDPOTBLOCK"
                If Not dict.Exists("IDPOT") Then dict.Add "IDPOT", True

            Case "ODPOT", "ODPOTBLOCK"
                If Not dict.Exists("ODPOT") Then dict.Add "ODPOT", True
        End Select

        If InStr(d, "ID HOLDER") > 0 Or InStr(d, "TOP HOLDER") > 0 Or InStr(d, "ID MOLD BASE") > 0 Then
            If Not dict.Exists("IDHOLDER") Then dict.Add "IDHOLDER", True
        End If

        If InStr(d, "OD HOLDER") > 0 Or InStr(d, "BOTTOM HOLDER") > 0 Or InStr(d, "BOT HOLDER") > 0 Or InStr(d, "OD MOLD BASE") > 0 Then
            If Not dict.Exists("ODHOLDER") Then dict.Add "ODHOLDER", True
        End If

        If InStr(d, "ID POT") > 0 Or InStr(d, "TOP POT") > 0 Or InStr(d, "TCP POT") > 0 Then
            If Not dict.Exists("IDPOT") Then dict.Add "IDPOT", True
        End If

        If InStr(d, "OD POT") > 0 Or InStr(d, "BOTTOM POT") > 0 Or InStr(d, "BOT POT") > 0 Or InStr(d, "BCP POT") > 0 Then
            If Not dict.Exists("ODPOT") Then dict.Add "ODPOT", True
        End If

        If InStr(d, "SMED") > 0 Then
            If IsLikelyIdSideName(d) Then
                If Not dict.Exists("TCP_SMED") Then dict.Add "TCP_SMED", True
            ElseIf IsLikelyOdSideName(d) Then
                If Not dict.Exists("BCP_SMED") Then dict.Add "BCP_SMED", True
            Else
                If Not dict.Exists("SMED") Then dict.Add "SMED", True
            End If
        End If

        If InStr(d, "POT BLOCK") > 0 Then
            If Not dict.Exists("POTBLOCK") Then dict.Add "POTBLOCK", True
        End If

        If InStr(d, "HOLDER BLOCK") > 0 Then
            If Not dict.Exists("HOLDERBLOCK") Then dict.Add "HOLDERBLOCK", True
        End If

    Next i

    CountDistinctBmsBomRoles = dict.Count
    Exit Function

ErrHandler:
    CountDistinctBmsBomRoles = 0
End Function

Private Function CountDistinctStandardBomRoles() As Long
On Error GoTo ErrHandler

    CountDistinctStandardBomRoles = 0

    If BomCount < 1 Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim nm As String
    Dim k As String

    For i = 1 To BomCount

        If Not IsHardwareName(BomRows(i).Description) Then

            nm = StandardPlateNameStd(BomRows(i).Description)

            If nm <> "" Then
                k = NormalizeKey(nm)

                Select Case k
                    Case "TOPCLAMPPLATE", "BOTTOMCLAMPPLATE", _
                         "APLATE", "BPLATE", _
                         "STRIPPERPLATE", "SUPPORTPLATE", _
                         "MANIFOLDPLATE", "EJECTORPLATE", "BOTTOMEJECTORPLATE", _
                         "SCRETAINERPLATE", "SCBACKUPPLATE", _
                         "DIEPLATE", "DIEBACKUPPLATE", _
                         "XPLATE", "YPLATE"

                        If Not dict.Exists(k) Then dict.Add k, True
                End Select
            End If

        End If

    Next i

    CountDistinctStandardBomRoles = dict.Count
    Exit Function

ErrHandler:
    CountDistinctStandardBomRoles = 0
End Function

Private Sub GetBaseTypeGeometryStats(ByRef nFullThin As Long, _
                                     ByRef nThickInner As Long, _
                                     ByRef nPotLike As Long, _
                                     ByRef nThinSheet As Long)
On Error GoTo ErrHandler

    nFullThin = 0
    nThickInner = 0
    nPotLike = 0
    nThinSheet = 0

    If PartCount < 1 Then Exit Sub

    Dim maxFp As Double
    Dim i As Long
    Dim fp As Double

    maxFp = 0#

    For i = 1 To PartCount
        If Not IsPyropelPartIndex(i) Then
            fp = parts(i).Width * parts(i).Length
            If fp > maxFp Then maxFp = fp
        End If
    Next i

    If maxFp <= 0# Then Exit Sub

    For i = 1 To PartCount

        If IsPyropelPartIndex(i) Then GoTo NextPart

        fp = parts(i).Width * parts(i).Length

        If fp >= 0.8 * maxFp Then
            If parts(i).Thickness >= 0.5 And parts(i).Thickness <= 2.75 Then
                nFullThin = nFullThin + 1
            End If
        End If

        If parts(i).Thickness >= 3# Then
            If fp >= 0.15 * maxFp And fp < 0.85 * maxFp Then
                nThickInner = nThickInner + 1
            End If
        End If

        If IsPotBlockGeometry(parts(i).Thickness, parts(i).Width, parts(i).Length, maxFp) Then
            nPotLike = nPotLike + 1
        End If

        If Abs(parts(i).Thickness - 0.25) <= 0.06 Then
            nThinSheet = nThinSheet + 1
        End If

NextPart:
    Next i

    Exit Sub

ErrHandler:
    LogLine "GetBaseTypeGeometryStats error: " & Err.Description
End Sub

' Count CAD components whose names look like PCS/DME structural plates.
Private Function CountPcsStandardPlateNameHits() As Long
    Dim i As Long, n As Long, s As String, u As String
    n = 0
    For i = 1 To PartCount
        s = " " & NormalizeText(parts(i).componentName) & " "
        u = UCase$(parts(i).componentName)
        If InStr(u, "A-PLATE") > 0 Or InStr(u, "A_PLATE") > 0 Or InStr(s, " A PLATE ") > 0 Or InStr(s, " A PLT ") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "B-PLATE") > 0 Or InStr(u, "B_PLATE") > 0 Or InStr(s, " B PLATE ") > 0 Or InStr(s, " B PLT ") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "EJ-RET") > 0 Or InStr(u, "EJ_RET") > 0 Or InStr(u, "EJ RET") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "EJ-BACKUP") > 0 Or InStr(u, "EJ_BACKUP") > 0 Or InStr(u, "EJ BACKUP") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "EJECTOR") > 0 And (InStr(u, "PLATE") > 0 Or InStr(u, "RET") > 0 Or InStr(u, "BACKUP") > 0) Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "SC-RETAINER") > 0 Or InStr(u, "SC-BACKUP") > 0 Or InStr(u, "SC RETAINER") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "TOP CLAMP") > 0 Or InStr(u, "BOTTOM CLAMP") > 0 Or InStr(u, "TOPCLAMP") > 0 Or InStr(u, "BOTCLAMP") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "CLAMP") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "SUPPORT") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "STRIPPER") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextPcsHit
        If InStr(u, "RAIL") > 0 Then n = n + 1: GoTo NextPcsHit
        If StandardPlateNameStd(parts(i).componentName) <> "" Then n = n + 1
NextPcsHit:
    Next i
    CountPcsStandardPlateNameHits = n
End Function

' Strong PCS tokens only (A/B plates, clamp plates, EJ plates) — not rails alone.
Private Function CountPcsStrongPlateNameHits() As Long
    Dim i As Long, n As Long, s As String, u As String, nm As String
    n = 0
    For i = 1 To PartCount
        s = " " & NormalizeText(parts(i).componentName) & " "
        u = UCase$(parts(i).componentName)
        If InStr(u, "A-PLATE") > 0 Or InStr(u, "A_PLATE") > 0 Or InStr(s, " A PLATE ") > 0 Or InStr(s, " A PLT ") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "B-PLATE") > 0 Or InStr(u, "B_PLATE") > 0 Or InStr(s, " B PLATE ") > 0 Or InStr(s, " B PLT ") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "EJ-RET") > 0 Or InStr(u, "EJ_RET") > 0 Or InStr(u, "EJ RET") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "EJ-BACKUP") > 0 Or InStr(u, "EJ_BACKUP") > 0 Or InStr(u, "EJ BACKUP") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "EJECTOR") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "TOP CLAMP") > 0 Or InStr(u, "BOTTOM CLAMP") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "CLAMP") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "SUPPORT") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextStrong
        If InStr(u, "STRIPPER") > 0 And InStr(u, "PLATE") > 0 Then n = n + 1: GoTo NextStrong
        nm = StandardPlateNameStd(parts(i).componentName)
        If nm <> "" Then
            If InStr(UCase$(nm), "PLATE") > 0 Or InStr(UCase$(nm), "CLAMP") > 0 Then n = n + 1
        End If
NextStrong:
    Next i
    CountPcsStrongPlateNameHits = n
End Function

' Folder / customer / CAD-file BMS markers. Most BMS jobs have "BMS" in the name.
Private Function LooksLikeBmsJobFromName() As Boolean
    LooksLikeBmsJobFromName = False
    Dim blob As String
    blob = UCase$(CurrentJobNumber & "|" & CustomerJobNumber & "|" & CustomerPrefix & "|" & _
                  CustomerDisplayName & "|" & gExactJobFolderName & "|" & NetworkJobFolder & "|" & _
                  CurrentJobFolder & "|" & gHandoffAttachDir & "|" & JobBaseName)

    ' Also include the open CAD file path/title — BMS is often in the .sldasm name.
    On Error Resume Next
    If Not swModel Is Nothing Then
        blob = blob & "|" & UCase$(swModel.GetTitle)
        blob = blob & "|" & UCase$(swModel.GetPathName)
    End If
    On Error GoTo 0

    If InStr(blob, "BMS") > 0 Then LooksLikeBmsJobFromName = True: Exit Function
    If InStr(blob, "POTBLOCK") > 0 Or InStr(blob, "POT-BLOCK") > 0 Or InStr(blob, "POT_BLOCK") > 0 Then
        LooksLikeBmsJobFromName = True: Exit Function
    End If
    ' Tempcraft / Howmet pot-block RFQs (often no "BMS" in the folder name).
    If InStr(blob, "TEMPCRAFT") > 0 Or InStr(blob, "HOWMET") > 0 Then
        LooksLikeBmsJobFromName = True: Exit Function
    End If
    ' HTE as a whole-token customer prefix (avoid matching random substrings).
    If InStr(blob, "|HTE|") > 0 Or InStr(blob, "|HTE-") > 0 Or InStr(blob, "-HTE|") > 0 Or InStr(blob, "\HTE\") > 0 Then
        LooksLikeBmsJobFromName = True: Exit Function
    End If
    If InStr(blob, "RFQ_MB_ASM") > 0 Or InStr(blob, "MB_ASM") > 0 Then
        LooksLikeBmsJobFromName = True: Exit Function
    End If
End Function

' BOM already names pot-block plates. TCP/BCP alone are NOT enough.
Private Function LooksLikeBmsJobFromBom() As Boolean
    LooksLikeBmsJobFromBom = False
    Dim i As Long, d As String, q As String
    For i = 1 To BomCount
        q = NormalizeKey(BomRows(i).quoteName)
        Select Case q
            Case "IDHOLDER", "ODHOLDER", "IDPOT", "IDPOTBLOCK", "ODPOT", "ODPOTBLOCK", "SMEDPLATE"
                LooksLikeBmsJobFromBom = True: Exit Function
        End Select
        d = NormalizeText(BomRows(i).Description)
        If InStr(d, "SMED") > 0 Or InStr(d, "POT BLOCK") > 0 Or InStr(d, "HOLDER BLOCK") > 0 Then
            LooksLikeBmsJobFromBom = True: Exit Function
        End If
        If InStr(d, "ID HOLDER") > 0 Or InStr(d, "OD HOLDER") > 0 Or InStr(d, "TOP HOLDER") > 0 Or InStr(d, "BOTTOM HOLDER") > 0 Then
            LooksLikeBmsJobFromBom = True: Exit Function
        End If
        If InStr(d, "ID POT") > 0 Or InStr(d, "OD POT") > 0 Or InStr(d, "POT BLOCK") > 0 Then
            LooksLikeBmsJobFromBom = True: Exit Function
        End If
    Next i
End Function

' Geometry-only pot-block: ~2 thin full clamps + >=2 thick non-full holders
' + distinguishable pots (thick, chunky, footprint << mold) and/or 0.25" sheets.
' NEVER use gIdxIDH/ODH/IDP/ODP alone — those can be wrong on standard molds.
Private Function LooksLikeBmsJobFromGeometry() As Boolean
    LooksLikeBmsJobFromGeometry = False
    If PartCount < 6 Then Exit Function

    ' Standard mold stacks have many full-footprint plates — never BMS.
    If CountFullFootprintPlates() >= 5 Then Exit Function

    Dim maxFp As Double, i2 As Long, fp As Double
    Dim nFullThin As Long, nThickInner As Long, nThinSheet As Long, nPotLike As Long
    maxFp = 0#
    For i2 = 1 To PartCount
        fp = parts(i2).Width * parts(i2).Length
        If fp > maxFp Then maxFp = fp
    Next i2
    If maxFp <= 0# Then Exit Function

    For i2 = 1 To PartCount
        fp = parts(i2).Width * parts(i2).Length
        If fp >= 0.85 * maxFp And parts(i2).Thickness >= 0.75 And parts(i2).Thickness <= 2.5 Then
            nFullThin = nFullThin + 1
        End If
        If parts(i2).Thickness >= 3# And fp < 0.85 * maxFp And fp >= 0.15 * maxFp Then
            nThickInner = nThickInner + 1
        End If
        If Abs(parts(i2).Thickness - 0.25) <= 0.06 Then nThinSheet = nThinSheet + 1
        If IsPotBlockGeometry(parts(i2).Thickness, parts(i2).Width, parts(i2).Length, maxFp) Then
            nPotLike = nPotLike + 1
        End If
    Next i2

    ' Prefer real pots; insulation sheets alone are not enough without holders.
    If nFullThin <= 2 And nThickInner >= 2 And (nPotLike >= 2 Or (nThinSheet >= 2 And nPotLike >= 1)) Then
        LooksLikeBmsJobFromGeometry = True
    End If
End Function

' Combined helper for callers that only need a yes/no BMS check.
Private Function LooksLikeBmsJob() As Boolean
    LooksLikeBmsJob = LooksLikeBmsJobFromName() Or LooksLikeBmsJobFromBom() Or LooksLikeBmsJobFromGeometry()
End Function

Private Sub AddStdPlate(ByVal nm As String, _
                        ByVal t As Double, _
                        ByVal w As Double, _
                        ByVal l As Double, _
                        ByVal qty As Long, _
                        Optional ByVal gradeHint As String = "", _
                        Optional ByVal cadIdx As Long = 0)

    Dim e As Long
    e = FindStdByName(nm)

    ' Combine only if same name AND same size.
    If e > 0 Then
        If Abs(StdT(e) - t) <= 0.01 And _
           Abs(StdW(e) - w) <= 0.01 And _
           Abs(StdL(e) - l) <= 0.01 Then

            StdQty(e) = StdQty(e) + IIf(qty < 1, 1, qty)

            If StdCadIndex(e) = 0 And cadIdx > 0 Then
                StdCadIndex(e) = cadIdx
            End If

            Exit Sub
        End If
    End If

    StdCount = StdCount + 1

    stdName(StdCount) = nm
    StdT(StdCount) = t
    StdW(StdCount) = w
    StdL(StdCount) = l
    StdQty(StdCount) = IIf(qty < 1, 1, qty)
    StdCadIndex(StdCount) = cadIdx

    Dim g As String
    Dim qr As Long

    StdTargetForName nm, gradeHint, g, qr

    StdGrade(StdCount) = g
    StdQuoteRow(StdCount) = qr
End Sub

Private Sub AddStdPlateFromCad(ByVal idx As Long, ByVal nm As String)
    If idx < 1 Or idx > PartCount Then Exit Sub

    AddStdPlate nm, _
                parts(idx).Thickness, _
                parts(idx).Width, _
                parts(idx).Length, _
                1, _
                "", _
                idx
End Sub

' Grade + Quote row for a standard plate name (gradeHint from BOM material, "" if unknown).
Private Sub StdTargetForName(ByVal nm As String, ByVal gradeHint As String, ByRef grade As String, ByRef quoteRow As Long)
    Dim slot As String
    slot = StdSlotForName(nm)
    grade = ResolveStdGrade(gradeHint, slot)
    quoteRow = StdQuoteRowFor(slot, grade)
    ' A structural plate may have no row in the chosen block (e.g. P20 has no
    ' clamp/rails/ejector rows) - fall back to the A-36 block.
    If quoteRow = 0 And UCase(grade) <> "A36" Then
        grade = "A36"
        quoteRow = StdQuoteRowFor(slot, "A36")
    End If
End Sub

' Name a full-footprint plate by its position in the top->bottom stack.
' Names are hints only: if CAD/BOM names are useful, use them; otherwise use
' standard mold-base stack patterns from the plate count and order.
Private Function StdFullPlateName(ByVal pos As Long, ByVal nFull As Long) As String
    If pos = 1 Then StdFullPlateName = "Top Clamp Plate": Exit Function
    If pos = nFull Then StdFullPlateName = "Bottom Clamp Plate": Exit Function

    Select Case nFull
        Case 3
            If pos = 2 Then StdFullPlateName = "Stripper Plate": Exit Function

        Case 4
            If pos = 2 Then StdFullPlateName = "A Plate": Exit Function
            If pos = 3 Then StdFullPlateName = "B Plate": Exit Function

        Case 5
            If pos = 2 Then StdFullPlateName = "A Plate": Exit Function
            If pos = 3 Then StdFullPlateName = "B Plate": Exit Function
            If pos = 4 Then StdFullPlateName = "Support Plate": Exit Function

        Case 6
            If pos = 2 Then StdFullPlateName = "A Plate": Exit Function
            If pos = 3 Then StdFullPlateName = "Stripper Plate": Exit Function
            If pos = 4 Then StdFullPlateName = "B Plate": Exit Function
            If pos = 5 Then StdFullPlateName = "Support Plate": Exit Function

        Case 7
            If pos = 2 Then StdFullPlateName = "A Plate": Exit Function
            If pos = 3 Then StdFullPlateName = "Stripper Plate": Exit Function
            If pos = 4 Then StdFullPlateName = "B Plate": Exit Function
            If pos = 5 Then StdFullPlateName = "Die Plate": Exit Function
            If pos = 6 Then StdFullPlateName = "Die Backup Plate": Exit Function

        Case Else
            Select Case pos
                Case 2: StdFullPlateName = "A Plate"
                Case 3: StdFullPlateName = "Stripper Plate"
                Case 4: StdFullPlateName = "B Plate"
                Case 5: StdFullPlateName = "Support Plate"
                Case Else: StdFullPlateName = "Plate " & pos
            End Select
            Exit Function
    End Select

    StdFullPlateName = "Plate " & pos
End Function

Private Function StdFullPlateNameFromGeometry(ByVal pos As Long, ByVal nFull As Long, ByVal cadIdx As Long, _
                                              ByRef fullIdx() As Long, ByVal hasRails As Boolean, _
                                              ByVal hasEjectorStack As Boolean) As String
    Dim hinted As String
    Dim guessed As String

    guessed = StdFullPlateNameSmart(pos, nFull, fullIdx, hasRails, hasEjectorStack)
    hinted = ""
    If cadIdx > 0 And cadIdx <= PartCount Then
        hinted = StandardPlateNameStd(parts(cadIdx).componentName)
    End If

    ' Exact shop-standard tokens from imported STEP files are stronger than the
    ' generic stack pattern. Keep using geometry for generic/mixed names.
    Select Case NormalizeKey(hinted)
        Case "APLATE", "BPLATE", "SCRETAINERPLATE", "SCBACKUPPLATE", _
             "BOTTOMCLAMPPLATE", "TOPCLAMPPLATE", "SUPPORTPLATE", "STRIPPERPLATE"
            StdFullPlateNameFromGeometry = hinted
            Exit Function
    End Select

    ' Sequenced / latch-lock SC stack (Qwen + mold_geometry_knowledge):
    ' A / B / SC Retainer / SC Backup / Bottom Clamp when no top clamp.
    If gStdSequencedLatchLock And Not StdTopClampAppearsPresent(fullIdx, nFull) And nFull >= 5 Then
        Select Case pos
            Case 1: StdFullPlateNameFromGeometry = "A Plate": Exit Function
            Case 2: StdFullPlateNameFromGeometry = "B Plate": Exit Function
            Case 3: StdFullPlateNameFromGeometry = "SC Retainer Plate": Exit Function
            Case 4: StdFullPlateNameFromGeometry = "SC Backup Plate": Exit Function
            Case nFull: StdFullPlateNameFromGeometry = "Bottom Clamp Plate": Exit Function
        End Select
    End If

    If Not STD_TRUST_CAD_NAMES_FOR_STANDARD_STACK Then
        StdFullPlateNameFromGeometry = guessed
        Exit Function
    End If

    ' CAD names are useful hints, but they are not allowed to flip the A/B side
    ' against the mold-base stack. Leader pins/rails/ejector stack define the B
    ' side; the plate immediately toward that side is the core/B plate.
    If StdNameHintIsSafeForPosition(hinted, guessed, pos, nFull) Then
        StdFullPlateNameFromGeometry = hinted
    Else
        StdFullPlateNameFromGeometry = guessed
    End If
End Function

Private Function StdFullPlateNameSmart(ByVal pos As Long, ByVal nFull As Long, _
                                       ByRef fullIdx() As Long, ByVal hasRails As Boolean, _
                                       ByVal hasEjectorStack As Boolean) As String
    If nFull < 1 Then Exit Function

    Dim topClampPresent As Boolean
    topClampPresent = StdTopClampAppearsPresent(fullIdx, nFull)

    If topClampPresent Then
        If pos = 1 Then StdFullPlateNameSmart = "Top Clamp Plate": Exit Function
        If pos = nFull Then StdFullPlateNameSmart = "Bottom Clamp Plate": Exit Function
        StdFullPlateNameSmart = StdInnerStackName(pos - 1, nFull - 2, hasRails, hasEjectorStack)
    Else
        ' Some exports/jobs do not include the TCP. In that case the first big
        ' plate is normally A/cavity, not a top clamp.
        If pos = nFull Then StdFullPlateNameSmart = "Bottom Clamp Plate": Exit Function
        StdFullPlateNameSmart = StdInnerStackName(pos, nFull - 1, hasRails, hasEjectorStack)
    End If

    If StdFullPlateNameSmart = "" Then StdFullPlateNameSmart = StdFullPlateName(pos, nFull)
End Function

Private Function StdTopClampAppearsPresent(ByRef fullIdx() As Long, ByVal nFull As Long) As Boolean
    StdTopClampAppearsPresent = True
    If nFull < 3 Then Exit Function

    If Not STD_TRUST_CAD_NAMES_FOR_STANDARD_STACK Then
        Dim avgNoName As Double
        avgNoName = StdAverageInnerFullPlateThickness(fullIdx, nFull)
        If avgNoName > 0# Then
            If parts(fullIdx(1)).Thickness > avgNoName * 1.35 Then
                StdTopClampAppearsPresent = False
                Exit Function
            End If
        End If
        If nFull >= 5 Then
            StdTopClampAppearsPresent = True
            Exit Function
        End If
        StdTopClampAppearsPresent = True
        Exit Function
    End If

    Dim firstHint As String
    firstHint = StandardPlateNameStd(parts(fullIdx(1)).componentName)
    If InStr(UCase(firstHint), "TOP CLAMP") > 0 Then StdTopClampAppearsPresent = True: Exit Function
    If InStr(UCase(firstHint), "A PLATE") > 0 Or InStr(UCase(firstHint), "B PLATE") > 0 Or _
       InStr(UCase(firstHint), "CAVITY") > 0 Or InStr(UCase(firstHint), "CORE") > 0 Or _
       InStr(UCase(firstHint), "SUPPORT") > 0 Then
        StdTopClampAppearsPresent = False
        Exit Function
    End If

    ' Without a useful name, a much-thicker first full plate usually means the
    ' TCP is missing and the first full plate is the A/cavity plate.
    Dim avgInner As Double
    avgInner = StdAverageInnerFullPlateThickness(fullIdx, nFull)
    If avgInner > 0# Then
        If parts(fullIdx(1)).Thickness > avgInner * 1.15 Then
            StdTopClampAppearsPresent = False
            Exit Function
        End If
    End If

    ' Five or more full plates usually include a top clamp in a standard stack.
    If nFull >= 5 Then StdTopClampAppearsPresent = True
End Function

Private Function StdAverageInnerFullPlateThickness(ByRef fullIdx() As Long, ByVal nFull As Long) As Double
    Dim i As Long
    Dim total As Double
    Dim cnt As Long
    If nFull <= 2 Then Exit Function
    For i = 2 To nFull - 1
        total = total + parts(fullIdx(i)).Thickness
        cnt = cnt + 1
    Next i
    If cnt > 0 Then StdAverageInnerFullPlateThickness = total / cnt
End Function

Private Function StdInnerStackName(ByVal innerPos As Long, ByVal innerCount As Long, _
                                   ByVal hasRails As Boolean, ByVal hasEjectorStack As Boolean) As String
    If innerPos < 1 Or innerCount < 1 Then Exit Function

    Select Case innerCount
        Case 1
            If hasRails Or hasEjectorStack Then
                StdInnerStackName = "B Plate"
            Else
                StdInnerStackName = "A Plate"
            End If

        Case 2
            If innerPos = 1 Then StdInnerStackName = "A Plate": Exit Function
            If innerPos = 2 Then StdInnerStackName = "B Plate": Exit Function

        Case 3
            If innerPos = 1 Then StdInnerStackName = "A Plate": Exit Function
            If innerPos = 2 Then StdInnerStackName = "B Plate": Exit Function
            If innerPos = 3 Then StdInnerStackName = "Support Plate": Exit Function

        Case 4
            If innerPos = 1 Then StdInnerStackName = "A Plate": Exit Function
            If innerPos = 2 Then StdInnerStackName = "Stripper Plate": Exit Function
            If innerPos = 3 Then StdInnerStackName = "B Plate": Exit Function
            If innerPos = 4 Then StdInnerStackName = "Support Plate": Exit Function

        Case Else
            Select Case innerPos
                Case 1: StdInnerStackName = "A Plate"
                Case 2: StdInnerStackName = "Stripper Plate"
                Case 3: StdInnerStackName = "B Plate"
                Case 4: StdInnerStackName = "Support Plate"
                Case Else: StdInnerStackName = "Plate " & innerPos
            End Select
    End Select
End Function

Private Function StdNameHintIsSafeForPosition(ByVal hinted As String, ByVal guessed As String, _
                                              ByVal pos As Long, ByVal nFull As Long) As Boolean
    StdNameHintIsSafeForPosition = False
    If hinted = "" Then Exit Function

    Dim h As String
    Dim g As String
    h = NormalizeKey(hinted)
    g = NormalizeKey(guessed)
    If h = g Then StdNameHintIsSafeForPosition = True: Exit Function

    ' Only let strong clamp/special names override the stack when they are at a
    ' sensible end of the full-plate stack.
    If h = "TOPCLAMPPLATE" And pos <= 2 Then StdNameHintIsSafeForPosition = True: Exit Function
    If h = "BOTTOMCLAMPPLATE" And pos >= nFull - 1 Then StdNameHintIsSafeForPosition = True: Exit Function
    If (h = "MANIFOLDPLATE" Or h = "STRIPPERPLATE") And pos > 1 And pos < nFull Then StdNameHintIsSafeForPosition = True: Exit Function

    ' Do not let generic CAD names flip cavity/core. Those two are decided by
    ' position relative to the ejector/rail/leader-pin side.
End Function

Private Function StdDmeStackFamilyName(ByVal nFull As Long, ByVal topClampPresent As Boolean, _
                                       ByVal hasRails As Boolean, ByVal hasEjectorStack As Boolean, _
                                       ByVal nLeaderStack As Long) As String
    Dim innerCount As Long
    If topClampPresent Then
        innerCount = nFull - 2
    Else
        innerCount = nFull - 1
    End If

    Select Case innerCount
        Case 1
            StdDmeStackFamilyName = "DME minimal/core-only stack"
        Case 2
            StdDmeStackFamilyName = "DME 2-plate A/B stack"
        Case 3
            StdDmeStackFamilyName = "DME A/B/support stack"
        Case 4
            StdDmeStackFamilyName = "DME stripper A/B/support stack"
        Case 5
            StdDmeStackFamilyName = "DME manifold/stripper A/B/support stack"
        Case Else
            StdDmeStackFamilyName = "DME extended standard stack"
    End Select

    If Not topClampPresent Then StdDmeStackFamilyName = StdDmeStackFamilyName & " (top clamp missing)"
    If hasRails Then StdDmeStackFamilyName = StdDmeStackFamilyName & " + rails"
    If hasEjectorStack Then StdDmeStackFamilyName = StdDmeStackFamilyName & " + ejector stack"
    If nLeaderStack > 0 Then StdDmeStackFamilyName = StdDmeStackFamilyName & " + leader-pin stack"
End Function

Private Function StdLeaderPinOrientationTopIsFirst(ByRef fullIdx() As Long, ByVal nFull As Long, _
                                                   ByRef lpIdx() As Long, ByVal nLp As Long, _
                                                   ByVal ax As Integer, ByRef topIsFirstOut As Boolean) As Boolean
    ' Fallback ONLY when rails/ejector anchors are missing (Qwen rule).
    ' Prefer PRIMARY leader pins (matched to shoulder/LBB bushings). Never let
    ' SECONDARY (guided-ejector) pins decide stack orientation.
    StdLeaderPinOrientationTopIsFirst = False
    If nFull < 1 Or nLp < 1 Then Exit Function

    Dim i As Long
    Dim roleKey As String
    Dim pinMean As Double, pinCount As Long
    Dim bushMean As Double, bushCount As Long
    Dim setTag As String

    For i = 1 To nLp
        roleKey = NormalizeKey(StandardRoundComponentRole(lpIdx(i)))
        setTag = ""
        On Error Resume Next
        setTag = gStdLeaderPinSetByPart(lpIdx(i))
        On Error GoTo 0
        Select Case roleKey
            Case "LEADERPIN"
                If setTag = "SECONDARY" Then GoTo nextLp
                pinMean = pinMean + PartAxisCenter(lpIdx(i), ax)
                pinCount = pinCount + 1
            Case "LEADERPINBUSHING"
                bushMean = bushMean + PartAxisCenter(lpIdx(i), ax)
                bushCount = bushCount + 1
            ' Guided-ejector bushings intentionally ignored for orientation.
        End Select
nextLp:
    Next i

    Dim firstC As Double
    Dim lastC As Double
    firstC = PartAxisCenter(fullIdx(1), ax)
    lastC = PartAxisCenter(fullIdx(nFull), ax)

    ' DME/common mold-base rule: leader pins are on the B/core side, bushings
    ' are on the A/cavity side. Prefer pins first because the B side also has
    ' the rails/ejector hardware and is the most important orientation anchor.
    If pinCount > 0 Then
        pinMean = pinMean / pinCount
        topIsFirstOut = (Abs(firstC - pinMean) > Abs(lastC - pinMean))
        StdLeaderPinOrientationTopIsFirst = True
        Exit Function
    End If

    If bushCount > 0 Then
        bushMean = bushMean / bushCount
        topIsFirstOut = (Abs(firstC - bushMean) < Abs(lastC - bushMean))
        StdLeaderPinOrientationTopIsFirst = True
    End If
End Function

Private Sub StdSetPartingLineFromRoles(ByVal ax As Integer)
    gStdPartingLineAxis = 0
    gStdPartingLinePos = 0#
    If gStdCavityCadIndex < 1 Or gStdCoreCadIndex < 1 Then Exit Sub

    gStdPartingLineAxis = ax
    gStdPartingLinePos = (PartAxisCenter(gStdCavityCadIndex, ax) + PartAxisCenter(gStdCoreCadIndex, ax)) / 2#
    LogLine "Standard parting line rule: A/cavity idx " & gStdCavityCadIndex & _
            " and B/core idx " & gStdCoreCadIndex & _
            " -> axis " & ax & " pos " & FormatNumberForCsv(gStdPartingLinePos)
End Sub

Private Function StdPartingSideForCadIndex(ByVal idx As Long) As String
    StdPartingSideForCadIndex = ""
    If idx < 1 Or idx > PartCount Then Exit Function
    If gStdPartingLineAxis < 1 Or gStdPartingLineAxis > 3 Then Exit Function

    Dim v As Double
    v = PartAxisCenter(idx, gStdPartingLineAxis)
    If Abs(v - gStdPartingLinePos) <= 0.05 Then
        StdPartingSideForCadIndex = "PARTING_LINE"
    ElseIf gStdTopIsFirst Then
        If v > gStdPartingLinePos Then StdPartingSideForCadIndex = "A/CAVITY_SIDE" Else StdPartingSideForCadIndex = "B/CORE_SIDE"
    Else
        If v < gStdPartingLinePos Then StdPartingSideForCadIndex = "A/CAVITY_SIDE" Else StdPartingSideForCadIndex = "B/CORE_SIDE"
    End If
End Function

Private Function IsStandardRailCandidate(ByVal idx As Long, ByVal baseW As Double, ByVal baseL As Double) As Boolean
    IsStandardRailCandidate = False
    If idx < 1 Or idx > PartCount Then Exit Function

    Dim t As Double
    Dim w As Double
    Dim l As Double
    Dim crossRatio As Double
    Dim slenderRatio As Double

    t = parts(idx).Thickness
    w = parts(idx).Width
    l = parts(idx).Length
    If t <= 0# Or w <= 0# Or l <= 0# Then Exit Function

    If l < STD_RAIL_MIN_LENGTH_FRAC * baseL Then Exit Function
    If w > STD_RAIL_MAX_WIDTH_FRAC * baseW Then Exit Function
    If t < 0.5 Then Exit Function

    ' Rails are flat/rectangular support blocks. Four near-square long parts are
    ' usually rods, pins, or pillars from the CAD bounding box, not rail steel.
    crossRatio = w / t
    If crossRatio < 1# Then crossRatio = 1# / crossRatio
    If crossRatio < 1.35 Then Exit Function

    slenderRatio = l / w
    If slenderRatio < 2.25 Then Exit Function

    IsStandardRailCandidate = True
End Function

Private Function IsStandardEjectorPlateCandidate(ByVal idx As Long, ByVal baseW As Double, ByVal baseL As Double, ByVal baseFoot As Double) As Boolean
    IsStandardEjectorPlateCandidate = False
    If idx < 1 Or idx > PartCount Then Exit Function

    Dim t As Double
    Dim w As Double
    Dim l As Double
    Dim fp As Double

    t = parts(idx).Thickness
    w = parts(idx).Width
    l = parts(idx).Length
    If t <= 0# Or w <= 0# Or l <= 0# Then Exit Function

    ' Ejector retainer/plate steel is long and flat in the mold footprint. Small
    ' cavity/core inserts can have enough area to pass a simple footprint test,
    ' but they are not long base-stack plates.
    If l < 0.6 * baseL Then Exit Function
    If w < 0.35 * baseW Then Exit Function

    fp = w * l
    If fp < STD_EJECTOR_MIN_FOOT_FRAC * baseFoot Then Exit Function

    IsStandardEjectorPlateCandidate = True
End Function
Private Sub BuildStdFromGeometry()
    If PartCount < 1 Then Exit Sub

    ' ================================================================
    ' Qwen classify_geometry parity — full offline stack thinking:
    '   1. Strong shop-name tokens (A-PLATE, LDR-PIN, LBB, RAIL, EJ-*, SC-*)
    '   2. Full-footprint plates -> stack axis -> top-to-bottom order
    '   3. Bottom-up orientation from rails/ejector (pins only if missing)
    '   4. Two-half mold pattern when only 2 full plates
    '   5. Zone-based rails + ejector plates near support/bottom
    '   6. Round hardware by diameter/length + pin-bushing plane match
    '   7. Latch-lock sequenced / SC stack naming
    '   8. Measure pin top/bottom direction WITHOUT flipping A/B
    ' ================================================================

    Dim i As Long, j As Long, fp As Double, baseFoot As Double, baseW As Double, baseL As Double
    Dim maxW As Double, maxL As Double
    baseFoot = 0: baseW = 0: baseL = 0: maxW = 0: maxL = 0
    For i = 1 To PartCount
        fp = parts(i).Width * parts(i).Length
        If parts(i).Width > maxW Then maxW = parts(i).Width
        If parts(i).Length > maxL Then maxL = parts(i).Length
        If fp > baseFoot Then baseFoot = fp: baseW = parts(i).Width: baseL = parts(i).Length
    Next i
    If baseFoot <= 0 Then Exit Sub

    Dim fullIdx(1 To 60) As Long, nFull As Long
    Dim railIdx(1 To 60) As Long, nRail As Long
    Dim ejIdx(1 To 60) As Long, nEj As Long
    Dim lpIdx(1 To 120) As Long, nLp As Long
    Dim shopLocked() As Boolean
    Dim shopPlateName As String
    Dim t As Double, w As Double, l As Double
    Dim rr As String, uName As String
    Dim already As Boolean
    Dim alreadyFull As Boolean
    nFull = 0: nRail = 0: nEj = 0: nLp = 0
    ReDim shopLocked(1 To PartCount)

    ' --- Pass 0: latch-lock / sequenced detection ---
    For i = 1 To PartCount
        If IsLatchLockName(parts(i).componentName) Then
            gStdSequencedLatchLock = True
            Exit For
        End If
    Next i

    ' --- Pass 1: collect candidates; shop tokens lock roles early ---
    For i = 1 To PartCount
        t = parts(i).Thickness: w = parts(i).Width: l = parts(i).Length
        uName = UCase(parts(i).componentName)
        shopPlateName = StandardPlateNameStd(parts(i).componentName)
        fp = w * l

        ' Latch-lock hardware
        If IsLatchLockName(parts(i).componentName) Then
            SetStdCadRole i, "Latch Lock / Safety Strap"
            shopLocked(i) = True
        End If

        ' Shop-token rails
        If (InStr(uName, "RAIL-") > 0 Or InStr(uName, "_RAIL") > 0 Or InStr(uName, "/RAIL") > 0 Or _
            InStr(uName, " RAIL") > 0) And NormalizeKey(shopPlateName) <> "APLATE" Then
            If nRail < UBound(railIdx) Then
                nRail = nRail + 1: railIdx(nRail) = i
                SetStdCadRole i, "Rails"
                shopLocked(i) = True
            End If
            GoTo nextCollect
        End If

        ' Shop-token ejector stack plates
        If NormalizeKey(shopPlateName) = "EJECTORPLATE" Or NormalizeKey(shopPlateName) = "BOTTOMEJECTORPLATE" Then
            If nEj < UBound(ejIdx) Then
                nEj = nEj + 1: ejIdx(nEj) = i
                SetStdCadRole i, shopPlateName
                shopLocked(i) = True
            End If
            GoTo nextCollect
        End If

        ' Shop-token structural plates (A/B/SC/clamp/support) — lock role even
        ' when footprint is under the full-plate threshold (Qwen applies tokens
        ' before geometry filters). Only add to fullIdx when footprint qualifies.
        If shopPlateName <> "" Then
            Select Case NormalizeKey(shopPlateName)
                Case "APLATE", "BPLATE", "TOPCLAMPPLATE", "BOTTOMCLAMPPLATE", _
                     "SUPPORTPLATE", "SCRETAINERPLATE", "SCBACKUPPLATE", "STRIPPERPLATE"
                    SetStdCadRole i, shopPlateName
                    shopLocked(i) = True
                    If (w >= maxW * 0.85 And l >= maxL * 0.85 And t >= 0.5) Or _
                       (fp >= (1 - STD_FOOTPRINT_TOL) * baseFoot And t >= STD_MIN_PLATE_THICKNESS) Then
                        alreadyFull = False
                        For j = 1 To nFull
                            If fullIdx(j) = i Then alreadyFull = True: Exit For
                        Next j
                        If Not alreadyFull And nFull < UBound(fullIdx) Then
                            nFull = nFull + 1
                            fullIdx(nFull) = i
                        End If
                    End If
                    GoTo nextCollect
            End Select
        End If

        If t >= STD_MIN_PLATE_THICKNESS Then
            ' Full-footprint: Qwen uses >= 85% of max W AND max L; macro uses footprint tol.
            If (w >= maxW * 0.85 And l >= maxL * 0.85 And t >= 0.5) Or _
               (fp >= (1 - STD_FOOTPRINT_TOL) * baseFoot And t >= STD_MIN_PLATE_THICKNESS) Then
                If nFull < UBound(fullIdx) Then
                    nFull = nFull + 1
                    fullIdx(nFull) = i
                End If
            ElseIf Not shopLocked(i) And IsStandardRailCandidate(i, baseW, baseL) Then
                If nRail < UBound(railIdx) Then
                    nRail = nRail + 1
                    railIdx(nRail) = i
                End If
            ElseIf Not shopLocked(i) And IsStandardEjectorPlateCandidate(i, baseW, baseL, baseFoot) Then
                If nEj < UBound(ejIdx) Then
                    nEj = nEj + 1
                    ejIdx(nEj) = i
                End If
            End If
        End If

        ' Round guide hardware (leader / bushing / return / pillar)
        If IsRoundBarLike(i) Then
            rr = NormalizeKey(StandardRoundComponentRole(i))
            If rr = "LEADERPIN" Or rr = "LEADERPINBUSHING" Or rr = "GUIDEDEJECTORBUSHING" Or _
               rr = "RETURNPIN" Or rr = "EJECTORRETURNPIN" Or rr = "SUPPORTPILLAR" Then
                If rr = "LEADERPIN" Or rr = "LEADERPINBUSHING" Or rr = "GUIDEDEJECTORBUSHING" Then
                    If nLp < UBound(lpIdx) Then
                        nLp = nLp + 1
                        lpIdx(nLp) = i
                    End If
                End If
                If InStr(uName, "LDR-PIN") > 0 Or InStr(uName, "LDR_PIN") > 0 Or _
                   InStr(uName, "LBB_") > 0 Or InStr(uName, "/LBB_") > 0 Then
                    shopLocked(i) = True
                End If
            End If
        End If
nextCollect:
    Next i
    If nFull < 1 Then Exit Sub

    ' --- Stack axis = greatest center spread among full plates (Qwen) ---
    Dim ax As Integer, bestRange As Double, a As Integer, mn As Double, mx As Double, v As Double
    bestRange = -1: ax = 3
    For a = 1 To 3
        mn = 1E+30: mx = -1E+30
        For i = 1 To nFull
            v = PartAxisCenter(fullIdx(i), a)
            If v < mn Then mn = v
            If v > mx Then mx = v
        Next i
        If (mx - mn) > bestRange Then bestRange = (mx - mn): ax = a
    Next a

    StdSortByAxisDesc fullIdx, nFull, ax

    ' Pre-classify leader-pin PRIMARY/SECONDARY before orientation.
    Dim supportPosGuess As Double
    supportPosGuess = 0#
    If nFull >= 4 Then supportPosGuess = PartAxisCenter(fullIdx(nFull - 1), ax)
    If nLp > 0 Then ClassifyLeaderPinSetsByBushingPlane lpIdx, nLp, ax, supportPosGuess

    ' --- Orientation: rails/ejector first; pins only if missing (Qwen) ---
    Dim topIsFirst As Boolean
    topIsFirst = True
    Dim leaderOriented As Boolean
    leaderOriented = False
    If nEj > 0 Or nRail > 0 Then
        Dim anchorMean As Double
        Dim anchorCount As Long
        anchorMean = 0: anchorCount = 0
        If nEj > 0 Then
            For i = 1 To nEj
                anchorMean = anchorMean + PartAxisCenter(ejIdx(i), ax)
                anchorCount = anchorCount + 1
            Next i
        Else
            For i = 1 To nRail
                anchorMean = anchorMean + PartAxisCenter(railIdx(i), ax)
                anchorCount = anchorCount + 1
            Next i
        End If
        If anchorCount > 0 Then
            anchorMean = anchorMean / anchorCount
            If Abs(PartAxisCenter(fullIdx(1), ax) - anchorMean) < Abs(PartAxisCenter(fullIdx(nFull), ax) - anchorMean) Then topIsFirst = False
            LogLine "Standard orientation rule: rails/ejector stack set topIsFirst=" & CStr(topIsFirst)
        End If
    ElseIf StdLeaderPinOrientationTopIsFirst(fullIdx, nFull, lpIdx, nLp, ax, topIsFirst) Then
        leaderOriented = True
        LogLine "Standard orientation rule: PRIMARY leader pins/bushings set topIsFirst=" & CStr(topIsFirst)
    Else
        Dim firstNm As String, lastNm As String
        firstNm = StandardPlateNameStd(parts(fullIdx(1)).componentName)
        lastNm = StandardPlateNameStd(parts(fullIdx(nFull)).componentName)
        If InStr(UCase(firstNm), "BOTTOM CLAMP") > 0 Or InStr(UCase(lastNm), "TOP CLAMP") > 0 Then
            topIsFirst = False
        ElseIf InStr(UCase(firstNm), "TOP CLAMP") > 0 Or InStr(UCase(lastNm), "BOTTOM CLAMP") > 0 Then
            topIsFirst = True
        ElseIf parts(fullIdx(1)).Thickness > parts(fullIdx(nFull)).Thickness + 0.25 Then
            topIsFirst = False
        ElseIf parts(fullIdx(nFull)).Thickness > parts(fullIdx(1)).Thickness + 0.25 Then
            topIsFirst = True
        End If
    End If
    If Not topIsFirst Then StdReverse fullIdx, nFull
    gStdStackAxis = ax
    gStdTopIsFirst = topIsFirst

    Dim roleName As String
    Dim topPos As Double, bottomPos As Double, supportPos As Double
    Dim retainerIdx As Long
    Dim minEjT As Double
    Dim ejRole As String
    Dim nLatch As Long

    ' --- TWO-HALF mold pattern (Qwen: len(full_plates) == 2) ---
    If nFull = 2 Then
        BuildStdTwoHalfMoldPattern fullIdx, nFull, ax, maxW, maxL, baseFoot, railIdx, nRail, ejIdx, nEj
        gStdDmeStackFamily = "Two-half mold pattern" & IIf(nRail > 0, " + rails", "") & IIf(nEj > 0, " + ejector", "")
        LogLine "Standard DME stack family: " & gStdDmeStackFamily
        GoTo afterPlates
    End If

    gStdDmeStackFamily = StdDmeStackFamilyName(nFull, StdTopClampAppearsPresent(fullIdx, nFull), (nRail > 0), (nEj > 0), nLp)
    If gStdSequencedLatchLock Then gStdDmeStackFamily = gStdDmeStackFamily & " + latch-lock sequenced"
    LogLine "Standard DME stack family: " & gStdDmeStackFamily

    ' --- Name full plates top -> bottom (shop tokens win; else stack pattern / SC) ---
    For i = 1 To nFull
        If shopLocked(fullIdx(i)) And StdCadRole(fullIdx(i)) <> "" Then
            roleName = StdCadRole(fullIdx(i))
        Else
            roleName = StdFullPlateNameFromGeometry(i, nFull, fullIdx(i), fullIdx, (nRail > 0), (nEj > 0))
            SetStdCadRole fullIdx(i), roleName
        End If
        Select Case NormalizeKey(roleName)
            Case "APLATE", "CAVITYPLATE"
                gStdCavityCadIndex = fullIdx(i)
            Case "BPLATE", "COREPLATE"
                gStdCoreCadIndex = fullIdx(i)
        End Select
        AddStdPlateFromCad fullIdx(i), roleName
        LogLine "Standard full plate rule: pos " & i & "/" & nFull & _
                " idx " & fullIdx(i) & " -> " & roleName & _
                " | T=" & parts(fullIdx(i)).Thickness & " W=" & parts(fullIdx(i)).Width & " L=" & parts(fullIdx(i)).Length & _
                " | name=" & parts(fullIdx(i)).componentName & _
                IIf(shopLocked(fullIdx(i)), " [SHOP TOKEN]", "")
    Next i

afterPlates:
    StdSetPartingLineFromRoles ax

    ' Quote any shop-token structural plates that were locked but not already
    ' added via the full-footprint naming loop (e.g. SC plates under footprint).
    For i = 1 To PartCount
        If shopLocked(i) Then
            roleName = StdCadRole(i)
            Select Case NormalizeKey(roleName)
                Case "APLATE", "BPLATE", "TOPCLAMPPLATE", "BOTTOMCLAMPPLATE", _
                     "SUPPORTPLATE", "SCRETAINERPLATE", "SCBACKUPPLATE", "STRIPPERPLATE"
                    If FindStdByName(roleName) = 0 Then
                        AddStdPlateFromCad i, roleName
                        Select Case NormalizeKey(roleName)
                            Case "APLATE": gStdCavityCadIndex = i
                            Case "BPLATE": gStdCoreCadIndex = i
                        End Select
                        LogLine "Shop-token plate (non-full or unlocked stack): idx " & i & " -> " & roleName
                    End If
            End Select
        End If
    Next i

    ' --- Zone-based rail / ejector refinement (Qwen side_offset / centered) ---
    topPos = PartAxisCenter(fullIdx(1), ax)
    bottomPos = PartAxisCenter(fullIdx(nFull), ax)
    supportPos = 0#
    For i = 1 To nFull
        rr = NormalizeKey(StdCadRole(fullIdx(i)))
        If rr = "SUPPORTPLATE" Or rr = "SCBACKUPPLATE" Then
            supportPos = PartAxisCenter(fullIdx(i), ax)
            Exit For
        End If
    Next i
    If supportPos = 0# And nFull >= 2 Then supportPos = PartAxisCenter(fullIdx(nFull - 1), ax)

    RefineRailsAndEjectorsByZone ax, maxW, maxL, topPos, bottomPos, supportPos, _
                                 railIdx, nRail, ejIdx, nEj, shopLocked

    ' Rails
    If nRail > 0 Then
        For i = 1 To nRail
            If Not shopLocked(railIdx(i)) Or StdCadRole(railIdx(i)) = "" Then SetStdCadRole railIdx(i), "Rails"
        Next i
        AddStdPlate "Rails", parts(railIdx(1)).Thickness, parts(railIdx(1)).Width, parts(railIdx(1)).Length, nRail
        LogLine "Standard rail rule: qty " & nRail & " using idx " & railIdx(1) & _
                " | T=" & parts(railIdx(1)).Thickness & " W=" & parts(railIdx(1)).Width & " L=" & parts(railIdx(1)).Length
    End If

    ' Ejector stack: thinner = Ejector Plate; thicker/lower = Bottom Ejector Plate
    If nEj > 0 Then
        StdSortByAxisDesc ejIdx, nEj, ax
        If Not topIsFirst Then StdReverse ejIdx, nEj
        retainerIdx = ejIdx(1)
        minEjT = parts(ejIdx(1)).Thickness
        For j = 2 To nEj
            If parts(ejIdx(j)).Thickness < minEjT Then
                minEjT = parts(ejIdx(j)).Thickness
                retainerIdx = ejIdx(j)
            End If
        Next j
        For j = 1 To nEj
            ' Preserve shop-token EJ-RET / EJ-BACKUP names when present.
            If shopLocked(ejIdx(j)) And StdCadRole(ejIdx(j)) <> "" Then
                ejRole = StdCadRole(ejIdx(j))
            ElseIf ejIdx(j) = retainerIdx Then
                ejRole = "Ejector Plate"
            Else
                ejRole = "Bottom Ejector Plate"
            End If
            SetStdCadRole ejIdx(j), ejRole
            AddStdPlateFromCad ejIdx(j), ejRole
            LogLine "Standard ejector-stack rule: idx " & ejIdx(j) & " -> " & ejRole & _
                    " | T=" & parts(ejIdx(j)).Thickness & " | name=" & parts(ejIdx(j)).componentName
        Next j
    End If

    ' Round hardware roles (all guide / return / pillar)
    For i = 1 To PartCount
        If IsRoundBarLike(i) Then
            rr = StandardRoundComponentRole(i)
            If rr <> "" Then
                If StdCadRole(i) = "" Or Not shopLocked(i) Then SetStdCadRole i, rr
                If NormalizeKey(rr) = "LEADERPIN" Or NormalizeKey(rr) = "LEADERPINBUSHING" Or _
                   NormalizeKey(rr) = "GUIDEDEJECTORBUSHING" Then
                    already = False
                    For j = 1 To nLp
                        If lpIdx(j) = i Then already = True: Exit For
                    Next j
                    If Not already And nLp < UBound(lpIdx) Then
                        nLp = nLp + 1: lpIdx(nLp) = i
                    End If
                End If
            End If
        End If
    Next i
    If nLp > 0 Then LogLine "Standard leader-pin stack rule: round leader/bushing components=" & nLp

    ' Qwen: short bushings above support = leader_pin_bushing; below = guided_ejector_bushing
    ClassifyBushingsBySupportZone ax, supportPos

    ' Latch-lock tags
    nLatch = 0
    For i = 1 To PartCount
        If IsLatchLockName(parts(i).componentName) Then
            SetStdCadRole i, "Latch Lock / Safety Strap"
            nLatch = nLatch + 1
        End If
    Next i
    If nLatch > 0 Then
        gStdSequencedLatchLock = True
        LogLine "Standard latch-lock rule: " & nLatch & " PLC/latch-lock/safety-strap parts (sequenced base; must not flip A/B)"
    End If

    ' Pin-bushing plane match + top/bottom direction (never flips A/B)
    ClassifyLeaderPinSetsByBushingPlane lpIdx, nLp, ax, supportPos
    MeasureLeaderPinTopBottomDirection ax
    BuildStdStackAnalysisText ax, nFull, nRail, nEj, nLp

    LogLine "Standard base (geometry): full=" & nFull & " rails=" & nRail & " ejector=" & nEj & " leaderStack=" & nLp & " (stack axis " & StdStackAxisName(ax) & ")"
    LogLine "Standard parting_line: " & gStdPartingLineText
End Sub

' Qwen two-half mold pattern: 2 full-footprint clamps + large inner A/B blocks
' + thin rails + ejector-stack plates from remaining thin-large parts.
Private Sub BuildStdTwoHalfMoldPattern(ByRef fullIdx() As Long, ByVal nFull As Long, _
                                       ByVal ax As Integer, ByVal maxW As Double, ByVal maxL As Double, _
                                       ByVal baseFoot As Double, _
                                       ByRef railIdx() As Long, ByRef nRail As Long, _
                                       ByRef ejIdx() As Long, ByRef nEj As Long)
    Dim i As Long, j As Long
    Dim hiFull As Long, loFull As Long
    Dim roleName As String

    If PartAxisCenter(fullIdx(1), ax) >= PartAxisCenter(fullIdx(2), ax) Then
        hiFull = fullIdx(1): loFull = fullIdx(2)
    Else
        hiFull = fullIdx(2): loFull = fullIdx(1)
    End If

    ' Lower full plate = BCP; opposite = top clamp (Qwen).
    SetStdCadRole loFull, "Bottom Clamp Plate"
    AddStdPlateFromCad loFull, "Bottom Clamp Plate"
    SetStdCadRole hiFull, "Top Clamp Plate"
    AddStdPlateFromCad hiFull, "Top Clamp Plate"
    LogLine "Two-half: full clamps hi=" & hiFull & " (Top Clamp) lo=" & loFull & " (Bottom Clamp)"

    ' Largest non-full thick blocks -> A (high) / B (low) along stack axis.
    Dim blockIdx(1 To 40) As Long, nBlock As Long
    Dim t As Double, w As Double, l As Double, fp As Double
    nBlock = 0
    For i = 1 To PartCount
        If i = hiFull Or i = loFull Then GoTo nextBlk
        If StdCadRole(i) <> "" Then GoTo nextBlk
        t = parts(i).Thickness: w = parts(i).Width: l = parts(i).Length
        If t < 3# Then GoTo nextBlk
        If w < maxW * 0.3 Or l < maxL * 0.3 Then GoTo nextBlk
        If nBlock < UBound(blockIdx) Then
            nBlock = nBlock + 1
            blockIdx(nBlock) = i
        End If
nextBlk:
    Next i

    ' Sort blocks by volume desc, take top 2, then assign by axis.
    Dim tmp As Long
    For i = 1 To nBlock - 1
        For j = i + 1 To nBlock
            If parts(blockIdx(j)).BBoxVolume > parts(blockIdx(i)).BBoxVolume Then
                tmp = blockIdx(i): blockIdx(i) = blockIdx(j): blockIdx(j) = tmp
            End If
        Next j
    Next i

    If nBlock >= 2 Then
        Dim aIdx As Long, bIdx As Long
        If PartAxisCenter(blockIdx(1), ax) >= PartAxisCenter(blockIdx(2), ax) Then
            aIdx = blockIdx(1): bIdx = blockIdx(2)
        Else
            aIdx = blockIdx(2): bIdx = blockIdx(1)
        End If
        SetStdCadRole aIdx, "A Plate"
        AddStdPlateFromCad aIdx, "A Plate"
        gStdCavityCadIndex = aIdx
        SetStdCadRole bIdx, "B Plate"
        AddStdPlateFromCad bIdx, "B Plate"
        gStdCoreCadIndex = bIdx
        LogLine "Two-half: inner A idx=" & aIdx & " B idx=" & bIdx
    End If

    ' Thin large remaining -> rails (first 2) then ejector plates (next 2).
    Dim thinIdx(1 To 40) As Long, nThin As Long
    nThin = 0
    For i = 1 To PartCount
        If StdCadRole(i) <> "" Then GoTo nextThin
        t = parts(i).Thickness: w = parts(i).Width: l = parts(i).Length
        If t > 0.75 Then GoTo nextThin
        If l < maxL * 0.35 Or w < maxW * 0.2 Then GoTo nextThin
        If nThin < UBound(thinIdx) Then
            nThin = nThin + 1
            thinIdx(nThin) = i
        End If
nextThin:
    Next i
    For i = 1 To nThin - 1
        For j = i + 1 To nThin
            If parts(thinIdx(j)).BBoxVolume > parts(thinIdx(i)).BBoxVolume Then
                tmp = thinIdx(i): thinIdx(i) = thinIdx(j): thinIdx(j) = tmp
            End If
        Next j
    Next i

    nRail = 0: nEj = 0
    If nThin >= 2 Then
        For i = 1 To 2
            If nRail < UBound(railIdx) Then
                nRail = nRail + 1
                railIdx(nRail) = thinIdx(i)
                SetStdCadRole thinIdx(i), "Rails"
            End If
        Next i
        LogLine "Two-half: rails from thin-large qty=" & nRail
    End If
    If nThin >= 4 Then
        Dim ejA As Long, ejB As Long
        ejA = thinIdx(3): ejB = thinIdx(4)
        If parts(ejA).Thickness <= parts(ejB).Thickness Then
            SetStdCadRole ejA, "Ejector Plate"
            SetStdCadRole ejB, "Bottom Ejector Plate"
        Else
            SetStdCadRole ejB, "Ejector Plate"
            SetStdCadRole ejA, "Bottom Ejector Plate"
        End If
        nEj = 2
        ejIdx(1) = ejA: ejIdx(2) = ejB
        LogLine "Two-half: ejector stack from remaining thin-large"
    End If
End Sub

' Qwen zone rules: rails = long narrow side-offset in ejector/rail zone;
' ejector plates = long centered medium-width between bottom and support.
Private Sub RefineRailsAndEjectorsByZone(ByVal ax As Integer, ByVal maxW As Double, ByVal maxL As Double, _
                                         ByVal topPos As Double, ByVal bottomPos As Double, ByVal supportPos As Double, _
                                         ByRef railIdx() As Long, ByRef nRail As Long, _
                                         ByRef ejIdx() As Long, ByRef nEj As Long, _
                                         ByRef shopLocked() As Boolean)
    Dim i As Long
    Dim axisPos As Double
    Dim sideOffset As Double
    Dim longFull As Boolean
    Dim narrowWidth As Boolean
    Dim ejectorWidth As Boolean
    Dim sideBlock As Boolean
    Dim centeredSide As Boolean
    Dim already As Boolean
    Dim j As Long
    Dim t As Double, w As Double, l As Double
    Dim a1 As Double, a2 As Double

    For i = 1 To PartCount
        If shopLocked(i) Then GoTo nextZone
        If StdCadRole(i) <> "" Then
            ' Skip already-named full plates / latch locks
            Select Case NormalizeKey(StdCadRole(i))
                Case "APLATE", "BPLATE", "TOPCLAMPPLATE", "BOTTOMCLAMPPLATE", "SUPPORTPLATE", _
                     "SCRETAINERPLATE", "SCBACKUPPLATE", "STRIPPERPLATE", "LATCHLOCK/SAFETYSTRAP", "LATCHLOCKSAFETYSTRAP"
                    GoTo nextZone
            End Select
        End If
        t = parts(i).Thickness: w = parts(i).Width: l = parts(i).Length
        If t < 0.4 Or l <= 0 Or w <= 0 Then GoTo nextZone

        longFull = (l >= maxL * 0.85)
        If Not longFull Then GoTo nextZone

        ' Lateral offset from mold center in the two non-stack axes.
        Select Case ax
            Case 1: a1 = parts(i).AsmCenterY: a2 = parts(i).AsmCenterZ
            Case 2: a1 = parts(i).AsmCenterX: a2 = parts(i).AsmCenterZ
            Case Else: a1 = parts(i).AsmCenterX: a2 = parts(i).AsmCenterY
        End Select
        sideOffset = Abs(a1)
        If Abs(a2) > sideOffset Then sideOffset = Abs(a2)

        narrowWidth = (w >= maxW * 0.18 And w <= maxW * 0.72)
        ejectorWidth = (w >= maxW * 0.58 And w <= maxW * 0.86)
        sideBlock = (sideOffset >= maxW * 0.25)
        centeredSide = (sideOffset <= maxW * 0.15)
        axisPos = PartAxisCenter(i, ax)

        already = False
        For j = 1 To nRail
            If railIdx(j) = i Then already = True: Exit For
        Next j
        For j = 1 To nEj
            If ejIdx(j) = i Then already = True: Exit For
        Next j

        If longFull And narrowWidth And sideBlock And _
           axisPos >= bottomPos - 0.5 And axisPos <= supportPos + 1# Then
            If Not already And nRail < UBound(railIdx) Then
                nRail = nRail + 1
                railIdx(nRail) = i
                LogLine "Zone rail: idx " & i & " sideOffset=" & FormatNumberForCsv(sideOffset)
            End If
        ElseIf longFull And ejectorWidth And centeredSide And _
               axisPos >= bottomPos - 0.5 And axisPos <= supportPos + 2# Then
            If Not already And nEj < UBound(ejIdx) Then
                nEj = nEj + 1
                ejIdx(nEj) = i
                LogLine "Zone ejector: idx " & i & " T=" & FormatNumberForCsv(t)
            End If
        End If
nextZone:
    Next i
End Sub

' Qwen: short round cylinders at/above support = leader_pin_bushing;
' below support = guided_ejector_bushing. Does not override LDR-PIN/LBB tokens.
Private Sub ClassifyBushingsBySupportZone(ByVal ax As Integer, ByVal supportPos As Double)
    Dim i As Long
    Dim dia As Double, axisLen As Double, ratio As Double
    Dim roleKey As String
    Dim axisPos As Double
    If supportPos = 0# Then Exit Sub

    For i = 1 To PartCount
        If Not IsRoundBarLike(i) Then GoTo nextBush
        dia = RoundBarDiameter(i)
        axisLen = RoundBarAxisLength(i)
        If dia <= 0# Or dia > 4# Then GoTo nextBush
        ratio = axisLen / dia
        ' Short cylinder only
        If ratio > 2# Or ratio < 0.6 Then GoTo nextBush
        If dia < 1# Or dia > 2.6 Then GoTo nextBush

        roleKey = NormalizeKey(StdCadRole(i))
        If roleKey = "" Then roleKey = NormalizeKey(StandardRoundComponentRole(i))
        ' Don't reclassify long pins / pillars / shop-token LDR-PIN
        If roleKey = "LEADERPIN" Or roleKey = "SUPPORTPILLAR" Or roleKey = "RETURNPIN" Or _
           roleKey = "EJECTORRETURNPIN" Or roleKey = "EJECTORPIN" Then GoTo nextBush
        If InStr(UCase(parts(i).componentName), "LDR-PIN") > 0 Or InStr(UCase(parts(i).componentName), "LDR_PIN") > 0 Then GoTo nextBush
        If InStr(UCase(parts(i).componentName), "LBB_") > 0 Or InStr(UCase(parts(i).componentName), "/LBB_") > 0 Then
            SetStdCadRole i, "Leader Pin Bushing"
            GoTo nextBush
        End If

        axisPos = PartAxisCenter(i, ax)
        If axisPos >= supportPos Then
            SetStdCadRole i, "Leader Pin Bushing"
        Else
            SetStdCadRole i, "Guided Ejector Bushing"
        End If
nextBush:
    Next i
End Sub

Private Function StdSteelTypeFor(ByVal grade As String) As String
    Select Case UCase(grade)
        Case "P20": StdSteelTypeFor = "#3 P20"
        Case "4140": StdSteelTypeFor = "#2 4140"
        Case "420SS": StdSteelTypeFor = "#7 420-SS"
        Case "6061": StdSteelTypeFor = "ALM 6061"
        Case "H13": StdSteelTypeFor = "#5 H13"
        Case Else: StdSteelTypeFor = "#1 A-36"
    End Select
End Function

Private Function NextStdSpareQuoteRow(ByRef usedRows() As Boolean, ByVal grade As String) As Long
On Error GoTo Done
    Dim rows As Variant, i As Long, r As Long
    Select Case UCase(grade)
        Case "4140"
            rows = Array(24, 25, 30, 31, 32, 33, 34, 26, 27, 28, 29)
        Case "P20"
            rows = Array(41, 42, 43, 38, 39, 40)
        Case "420SS"
            rows = Array(52, 53, 54, 55, 56, 57, 58, 59, 60, 61)
        Case "6061"
            rows = Array(68, 69, 70, 71, 72, 73, 74, 75, 76, 77)
        Case Else
            rows = Array(13, 6, 7, 8, 9, 10, 11, 12, 14, 15)
    End Select
    For i = LBound(rows) To UBound(rows)
        r = CLng(rows(i))
        If r >= LBound(usedRows) And r <= UBound(usedRows) Then
            If Not usedRows(r) Then NextStdSpareQuoteRow = r: Exit Function
        End If
    Next i
Done:
End Function
Private Sub FillStandardBaseQuote()
On Error GoTo ErrHandler
    If StdCount < 1 Then LogLine "Standard base: no plates identified; skipping Quote.": Exit Sub
    Dim templatePath As String
    templatePath = FindQuoteTemplateAnywhere()
    If templatePath = "" Then LogLine "Quote template not found in Downloads; skipping.": Exit Sub
    Dim quotePath As String
    quotePath = CopyTemplateToJobFolder(templatePath)
    If quotePath = "" Then quotePath = templatePath
    LogLine "Quote workbook (job copy): " & quotePath

    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False: xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    Set xlWb = xlApp.Workbooks.Open(quotePath)
    xlApp.Calculation = -4135  ' manual calc - valid only after a workbook is open
    On Error Resume Next
    Set xlWs = xlWb.Worksheets(QUOTE_SHEET_NAME)
    On Error GoTo ErrHandler
    If xlWs Is Nothing Then
        xlWb.Close False: xlApp.Quit
        LogLine "QuoteWorksheet missing."
        Exit Sub
    End If

    ' Clear every plate block we might touch so nothing carries a stale quantity
    ' (e.g. the #2 pot block ships with qty 1 on TCP/BCP/holders/pots).
    Dim cr As Long
    For cr = 6 To 15            ' #1 A-36
        xlWs.Cells(cr, 3).value = "": xlWs.Cells(cr, 4).value = ""
        xlWs.Cells(cr, 5).value = "": xlWs.Cells(cr, 6).value = ""
    Next cr
    For cr = 22 To 34           ' #2 4140 (pot block)
        xlWs.Cells(cr, 3).value = "": xlWs.Cells(cr, 4).value = ""
        xlWs.Cells(cr, 5).value = "": xlWs.Cells(cr, 6).value = ""
    Next cr
    For cr = 38 To 43           ' #3 P20
        xlWs.Cells(cr, 3).value = "": xlWs.Cells(cr, 4).value = ""
        xlWs.Cells(cr, 5).value = "": xlWs.Cells(cr, 6).value = ""
    Next cr
    For cr = 52 To 61           ' #7 420-SS
        xlWs.Cells(cr, 3).value = "": xlWs.Cells(cr, 4).value = ""
        xlWs.Cells(cr, 5).value = "": xlWs.Cells(cr, 6).value = ""
    Next cr
    For cr = 68 To 77           ' ALM 6061
        xlWs.Cells(cr, 3).value = "": xlWs.Cells(cr, 4).value = ""
        xlWs.Cells(cr, 5).value = "": xlWs.Cells(cr, 6).value = ""
    Next cr

    Dim usedQuoteRows(1 To 200) As Boolean
    Dim i As Long, qt As Double, qw As Double, ql As Double, targetRow As Long, baseRow As Long
    For i = 1 To StdCount
        baseRow = StdQuoteRow(i)
        targetRow = baseRow
        If targetRow > 0 And usedQuoteRows(targetRow) Then
            targetRow = NextStdSpareQuoteRow(usedQuoteRows, StdGrade(i))
            If targetRow > 0 Then LogLine "Quote: moved duplicate " & StdGrade(i) & " item from row " & baseRow & " to open row " & targetRow
        ElseIf targetRow = 0 Then
            targetRow = NextStdSpareQuoteRow(usedQuoteRows, StdGrade(i))
            If targetRow > 0 Then LogLine "Quote: placed extra " & StdGrade(i) & " item on open row " & targetRow
        End If

        If targetRow > 0 Then
            usedQuoteRows(targetRow) = True
            qt = StdT(i): qw = RoundUpToNickel(StdW(i)): ql = RoundUpToNickel(StdL(i))
            If QUOTE_ROUND_UP_TO_QUARTER Then qt = SteelStockThickness(qt)
            xlWs.Cells(targetRow, 1).value = stdName(i)
            xlWs.Cells(targetRow, 3).value = StdQty(i)
            xlWs.Cells(targetRow, 4).value = qt
            xlWs.Cells(targetRow, 5).value = qw
            xlWs.Cells(targetRow, 6).value = ql
            LogLine "Quote " & StdGrade(i) & " row " & targetRow & " <- " & stdName(i) & " T=" & qt & " W=" & qw & " L=" & ql
        Else
            LogLine "Quote: no open target row for " & stdName(i) & " (on steel sheet only)"
        End If
    Next i

    If PcCount > 0 Then WritePullcoreCategoryToSheet xlWs
    WritePullcoreTotalToSummary xlWs
    If PpCount > 0 Then
        WritePurchasedToComponentsArea xlWs
        WritePurchasedCategoryToSheet xlWs
    End If
    StampWorkbookDateAndRef xlWb, FormatRefNumber
    xlWb.Save
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    LogLine "Quote workbook saved (standard base)."
    Exit Sub
ErrHandler:
    LogLine "FillStandardBaseQuote error: " & Err.Description
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

Private Sub FillStandardBaseSteel()
On Error GoTo ErrHandler
    If StdCount < 1 Then LogLine "Standard base: no plates identified; skipping J000.": Exit Sub
    Dim templatePath As String
    templatePath = FindJ000TemplateAnywhere()
    If templatePath = "" Then LogLine "J000 template not found in Downloads; skipping.": Exit Sub
    Dim jPath As String
    jPath = CopyTemplateToJobFolder(templatePath)
    If jPath = "" Then jPath = templatePath
    LogLine "J000 workbook (job copy): " & jPath

    Dim xlApp As Object, xlWb As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False: xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    Set xlWb = xlApp.Workbooks.Open(jPath)
    xlApp.Calculation = -4135  ' manual calc - valid only after a workbook is open

    Dim sheetNames As Variant
    sheetNames = Array("Steel Order", "Machining Sheet")
    Dim sName As Variant, ws As Object, writeRow As Long, i As Long
    For Each sName In sheetNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = xlWb.Worksheets(CStr(sName))
        On Error GoTo ErrHandler
        If Not ws Is Nothing Then
            writeRow = 19
            For i = 1 To StdCount
                ws.Cells(writeRow, 1).value = StdQty(i)
                ws.Cells(writeRow, 2).value = Replace(stdName(i), Chr(34), "")
                ws.Cells(writeRow, 3).value = StdT(i)
                ws.Cells(writeRow, 5).value = StdW(i)
                ws.Cells(writeRow, 7).value = StdL(i)
                ws.Cells(writeRow, 8).value = StdSteelTypeFor(StdGrade(i))
                writeRow = writeRow + 1
            Next i
            LogLine "Filled '" & CStr(sName) & "' rows 19.." & (writeRow - 1)
        End If
    Next sName

    StampWorkbookDateAndRef xlWb, FormatRefNumber
    ' REF # in I11 (right next to its label); clear the stray J12 placement.
    On Error Resume Next
    Dim refSh As Object, refNm As Variant
    For Each refNm In Array("Steel Order", "Machining Sheet")
        Set refSh = Nothing
        Set refSh = xlWb.Sheets(CStr(refNm))
        If Not refSh Is Nothing Then
            refSh.Cells(11, 9).Value = FormatRefNumber
            refSh.Cells(12, 10).ClearContents
        End If
    Next refNm
    On Error GoTo ErrHandler
    xlWb.Save
    xlWb.Close False
    xlApp.Quit
    Set xlWb = Nothing: Set xlApp = Nothing
    LogLine "J000 steel sheet saved (standard base)."
    Exit Sub
ErrHandler:
    LogLine "FillStandardBaseSteel error: " & Err.Description
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

' ============================================================
' STANDARD BASE: name recognition (CAD names + BOM), fractional
' inch parsing, material->grade->block routing, source builders,
' and the pullcore/key volume quote.
' ============================================================

' (PULLCORE_RATE constant is declared with the other settings at the top.)

Private Function StdCleanName(ByVal raw As String) As String
    Dim s As String
    s = UCase(raw)
    s = Replace(s, "_", " "): s = Replace(s, "-", " "): s = Replace(s, "/", " ")
    s = Replace(s, ".", " "): s = Replace(s, ",", " "): s = Replace(s, Chr(34), " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    StdCleanName = " " & Trim(s) & " "
End Function

' Canonical standard-plate name from a CAD component name OR a BOM description.
' Returns "" if it isn't a recognizable structural plate.
' Exact shop STEP tokens (A-PLATE, EJ-RET-PLATE, SC-RETAINER-PLATE, ...) are
' checked first — same priority as qwen_classify_xt_csv.apply_strong_shop_name_hints.
Private Function StandardPlateNameStd(ByVal raw As String) As String
    Dim s As String
    Dim u As String
    s = StdCleanName(raw)
    u = UCase(raw)
    StandardPlateNameStd = ""
    If IsHardwareName(raw) Then Exit Function

    ' --- Strong shop STEP tokens (hyphen/underscore forms before spaced cleanup) ---
    If InStr(u, "A-PLATE") > 0 Or InStr(u, "A_PLATE") > 0 Then StandardPlateNameStd = "A Plate": Exit Function
    If InStr(u, "B-PLATE") > 0 Or InStr(u, "B_PLATE") > 0 Then StandardPlateNameStd = "B Plate": Exit Function
    If InStr(u, "SC-RETAINER-PLATE") > 0 Or InStr(u, "SC_RETAINER_PLATE") > 0 Then StandardPlateNameStd = "SC Retainer Plate": Exit Function
    If InStr(u, "SC-BACKUP-PLATE") > 0 Or InStr(u, "SC_BACKUP_PLATE") > 0 Then StandardPlateNameStd = "SC Backup Plate": Exit Function
    If InStr(u, "EJ-BACKUP-PLATE") > 0 Or InStr(u, "EJ_BACKUP_PLATE") > 0 Then StandardPlateNameStd = "Bottom Ejector Plate": Exit Function
    If InStr(u, "EJ-RET-PLATE") > 0 Or InStr(u, "EJ_RET_PLATE") > 0 Then StandardPlateNameStd = "Ejector Plate": Exit Function
    If InStr(u, "CLAMP-PLATE") > 0 Or InStr(u, "CLAMP_PLATE") > 0 Then
        If InStr(u, "TOP") = 0 Then StandardPlateNameStd = "Bottom Clamp Plate": Exit Function
    End If

    If InStr(s, " LOWER BASE ") > 0 Then StandardPlateNameStd = "Bottom Clamp Plate": Exit Function
    If InStr(s, " UPPER BASE ") > 0 Then StandardPlateNameStd = "Top Clamp Plate": Exit Function
    If InStr(s, " BASE PLATE ") > 0 Or InStr(s, " BASE ") > 0 Then StandardPlateNameStd = ProperCaseText(raw): Exit Function
    If InStr(s, " BALANCE PLATE ") > 0 Or InStr(s, " BALANCE PLATES ") > 0 Then StandardPlateNameStd = ProperCaseText(raw): Exit Function
    If InStr(s, " SC RETAINER ") > 0 Or InStr(s, " SC RETAINER PLATE ") > 0 Then StandardPlateNameStd = "SC Retainer Plate": Exit Function
    If InStr(s, " SC BACKUP ") > 0 Or InStr(s, " SC BACKUP PLATE ") > 0 Or InStr(s, " SC BACK UP ") > 0 Then StandardPlateNameStd = "SC Backup Plate": Exit Function
    If InStr(s, " EJECTOR RETAINER ") > 0 Or InStr(s, " EJ RET ") > 0 Or InStr(s, " KO RET ") > 0 Then StandardPlateNameStd = "Ejector Plate": Exit Function
    ' CMS naming: the thicker/lower backing plate is the "Bottom Ejector Plate"
    ' (never "Ejector Retainer Plate"). Same PIN quote row as before.
    If InStr(s, " EJECTOR BACKUP ") > 0 Or InStr(s, " EJECTOR BACK UP ") > 0 Or InStr(s, " EJ BACKUP ") > 0 Then StandardPlateNameStd = "Bottom Ejector Plate": Exit Function
    If InStr(s, " RETAINER ") > 0 Or InStr(s, " RETAINER PLATE ") > 0 Then StandardPlateNameStd = ProperCaseText(raw): Exit Function
    If InStr(s, " HOLDER MOUNT ") > 0 Or InStr(s, " MOUNT ") > 0 Then StandardPlateNameStd = ProperCaseText(raw): Exit Function
    If InStr(s, " RUNNER STRIPPER ") > 0 Then StandardPlateNameStd = "Runner Stripper Plate": Exit Function
    If InStr(s, " STRIPPER ") > 0 Or InStr(s, " STRIPPER PLT ") > 0 Then StandardPlateNameStd = "Stripper Plate": Exit Function
    If InStr(s, " TCP ") > 0 Or InStr(s, " TOP CLAMP ") > 0 Or InStr(s, " TOP CLP ") > 0 Then StandardPlateNameStd = "Top Clamp Plate": Exit Function
    If InStr(s, " BCP ") > 0 Or InStr(s, " BOTTOM CLAMP ") > 0 Or InStr(s, " BOT CLAMP ") > 0 Then StandardPlateNameStd = "Bottom Clamp Plate": Exit Function
    If InStr(s, " MANIFOLD ") > 0 Or InStr(s, " MAN PLT ") > 0 Then StandardPlateNameStd = "Manifold Plate": Exit Function
    If InStr(s, " CAVITY PLT ") > 0 Or InStr(s, " CAVITY PLATE ") > 0 Or InStr(s, " CAVITY RETAINER ") > 0 Or InStr(s, " STATIONARY RETAINER ") > 0 Then StandardPlateNameStd = "A Plate": Exit Function
    If InStr(s, " CORE PLT ") > 0 Or InStr(s, " CORE PLATE ") > 0 Or InStr(s, " CORE RETAINER ") > 0 Or InStr(s, " MOVABLE RETAINER ") > 0 Or InStr(s, " MOVEABLE RETAINER ") > 0 Then StandardPlateNameStd = "B Plate": Exit Function
    If InStr(s, " DIE BACK UP ") > 0 Or InStr(s, " DIE BACKUP ") > 0 Or InStr(s, " BACK UP PLT ") > 0 Or InStr(s, " BACKUP PLT ") > 0 Then StandardPlateNameStd = "Die Backup Plate": Exit Function
    If InStr(s, " DIE PLATE ") > 0 Or InStr(s, " DIE PLT ") > 0 Then StandardPlateNameStd = "Die Plate": Exit Function
    If InStr(s, " PIN PLATE ") > 0 Or InStr(s, " PIN PLT ") > 0 Then StandardPlateNameStd = "Ejector Plate": Exit Function
    If InStr(s, " BOTTOM EJECTOR ") > 0 Or InStr(s, " EJECTOR BACKUP ") > 0 Or InStr(s, " EJECTOR BACK UP ") > 0 Then StandardPlateNameStd = "Bottom Ejector Plate": Exit Function
    If InStr(s, " EJECTOR PLATE ") > 0 Or InStr(s, " EJECTOR PLT ") > 0 Or InStr(s, " KNOCKOUT ") > 0 Or InStr(s, " KO PLT ") > 0 Or InStr(s, " EJ PLT ") > 0 Then StandardPlateNameStd = "Bottom Ejector Plate": Exit Function
    If InStr(s, " SUPPORT PLATE ") > 0 Or InStr(s, " SUPPORT PLT ") > 0 Or InStr(s, " SUP PLT ") > 0 Or InStr(s, " SUPPORT PLAT ") > 0 Then StandardPlateNameStd = "Support Plate": Exit Function
    If InStr(s, " RAIL ") > 0 Or InStr(s, " RAILS ") > 0 Or InStr(s, " RISER ") > 0 Or InStr(s, " RISERS ") > 0 Or InStr(s, " SPACER ") > 0 Then StandardPlateNameStd = ProperCaseText(raw): Exit Function
    If InStr(s, " A PLATE ") > 0 Or InStr(s, " A PLT ") > 0 Or InStr(s, " A SIDE ") > 0 Then StandardPlateNameStd = "A Plate": Exit Function
    If InStr(s, " B PLATE ") > 0 Or InStr(s, " B PLT ") > 0 Or InStr(s, " B SIDE ") > 0 Then StandardPlateNameStd = "B Plate": Exit Function
    If InStr(s, " X PLATE ") > 0 Or InStr(s, " X PLT ") > 0 Then StandardPlateNameStd = """X"" Plate": Exit Function
    If InStr(s, " Y PLATE ") > 0 Or InStr(s, " Y PLT ") > 0 Then StandardPlateNameStd = """Y"" Plate": Exit Function
End Function

' Slot keyword for a canonical name (used to find the row in a grade block).
Private Function StdSlotForName(ByVal nm As String) As String
    Dim s As String
    s = StdCleanName(nm)
    StdSlotForName = ""
    If InStr(s, " RUNNER STRIPPER ") > 0 Then StdSlotForName = "X": Exit Function
    If InStr(s, " STRIPPER ") > 0 Then StdSlotForName = "STRIPPER": Exit Function
    If InStr(s, " TOP CLAMP ") > 0 Then StdSlotForName = "TOPCLAMP": Exit Function
    If InStr(s, " BOTTOM CLAMP ") > 0 Then StdSlotForName = "BOTTOMCLAMP": Exit Function
    If InStr(s, " A PLATE ") > 0 Or InStr(s, " CAVITY ") > 0 Then StdSlotForName = "A": Exit Function
    If InStr(s, " B PLATE ") > 0 Or InStr(s, " CORE ") > 0 Then StdSlotForName = "B": Exit Function
    If InStr(s, " SC RETAINER ") > 0 Then StdSlotForName = "B": Exit Function
    If InStr(s, " SC BACKUP ") > 0 Or InStr(s, " SC BACK UP ") > 0 Then StdSlotForName = "SUPPORT": Exit Function
    If InStr(s, " EJECTOR BACKUP ") > 0 Or InStr(s, " EJECTOR BACK UP ") > 0 Or InStr(s, " EJ BACKUP ") > 0 Then StdSlotForName = "PIN": Exit Function
    If InStr(s, " EJ RET ") > 0 Then StdSlotForName = "EJECTOR": Exit Function
    ' CMS rule: "Bottom Ejector Plate" is the thicker/lower plate (same physical
    ' plate the EJ-BACKUP token names) -> same PIN row it always used. The
    ' thinner "Ejector Plate" keeps the EJECTOR row below.
    If InStr(s, " BOTTOM EJECTOR ") > 0 Then StdSlotForName = "PIN": Exit Function
    If InStr(s, " EJECTOR RETAINER ") > 0 Or InStr(s, " PIN PLATE ") > 0 Or InStr(s, " PIN ") > 0 Or InStr(s, " RETAINER ") > 0 Then StdSlotForName = "PIN": Exit Function
    If InStr(s, " EJECTOR BACKUP ") > 0 Or InStr(s, " EJECTOR BACK UP ") > 0 Or InStr(s, " EJECTOR ") > 0 Then StdSlotForName = "EJECTOR": Exit Function
    If InStr(s, " SUPPORT ") > 0 Or InStr(s, " PILLAR ") > 0 Then StdSlotForName = "SUPPORT": Exit Function
    If InStr(s, " RAIL ") > 0 Or InStr(s, " RAILS ") > 0 Or InStr(s, " RISER ") > 0 Or InStr(s, " RISERS ") > 0 Then StdSlotForName = "RAILS": Exit Function
    If InStr(s, " DIE BACKUP ") > 0 Or InStr(s, " DIE BACK UP ") > 0 Then StdSlotForName = "SUPPORT": Exit Function
    If InStr(s, " DIE ") > 0 Then StdSlotForName = "X": Exit Function
    If InStr(s, " MANIFOLD ") > 0 Then StdSlotForName = "MANIFOLD": Exit Function
    If InStr(s, " BALANCE ") > 0 Then StdSlotForName = "A": Exit Function
    If InStr(s, " X PLATE ") > 0 Then StdSlotForName = "X": Exit Function
    If InStr(s, " Y PLATE ") > 0 Then StdSlotForName = "Y": Exit Function
    If InStr(s, " PLATE ") > 0 Then StdSlotForName = "A": Exit Function
End Function

' Decide the steel block for a plate from its material hint (and slot if unknown).
Private Function DefaultStandardGradeForSlot(ByVal slot As String) As String
    Select Case UCase(Trim(slot))
        Case "A", "B"
            DefaultStandardGradeForSlot = STD_A_B_GRADE   ' normally P20

        Case "X", "Y", "MANIFOLD", "STRIPPER"
            DefaultStandardGradeForSlot = STD_A_B_GRADE

        Case Else
            DefaultStandardGradeForSlot = "A36"
    End Select
End Function

Private Function ResolveStdGrade(ByVal gradeHint As String, ByVal slot As String) As String
    Dim g As String
    g = NormalizeSteelType(gradeHint)

    If g = "" Or g = "STEEL" Or g = "MATERIAL" Or g = "OUTSOURCE" Then
        ResolveStdGrade = DefaultStandardGradeForSlot(slot)
        Exit Function
    End If

    Select Case g
        Case "P20", "420SS", "6061", "4140", "A2", "O1"
            ResolveStdGrade = g

        Case "A36"
            ResolveStdGrade = "A36"

        Case Else
            ResolveStdGrade = DefaultStandardGradeForSlot(slot)
    End Select
End Function

' Quote row for (slot, grade). 0 if that block has no row for the slot.
Private Function StdQuoteRowFor(ByVal slot As String, ByVal grade As String) As Long
    Dim sl As String
    sl = slot
    If sl = "STRIPPER" Then sl = "B"     ' stripper quoted in the B row (per shop example)
    StdQuoteRowFor = 0
    Select Case UCase(grade)
        Case "4140"     ' #2 block (shares the pot block; clamps + spare rows)
            Select Case sl
                Case "TOPCLAMP": StdQuoteRowFor = 22
                Case "BOTTOMCLAMP": StdQuoteRowFor = 23
                Case "RAILS": StdQuoteRowFor = 24
                Case "PIN": StdQuoteRowFor = 25
                Case "SUPPORT": StdQuoteRowFor = 26
                Case "A": StdQuoteRowFor = 27
                Case "B": StdQuoteRowFor = 28
                Case "MANIFOLD": StdQuoteRowFor = 29
                Case "EJECTOR": StdQuoteRowFor = 30
                Case "X": StdQuoteRowFor = 31
                Case "Y": StdQuoteRowFor = 32
            End Select
        Case "P20"
            Select Case sl
                Case "TOPCLAMP": StdQuoteRowFor = 38
                Case "A": StdQuoteRowFor = 39
                Case "B": StdQuoteRowFor = 40
                Case "X": StdQuoteRowFor = 41
                Case "Y": StdQuoteRowFor = 42
                Case "SUPPORT": StdQuoteRowFor = 43
            End Select
        Case "420SS"
            Select Case sl
                Case "TOPCLAMP": StdQuoteRowFor = 52
                Case "MANIFOLD": StdQuoteRowFor = 53
                Case "A": StdQuoteRowFor = 54
                Case "B": StdQuoteRowFor = 55
                Case "Y": StdQuoteRowFor = 56
                Case "SUPPORT": StdQuoteRowFor = 57
                Case "RAILS": StdQuoteRowFor = 58
                Case "BOTTOMCLAMP": StdQuoteRowFor = 59
                Case "PIN": StdQuoteRowFor = 60
                Case "EJECTOR": StdQuoteRowFor = 61
            End Select
        Case "6061"
            Select Case sl
                Case "TOPCLAMP": StdQuoteRowFor = 68
                Case "A": StdQuoteRowFor = 69
                Case "B": StdQuoteRowFor = 70
                Case "X": StdQuoteRowFor = 71
                Case "Y": StdQuoteRowFor = 72
                Case "SUPPORT": StdQuoteRowFor = 73
                Case "RAILS": StdQuoteRowFor = 74
                Case "BOTTOMCLAMP": StdQuoteRowFor = 75
                Case "PIN": StdQuoteRowFor = 76
                Case "EJECTOR": StdQuoteRowFor = 77
            End Select
        Case "A2", "O1"
            Select Case sl
                Case "TOPCLAMP": StdQuoteRowFor = 6
                Case "MANIFOLD": StdQuoteRowFor = 7
                Case "A": StdQuoteRowFor = 8
                Case "B": StdQuoteRowFor = 9
                Case "SUPPORT": StdQuoteRowFor = 10
                Case "BOTTOMCLAMP": StdQuoteRowFor = 11
                Case "RAILS": StdQuoteRowFor = 12
                Case "PIN": StdQuoteRowFor = 14
                Case "EJECTOR": StdQuoteRowFor = 15
            End Select
        Case Else      ' A-36 (#1 block) for A36 / 1045 / H13 / unknown
            Select Case sl
                Case "TOPCLAMP": StdQuoteRowFor = 6
                Case "MANIFOLD": StdQuoteRowFor = 7
                Case "A": StdQuoteRowFor = 8
                Case "B": StdQuoteRowFor = 9
                Case "SUPPORT": StdQuoteRowFor = 10
                Case "BOTTOMCLAMP": StdQuoteRowFor = 11
                Case "RAILS": StdQuoteRowFor = 12
                Case "PIN": StdQuoteRowFor = 14
                Case "EJECTOR": StdQuoteRowFor = 15
            End Select
    End Select
End Function

' Parse one inch token that may be a decimal or a (whole-)fraction: 9-7/8, 7/8,
' 1-3/8, 1.375, .875, 1-9/16, etc. Returns inches.
Private Function ParseInchToken(ByVal tok As String) As Double
    Dim s As String, j As Long, ch As String, t As String
    s = Replace(tok, "-", " ")
    t = ""
    For j = 1 To Len(s)
        ch = Mid(s, j, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "/" Or ch = " " Then t = t & ch
    Next j
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    t = Trim(t)
    If t = "" Then Exit Function
    Dim total As Double, parts() As String, i As Long, p As String, fp() As String
    total = 0
    parts = Split(t, " ")
    For i = LBound(parts) To UBound(parts)
        p = parts(i)
        If p <> "" Then
            If InStr(p, "/") > 0 Then
                fp = Split(p, "/")
                If UBound(fp) = 1 Then
                    If IsNumeric(fp(0)) And IsNumeric(fp(1)) Then
                        If CDbl(fp(1)) <> 0 Then total = total + CDbl(fp(0)) / CDbl(fp(1))
                    End If
                End If
            ElseIf IsNumeric(p) Then
                total = total + CDbl(p)
            End If
        End If
    Next i
    ParseInchToken = total
End Function

' Parse "name, T x W x L" (decimals or fractions) -> T,W,L. True if 3 dims found.
Private Function ParseInchDimsFromText(ByVal text As String, ByRef t As Double, ByRef w As Double, ByRef l As Double) As Boolean
    Dim s As String
    s = text
    Dim cpos As Long
    cpos = InStr(s, ",")
    If cpos > 0 Then s = Mid(s, cpos + 1)
    s = Replace(s, "?", "x"): s = Replace(s, "X", "x")
    Dim parts() As String
    parts = Split(s, "x")
    Dim vals() As Double
    ReDim vals(1 To 12)
    Dim n As Long, i As Long, v As Double
    n = 0
    For i = LBound(parts) To UBound(parts)
        v = ParseInchToken(parts(i))
        If v > 0 Then n = n + 1: If n <= 12 Then vals(n) = v
    Next i
    If n < 3 Then ParseInchDimsFromText = False: Exit Function
    Dim a As Double, b As Double, c As Double
    PickThreeLargest vals, n, a, b, c
    SortThreeDimensions a, b, c, l, w, t
    ParseInchDimsFromText = True
End Function

Private Sub StdResetArrays()
    StdCount = 0
    ReDim stdName(1 To 80)
    ReDim StdT(1 To 80)
    ReDim StdW(1 To 80)
    ReDim StdL(1 To 80)
    ReDim StdQty(1 To 80)
    ReDim StdGrade(1 To 80)
    ReDim StdQuoteRow(1 To 80)
    ReDim StdCadIndex(1 To 80)
    If PartCount > 0 Then
        ReDim gStdRoleByPart(1 To PartCount)
        ReDim gStdLeaderPinSetByPart(1 To PartCount)
    Else
        Erase gStdRoleByPart
        Erase gStdLeaderPinSetByPart
    End If
    gStdStackAxis = 0
    gStdTopIsFirst = True
    gStdDmeStackFamily = ""
    gStdPartingLineAxis = 0
    gStdPartingLinePos = 0#
    gStdCavityCadIndex = 0
    gStdCoreCadIndex = 0
    gStdLeaderPinFromTop = False
    gStdLeaderPinFromKnown = False
    gStdLeaderPinReversed = False
    gStdSequencedLatchLock = False
    gStdStackRules = ""
    gStdPartingLineText = ""
End Sub

Private Sub SetStdCadRole(ByVal idx As Long, ByVal roleName As String)
    On Error Resume Next
    If idx < 1 Or idx > PartCount Then Exit Sub
    If Not StdRoleArrayReady() Then ReDim gStdRoleByPart(1 To PartCount)
    gStdRoleByPart(idx) = roleName
End Sub

Private Function StdCadRole(ByVal idx As Long) As String
    On Error Resume Next
    If idx < 1 Or idx > PartCount Then Exit Function
    If Not StdRoleArrayReady() Then Exit Function
    StdCadRole = gStdRoleByPart(idx)
End Function

Private Function StdRoleArrayReady() As Boolean
    On Error GoTo nope
    Dim hi As Long
    hi = UBound(gStdRoleByPart)
    StdRoleArrayReady = (hi >= 1)
    Exit Function
nope:
    StdRoleArrayReady = False
End Function

Private Function FindStdByName(ByVal nm As String) As Long
    Dim i As Long, k As String
    k = NormalizeKey(nm)
    For i = 1 To StdCount
        If NormalizeKey(stdName(i)) = k Then FindStdByName = i: Exit Function
    Next i
End Function

Private Function BomMaterialRowHasCadMatch(ByVal bomIdx As Long) As Boolean
    BomMaterialRowHasCadMatch = False
    If bomIdx < 1 Or bomIdx > BomCount Then Exit Function
    If BomRows(bomIdx).hasDims = False Then Exit Function
    Dim i As Long, d As Double
    For i = 1 To PartCount
        d = Abs(parts(i).Length - BomRows(bomIdx).BomLength) + _
            Abs(parts(i).Width - BomRows(bomIdx).BomWidth) + _
            Abs(parts(i).Thickness - BomRows(bomIdx).BomThickness)
        If d <= DIM_REVIEW_TOL * 3# Then BomMaterialRowHasCadMatch = True: Exit Function
    Next i
End Function

' Build the plate list from the BOM (names, dims, material).
Private Function BuildStdFromBom() As Boolean
    Dim i As Long, nm As String, added As Long
    added = 0
    For i = 1 To BomCount
        nm = StandardPlateNameStd(BomRows(i).Description)
        If nm <> "" And BomRows(i).hasDims Then
            If BomMaterialRowHasCadMatch(i) Then
                AddStdPlate nm, BomRows(i).BomThickness, BomRows(i).BomWidth, BomRows(i).BomLength, BomRows(i).Quantity, BomRows(i).material
                added = added + 1
            Else
                LogLine "BOM material not quoted (no CAD size match): " & BomRows(i).Description
            End If
        End If
    Next i
    BuildStdFromBom = (added > 0)
End Function

' Build the plate list from CAD component names. True only if it found enough.
Private Function BuildStdFromCadNames() As Boolean
    Dim i As Long, nm As String, added As Long
    added = 0
    For i = 1 To PartCount
        nm = StandardPlateNameStd(parts(i).componentName)
        If nm <> "" Then
            AddStdPlate nm, parts(i).Thickness, parts(i).Width, parts(i).Length, 1, ""
            added = added + 1
        End If
    Next i
    BuildStdFromCadNames = (added >= 3)
End Function

' ============================================================
' AI BRIDGE implementation
' ============================================================

' Job token used for bridge file names and web-app job registration.
Private Function AiBridgeJobToken() As String
    Dim t As String
    t = Trim(CurrentJobNumber)
    If t = "" Then t = Trim(AssignedQuoteNumber)
    If t = "" And CurrentJobFolder <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        t = fso.GetFileName(CurrentJobFolder)
    End If
    If t = "" Then t = "ACTIVE"
    AiBridgeJobToken = t
End Function

Private Sub AiBridgeReset()
    gAiRoleCount = 0
    gAiBridgeUsed = False
    gAiSequencedLatchLock = False
    If PartCount >= 1 Then
        ReDim gAiRoleByPart(1 To PartCount)
        ReDim gAiConfByPart(1 To PartCount)
    End If
End Sub

Private Function AiJsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr(34), "\" & Chr(34))
    AiJsonEscape = s
End Function

' POST the CAD CSV path to the LOCAL AI service. Returns the bridge CSV text
' (Index,Component,Role,ResolvedName,Confidence,Quote,Price,SecondaryPartingLine)
' or "" when the service is unreachable.
Private Function AiBridgeHttpClassify(ByVal csvPath As String) As String
    AiBridgeHttpClassify = ""
    On Error GoTo done
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 3000, 3000, AI_BRIDGE_TIMEOUT_MS, AI_BRIDGE_TIMEOUT_MS
    http.Open "POST", AI_BRIDGE_URL & "/api/vba/classify", False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send "{""job_id"":""" & AiJsonEscape(AiBridgeJobToken()) & _
              """,""csv_path"":""" & AiJsonEscape(csvPath) & _
              """,""base_type"":""standard""}"
    If http.Status = 200 Then AiBridgeHttpClassify = http.responseText
done:
End Function

' Fallback: read a bridge CSV the web app exported earlier for this job.
Private Function AiBridgeReadFallbackFile() As String
    AiBridgeReadFallbackFile = ""
    On Error GoTo done
    Dim fso As Object, p As String, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = AI_BRIDGE_FILE_DIR & "\" & AiBridgeJobToken() & "_part_names.csv"
    If Not fso.FileExists(p) Then Exit Function
    Set ts = fso.OpenTextFile(p, 1)
    AiBridgeReadFallbackFile = ts.ReadAll
    ts.Close
    LogLine "AI bridge: using fallback file " & p
done:
End Function

' Parse the bridge CSV into gAiRoleByPart/gAiConfByPart (by CAD part Index).
Private Sub AiBridgeParseCsv(ByVal csvText As String)
    Dim lines() As String, i As Long, f() As String
    Dim idx As Long, roleCol As Long, confCol As Long
    csvText = Replace(csvText, vbCrLf, vbLf)
    csvText = Replace(csvText, vbCr, vbLf)
    lines = Split(csvText, vbLf)
    If UBound(lines) < 1 Then Exit Sub

    ' Locate the Role / Confidence columns from the header so the format can
    ' grow new columns without breaking this parser.
    Dim hdr() As String, c As Long
    hdr = Split(lines(0), ",")
    roleCol = -1: confCol = -1
    For c = 0 To UBound(hdr)
        Select Case UCase(Trim(hdr(c)))
            Case "ROLE": roleCol = c
            Case "CONFIDENCE": confCol = c
        End Select
    Next c
    If roleCol < 0 Then Exit Sub

    For i = 1 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            f = Split(lines(i), ",")
            If UBound(f) >= roleCol Then
                idx = Val(f(0))
                If idx >= 1 And idx <= PartCount Then
                    gAiRoleByPart(idx) = LCase(Trim(f(roleCol)))
                    If confCol >= 0 And UBound(f) >= confCol Then
                        gAiConfByPart(idx) = UCase(Trim(f(confCol)))
                    Else
                        gAiConfByPart(idx) = "MEDIUM"
                    End If
                    gAiRoleCount = gAiRoleCount + 1
                    If gAiRoleByPart(idx) = "latch_lock" Then gAiSequencedLatchLock = True
                End If
            End If
        End If
    Next i
End Sub

' Map an AI role key to the macro's standard plate name. "" = not a quoted plate.
Private Function AiPlateNameForRole(ByVal roleKey As String) As String
    Select Case LCase(Trim(roleKey))
        Case "top_clamp_plate": AiPlateNameForRole = "Top Clamp Plate"
        Case "a_plate": AiPlateNameForRole = "A Plate"
        Case "b_plate": AiPlateNameForRole = "B Plate"
        Case "stripper_plate": AiPlateNameForRole = "Stripper Plate"
        Case "sc_retainer_plate": AiPlateNameForRole = "SC Retainer Plate"
        Case "sc_backup_plate": AiPlateNameForRole = "SC Backup Plate"
        Case "support_plate": AiPlateNameForRole = "Support Plate"
        Case "bottom_clamp_plate": AiPlateNameForRole = "Bottom Clamp Plate"
        Case "rail", "rail_1", "rail_2": AiPlateNameForRole = "Rails"
        Case "pin_plate": AiPlateNameForRole = "Pin Plate"
        Case "ejector_plate": AiPlateNameForRole = "Ejector Plate"
        ' CMS naming: the thicker/lower ejector-stack plate is the Bottom
        ' Ejector Plate (never "Ejector Retainer Plate").
        Case "bottom_ejector_plate", "ejector_retainer_plate", "ejector_backup_plate"
            AiPlateNameForRole = "Bottom Ejector Plate"
        Case Else: AiPlateNameForRole = ""
    End Select
End Function

' Round-hardware role names for SetStdCadRole (naming analysis / logs only).
Private Function AiHardwareNameForRole(ByVal roleKey As String) As String
    Select Case LCase(Trim(roleKey))
        Case "leader_pin": AiHardwareNameForRole = "Leader Pin"
        Case "leader_pin_bushing": AiHardwareNameForRole = "Leader Pin Bushing"
        Case "guided_ejector_bushing": AiHardwareNameForRole = "Guided Ejector Bushing"
        Case "return_pin": AiHardwareNameForRole = "Return Pin"
        Case "ejector_pin": AiHardwareNameForRole = "Ejector Pin"
        Case "support_pillar": AiHardwareNameForRole = "Support Pillar"
        Case "latch_lock": AiHardwareNameForRole = "Latch Lock / Safety Strap"
        Case Else: AiHardwareNameForRole = ""
    End Select
End Function

' Main entry: classify this job through the local AI service.
' HARD GUARD: BMS / pot-block jobs never touch the AI - their BOM-driven
' flow already works and must not be disturbed.
Public Sub RunAiBridgeClassification(ByVal csvPath As String, ByVal isStandardBase As Boolean)
    On Error GoTo eh
    AiBridgeReset
    If Not AI_BRIDGE_ENABLED Then Exit Sub
    If Not isStandardBase Then
        LogLine "AI bridge: SKIPPED - BMS/pot-block base keeps the BOM-driven flow (AI is for standard bases only)."
        ' Still register the job with the local app (marked BMS) so it shows in
        ' the dashboard, but no classification is applied.
        AiBridgeNotifyBms csvPath
        Exit Sub
    End If

    Dim csvText As String
    csvText = AiBridgeHttpClassify(csvPath)
    If csvText = "" Then csvText = AiBridgeReadFallbackFile()
    If csvText = "" Then
        LogLine "AI bridge: local AI service not reachable at " & AI_BRIDGE_URL & " and no fallback file - macro geometry rules will run alone."
        Exit Sub
    End If

    AiBridgeParseCsv csvText
    LogLine "AI bridge: " & gAiRoleCount & " part roles received." & _
            IIf(gAiSequencedLatchLock, " Plate-sequenced LATCH-LOCK base detected: latch locks mark secondary parting lines and must not flip A/B.", "")
    Exit Sub
eh:
    LogLine "AI bridge error (ignored, geometry rules continue): " & Err.Description
End Sub

' Register a BMS/pot-block job with the local app dashboard (no classification).
Private Sub AiBridgeNotifyBms(ByVal csvPath As String)
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 2000, 2000, 5000, 5000
    http.Open "POST", AI_BRIDGE_URL & "/api/vba/classify", False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send "{""job_id"":""" & AiJsonEscape(AiBridgeJobToken()) & _
              """,""csv_path"":""" & AiJsonEscape(csvPath) & _
              """,""base_type"":""bms""}"
End Sub

' Tell the local webapp the job finished so it can import quote/steel/CSV outputs.
Private Sub AiBridgeNotifyJobComplete(ByVal baseType As String)
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 2000, 2000, 15000, 15000
    http.Open "POST", AI_BRIDGE_URL & "/api/vba/job-complete", False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send "{""job_id"":""" & AiJsonEscape(AiBridgeJobToken()) & _
              """,""folder_path"":""" & AiJsonEscape(CurrentJobFolder) & _
              """,""base_type"":""" & AiJsonEscape(baseType) & _
              """,""status"":""completed""}"
    If http.Status = 200 Then
        LogLine "AI bridge: job-complete synced to webapp."
    End If
End Sub

' Build the standard plate list from AI roles. Returns True when the AI gave
' enough plate roles to define the stack; otherwise the caller falls back to
' the macro's own geometry pass.
Private Function BuildStdFromAiBridge() As Boolean
    BuildStdFromAiBridge = False
    If gAiRoleCount < 1 Then Exit Function

    Dim i As Long, nm As String, hw As String, plateCount As Long
    Dim railT As Double, railW As Double, railL As Double, nRail As Long

    ' First pass: count distinct AI plate roles so we only take over when the
    ' AI actually resolved a stack (>= AI_BRIDGE_MIN_PLATES full plates).
    For i = 1 To PartCount
        nm = AiPlateNameForRole(gAiRoleByPart(i))
        If nm <> "" And nm <> "Rails" Then plateCount = plateCount + 1
    Next i
    If plateCount < AI_BRIDGE_MIN_PLATES Then
        LogLine "AI bridge: only " & plateCount & " plate roles (< " & AI_BRIDGE_MIN_PLATES & ") - macro geometry rules stay in charge."
        Exit Function
    End If

    nRail = 0
    For i = 1 To PartCount
        nm = AiPlateNameForRole(gAiRoleByPart(i))
        If nm = "Rails" Then
            SetStdCadRole i, "Rails"
            If nRail = 0 Then
                railT = parts(i).Thickness: railW = parts(i).Width: railL = parts(i).Length
            End If
            nRail = nRail + 1
        ElseIf nm <> "" Then
            SetStdCadRole i, nm
            Select Case NormalizeKey(nm)
                Case "APLATE": gStdCavityCadIndex = i
                Case "BPLATE": gStdCoreCadIndex = i
            End Select
            AddStdPlateFromCad i, nm
            LogLine "AI bridge plate: idx " & i & " -> " & nm & " [" & gAiConfByPart(i) & "]" & _
                    " | T=" & parts(i).Thickness & " W=" & parts(i).Width & " L=" & parts(i).Length & _
                    " | name=" & parts(i).cleanName
        Else
            hw = AiHardwareNameForRole(gAiRoleByPart(i))
            If hw <> "" Then SetStdCadRole i, hw
        End If
    Next i
    If nRail > 0 Then
        AddStdPlate "Rails", railT, railW, railL, nRail
        LogLine "AI bridge rails: qty " & nRail
    End If

    ' Carry latch-lock flag + Qwen-parity leader-pin set / top-bottom direction
    ' into the same globals the offline geometry path fills.
    If gAiSequencedLatchLock Then gStdSequencedLatchLock = True

    Dim ax As Integer, bestRange As Double, a As Integer, mn As Double, mx As Double, v As Double
    Dim fullIdx(1 To 60) As Long, nFull As Long
    Dim lpIdx(1 To 120) As Long, nLp As Long
    Dim nEj As Long
    Dim baseFoot As Double, fp As Double
    baseFoot = 0#: nFull = 0: nLp = 0: nEj = 0
    For i = 1 To PartCount
        fp = parts(i).Width * parts(i).Length
        If fp > baseFoot Then baseFoot = fp
    Next i
    For i = 1 To PartCount
        nm = AiPlateNameForRole(gAiRoleByPart(i))
        hw = AiHardwareNameForRole(gAiRoleByPart(i))
        If nm <> "" And nm <> "Rails" And baseFoot > 0# Then
            If parts(i).Width * parts(i).Length >= (1 - STD_FOOTPRINT_TOL) * baseFoot Then
                If nFull < UBound(fullIdx) Then nFull = nFull + 1: fullIdx(nFull) = i
            End If
        End If
        If NormalizeKey(nm) = "EJECTORPLATE" Or NormalizeKey(nm) = "BOTTOMEJECTORPLATE" Then nEj = nEj + 1
        If hw = "Leader Pin" Or hw = "Leader Pin Bushing" Or hw = "Guided Ejector Bushing" Then
            If nLp < UBound(lpIdx) Then nLp = nLp + 1: lpIdx(nLp) = i
        End If
    Next i

    bestRange = -1: ax = 3
    If nFull >= 2 Then
        For a = 1 To 3
            mn = 1E+30: mx = -1E+30
            For i = 1 To nFull
                v = PartAxisCenter(fullIdx(i), a)
                If v < mn Then mn = v
                If v > mx Then mx = v
            Next i
            If (mx - mn) > bestRange Then bestRange = (mx - mn): ax = a
        Next a
    End If
    gStdStackAxis = ax
    gStdTopIsFirst = True
    If gStdCavityCadIndex > 0 And gStdCoreCadIndex > 0 Then StdSetPartingLineFromRoles ax

    Dim supportPos As Double
    supportPos = 0#
    For i = 1 To PartCount
        If NormalizeKey(StdCadRole(i)) = "SUPPORTPLATE" Or NormalizeKey(StdCadRole(i)) = "SCBACKUPPLATE" Then
            supportPos = PartAxisCenter(i, ax)
            Exit For
        End If
    Next i
    ClassifyLeaderPinSetsByBushingPlane lpIdx, nLp, ax, supportPos
    MeasureLeaderPinTopBottomDirection ax
    BuildStdStackAnalysisText ax, nFull, nRail, nEj, nLp
    gStdDmeStackFamily = "AI bridge stack (" & StdStackAxisName(ax) & ")" & _
        IIf(gStdSequencedLatchLock, " + latch-lock sequenced", "") & _
        IIf(nLp > 0, " + leader-pin stack", "")

    gAiBridgeUsed = True
    BuildStdFromAiBridge = True
End Function

' Standard-base plate source priority: AI bridge roles first (shop-token +
' latch-lock aware), then geometry/layout, then BOM/names if geometry is
' unavailable. Names help label plates, but they do not define the stack.
Private Sub ClassifyStandardBasePlates()
    StdResetArrays
    If BuildStdFromAiBridge() Then
        LogLine "Standard base plates: AI bridge classification in charge."
        GoTo finishStd
    End If

    StdResetArrays
    BuildStdFromGeometry
    If StdCount >= 3 Then GoTo finishStd

    StdResetArrays
    If BuildStdFromBom() Then
        If StdCount >= 3 Then GoTo finishStd
    End If

    StdResetArrays
    If BuildStdFromCadNames() Then GoTo finishStd

finishStd:
    Dim i As Long
    LogLine "Standard base plates identified: " & StdCount
    For i = 1 To StdCount
        LogLine "  STD " & Replace(stdName(i), Chr(34), "") & " qty " & StdQty(i) & _
                " T=" & StdT(i) & " W=" & StdW(i) & " L=" & StdL(i) & " -> " & StdGrade(i) & " row " & StdQuoteRow(i)
    Next i
    If WRITE_STACK_LEADERPIN_ANALYSIS And CurrentJobFolder <> "" Then
        WriteStackLeaderPinAnalysis CurrentJobFolder & "\Stack_LeaderPin_Analysis.csv", True
    End If
End Sub

' ---- Pullcore / key straight quote: total volume x rate ----
' ---- Pull-core / key name matching (ported from the shop's pullcore logic) ----
Private Function GetPullcoreLocationCode(ByVal text As String) As String
    Dim s As String
    s = NormalizeText(text)
    If InStr(s, "IDTE") > 0 Or InStr(s, "ID TE") > 0 Then GetPullcoreLocationCode = "IDTE": Exit Function
    If InStr(s, "IDLE") > 0 Or InStr(s, "ID LE") > 0 Then GetPullcoreLocationCode = "IDLE": Exit Function
    If InStr(s, "ODTE") > 0 Or InStr(s, "OD TE") > 0 Then GetPullcoreLocationCode = "ODTE": Exit Function
    If InStr(s, "ODLE") > 0 Or InStr(s, "OD LE") > 0 Then GetPullcoreLocationCode = "ODLE": Exit Function
    If InStr(s, "ID") > 0 And InStr(s, "OD") = 0 Then GetPullcoreLocationCode = "ID": Exit Function
    If InStr(s, "OD") > 0 And InStr(s, "ID") = 0 Then GetPullcoreLocationCode = "OD": Exit Function
    Dim toks() As String, i As Long
    toks = Split(s, " ")
    For i = LBound(toks) To UBound(toks)
        If toks(i) = "TE" Then GetPullcoreLocationCode = "TE": Exit Function
        If toks(i) = "LE" Then GetPullcoreLocationCode = "LE": Exit Function
    Next i
    GetPullcoreLocationCode = ""
End Function

Private Function HasPullcoreLocationToken(ByVal d As String) As Boolean
    If GetPullcoreLocationCode(d) <> "" Then HasPullcoreLocationToken = True: Exit Function
    Dim toks() As String, i As Long
    toks = Split(d, " ")
    For i = LBound(toks) To UBound(toks)
        Select Case toks(i)
            Case "TE", "LE", "ID", "OD", "IDTE", "IDLE", "ODTE", "ODLE": HasPullcoreLocationToken = True: Exit Function
        End Select
    Next i
End Function

Private Function CleanPullcoreDisplayName(ByVal s As String) As String
    s = Trim(s)
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    CleanPullcoreDisplayName = s
End Function

' A pull-core cam or key: must say CAM or KEY, must NOT be an ejector/flipper/
' cover/dirt/stop/J-block/holder/plate/smed/pot, and either says PULLCORE or
' carries a pull-core location token (ID/OD/TE/LE...).
Private Function IsPullcoreDesc(ByVal raw As String) As Boolean
    Dim d As String
    d = NormalizeText(raw)
    If InStr(d, "CAM") = 0 And InStr(d, "KEY") = 0 Then Exit Function
    If InStr(d, "EJECTOR") > 0 Then Exit Function
    If InStr(d, "FLIPPER") > 0 Then Exit Function
    If InStr(d, "COVER") > 0 Then Exit Function
    If InStr(d, "DIRT") > 0 Then Exit Function
    If InStr(d, "STOP") > 0 Then Exit Function
    If InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Or InStr(d, "JBLOCK") > 0 Then Exit Function
    If InStr(d, "HOLDER") > 0 Then Exit Function
    If InStr(d, "PLATE") > 0 Then Exit Function
    If InStr(d, "SMED") > 0 Then Exit Function
    If InStr(d, "POT") > 0 Then Exit Function
    If InStr(d, "PULLCORE") > 0 Or InStr(d, "PULL CORE") > 0 Then IsPullcoreDesc = True: Exit Function
    If HasPullcoreLocationToken(d) Then IsPullcoreDesc = True
End Function

Private Function IsPullcoreOrKeyName(ByVal raw As String) As Boolean
    IsPullcoreOrKeyName = IsPullcoreDesc(raw)
End Function

Private Function IsPullcoreBomInfo(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)

    If IsPullcoreDesc(d) Then IsPullcoreBomInfo = True: Exit Function

    If InStr(d, "EJECTOR") > 0 Then Exit Function
    If InStr(d, "FLIPPER") > 0 Then Exit Function
    If InStr(d, "COVER") > 0 Then Exit Function
    If InStr(d, "DIRT") > 0 Then Exit Function
    If InStr(d, "STOP") > 0 Then Exit Function
    If InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Or InStr(d, "JBLOCK") > 0 Then Exit Function
    If InStr(d, "HOLDER") > 0 Then Exit Function
    If InStr(d, "MOLD BASE") > 0 Or InStr(d, "MOLDBASE") > 0 Then Exit Function
    If InStr(d, "PLATE") > 0 Then Exit Function
    If InStr(d, "SMED") > 0 Then Exit Function
    If InStr(d, "POT") > 0 Then Exit Function

    If InStr(d, "PULLCORE") > 0 Or InStr(d, "PULL CORE") > 0 Then
        IsPullcoreBomInfo = True
        Exit Function
    End If

    If HasPullcoreLocationToken(d) Then
        If InStr(d, "CAM") > 0 Or InStr(d, "KEY") > 0 Then
            IsPullcoreBomInfo = True
            Exit Function
        End If
    End If

    If b.hasDims Then
        If IsPullcoreCamBomInfo(b) And LooksLikePullcoreCamBomDims(b) Then IsPullcoreBomInfo = True: Exit Function
        If IsPullcoreKeyBomInfo(b) And LooksLikePullcoreKeyBomDims(b) Then IsPullcoreBomInfo = True: Exit Function
    End If
End Function

Private Function IsPullcoreCamBomInfo(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)
    If InStr(d, "KEY") > 0 And InStr(d, "CAM") = 0 Then Exit Function
    If InStr(d, "CAM") > 0 Then IsPullcoreCamBomInfo = True: Exit Function
    If NormalizeSteelType(b.material) = "H13" Then IsPullcoreCamBomInfo = True
End Function

Private Function IsPullcoreKeyBomInfo(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)
    If InStr(d, "CAM") > 0 And InStr(d, "KEY") = 0 Then Exit Function
    If InStr(d, "KEY") > 0 Then IsPullcoreKeyBomInfo = True: Exit Function
    If NormalizeSteelType(b.material) = "A2" Then IsPullcoreKeyBomInfo = True
End Function

Private Function LooksLikePullcoreCamBomDims(ByRef b As BomInfo) As Boolean
    If b.hasDims = False Then Exit Function
    If b.BomLength < 3.4 Or b.BomLength > 4.6 Then Exit Function
    If b.BomWidth < 1.6 Or b.BomWidth > 3.2 Then Exit Function
    If b.BomThickness < 0.75 Or b.BomThickness > 1.8 Then Exit Function
    LooksLikePullcoreCamBomDims = True
End Function

Private Function LooksLikePullcoreKeyBomDims(ByRef b As BomInfo) As Boolean
    If b.hasDims = False Then Exit Function
    If b.BomLength < 2# Or b.BomLength > 3.6 Then Exit Function
    If b.BomWidth < 0.7 Or b.BomWidth > 1.8 Then Exit Function
    If b.BomThickness < 0.45 Or b.BomThickness > 1.2 Then Exit Function
    LooksLikePullcoreKeyBomDims = True
End Function

Private Sub AddPullcore(ByVal nm As String, ByVal q As Long, ByVal t As Double, ByVal w As Double, ByVal l As Double, ByVal mat As String)
    PcCount = PcCount + 1
    PcName(PcCount) = nm
    PcQty(PcCount) = IIf(q < 1, 1, q)
    PcT(PcCount) = t: PcW(PcCount) = w: PcL(PcCount) = l
    PcMat(PcCount) = mat
    PcVol(PcCount) = t * w * l * PcQty(PcCount)
End Sub

' Match a dimensionless BOM pullcore row to a CAD part (by pullcore name +
' matching location code) so it can be sized from CAD geometry.
Private Function FindPullcoreCadForBom(ByRef b As BomInfo, ByRef usedCad() As Boolean) As Long
    Dim i As Long, bloc As String
    bloc = GetPullcoreLocationCode(b.Description)
    For i = 1 To PartCount
        If Not usedCad(i) Then
            If IsPullcoreDesc(parts(i).componentName) Then
                If bloc = "" Or GetPullcoreLocationCode(parts(i).componentName) = bloc Then
                    FindPullcoreCadForBom = i: Exit Function
                End If
            End If
        End If
    Next i
End Function

' Build the pullcore/key list: names + sizing (BOM dims, else matched CAD bbox).
Private Function CadPurchasePartToken(ByVal componentName As String) As String
On Error GoTo ErrHandler

    Dim s As String
    Dim p As Long

    s = Trim(componentName)

    ' Use the leaf component name, not the parent assembly path.
    p = InStrRev(s, "/")
    If p > 0 Then s = Mid$(s, p + 1)

    p = InStrRev(s, Chr$(92))
    If p > 0 Then s = Mid$(s, p + 1)

    p = InStr(1, UCase$(s), ".STEP", vbTextCompare)
    If p > 1 Then s = Left$(s, p - 1)

    p = InStr(1, UCase$(s), ".SLDPRT", vbTextCompare)
    If p > 1 Then s = Left$(s, p - 1)

    p = InStr(1, UCase$(s), ".X_T", vbTextCompare)
    If p > 1 Then s = Left$(s, p - 1)

    CadPurchasePartToken = Trim(s)
    Exit Function

ErrHandler:
    CadPurchasePartToken = Trim(componentName)
End Function

Private Function LooksLikeFullBasePlateCad(ByVal idx As Long) As Boolean
    LooksLikeFullBasePlateCad = False
    If idx < 1 Or idx > PartCount Then Exit Function

    Dim baseFoot As Double
    Dim i As Long
    Dim fp As Double

    baseFoot = 0#
    For i = 1 To PartCount
        fp = parts(i).Width * parts(i).Length
        If fp > baseFoot Then baseFoot = fp
    Next i

    If baseFoot <= 0# Then Exit Function
    LooksLikeFullBasePlateCad = (parts(idx).Thickness >= STD_MIN_PLATE_THICKNESS And _
                                 parts(idx).Width * parts(idx).Length >= (1 - STD_FOOTPRINT_TOL) * baseFoot)
End Function

Private Function TryClassifyStandardCadPurchased(ByVal idx As Long, _
                                                 ByRef desc As String, _
                                                 ByRef vendor As String, _
                                                 ByRef partNo As String) As Boolean
    TryClassifyStandardCadPurchased = False
    desc = "": vendor = "": partNo = ""

    If idx < 1 Or idx > PartCount Then Exit Function
    If LooksLikeFullBasePlateCad(idx) Then Exit Function

    Dim raw As String
    Dim u As String
    raw = parts(idx).componentName
    u = UCase(raw)
    partNo = CadPurchasePartToken(raw)

    ' If the token is still just the parent assembly, do not use it as a part number.
    If InStr(UCase$(partNo), "MOLDBASE_ASM") > 0 Or InStr(UCase$(partNo), "MOLD_BASE_ASM") > 0 Then
        partNo = ""
    End If

    If IsRoundBarLike(idx) Then
        Dim rr As String
        rr = StandardRoundComponentRole(idx)
        Select Case NormalizeKey(rr)
            Case "LEADERPIN"
                desc = "Leader Pin"
                vendor = "PCS"
                TryClassifyStandardCadPurchased = True
                Exit Function
            Case "LEADERPINBUSHING"
                desc = "Leader Pin Bushing"
                vendor = "PCS"
                TryClassifyStandardCadPurchased = True
                Exit Function
            Case "GUIDEDEJECTORBUSHING"
                desc = "Guided Ejector Bushing"
                vendor = "PCS"
                TryClassifyStandardCadPurchased = True
                Exit Function
            Case "RETURNPIN", "EJECTORRETURNPIN"
                desc = "Return Pin"
                vendor = "PCS"
                TryClassifyStandardCadPurchased = True
                Exit Function
            Case "EJECTORPIN"
                desc = "Ejector Pin"
                vendor = "PCS"
                TryClassifyStandardCadPurchased = True
                Exit Function
            Case "SUPPORTPILLAR"
                desc = "Support Pillar"

                If InStr(UCase$(partNo), "_PCS") > 0 Or InStr(UCase$(raw), "_PCS") > 0 Then
                    vendor = "PCS"
                ElseIf InStr(UCase$(partNo), "_DME") > 0 Or InStr(UCase$(raw), "_DME") > 0 Then
                    vendor = "DME"
                Else
                    vendor = ""
                End If

                TryClassifyStandardCadPurchased = True
                Exit Function
        End Select
    End If

    If Left(UCase(partNo), 2) = "LR" Then
        desc = "Locating Ring"
        vendor = "PCS"
        TryClassifyStandardCadPurchased = True
        Exit Function
    End If

    If InStr(u, "PULL STRAP") > 0 Or InStr(u, "SAFETY STRAP") > 0 Then
        desc = "Safety Strap"
        vendor = "PCS"
        TryClassifyStandardCadPurchased = True
        Exit Function
    End If

    If InStr(u, "SHCS") > 0 Or InStr(u, "SOCKET HEAD") > 0 Then
        desc = "Socket Head Cap Screw"
        vendor = "PCS"
        TryClassifyStandardCadPurchased = True
        Exit Function
    End If

    If InStr(u, "SHSS") > 0 Or InStr(u, "SET SCREW") > 0 Then
        desc = "Socket Head Set Screw"
        vendor = "PCS"
        TryClassifyStandardCadPurchased = True
        Exit Function
    End If

    If InStr(u, "DOWEL") > 0 Then
        desc = "Dowel"
        vendor = "PCS"
        TryClassifyStandardCadPurchased = True
        Exit Function
    End If
End Function

Private Sub CaptureStandardPurchasedFromCadIfNeeded()
On Error GoTo ErrHandler

    If PartCount < 1 Then Exit Sub
    If PpCount > 0 Then Exit Sub

    Dim key(1 To 300) As String
    Dim desc(1 To 300) As String
    Dim vendor(1 To 300) As String
    Dim partNo(1 To 300) As String
    Dim qty(1 To 300) As Long
    Dim wid(1 To 300) As Double
    Dim lng(1 To 300) As Double
    Dim thk(1 To 300) As Double
    Dim n As Long

    Dim i As Long
    Dim d As String
    Dim v As String
    Dim pn As String
    Dim k As String
    Dim j As Long
    Dim hit As Long

    n = 0

    For i = 1 To PartCount
        If TryClassifyStandardCadPurchased(i, d, v, pn) Then
            ' Group by description + part number + SIZE so different-size parts
            ' with the same generic name (e.g. 1" ejector leader pins vs 1-1/2"
            ' main leader pins, both with no part number) stay separate quote
            ' lines instead of collapsing into one mispriced line.
            k = NormalizeKey(d & "|" & pn & "|" & _
                             Format(parts(i).Thickness, "0.000") & "x" & _
                             Format(parts(i).Width, "0.000") & "x" & _
                             Format(parts(i).Length, "0.000"))
            hit = 0
            For j = 1 To n
                If key(j) = k Then hit = j: Exit For
            Next j

            If hit = 0 Then
                n = n + 1
                If n > 300 Then Exit For
                key(n) = k
                desc(n) = d
                vendor(n) = v
                partNo(n) = pn
                qty(n) = parts(i).Quantity
                wid(n) = parts(i).Width
                lng(n) = parts(i).Length
                thk(n) = parts(i).Thickness
            Else
                qty(hit) = qty(hit) + parts(i).Quantity
            End If
        End If
    Next i

    For i = 1 To n
        CapturePurchased desc(i), qty(i), "", thk(i), wid(i), lng(i), partNo(i), vendor(i), "", "Purchase-CAD"
        LogLine "CAD standard purchased component: " & desc(i) & " " & partNo(i) & " qty " & qty(i)
    Next i

    If n > 0 Then LogLine "CAD standard purchased components captured: " & n
    Exit Sub

ErrHandler:
    LogLine "CaptureStandardPurchasedFromCadIfNeeded error: " & Err.Description
End Sub
Private Sub BuildPullcoreList()
    PcCount = 0
    ReDim PcName(1 To 80): ReDim PcQty(1 To 80): ReDim PcT(1 To 80)
    ReDim PcW(1 To 80): ReDim PcL(1 To 80): ReDim PcMat(1 To 80): ReDim PcVol(1 To 80)
    Dim usedCad() As Boolean
    If PartCount > 0 Then ReDim usedCad(1 To PartCount)
    Dim i As Long, t As Double, w As Double, l As Double, ci As Long, q As Long
    Dim baseNm As String, handled As Boolean
    ' 1) BOM pullcore rows
    For i = 1 To BomCount
        If IsPullcoreBomInfo(BomRows(i)) Then
            q = BomRows(i).Quantity
            baseNm = CleanPullcoreDisplayName(ProperCaseText(BomRows(i).Description))
            handled = False
            ' Required qty >= 2: the two parts are DIFFERENT. Match them to CAD
            ' and tell them apart by Y center - the higher Y is the ID pull core.
            If q >= 2 And PartCount > 0 Then
                Dim idxArr() As Long, n As Long, k As Long, lab As String
                n = CollectPullcoreCad(BomRows(i), usedCad, q, idxArr)
                If n >= 2 Then
                    SortIdxByYDesc idxArr, n
                    For k = 1 To n
                        If k = 1 Then
                            lab = "ID " & baseNm
                        ElseIf k = 2 Then
                            lab = "OD " & baseNm
                        Else
                            lab = "#" & k & " " & baseNm
                        End If
                        If BomRows(i).hasDims Then
                            AddPullcore lab, 1, BomRows(i).BomThickness, BomRows(i).BomWidth, BomRows(i).BomLength, BomRows(i).material
                        Else
                            AddPullcore lab, 1, parts(idxArr(k)).Thickness, parts(idxArr(k)).Width, parts(idxArr(k)).Length, BomRows(i).material
                        End If
                        usedCad(idxArr(k)) = True
                        LogLine "Pullcore split " & lab & " (CenterY=" & FormatNumberForCsv(parts(idxArr(k)).AsmCenterY) & ")"
                    Next k
                    handled = True
                ElseIf BomRows(i).hasDims Then
                    For k = 1 To q
                        If k = 1 Then
                            lab = "ID " & baseNm
                        ElseIf k = 2 Then
                            lab = "OD " & baseNm
                        Else
                            lab = "#" & k & " " & baseNm
                        End If
                        AddPullcore lab, 1, BomRows(i).BomThickness, BomRows(i).BomWidth, BomRows(i).BomLength, BomRows(i).material
                        LogLine "Pullcore split from BOM dims " & lab
                    Next k
                    handled = True
                End If
            End If
            If Not handled Then
                t = 0: w = 0: l = 0
                If BomRows(i).hasDims Then
                    t = BomRows(i).BomThickness: w = BomRows(i).BomWidth: l = BomRows(i).BomLength
                ElseIf PartCount > 0 Then
                    ci = FindPullcoreCadForBom(BomRows(i), usedCad)
                    If ci > 0 Then
                        t = parts(ci).Thickness: w = parts(ci).Width: l = parts(ci).Length
                        usedCad(ci) = True
                    End If
                End If
                If t > 0 And w > 0 And l > 0 Then
                    AddPullcore baseNm, q, t, w, l, BomRows(i).material
                Else
                    LogLine "Pullcore with no size (skipped from quote): " & BomRows(i).Description
                End If
            End If
        End If
    Next i
    ' 2) No BOM pullcores -> scan CAD component names
    If PcCount = 0 And PartCount > 0 Then
        For i = 1 To PartCount
            If IsPullcoreDesc(parts(i).componentName) Then
                AddPullcore CleanPullcoreDisplayName(parts(i).cleanName), parts(i).Quantity, _
                            parts(i).Thickness, parts(i).Width, parts(i).Length, ""
            End If
        Next i
    End If
End Sub

' Collect up to maxN unused CAD parts that look like this BOM pull-core row.
Private Function CollectPullcoreCad(ByRef b As BomInfo, ByRef usedCad() As Boolean, ByVal maxN As Long, ByRef outIdx() As Long) As Long
    Dim bloc As String, n As Long, i As Long
    bloc = GetPullcoreLocationCode(b.Description)
    ReDim outIdx(1 To IIf(maxN < 1, 1, maxN))
    n = 0
    For i = 1 To PartCount
        If n >= maxN Then Exit For
        If Not usedCad(i) Then
            If IsPullcoreDesc(parts(i).componentName) Then
                If bloc = "" Or GetPullcoreLocationCode(parts(i).componentName) = bloc Then
                    n = n + 1: outIdx(n) = i
                End If
            End If
        End If
    Next i
    CollectPullcoreCad = n
End Function

Private Sub SortIdxByYDesc(ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If parts(idx(j)).AsmCenterY > parts(idx(i)).AsmCenterY Then
                t = idx(i): idx(i) = idx(j): idx(j) = t
            End If
        Next j
    Next i
End Sub

Private Sub WritePullcorePriceFile()
On Error GoTo eh
    Dim csv As String, i As Long, totVol As Double, totPrice As Double, price As Double
    csv = "Pull Core / Key,Qty,Thickness,Width,Length,Material,Cu In,Price USD" & vbCrLf
    For i = 1 To PcCount
        price = PcVol(i) * PULLCORE_RATE
        totVol = totVol + PcVol(i): totPrice = totPrice + price
        csv = csv & CsvText(PcName(i)) & "," & PcQty(i) & "," & FormatNumberForCsv(PcT(i)) & "," & _
              FormatNumberForCsv(PcW(i)) & "," & FormatNumberForCsv(PcL(i)) & "," & CsvText(PcMat(i)) & "," & _
              FormatNumberForCsv(PcVol(i)) & "," & FormatNumberForCsv(price) & vbCrLf
    Next i
    csv = csv & "TOTAL,,,,,," & FormatNumberForCsv(totVol) & "," & FormatNumberForCsv(totPrice) & vbCrLf
    csv = csv & "RATE ($/in3),,,,,,," & FormatNumberForCsv(PULLCORE_RATE) & vbCrLf
    Dim p As String, f As Integer
    p = GetWritableCsvPath(CurrentJobFolder & "\" & PULLCORE_PRICE_FILE)
    f = FreeFile
    Open p For Output As #f
    Print #f, csv
    Close #f
    LogLine "Pullcore prices file: " & p & "  (total $" & FormatNumberForCsv(totPrice) & ")"
    Exit Sub
eh:
    LogLine "WritePullcorePriceFile error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

' Add a professional PULL CORES & KEYS category to the Quote sheet (in the
' empty space below the existing sections). Prices = volume x rate.
Private Sub WritePullcoreCategoryToSheet(ByVal ws As Object)
On Error GoTo eh
    If PcCount < 1 Or ws Is Nothing Then Exit Sub
    Dim r As Long, hr As Long, rr As Long, i As Long
    r = PULLCORE_QUOTE_START_ROW

    ' Title bar (merged A:H)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Merge
    ws.Cells(r, 1).value = "PULL CORES & KEYS  (volume x $" & PULLCORE_RATE & " / in" & Chr(179) & ")"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).HorizontalAlignment = -4108
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(31, 78, 121)

    ' Header row
    hr = r + 1
    ws.Cells(hr, 1).value = "Description"
    ws.Cells(hr, 3).value = "QTY"
    ws.Cells(hr, 4).value = "Thickness"
    ws.Cells(hr, 5).value = "Width"
    ws.Cells(hr, 6).value = "Length"
    ws.Cells(hr, 7).value = "Cu. In."
    ws.Cells(hr, 8).value = "Price"
    With ws.Range(ws.Cells(hr, 1), ws.Cells(hr, 8))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = -4108
    End With

    ' Item rows
    rr = hr + 1
    For i = 1 To PcCount
        ws.Cells(rr, 1).value = PcName(i)
        ws.Cells(rr, 3).value = PcQty(i)
        ws.Cells(rr, 4).value = PcT(i)
        ws.Cells(rr, 5).value = PcW(i)
        ws.Cells(rr, 6).value = PcL(i)
        ws.Cells(rr, 7).Formula = "=C" & rr & "*D" & rr & "*E" & rr & "*F" & rr
        ws.Cells(rr, 8).Formula = "=G" & rr & "*" & Replace(CStr(PULLCORE_RATE), ",", ".")
        ws.Cells(rr, 4).NumberFormat = "0.000"
        ws.Cells(rr, 5).NumberFormat = "0.000"
        ws.Cells(rr, 6).NumberFormat = "0.000"
        ws.Cells(rr, 7).NumberFormat = "0.00"
        ws.Cells(rr, 8).NumberFormat = "$#,##0.00"
        rr = rr + 1
    Next i

    ' Total row
    ws.Cells(rr, 1).value = "Total"
    ws.Cells(rr, 1).Font.Bold = True
    ws.Cells(rr, 7).Formula = "=SUM(G" & (hr + 1) & ":G" & (rr - 1) & ")"
    ws.Cells(rr, 8).Formula = "=SUM(H" & (hr + 1) & ":H" & (rr - 1) & ")"
    ws.Cells(rr, 7).NumberFormat = "0.00"
    ws.Cells(rr, 8).NumberFormat = "$#,##0.00"
    ws.Cells(rr, 7).Font.Bold = True
    ws.Cells(rr, 8).Font.Bold = True

    ' Outline the table
    With ws.Range(ws.Cells(hr, 1), ws.Cells(rr, 8)).Borders
        .LineStyle = 1
        .Weight = 2
    End With

    LogLine "Pullcore category written to Quote sheet starting row " & PULLCORE_QUOTE_START_ROW
    Exit Sub
eh:
    LogLine "WritePullcoreCategoryToSheet error: " & Err.Description
End Sub

' Put the pull-core dollar total on the QuoteWorksheet summary line C196/D196.
' Row 196 sits inside the grand-total range SUM(C194:C210), so it rolls into the
' Total Price and the 6% commission automatically. The detailed per-core block
' (WritePullcoreCategoryToSheet) is still written lower down for reference.
Private Sub WritePullcoreTotalToSummary(ByVal ws As Object)
On Error GoTo eh
    If PcCount < 1 Or ws Is Nothing Then Exit Sub
    Dim i As Long, totVol As Double
    For i = 1 To PcCount: totVol = totVol + PcVol(i): Next i
    Dim dollars As Double
    dollars = totVol * PULLCORE_RATE
    ws.Cells(196, 1).value = "Pull Cores"   ' A196 label
    ws.Cells(196, 3).value = dollars         ' C196 (rough col)  -> in SUM(C194:C210)
    ws.Cells(196, 4).value = dollars         ' D196 (finish col) -> in SUM(D194:D210)
    LogLine "Quote summary: Pull Cores $" & FormatNumberForCsv(dollars) & " -> C196/D196 (in grand total)"
    Exit Sub
eh:
    LogLine "WritePullcoreTotalToSummary error: " & Err.Description
End Sub

' Place the purchased hardware into the built-in "Components" area (cols K:S).
' Each row: M=QTY, R=Unit Price (S=M*R auto), Q=Part No. S48 sums S4:S47 and feeds
' the Components line C195, which is already inside the grand total. Items map to
' their labeled category group (4 rows each); anything unmatched spills to a free row
' so its cost is still counted.
Private Sub WritePurchasedToComponentsArea(ByVal ws As Object)
On Error GoTo eh
    If PpCount < 1 Or ws Is Nothing Then Exit Sub
    ' List each purchased part on its own row with the real BOM name. The
    ' pre-printed category labels (Leader Pins, L P Bushings, ...) are replaced
    ' with the actual component names. Columns: K=Name L=STD M=QTY N=Dia O=Length
    ' P=Width Q=Part No. R=Price (S=Total is a template formula M*R).
    Const FIRST_ROW As Long = 4
    Const LAST_ROW As Long = 47
    Dim i As Long, target As Long, placed As Long, nm As String
    target = FIRST_ROW
    For i = 1 To PpCount
        If target > LAST_ROW Then Exit For
        nm = Trim(PpDesc(i))
        If nm = "" Then nm = PpComp(i)
        ws.Cells(target, 11).value = nm                                 ' K = Components (name)
        ws.Cells(target, 13).value = PpQty(i)                           ' M = QTY (No. Req'd)
        If PpW(i) > 0 Then ws.Cells(target, 14).value = PpW(i)          ' N = Dia. (BOM O.D./width)
        If PpL(i) > 0 Then ws.Cells(target, 15).value = PpL(i)          ' O = Length
        If PpT(i) > 0 Then ws.Cells(target, 16).value = PpT(i)          ' P = Width (thickness/height)
        If PpPartNo(i) <> "" Then ws.Cells(target, 17).value = PpPartNo(i)   ' Q = Part No.
        ws.Cells(target, 18).value = PpPrice(i)                         ' R = Unit Price
        LogLine "Components row " & target & " <- " & nm & " qty " & PpQty(i) & _
                " part " & PpPartNo(i) & " @ $" & FormatNumberForCsv(PpPrice(i))
        target = target + 1
        placed = placed + 1
    Next i
    LogLine "Purchased components written: " & placed & " of " & PpCount
    Exit Sub
eh:
    LogLine "WritePurchasedToComponentsArea error: " & Err.Description
End Sub

' Map a component description to the start row of its Components-area category group.
' Returns 0 if no category matches.
Private Function PurchasedComponentRow(ByVal s As String) As Long
    Dim u As String
    u = UCase(s)
    PurchasedComponentRow = 0
    If InStr(u, "RETURN") > 0 Then PurchasedComponentRow = 12: Exit Function
    If InStr(u, "SPRUE") > 0 Then PurchasedComponentRow = 16: Exit Function
    If InStr(u, "LOCAT") > 0 Then PurchasedComponentRow = 20: Exit Function
    If InStr(u, "PILLAR") > 0 Then PurchasedComponentRow = 24: Exit Function
    If (InStr(u, "GUIDED EJEC") > 0 Or InStr(u, "GUIDE EJEC") > 0) And InStr(u, "BUSH") > 0 Then PurchasedComponentRow = 32: Exit Function
    If InStr(u, "GUIDED EJEC") > 0 Or InStr(u, "EJECTOR PIN") > 0 Then PurchasedComponentRow = 28: Exit Function
    If InStr(u, "SIDE LOCK") > 0 Then PurchasedComponentRow = 36: Exit Function
    If InStr(u, "SLIDE RETAIN") > 0 Or InStr(u, "RETAINER") > 0 Or InStr(u, "RETAIN") > 0 Or InStr(u, "RING") > 0 Or InStr(u, "CIRCLIP") > 0 Then PurchasedComponentRow = 40: Exit Function
    If InStr(u, "ANGLE PIN") > 0 Then PurchasedComponentRow = 44: Exit Function
    If InStr(u, "BUSH") > 0 Then PurchasedComponentRow = 8: Exit Function     ' L P / guide bushings
    If InStr(u, "LEADER") > 0 Then PurchasedComponentRow = 4: Exit Function   ' leader pins
End Function

Private Sub ComputePullcoreQuote()
    If PcCount < 1 Then BuildPullcoreList    ' normally already built before the fills
    If PcCount < 1 Then LogLine "Pullcore/key quote: no pull cores or keys found.": Exit Sub
    Dim i As Long, totVol As Double
    For i = 1 To PcCount: totVol = totVol + PcVol(i): Next i
    LogLine "PULLCORE/KEY QUOTE: " & PcCount & " item(s), " & FormatNumberForCsv(totVol) & _
            " cuin x $" & PULLCORE_RATE & " = $" & FormatNumberForCsv(totVol * PULLCORE_RATE)
    WritePullcorePriceFile
End Sub

' ============================================================
' NETWORK-AWARE PUBLISH
' Private local folder by default; public company share when on
' the company Netgear Wi-Fi. Publishes the job signature CSV (so
' Elgin's matcher imports it), plus the filled sheets, dimension
' CSVs and ISO images, so all the tools share one location.
' ============================================================

Private Function GetWifiSsidVba() As String
    On Error Resume Next
    Dim sh As Object, ex As Object, out As String
    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec("netsh wlan show interfaces")
    If ex Is Nothing Then Exit Function
    out = ex.StdOut.ReadAll
    Dim arr() As String, i As Long, s As String, p As Long
    arr = Split(out, vbCrLf)
    For i = LBound(arr) To UBound(arr)
        s = Trim(arr(i))
        ' Match a line that STARTS with "SSID" (so "BSSID" is ignored).
        If InStr(1, s, "SSID", vbTextCompare) = 1 Then
            p = InStr(s, ":")
            If p > 0 Then
                GetWifiSsidVba = Trim(Mid(s, p + 1))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsOnCompanyWifi() As Boolean
    If FORCE_LOCAL_PUBLISH Then Exit Function
    Dim ssid As String
    ssid = UCase$(GetWifiSsidVba())
    If ssid <> "" Then
        If InStr(ssid, UCase$(COMPANY_WIFI_SSID)) > 0 Then IsOnCompanyWifi = True
        Exit Function   ' SSID known but not the company one -> not company (no share probe)
    End If
    ' SSID unknown (e.g. wired): use share reachability as the signal.
    If PUBLIC_DATA_ROOT <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        If fso.FolderExists(PUBLIC_DATA_ROOT) Then IsOnCompanyWifi = True
    End If
End Function

Private Function ResolveMatchingRoot() As String
    Dim root As String
    If IsOnCompanyWifi() Then root = PUBLIC_DATA_ROOT Else root = PRIVATE_DATA_ROOT
    If root = "" Then root = PRIVATE_DATA_ROOT
    EnsureFolderDeep root
    ResolveMatchingRoot = root
End Function

Private Sub PublishJobOutputs()
    On Error GoTo eh
    If Not PUBLISH_OUTPUTS Then Exit Sub
    Dim onCo As Boolean
    onCo = IsOnCompanyWifi()
    Dim root As String
    root = ResolveMatchingRoot()
    If root = "" Then LogLine "Publish: no matching root resolved.": Exit Sub
    LogLine "Publish destination (" & IIf(onCo, "PUBLIC company share", "PRIVATE local folder") & "): " & root

    ' 1) Job signature CSV at the matching root (flat) so Elgin auto-imports it.
    Dim sigPath As String
    sigPath = root & "\XT_Export_Job_Signature_" & CleanFileName(CurrentJobNumber) & ".csv"
    WriteJobSignatureCsv sigPath

    ' 2) Copy this job's deliverables into root\<job>\ for Elgin's image/sheet finders.
    Dim jobOut As String
    jobOut = root & "\" & CleanFileName(CurrentJobNumber)
    EnsureFolderDeep jobOut
    CopyMatchingArtifacts CurrentJobFolder, jobOut
    LogLine "Published job outputs to: " & jobOut
    Exit Sub
eh:
    LogLine "PublishJobOutputs error: " & Err.Description
End Sub

Private Sub CopyMatchingArtifacts(ByVal srcFolder As String, ByVal dstFolder As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcFolder) Then Exit Sub
    Dim f As Object, nm As String, up As String, ext As String, take As Boolean
    For Each f In fso.GetFolder(srcFolder).Files
        nm = f.Name
        up = UCase$(nm)
        ext = UCase$(GetFileExtension(nm))
        take = False
        If ext = "JPG" Or ext = "JPEG" Or ext = "PNG" Then take = True
        If ext = "CSV" Then take = True
        If ext = "TXT" Then take = True
        If (ext = "XLS" Or ext = "XLSX" Or ext = "XLSM") And _
           (InStr(up, "QUOTE") > 0 Or InStr(up, "STEEL") > 0 Or InStr(up, "J000") > 0) Then take = True
        If take Then fso.CopyFile f.path, dstFolder & "\" & nm, True
    Next f
End Sub

' Write the 6-component pot/holder signature CSV in the header format Elgin reads.
Private Sub WriteJobSignatureCsv(ByVal destPath As String)
    On Error GoTo eh
    Dim haveAny As Boolean
    haveAny = (gIdxTCP > 0 Or gIdxBCP > 0 Or gIdxIDH > 0 Or gIdxODH > 0 Or gIdxIDP > 0 Or gIdxODP > 0)
    If Not haveAny Then LogLine "Signature CSV skipped (no pot/holder components identified).": Exit Sub
    Dim s As String
    s = "JobNumber,ComponentRole,QuoteName,CadComponent,CleanName,Length,Width,Thickness,Mass,CenterX,CenterY,CenterZ,HasCenter" & vbCrLf
    s = s & SigRow("TCP", gIdxTCP)
    s = s & SigRow("BCP", gIdxBCP)
    s = s & SigRow("ID HOLDER", gIdxIDH)
    s = s & SigRow("OD HOLDER", gIdxODH)
    s = s & SigRow("ID POT", gIdxIDP)
    s = s & SigRow("OD POT", gIdxODP)
    Dim f As Integer
    f = FreeFile
    Open destPath For Output As #f
    Print #f, s;
    Close #f
    LogLine "Wrote job signature CSV: " & destPath
    Exit Sub
eh:
    LogLine "WriteJobSignatureCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Function SigRow(ByVal role As String, ByVal idx As Long) As String
    If idx < 1 Or idx > PartCount Then Exit Function
    Dim hc As String
    hc = IIf(parts(idx).hasAsmCenter, "TRUE", "FALSE")
    SigRow = CsvText(CurrentJobNumber) & "," & CsvText(role) & "," & CsvText(role) & "," & _
             CsvText(parts(idx).componentName) & "," & CsvText(parts(idx).cleanName) & "," & _
             FormatNumberForCsv(parts(idx).Length) & "," & FormatNumberForCsv(parts(idx).Width) & "," & _
             FormatNumberForCsv(parts(idx).Thickness) & "," & FormatNumberForCsv(parts(idx).massValue) & "," & _
             FormatNumberForCsv(parts(idx).AsmCenterX) & "," & FormatNumberForCsv(parts(idx).AsmCenterY) & "," & _
             FormatNumberForCsv(parts(idx).AsmCenterZ) & "," & hc & vbCrLf
End Function

' Move stray native CAD parts created while saving the base into \base,
' and any PDFs into \pdf, so the job folder stays tidy.
' Move loose .sldprt part files from the job folder into the base subfolder.
' Run after the self-contained exports (.easm/.stl/.x_t/.igs) and after closing
' docs. NOTE: this can break the live base .sldasm's part references; the matching
' workflow uses the self-contained files, so that is acceptable.
Private Sub MoveLooseSolidWorksPartsToBaseFolder()
On Error GoTo eh
    If CurrentJobFolder = "" Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CurrentJobFolder) Then Exit Sub
    Dim baseDir As String
    baseDir = CurrentJobFolder & "\base"
    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir

    Dim paths() As String, n As Long, f As Object
    n = 0
    ReDim paths(0 To 0)
    For Each f In fso.GetFolder(CurrentJobFolder).Files
        If Left(f.Name, 2) <> "~$" Then
            If LCase(fso.GetExtensionName(f.path)) = "sldprt" Then
                ReDim Preserve paths(0 To n)
                paths(n) = f.path
                n = n + 1
            End If
        End If
    Next f

    Dim i As Long, dest As String
    For i = 0 To n - 1
        dest = GetUniqueFilePath(baseDir & "\" & fso.GetFileName(paths(i)))
        On Error Resume Next
        fso.MoveFile paths(i), dest
        If Err.Number = 0 Then
            LogLine "Moved part to base: " & fso.GetFileName(dest)
        Else
            Err.Clear
            fso.CopyFile paths(i), dest, True       ' locked? copy then delete
            If Err.Number = 0 Then
                fso.DeleteFile paths(i), True
                LogLine "Copied part to base: " & fso.GetFileName(dest)
            Else
                LogLine "WARNING could not move part: " & paths(i)
                Err.Clear
            End If
        End If
        On Error GoTo eh
    Next i
    LogLine "Loose .sldprt parts moved to base folder: " & n
    Exit Sub
eh:
    LogLine "MoveLooseSolidWorksPartsToBaseFolder error: " & Err.Description
End Sub

Private Sub OrganizeJobFiles()
    On Error GoTo eh
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CurrentJobFolder) Then Exit Sub
    Dim baseDir As String, pdfDir As String
    baseDir = CurrentJobFolder & "\base"
    pdfDir = CurrentJobFolder & "\pdf"
    EnsureFolderDeep baseDir
    EnsureFolderDeep pdfDir
    Dim names As Collection, f As Object
    Set names = New Collection
    For Each f In fso.GetFolder(CurrentJobFolder).Files
        names.Add f.Name
    Next f
    Dim idx As Long, nm As String, ext As String, dest As String, src As String, target As String
    For idx = 1 To names.Count

        nm = names(idx)

        If Left(nm, 2) = "~$" Then GoTo NextOrganizeFile

        ext = UCase$(GetFileExtension(nm))
        dest = ""

        If ext = "SLDPRT" Or ext = "SLDASM" Then
            dest = baseDir
        ElseIf ext = "PDF" Then
            dest = pdfDir
        End If

        If dest <> "" Then
            src = CurrentJobFolder & "\" & nm
            target = dest & "\" & nm
            On Error Resume Next
            If fso.FileExists(target) Then target = GetUniqueFilePath(target)
            fso.MoveFile src, target
            On Error GoTo eh
            LogLine "Organized: " & nm & " -> " & dest
        End If

NextOrganizeFile:
    Next idx
    Exit Sub
eh:
    LogLine "OrganizeJobFiles error: " & Err.Description
End Sub




' ============================================================
' PCS NAMING ANALYSIS  (geometry-first, names are only hints)
' ============================================================
Private Sub WritePcsNamingAnalysis(ByVal destPath As String, ByVal isStandardBase As Boolean)
On Error GoTo eh
    If PartCount < 1 Then Exit Sub

    Dim f As Integer
    f = FreeFile
    Open destPath For Output As #f
    Print #f, "Index,CurrentName,GeometryRole,SuggestedPCSName,Confidence,Reason,DMEStackFamily,PartingSide,PartingLineAxis,PartingLinePos,LeaderPinStackKey,StackAxisPos,Thickness,Width,Length,CenterX,CenterY,CenterZ"

    Dim i As Long
    Dim role As String
    Dim suggested As String
    Dim confidence As String
    Dim reason As String

    For i = 1 To PartCount
        role = GeometryRoleForCadIndex(i, isStandardBase)
        suggested = SuggestedPcsNameForCadIndex(i, role, confidence, reason)
        Print #f, i & "," & CsvText(parts(i).componentName) & "," & CsvText(role) & "," & _
                  CsvText(suggested) & "," & CsvText(confidence) & "," & CsvText(reason) & "," & _
                  CsvText(gStdDmeStackFamily) & "," & CsvText(StdPartingSideForCadIndex(i)) & "," & _
                  CStr(gStdPartingLineAxis) & "," & FormatNumberForCsv(gStdPartingLinePos) & "," & _
                  CsvText(LeaderPinStackKeyForCadIndex(i)) & "," & FormatNumberForCsv(StandardStackAxisPos(i)) & "," & _
                  FormatNumberForCsv(parts(i).Thickness) & "," & FormatNumberForCsv(parts(i).Width) & "," & _
                  FormatNumberForCsv(parts(i).Length) & "," & FormatNumberForCsv(parts(i).AsmCenterX) & "," & _
                  FormatNumberForCsv(parts(i).AsmCenterY) & "," & FormatNumberForCsv(parts(i).AsmCenterZ)
    Next i

    Close #f
    LogLine "Wrote PCS naming analysis: " & destPath
    Exit Sub
eh:
    LogLine "WritePcsNamingAnalysis error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

' Qwen-parity stack + leader-pin position dump (always-on for standard bases).
' Mirrors job_analysis + per-part roles/positions from the AI classifier output.
Private Sub WriteStackLeaderPinAnalysis(ByVal destPath As String, ByVal isStandardBase As Boolean)
On Error GoTo eh
    If PartCount < 1 Then Exit Sub
    If Not isStandardBase Then Exit Sub

    Dim f As Integer
    Dim i As Long
    Dim role As String
    Dim pinSet As String
    Dim pinDir As String
    Dim rules() As String
    Dim r As Long

    f = FreeFile
    Open destPath For Output As #f

    ' --- job_analysis header block (one row of metadata, then blank, then parts) ---
    Print #f, "SECTION,KEY,VALUE"
    Print #f, "job_analysis,stack_axis," & CsvText(StdStackAxisName(gStdStackAxis))
    Print #f, "job_analysis,top_is_first," & CStr(gStdTopIsFirst)
    Print #f, "job_analysis,dme_stack_family," & CsvText(gStdDmeStackFamily)
    Print #f, "job_analysis,parting_line," & CsvText(gStdPartingLineText)
    Print #f, "job_analysis,parting_line_axis," & CStr(gStdPartingLineAxis)
    Print #f, "job_analysis,parting_line_pos," & FormatNumberForCsv(gStdPartingLinePos)
    Print #f, "job_analysis,a_plate_idx," & CStr(gStdCavityCadIndex)
    Print #f, "job_analysis,b_plate_idx," & CStr(gStdCoreCadIndex)
    Print #f, "job_analysis,sequenced_latch_lock_base," & CStr(gStdSequencedLatchLock)
    If gStdLeaderPinFromKnown Then
        pinDir = IIf(gStdLeaderPinFromTop, "FROM_TOP_A", "FROM_BOTTOM_B")
        If gStdLeaderPinReversed Then pinDir = pinDir & "_REVERSED"
    Else
        pinDir = "UNKNOWN"
    End If
    Print #f, "job_analysis,leader_pin_direction," & CsvText(pinDir)
    Print #f, "job_analysis,leader_pin_reversed," & CStr(gStdLeaderPinReversed)
    Print #f, "job_analysis,ai_bridge_used," & CStr(gAiBridgeUsed)

    If gStdStackRules <> "" Then
        rules = Split(gStdStackRules, "|")
        For r = LBound(rules) To UBound(rules)
            If Trim(rules(r)) <> "" Then
                Print #f, "job_analysis,rule_" & (r - LBound(rules) + 1) & "," & CsvText(Trim(rules(r)))
            End If
        Next r
    End If

    Print #f, ""
    Print #f, "Index,CurrentName,GeometryRole,LeaderPinSet,PartingSide,LeaderPinStackKey,StackAxisPos,Thickness,Width,Length,CenterX,CenterY,CenterZ,ConfidenceHint"

    For i = 1 To PartCount
        role = GeometryRoleForCadIndex(i, True)
        pinSet = ""
        On Error Resume Next
        pinSet = gStdLeaderPinSetByPart(i)
        On Error GoTo eh
        ' Only emit structural / guide / latch rows to keep the file readable.
        If role = "" Then GoTo nextPart
        If NormalizeKey(role) = "HARDWAREOTHER" Or NormalizeKey(role) = "HARDWARE/OTHER" Then
            If pinSet = "" And Not IsLatchLockName(parts(i).componentName) Then GoTo nextPart
        End If
        If NormalizeKey(role) = "IGNORE" Then GoTo nextPart

        Print #f, i & "," & CsvText(parts(i).componentName) & "," & CsvText(role) & "," & _
                  CsvText(pinSet) & "," & CsvText(StdPartingSideForCadIndex(i)) & "," & _
                  CsvText(LeaderPinStackKeyForCadIndex(i)) & "," & FormatNumberForCsv(StandardStackAxisPos(i)) & "," & _
                  FormatNumberForCsv(parts(i).Thickness) & "," & FormatNumberForCsv(parts(i).Width) & "," & _
                  FormatNumberForCsv(parts(i).Length) & "," & FormatNumberForCsv(parts(i).AsmCenterX) & "," & _
                  FormatNumberForCsv(parts(i).AsmCenterY) & "," & FormatNumberForCsv(parts(i).AsmCenterZ) & "," & _
                  CsvText(IIf(pinSet = "PRIMARY", "HIGH primary pin-bushing plane", _
                         IIf(pinSet = "SECONDARY", "MEDIUM secondary guided-ejector plane", "")))
nextPart:
    Next i

    Close #f
    LogLine "Wrote stack/leader-pin analysis: " & destPath & _
            " | axis=" & StdStackAxisName(gStdStackAxis) & _
            " | pins=" & pinDir & _
            " | latchLock=" & CStr(gStdSequencedLatchLock)
    Exit Sub
eh:
    LogLine "WriteStackLeaderPinAnalysis error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Function StandardStackAxisPos(ByVal idx As Long) As Double
    If idx < 1 Or idx > PartCount Then Exit Function
    If gStdStackAxis < 1 Or gStdStackAxis > 3 Then
        StandardStackAxisPos = parts(idx).AsmCenterZ
    Else
        StandardStackAxisPos = PartAxisCenter(idx, gStdStackAxis)
    End If
End Function

Private Function LeaderPinStackKeyForCadIndex(ByVal idx As Long) As String
    If idx < 1 Or idx > PartCount Then Exit Function
    If IsRoundBarLike(idx) = False Then Exit Function

    Dim role As String
    role = NormalizeKey(StandardRoundComponentRole(idx))
    If role <> "LEADERPIN" And role <> "LEADERPINBUSHING" And role <> "GUIDEDEJECTORBUSHING" Then Exit Function

    Dim a1 As Double
    Dim a2 As Double
    Select Case gStdStackAxis
        Case 1
            a1 = parts(idx).AsmCenterY: a2 = parts(idx).AsmCenterZ
        Case 2
            a1 = parts(idx).AsmCenterX: a2 = parts(idx).AsmCenterZ
        Case Else
            a1 = parts(idx).AsmCenterX: a2 = parts(idx).AsmCenterY
    End Select

    LeaderPinStackKeyForCadIndex = "LPSTACK_" & Format$(Round(a1, 2), "0.00") & "_" & Format$(Round(a2, 2), "0.00")
End Function

Private Function GeometryRoleForCadIndex(ByVal idx As Long, ByVal isStandardBase As Boolean) As String
    If idx < 1 Or idx > PartCount Then Exit Function

    If isStandardBase Then
        Dim stdRole As String
        stdRole = StdCadRole(idx)
        If stdRole <> "" Then GeometryRoleForCadIndex = stdRole: Exit Function
    End If

    If idx = gIdxTCP Then GeometryRoleForCadIndex = "TCP": Exit Function
    If idx = gIdxBCP Then GeometryRoleForCadIndex = "BCP": Exit Function
    If idx = gIdxIDH Then GeometryRoleForCadIndex = "ID HOLDER": Exit Function
    If idx = gIdxODH Then GeometryRoleForCadIndex = "OD HOLDER": Exit Function
    If idx = gIdxIDP Then GeometryRoleForCadIndex = "ID POT": Exit Function
    If idx = gIdxODP Then GeometryRoleForCadIndex = "OD POT": Exit Function

    If IsRoundBarLike(idx) Then
        GeometryRoleForCadIndex = StandardRoundComponentRole(idx)
        Exit Function
    End If

    If IsLatchLockName(parts(idx).componentName) Then
        GeometryRoleForCadIndex = "Latch Lock / Safety Strap"
        Exit Function
    End If

    If LooksLikePlate(idx) Then
        If isStandardBase Then
            GeometryRoleForCadIndex = "STANDARD PLATE / RAIL"
        Else
            GeometryRoleForCadIndex = "PLATE-LIKE COMPONENT"
        End If
        Exit Function
    End If

    GeometryRoleForCadIndex = "HARDWARE / OTHER"
End Function

Private Function IsScrewFastenerName(ByVal raw As String) As Boolean
    Dim u As String
    u = UCase$(raw)

    IsScrewFastenerName = False

    If InStr(u, "SHCS") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "FHCS") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "BHCS") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "HHCS") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "SBHCS") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "UNC-") > 0 And InStr(u, "X-") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "SCREW") > 0 Then IsScrewFastenerName = True: Exit Function
    If InStr(u, "BOLT") > 0 Then IsScrewFastenerName = True: Exit Function
End Function

Private Function IsDowelOrMinorRoundHardwareName(ByVal raw As String) As Boolean
    Dim u As String
    u = UCase$(raw)

    IsDowelOrMinorRoundHardwareName = False

    If InStr(u, "DOWEL") > 0 Then IsDowelOrMinorRoundHardwareName = True: Exit Function
    If InStr(u, "TUBE_DOWEL") > 0 Then IsDowelOrMinorRoundHardwareName = True: Exit Function
    If InStr(u, "TLP") > 0 Then IsDowelOrMinorRoundHardwareName = True: Exit Function
    If InStr(u, "PLUG") > 0 Then IsDowelOrMinorRoundHardwareName = True: Exit Function
End Function

Private Function IsExplicitGuideHardwareName(ByVal raw As String) As Boolean
    Dim u As String
    u = UCase$(raw)

    IsExplicitGuideHardwareName = False

    If InStr(u, "LDR-PIN") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "LDR_PIN") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "LEADER PIN") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "GUIDE PIN") > 0 Then IsExplicitGuideHardwareName = True: Exit Function

    If InStr(u, "LBB_") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "/LBB") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "BUSHING") > 0 And InStr(u, "EJECTOR") = 0 Then IsExplicitGuideHardwareName = True: Exit Function

    If InStr(u, "GEB_") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "GUIDED EJECTOR") > 0 Then IsExplicitGuideHardwareName = True: Exit Function

    If InStr(u, "RETURN-PIN") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "RETURN PIN") > 0 Then IsExplicitGuideHardwareName = True: Exit Function

    If InStr(u, "PILLAR_D") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
    If InStr(u, "SUPPORT PILLAR") > 0 Then IsExplicitGuideHardwareName = True: Exit Function
End Function

Private Function IsGuideRoundCandidate(ByVal idx As Long) As Boolean
    IsGuideRoundCandidate = False

    If idx < 1 Or idx > PartCount Then Exit Function
    If Not IsRoundBarLike(idx) Then Exit Function

    Dim nm As String
    nm = parts(idx).componentName

    If IsScrewFastenerName(nm) Then Exit Function
    If IsDowelOrMinorRoundHardwareName(nm) Then Exit Function

    If IsExplicitGuideHardwareName(nm) Then
        IsGuideRoundCandidate = True
        Exit Function
    End If

    Dim dia As Double
    Dim lenA As Double

    dia = RoundBarDiameter(idx)
    lenA = RoundBarAxisLength(idx)

    If dia <= 0# Then Exit Function

    If dia >= 2# And lenA >= 3# Then
        IsGuideRoundCandidate = True
        Exit Function
    End If

    If dia >= 0.875 And lenA >= 6# Then
        IsGuideRoundCandidate = True
        Exit Function
    End If
End Function

Private Function StandardRoundComponentRole(ByVal idx As Long) As String
On Error GoTo ErrHandler

    StandardRoundComponentRole = ""

    If idx < 1 Or idx > PartCount Then Exit Function

    Dim raw As String
    Dim u As String
    Dim pn As String
    Dim dia As Double
    Dim axisLen As Double
    Dim ratio As Double

    raw = parts(idx).componentName
    u = UCase$(raw)
    pn = UCase$(CadPurchasePartToken(raw))

    If IsScrewFastenerName(raw) Then
        StandardRoundComponentRole = ""
        Exit Function
    End If

    If IsDowelOrMinorRoundHardwareName(raw) Then
        StandardRoundComponentRole = ""
        Exit Function
    End If

    If Not IsRoundBarLike(idx) Then Exit Function

    dia = RoundBarDiameter(idx)
    axisLen = RoundBarAxisLength(idx)

    If dia <= 0# Then Exit Function
    If dia > 4# Then Exit Function

    ratio = axisLen / dia

    If InStr(u, "LDR-PIN") > 0 Or InStr(u, "LDR_PIN") > 0 Then
        StandardRoundComponentRole = "Leader Pin"
        Exit Function
    End If

    If InStr(u, "/LBB_") > 0 Or InStr(u, "LBB_") > 0 Or InStr(u, "-LBB") > 0 Or InStr(u, "_LBB") > 0 Then
        StandardRoundComponentRole = "Leader Pin Bushing"
        Exit Function
    End If

    If InStr(u, "GEB_") > 0 Or InStr(u, "GUIDED EJECTOR") > 0 Then
        StandardRoundComponentRole = "Guided Ejector Bushing"
        Exit Function
    End If

    If InStr(u, "RETURN-PIN") > 0 Or InStr(u, "RETURN PIN") > 0 Then
        StandardRoundComponentRole = "Return Pin"
        Exit Function
    End If

    If InStr(u, "PILLAR_D") > 0 Or InStr(u, "SUPPORT PILLAR") > 0 Or InStr(u, "PILLAR") > 0 Then
        StandardRoundComponentRole = "Support Pillar"
        Exit Function
    End If

    If IsLatchLockName(raw) Then
        StandardRoundComponentRole = "Latch Lock / Safety Strap"
        Exit Function
    End If

    If dia >= 2# And axisLen >= 3# Then
        StandardRoundComponentRole = "Support Pillar"
        Exit Function
    End If

    If dia >= 1.35 And dia <= 2.5 And axisLen >= 6# And ratio >= 2# Then
        StandardRoundComponentRole = "Leader Pin"
        Exit Function
    End If

    If dia >= 0.75 And dia < 1.35 And axisLen >= 6# And ratio >= 4# Then
        If InStr(u, "RETURN") > 0 Or InStr(u, "EJECT") > 0 Then
            StandardRoundComponentRole = "Return Pin"
            Exit Function
        End If
    End If

    If ratio <= 2# And dia >= 1# And dia <= 2.75 Then
        If InStr(u, "BUSH") > 0 Or InStr(u, "SLEEVE") > 0 Then
            StandardRoundComponentRole = "Leader Pin Bushing"
            Exit Function
        End If
    End If

    Exit Function

ErrHandler:
    StandardRoundComponentRole = ""
End Function


Private Function IsLatchLockName(ByVal raw As String) As Boolean
    Dim u As String
    u = UCase(raw)
    IsLatchLockName = False
    If InStr(u, "LATCH-LOCK") > 0 Or InStr(u, "LATCH_LOCK") > 0 Or InStr(u, "LATCH LOCK") > 0 Then IsLatchLockName = True: Exit Function
    If InStr(u, "SAFETY-STRAP") > 0 Or InStr(u, "SAFETY_STRAP") > 0 Or InStr(u, "SAFETY STRAP") > 0 Then IsLatchLockName = True: Exit Function
    If InStr(u, "SAFTEY-STRAP") > 0 Or InStr(u, "SAFTEY_STRAP") > 0 Or InStr(u, "SAFTEY STRAP") > 0 Then IsLatchLockName = True: Exit Function
    If InStr(u, "PLC75") > 0 Then IsLatchLockName = True: Exit Function
    ' PLC + digit (PLC1, PLC2, ...) without matching random "PLC" substrings alone.
    Dim p As Long, ch As String
    p = InStr(u, "PLC")
    Do While p > 0
        If p + 3 <= Len(u) Then
            ch = Mid(u, p + 3, 1)
            If ch >= "0" And ch <= "9" Then IsLatchLockName = True: Exit Function
        End If
        p = InStr(p + 1, u, "PLC")
    Loop
End Function

' True when two round parts share a lateral center plane (orthogonal to stack axis).
' Mirrors qwen near_same_axis_plane: any two of X/Y/Z within tolerance.
Private Function NearSameAxisPlane(ByVal aIdx As Long, ByVal bIdx As Long, Optional ByVal tol As Double = LEADER_PIN_BUSHING_PLANE_TOL) As Boolean
    NearSameAxisPlane = False
    If aIdx < 1 Or bIdx < 1 Or aIdx > PartCount Or bIdx > PartCount Then Exit Function
    Dim dx As Double, dy As Double, dz As Double
    dx = Abs(parts(aIdx).AsmCenterX - parts(bIdx).AsmCenterX)
    dy = Abs(parts(aIdx).AsmCenterY - parts(bIdx).AsmCenterY)
    dz = Abs(parts(aIdx).AsmCenterZ - parts(bIdx).AsmCenterZ)
    NearSameAxisPlane = ((dx <= tol And dy <= tol) Or (dx <= tol And dz <= tol) Or (dy <= tol And dz <= tol))
End Function

' Classify leader-pin sets using bushing co-location AND distance from the
' ejector stack. When two guide-pin sets exist, PRIMARY = the set farther
' from the ejectors (main LDR-PIN / B-plate set); SECONDARY = the set nearer
' the ejectors (EJ_LDR_PIN / guided-ejector set). Secondary never decides A/B.
Private Sub ClassifyLeaderPinSetsByBushingPlane(ByRef lpIdx() As Long, ByVal nLp As Long, _
                                               ByVal ax As Integer, ByVal supportPos As Double)
    Dim i As Long, j As Long
    Dim roleKey As String
    Dim pinIdx(1 To 120) As Long, nPin As Long
    Dim shoulderIdx(1 To 120) As Long, nShoulder As Long
    Dim ejectorBushIdx(1 To 120) As Long, nEjectorBush As Long
    Dim longPinIdx(1 To 120) As Long, nLong As Long
    Dim shortBushIdx(1 To 120) As Long, nShort As Long
    Dim dia As Double, axisLen As Double, ratio As Double
    Dim matched As Boolean
    Dim ejMean As Double, ejCount As Long
    Dim uName As String

    If PartCount < 1 Then Exit Sub
    On Error Resume Next
    If UBound(gStdLeaderPinSetByPart) < 1 Then ReDim gStdLeaderPinSetByPart(1 To PartCount)
    On Error GoTo 0
    If Not StdRoleArrayReady() Then
        ReDim gStdRoleByPart(1 To PartCount)
    End If

    ' Clear previous set tags for a clean re-run.
    For i = 1 To PartCount
        gStdLeaderPinSetByPart(i) = ""
    Next i

    nPin = 0: nShoulder = 0: nEjectorBush = 0: nLong = 0: nShort = 0
    ejMean = 0#: ejCount = 0

    ' Ejector-stack anchor position (plates + rails) for "farther from ejectors".
    For i = 1 To PartCount
        roleKey = NormalizeKey(StdCadRole(i))
        uName = UCase(parts(i).componentName)
        If roleKey = "EJECTORPLATE" Or roleKey = "BOTTOMEJECTORPLATE" Or roleKey = "RAILS" Or _
           InStr(uName, "EJ-RET") > 0 Or InStr(uName, "EJ-BACKUP") > 0 Or _
           InStr(uName, "EJ_RET") > 0 Or InStr(uName, "EJ_BACKUP") > 0 Then
            ejMean = ejMean + PartAxisCenter(i, ax)
            ejCount = ejCount + 1
        End If
    Next i
    If ejCount > 0 Then
        ejMean = ejMean / ejCount
    ElseIf supportPos <> 0# Then
        ejMean = supportPos
        ejCount = 1
    End If

    For i = 1 To nLp
        roleKey = NormalizeKey(StandardRoundComponentRole(lpIdx(i)))
        Select Case roleKey
            Case "LEADERPIN"
                If nPin < UBound(pinIdx) Then nPin = nPin + 1: pinIdx(nPin) = lpIdx(i)
            Case "LEADERPINBUSHING"
                If nShoulder < UBound(shoulderIdx) Then nShoulder = nShoulder + 1: shoulderIdx(nShoulder) = lpIdx(i)
            Case "GUIDEDEJECTORBUSHING"
                If nEjectorBush < UBound(ejectorBushIdx) Then nEjectorBush = nEjectorBush + 1: ejectorBushIdx(nEjectorBush) = lpIdx(i)
        End Select
    Next i

    For i = 1 To PartCount

        If Not IsGuideRoundCandidate(i) Then GoTo nextRound

        dia = RoundBarDiameter(i)
        axisLen = RoundBarAxisLength(i)

        If dia <= 0# Or dia > 4# Then GoTo nextRound

        ratio = axisLen / dia

        If ratio >= 2# And axisLen >= 6# Then
            If nLong < UBound(longPinIdx) Then
                nLong = nLong + 1
                longPinIdx(nLong) = i
            End If

        ElseIf ratio >= 0.6 And ratio <= 1.6 And dia >= 1# And dia <= 2.75 Then
            If InStr(UCase$(parts(i).componentName), "BUSH") > 0 Or _
               InStr(UCase$(parts(i).componentName), "LBB") > 0 Or _
               InStr(UCase$(parts(i).componentName), "GEB") > 0 Then

                If nShort < UBound(shortBushIdx) Then
                    nShort = nShort + 1
                    shortBushIdx(nShort) = i
                End If
            End If
        End If

nextRound:
    Next i

    ' Ensure pin list includes all long guide pins (not only pre-tagged lpIdx).
    For i = 1 To nLong
        matched = False
        For j = 1 To nPin
            If pinIdx(j) = longPinIdx(i) Then matched = True: Exit For
        Next j
        If Not matched And nPin < UBound(pinIdx) Then
            nPin = nPin + 1
            pinIdx(nPin) = longPinIdx(i)
        End If
    Next i

    ' --- Cluster pins by stack-axis center (two sets = two distinct stack positions) ---
    Dim setPos(1 To 40) As Double
    Dim setCount(1 To 40) As Long
    Dim setSum(1 To 40) As Double
    Dim pinSetId(1 To 120) As Long
    Dim nSets As Long
    Dim pPos As Double
    Dim bestSet As Long
    Dim bestDist As Double
    Dim d As Double
    Const PIN_SET_CLUSTER_TOL As Double = 1.25

    nSets = 0
    For i = 1 To nPin
        pPos = PartAxisCenter(pinIdx(i), ax)
        bestSet = 0
        bestDist = 1E+30
        For j = 1 To nSets
            d = Abs(pPos - setPos(j))
            If d < bestDist Then bestDist = d: bestSet = j
        Next j
        If bestSet > 0 And bestDist <= PIN_SET_CLUSTER_TOL Then
            pinSetId(i) = bestSet
            setSum(bestSet) = setSum(bestSet) + pPos
            setCount(bestSet) = setCount(bestSet) + 1
            setPos(bestSet) = setSum(bestSet) / setCount(bestSet)
        ElseIf nSets < UBound(setPos) Then
            nSets = nSets + 1
            pinSetId(i) = nSets
            setSum(nSets) = pPos
            setCount(nSets) = 1
            setPos(nSets) = pPos
        End If
    Next i

    ' Bushing evidence per cluster: shoulder/LBB vs guided-ejector matches.
    Dim setShoulderHits(1 To 40) As Long
    Dim setEjectorHits(1 To 40) As Long
    Dim sid As Long
    For i = 1 To nPin
        sid = pinSetId(i)
        If sid < 1 Then GoTo nextBushEv
        For j = 1 To nShoulder
            If NearSameAxisPlane(pinIdx(i), shoulderIdx(j)) Then
                setShoulderHits(sid) = setShoulderHits(sid) + 1
                Exit For
            End If
        Next j
        For j = 1 To nEjectorBush
            If NearSameAxisPlane(pinIdx(i), ejectorBushIdx(j)) Then
                setEjectorHits(sid) = setEjectorHits(sid) + 1
                Exit For
            End If
        Next j
        ' Unnamed short bushings on same plane also count as shoulder-like evidence.
        For j = 1 To nShort
            If NearSameAxisPlane(pinIdx(i), shortBushIdx(j), LEADER_PIN_BUSHING_PLANE_TOL) Then
                If setShoulderHits(sid) = 0 And setEjectorHits(sid) = 0 Then
                    setShoulderHits(sid) = setShoulderHits(sid) + 1
                End If
                Dim bushEvRole As String
                bushEvRole = StandardRoundComponentRole(pinIdx(i))
                If NormalizeKey(bushEvRole) = "RETURNPIN" Or NormalizeKey(bushEvRole) = "EJECTORRETURNPIN" Then
                    SetStdCadRole pinIdx(i), "Return Pin"
                Else
                    SetStdCadRole pinIdx(i), "Leader Pin"
                End If
                If StdCadRole(shortBushIdx(j)) = "" Then SetStdCadRole shortBushIdx(j), "Leader Pin Bushing"
                Exit For
            End If
        Next j
nextBushEv:
    Next i

    ' Pick PRIMARY set:
    '   1) If 2+ clusters and we know ejector position -> farther from ejectors wins
    '   2) Else cluster with more shoulder/LBB hits
    '   3) Else single cluster / first cluster
    Dim primarySet As Long
    Dim secondarySet As Long
    Dim farDist As Double, nearDist As Double
    Dim farSet As Long, nearSet As Long
    Dim bestHits As Long
    primarySet = 0
    secondarySet = 0

    If nSets >= 2 And ejCount > 0 Then
        farDist = -1#
        nearDist = 1E+30
        farSet = 1: nearSet = 1
        For j = 1 To nSets
            d = Abs(setPos(j) - ejMean)
            If d > farDist Then farDist = d: farSet = j
            If d < nearDist Then nearDist = d: nearSet = j
        Next j
        primarySet = farSet
        If nearSet <> farSet Then secondarySet = nearSet
        LogLine "Leader-pin sets: " & nSets & " clusters; PRIMARY=set" & primarySet & _
                " (farther from ejectors, dist=" & FormatNumberForCsv(farDist) & _
                ") SECONDARY=set" & secondarySet & " (nearer ejectors, dist=" & FormatNumberForCsv(nearDist) & _
                ") ejMean=" & FormatNumberForCsv(ejMean)
    Else
        bestHits = -1
        primarySet = 1
        For j = 1 To nSets
            If setShoulderHits(j) > bestHits Then
                bestHits = setShoulderHits(j)
                primarySet = j
            End If
        Next j
        For j = 1 To nSets
            If j <> primarySet And (setEjectorHits(j) > 0 Or setCount(j) > 0) Then
                If secondarySet = 0 Or setEjectorHits(j) > setEjectorHits(secondarySet) Then secondarySet = j
            End If
        Next j
        LogLine "Leader-pin sets: " & nSets & " cluster(s); PRIMARY=set" & primarySet & _
                " (bushing/shoulder evidence) SECONDARY=set" & secondarySet
    End If

    Dim pinRole As String

    For i = 1 To nPin
        sid = pinSetId(i)

        pinRole = StandardRoundComponentRole(pinIdx(i))

        If NormalizeKey(pinRole) = "RETURNPIN" Or NormalizeKey(pinRole) = "EJECTORRETURNPIN" Then
            pinRole = "Return Pin"
        Else
            pinRole = "Leader Pin"
        End If

        If sid = primarySet Then
            gStdLeaderPinSetByPart(pinIdx(i)) = "PRIMARY"
            SetStdCadRole pinIdx(i), pinRole

        ElseIf sid = secondarySet Or (secondarySet = 0 And setEjectorHits(sid) > setShoulderHits(sid)) Then
            gStdLeaderPinSetByPart(pinIdx(i)) = "SECONDARY"
            SetStdCadRole pinIdx(i), pinRole

        ElseIf sid > 0 Then
            If ejCount > 0 And Abs(setPos(sid) - ejMean) + 0.25 < Abs(setPos(primarySet) - ejMean) Then
                gStdLeaderPinSetByPart(pinIdx(i)) = "SECONDARY"
            Else
                gStdLeaderPinSetByPart(pinIdx(i)) = "PRIMARY"
            End If

            SetStdCadRole pinIdx(i), pinRole
        End If
    Next i

    Dim nPri As Long, nSec As Long
    nPri = 0: nSec = 0
    For i = 1 To PartCount
        If gStdLeaderPinSetByPart(i) = "PRIMARY" Then nPri = nPri + 1
        If gStdLeaderPinSetByPart(i) = "SECONDARY" Then nSec = nSec + 1
    Next i
    LogLine "Leader-pin sets: PRIMARY=" & nPri & " (farther from ejectors / shoulder-LBB) SECONDARY=" & nSec & " (nearer ejectors; does not decide A/B)"
End Sub

' Measure whether primary leader pins enter from the top (A) or bottom (B) of
' the already-oriented stack. Reversed pins (seated in B, running toward A)
' are flagged but NEVER used to flip a confirmed A/B assignment.
Private Sub MeasureLeaderPinTopBottomDirection(ByVal ax As Integer)
    Dim i As Long
    Dim pinMean As Double, pinCount As Long
    Dim bushMean As Double, bushCount As Long
    Dim aPos As Double, bPos As Double
    Dim roleKey As String

    gStdLeaderPinFromKnown = False
    gStdLeaderPinFromTop = False
    gStdLeaderPinReversed = False
    If gStdCavityCadIndex < 1 Or gStdCoreCadIndex < 1 Then Exit Sub
    If PartCount < 1 Then Exit Sub

    aPos = PartAxisCenter(gStdCavityCadIndex, ax)
    bPos = PartAxisCenter(gStdCoreCadIndex, ax)

    For i = 1 To PartCount
        roleKey = NormalizeKey(StdCadRole(i))
        If roleKey = "" Then roleKey = NormalizeKey(StandardRoundComponentRole(i))
        If roleKey = "LEADERPIN" Then
            ' Only PRIMARY set decides guide direction (Qwen rule).
            If gStdLeaderPinSetByPart(i) = "SECONDARY" Then GoTo nextPin
            pinMean = pinMean + PartAxisCenter(i, ax)
            pinCount = pinCount + 1
        ElseIf roleKey = "LEADERPINBUSHING" Then
            bushMean = bushMean + PartAxisCenter(i, ax)
            bushCount = bushCount + 1
        End If
nextPin:
    Next i

    If pinCount < 1 Then Exit Sub
    pinMean = pinMean / pinCount
    gStdLeaderPinFromKnown = True

    ' Pins closer to B/core than A/cavity => seated on B side (typical / reversed-ok).
    ' "From top" means pin centers are nearer the A/cavity plate (pins enter from top).
    gStdLeaderPinFromTop = (Abs(pinMean - aPos) < Abs(pinMean - bPos))

    If bushCount > 0 Then
        bushMean = bushMean / bushCount
        ' Reversed: pin body on B side while bushings sit toward A (pins run upward).
        If (Abs(pinMean - bPos) < Abs(pinMean - aPos)) And (Abs(bushMean - aPos) < Abs(bushMean - bPos)) Then
            gStdLeaderPinReversed = True
        End If
    End If

    LogLine "Leader-pin direction: from_" & IIf(gStdLeaderPinFromTop, "TOP/A", "BOTTOM/B") & _
            IIf(gStdLeaderPinReversed, " (REVERSED — seated in B running toward A; A/B NOT flipped)", "") & _
            " | pinMean=" & FormatNumberForCsv(pinMean) & " A=" & FormatNumberForCsv(aPos) & " B=" & FormatNumberForCsv(bPos)
End Sub

Private Function StdStackAxisName(ByVal ax As Integer) As String
    Select Case ax
        Case 1: StdStackAxisName = "CenterX"
        Case 2: StdStackAxisName = "CenterY"
        Case Else: StdStackAxisName = "CenterZ"
    End Select
End Function

Private Sub BuildStdStackAnalysisText(ByVal ax As Integer, ByVal nFull As Long, _
                                      ByVal nRail As Long, ByVal nEj As Long, ByVal nLp As Long)
    Dim axisName As String
    axisName = StdStackAxisName(ax)
    gStdPartingLineText = "Between a_plate and b_plate from the full-footprint stack order."
    gStdStackRules = "Full-footprint plates were sorted by " & axisName & " from top to bottom."
    gStdStackRules = gStdStackRules & "|Rails and the ejector stack anchored the bottom of the stack first; leader-pin direction was not used to flip stack orientation."
    gStdStackRules = gStdStackRules & "|Ejector-stack plates: thinner = Ejector Plate, thicker/lower = Bottom Ejector Plate."
    gStdStackRules = gStdStackRules & "|Rails detected as long narrow side-offset blocks in the ejector/rail zone; ejector plates as centered medium-width plates near rails."
    gStdStackRules = gStdStackRules & "|Round guide hardware separated by diameter/length; when two leader-pin sets exist, PRIMARY = farther from ejectors (bushing-matched), SECONDARY = nearer ejectors."
    gStdStackRules = gStdStackRules & "|Exact shop-name tokens (A-PLATE, B-PLATE, SC-RETAINER, SC-BACKUP, EJ-RET, EJ-BACKUP, RAIL, LDR-PIN, LBB) applied before geometry-only rules."
    If nFull = 2 Then
        gStdStackRules = gStdStackRules & "|Two-half mold pattern: 2 full clamps + inner A/B blocks + thin rails/ejector from remaining thin-large parts."
    End If
    If gStdSequencedLatchLock Then
        gStdPartingLineText = "Primary parting line between a_plate and b_plate from the full-footprint stack order. " & _
            "Latch-lock/PLC/safety-strap hardware detected: plate-sequenced/latch-lock standard base with secondary opening/parting lines at latch attachment points."
        gStdStackRules = gStdStackRules & "|Latch-lock/PLC/safety-strap tokens detected. Latch locks mark secondary parting lines and do not set guide direction. " & _
            "Reversed/seated leader pins were not allowed to flip the A/B assignment."
        gStdStackRules = gStdStackRules & "|Sequenced/SC stack naming used when top clamp missing: A / B / SC Retainer / SC Backup / Bottom Clamp."
    End If
    If gStdLeaderPinFromKnown Then
        gStdStackRules = gStdStackRules & "|Leader pins enter from the " & IIf(gStdLeaderPinFromTop, "TOP/A", "BOTTOM/B") & " side" & _
            IIf(gStdLeaderPinReversed, " (reversed seating)", "") & "."
    End If
    gStdStackRules = gStdStackRules & "|Counts: full=" & nFull & " rails=" & nRail & " ejector=" & nEj & " leaderStack=" & nLp & "."
End Sub

Private Function SuggestedPcsNameForCadIndex(ByVal idx As Long, ByVal role As String, _
                                            ByRef confidence As String, ByRef reason As String) As String
    confidence = "LOW"
    reason = "No strong PCS geometry rule matched."
    SuggestedPcsNameForCadIndex = ""

    If idx < 1 Or idx > PartCount Then Exit Function
    Dim dia As Double
    Dim axisLen As Double

    ' Pot/holder assemblies can have generic CAD names and round-ish geometry, so
    ' trust the geometry/position role before applying generic PCS round-part rules.
    Select Case NormalizeKey(role)
        Case "TCP", "BCP", "IDHOLDER", "ODHOLDER", "IDPOT", "ODPOT"
            SuggestedPcsNameForCadIndex = CleanFileName(IIf(JobBaseName <> "", JobBaseName, CurrentJobNumber) & "_" & Replace(role, " ", "-"))
            confidence = "HIGH"
            reason = "Matched by pot-block geometry and assembly position, not by CAD file name."
            Exit Function
        Case "TOPCLAMPPLATE", "BOTTOMCLAMPPLATE", "APLATE", "BPLATE", "CAVITYPLATE", "COREPLATE", "SUPPORTPLATE", _
             "PINPLATE", "EJECTORPLATE", "EJECTORRETAINERPLATE", "BOTTOMEJECTORPLATE", _
             "SCRETAINERPLATE", "SCBACKUPPLATE", "STRIPPERPLATE", "MANIFOLDPLATE", "DIEPLATE", "DIEBACKUPPLATE"
            SuggestedPcsNameForCadIndex = CleanFileName(IIf(JobBaseName <> "", JobBaseName, CurrentJobNumber) & "_" & Replace(role, " ", "-"))
            confidence = "HIGH"
            reason = "Standard-base stack/size rule assigned this CAD component without needing a BOM."
            Exit Function
        Case "RAILS"
            SuggestedPcsNameForCadIndex = CleanFileName(IIf(JobBaseName <> "", JobBaseName, CurrentJobNumber) & "_Rails")
            confidence = "HIGH"
            reason = "Standard-base rail geometry: long narrow support blocks beside the ejector stack."
            Exit Function
        Case "LEADERPIN"
            dia = RoundBarDiameter(idx)
            axisLen = RoundBarAxisLength(idx)
            SuggestedPcsNameForCadIndex = "LEADER_PIN_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(axisLen) & "_PCS"
            confidence = "MEDIUM"
            reason = "Round long pin geometry matched leader-pin stack logic."
            Exit Function
        Case "LEADERPINBUSHING"
            dia = RoundBarDiameter(idx)
            axisLen = RoundBarAxisLength(idx)
            SuggestedPcsNameForCadIndex = "LP_BUSHING_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(axisLen) & "_PCS"
            confidence = "MEDIUM"
            reason = "Round sleeve/bushing geometry in the leader-pin stack."
            Exit Function
        Case "GUIDEDEJECTORBUSHING"
            dia = RoundBarDiameter(idx)
            axisLen = RoundBarAxisLength(idx)
            SuggestedPcsNameForCadIndex = "GUIDED_EJ_BUSHING_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(axisLen) & "_PCS"
            confidence = "MEDIUM"
            reason = "Round short bushing geometry near ejector-side stack."
            Exit Function
        Case "RETURNPIN", "EJECTORRETURNPIN"
            dia = RoundBarDiameter(idx)
            axisLen = RoundBarAxisLength(idx)
            SuggestedPcsNameForCadIndex = "RETURN_PIN_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(axisLen) & "_PCS"
            confidence = "MEDIUM"
            reason = "Small long round pin in the ejector-side stack."
            Exit Function
        Case "EJECTORPIN"
            dia = RoundBarDiameter(idx)
            axisLen = RoundBarAxisLength(idx)
            SuggestedPcsNameForCadIndex = "EJECTOR_PIN_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(axisLen) & "_PCS"
            confidence = "MEDIUM"
            reason = "Small long round-pin geometry."
            Exit Function
        Case "SUPPORTPILLAR"
            dia = RoundBarDiameter(idx)
            axisLen = RoundBarAxisLength(idx)
            SuggestedPcsNameForCadIndex = "SUPPORT_PILLAR_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(axisLen) & "_DME"
            confidence = "HIGH"
            reason = "Large round post geometry; kept separate from leader-pin logic."
            Exit Function
        Case "LATCHLOCK/SAFETYSTRAP", "LATCHLOCKSAFETYSTRAP", "LATCHLOCK"
            SuggestedPcsNameForCadIndex = "LATCH_LOCK_PLC"
            confidence = "HIGH"
            reason = "Shop latch-lock/PLC/safety-strap token; secondary parting marker, does not set guide direction."
            Exit Function
    End Select

    If IsRoundBarLike(idx) Then
        dia = RoundBarDiameter(idx)
        If parts(idx).Length / IIf(dia <= 0#, 1#, dia) >= 3.5 Then
            If dia >= 1.5 Then
                SuggestedPcsNameForCadIndex = "PILLAR_D" & PcsDimToken(dia) & "-X-" & PcsDimToken(parts(idx).Length) & "_PCS"
                confidence = "MEDIUM"
                reason = "Round-bar geometry with large diameter and long length; PCS examples use PILLAR_D<size>-X-<length>_PCS."
            Else
                SuggestedPcsNameForCadIndex = "EJ_LDR_PIN_D-" & PcsDimToken(dia) & "-X-" & PcsDimToken(parts(idx).Length) & "_PCS"
                confidence = "MEDIUM"
                reason = "Round-bar geometry with small diameter and long length; likely ejector/leader pin."
            End If
        Else
            SuggestedPcsNameForCadIndex = "BUSHING_D" & PcsDimToken(dia) & "-X-" & PcsDimToken(parts(idx).Length) & "_PCS"
            confidence = "LOW"
            reason = "Round short component; could be a bushing or sleeve, verify against BOM/vendor."
        End If
    End If
End Function

Private Function LooksLikePlate(ByVal idx As Long) As Boolean
    If idx < 1 Or idx > PartCount Then Exit Function
    If parts(idx).Thickness <= 0# Then Exit Function
    If parts(idx).Length <= 0# Or parts(idx).Width <= 0# Then Exit Function
    LooksLikePlate = (parts(idx).Width * parts(idx).Length >= PLATE_MIN_FOOTPRINT And parts(idx).Thickness <= parts(idx).Length * 0.35)
End Function

Private Function IsRoundBarLike(ByVal idx As Long) As Boolean
    If idx < 1 Or idx > PartCount Then Exit Function
    If parts(idx).Width <= 0# Or parts(idx).Thickness <= 0# Then Exit Function
    Dim dia As Double
    Dim axisLen As Double
    dia = RoundBarDiameter(idx)
    If dia <= 0# Then Exit Function
    If dia > 4# Then Exit Function
    axisLen = RoundBarAxisLength(idx)
    IsRoundBarLike = (axisLen >= dia * 0.25)
End Function

Private Function RoundBarDiameter(ByVal idx As Long) As Double
    If idx < 1 Or idx > PartCount Then Exit Function
    Dim a As Double, b As Double, c As Double
    Dim dTW As Double, dTL As Double, dWL As Double
    a = parts(idx).Thickness
    b = parts(idx).Width
    c = parts(idx).Length
    dTW = Abs(a - b)
    dTL = Abs(a - c)
    dWL = Abs(b - c)

    If dTW <= dTL And dTW <= dWL Then
        RoundBarDiameter = (a + b) / 2#
    ElseIf dTL <= dTW And dTL <= dWL Then
        RoundBarDiameter = (a + c) / 2#
    Else
        RoundBarDiameter = (b + c) / 2#
    End If

    If RoundBarDiameter <= 0# Then Exit Function
    If MinDouble3(dTW, dTL, dWL) > MaxDouble(0.08, RoundBarDiameter * 0.1) Then RoundBarDiameter = 0#
End Function

Private Function RoundBarAxisLength(ByVal idx As Long) As Double
    If idx < 1 Or idx > PartCount Then Exit Function
    Dim a As Double, b As Double, c As Double
    Dim dTW As Double, dTL As Double, dWL As Double
    a = parts(idx).Thickness
    b = parts(idx).Width
    c = parts(idx).Length
    dTW = Abs(a - b)
    dTL = Abs(a - c)
    dWL = Abs(b - c)

    If dTW <= dTL And dTW <= dWL Then
        RoundBarAxisLength = c
    ElseIf dTL <= dTW And dTL <= dWL Then
        RoundBarAxisLength = b
    Else
        RoundBarAxisLength = a
    End If
End Function

Private Function PcsDimToken(ByVal v As Double) As String
    If v <= 0# Then PcsDimToken = "0": Exit Function
    Dim whole As Long
    Dim frac As Double
    Dim num As Long
    Dim den As Long
    whole = Fix(v)
    frac = v - whole
    den = 16
    num = CLng(frac * den + 0.5)
    If num >= den Then
        whole = whole + 1
        num = 0
    End If
    If num = 0 Then
        PcsDimToken = CStr(whole)
    Else
        Do While num Mod 2 = 0 And den > 2
            num = num \ 2
            den = den \ 2
        Loop
        If whole > 0 Then
            PcsDimToken = CStr(whole) & "-" & CStr(num) & "-" & CStr(den)
        Else
            PcsDimToken = CStr(num) & "-" & CStr(den)
        End If
    End If
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    MinDouble3 = a
    If b < MinDouble3 Then MinDouble3 = b
    If c < MinDouble3 Then MinDouble3 = c
End Function
' ============================================================
' PURCHASED COMPONENTS  (DME / McMaster / Jaco hardware)
' Captured from the BOM "Purchase"/hardware rows and priced from
' Purchased Components Prices.csv (editable unit price + part #).
' ============================================================

Private Function FindPurchasedPriceFile() As String
    On Error Resume Next
    FindPurchasedPriceFile = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim cands(1 To 7) As String, i As Long
    cands(1) = CurrentJobFolder & "\" & PURCHASED_PRICE_FILE
    cands(2) = DOWNLOADS_FOLDER & "\" & PURCHASED_PRICE_FILE
    cands(3) = TRUSTED_FOLDER & "\" & PURCHASED_PRICE_FILE
    cands(4) = LOCAL_WORKSPACE_ROOT & "\" & PURCHASED_PRICE_FILE
    cands(5) = PRIVATE_DATA_ROOT & "\" & PURCHASED_PRICE_FILE
    cands(6) = DOWNLOADS_FOLDER & "\New folder (17)\" & PURCHASED_PRICE_FILE
    cands(7) = "C:\Users\lenovo\Downloads\New folder (17)\" & PURCHASED_PRICE_FILE

    For i = 1 To 7
        If cands(i) <> "" And Right(cands(i), 1) <> "\" Then
            If fso.FileExists(cands(i)) Then
                LogLine "Price file FOUND: " & cands(i)
                FindPurchasedPriceFile = cands(i)
                Exit Function
            Else
                LogLine "Price file not at: " & cands(i)
            End If
        End If
    Next i

    ' Last resort: search the common roots recursively for the file by name.
    Dim hit As String
    hit = SearchForFileByName(DOWNLOADS_FOLDER, PURCHASED_PRICE_FILE, fso, 0)
    If hit = "" Then hit = SearchForFileByName(LOCAL_WORKSPACE_ROOT, PURCHASED_PRICE_FILE, fso, 0)
    If hit = "" Then hit = SearchForFileByName(TRUSTED_FOLDER, PURCHASED_PRICE_FILE, fso, 0)
    If hit <> "" Then
        LogLine "Price file FOUND by search: " & hit
        FindPurchasedPriceFile = hit
    Else
        LogLine "Price file NOT FOUND anywhere - the quote will use the built-in list (all $0). " & _
                "Put '" & PURCHASED_PRICE_FILE & "' in " & DOWNLOADS_FOLDER
    End If
End Function

' Recursively find the first file named exactly fileName under root (max 4 levels).
Private Function SearchForFileByName(ByVal root As String, ByVal fileName As String, _
                                     ByVal fso As Object, ByVal depth As Long) As String
    On Error Resume Next
    SearchForFileByName = ""
    If depth > 4 Then Exit Function
    If Not fso.FolderExists(root) Then Exit Function
    Dim direct As String
    direct = root & "\" & fileName
    If fso.FileExists(direct) Then SearchForFileByName = direct: Exit Function
    Dim sub1 As Object, r As String
    For Each sub1 In fso.GetFolder(root).SubFolders
        r = SearchForFileByName(sub1.path, fileName, fso, depth + 1)
        If r <> "" Then SearchForFileByName = r: Exit Function
    Next sub1
End Function

Private Function ParsePriceToken(ByVal s As String) As Double
    Dim t As String
    t = Trim(s)
    t = Replace(t, "$", ""): t = Replace(t, ",", ""): t = Replace(t, " ", "")
    If IsNumeric(t) Then ParsePriceToken = CDbl(t) Else ParsePriceToken = 0
End Function

Private Sub LoadPurchasedPriceList()
    PpCount = 0
    PlCount = 0
    If Not FILL_PURCHASED_COMPONENTS Then Exit Sub
    Dim path As String
    path = FindPurchasedPriceFile()
    gPriceListPath = path
    If path = "" Then
        LogLine "Purchased price list file not found; using built-in default component list (prices 0)."
        SeedDefaultPriceList
        Exit Sub
    End If
    On Error GoTo eh
    ReDim PlComp(1 To 200): ReDim PlVendor(1 To 200): ReDim PlPartNo(1 To 200)
    ReDim PlDescr(1 To 200): ReDim PlUnit(1 To 200): ReDim PlPrice(1 To 200)
    Dim f As Integer, line As String, parts() As String
    Dim headerChecked As Boolean
    headerChecked = False
    Dim whole As String, allLines() As String, li As Long
    f = FreeFile
    Open path For Input As #f
    If LOF(f) > 0 Then whole = Input$(LOF(f), f)
    Close #f
    ' Normalize any line ending (Windows CRLF, Unix LF, old Mac CR) - VBA's
    ' Line Input only breaks on CR, so a tool that saved Unix LF would otherwise
    ' make the whole file read as one line (and parse 0 rows).
    whole = Replace(whole, vbCrLf, vbLf)
    whole = Replace(whole, vbCr, vbLf)
    allLines = Split(whole, vbLf)
    For li = 0 To UBound(allLines)
        line = Trim(allLines(li))
        If line <> "" And Left(line, 1) <> "#" Then
            If Not headerChecked Then
                headerChecked = True
                If InStr(1, line, "Component", vbTextCompare) > 0 And InStr(1, line, "Price", vbTextCompare) > 0 Then
                    line = ""   ' skip header row
                End If
            End If
            If line <> "" Then
                If InStr(line, vbTab) > 0 Then
                    parts = Split(line, vbTab)      ' tab-delimited (Excel "Save As Text")
                Else
                    parts = Split(line, ",")        ' comma CSV
                End If
                If UBound(parts) >= 5 Then
                    PlCount = PlCount + 1
                    If PlCount > UBound(PlComp) Then
                        ReDim Preserve PlComp(1 To PlCount + 50): ReDim Preserve PlVendor(1 To PlCount + 50)
                        ReDim Preserve PlPartNo(1 To PlCount + 50): ReDim Preserve PlDescr(1 To PlCount + 50)
                        ReDim Preserve PlUnit(1 To PlCount + 50): ReDim Preserve PlPrice(1 To PlCount + 50)
                    End If
                    PlComp(PlCount) = Trim(parts(0))
                    PlVendor(PlCount) = Trim(parts(1))
                    PlPartNo(PlCount) = Trim(parts(2))
                    PlDescr(PlCount) = Trim(parts(3))
                    PlUnit(PlCount) = Trim(parts(4))
                    PlPrice(PlCount) = ParsePriceToken(parts(5))
                End If
            End If
        End If
    Next li
    If PlCount = 0 Then
        LogLine "Price list parsed 0 rows from " & path & " - using built-in default list."
        SeedDefaultPriceList
    Else
        LogLine "Loaded purchased price list: " & PlCount & " line(s) from " & path
        Dim z As Long
        For z = 1 To PlCount
            LogLine "  pricelist[" & z & "] " & PlVendor(z) & " " & PlPartNo(z) & _
                    " = $" & FormatNumberForCsv(PlPrice(z))
        Next z
    End If
    Exit Sub
eh:
    LogLine "LoadPurchasedPriceList error: " & Err.Description
    On Error Resume Next
    Close #f
    If PlCount = 0 Then SeedDefaultPriceList
End Sub

' Built-in default component list so matching always works even with no CSV / bad
' delimiter. Prices are 0 - edit the CSV to set real prices. Matching is by keyword
' and by BOM part number, so 0-price parts still fill the Components row.
Private Sub SeedDefaultPriceList()
    ReDim PlComp(1 To 24): ReDim PlVendor(1 To 24): ReDim PlPartNo(1 To 24)
    ReDim PlDescr(1 To 24): ReDim PlUnit(1 To 24): ReDim PlPrice(1 To 24)
    PlCount = 0
    AddSeedRow "Leader Pin", "DME", "", "Shoulder leader pin", "EA", 0
    AddSeedRow "Guide Bushing", "DME", "", "Guide / leader bushing", "EA", 0
    AddSeedRow "Internal Retainer Ring", "PCS", "", "Internal retaining ring", "EA", 0
    AddSeedRow "Safety Strap", "DME", "", "Safety strap", "EA", 0
    AddSeedRow "Space Insulation", "Pyropel", "", "Pyropel insulation sheet", "Sq In", 0
    AddSeedRow "Sleeve Bearing", "Jaco", "", "Sleeve bearing / bushing", "EA", 0
    AddSeedRow "Leader Pin", "DME", "5211GL", "Leader pin", "EA", 0
    AddSeedRow "Leader Pin", "DME", "5210GL", "Leader pin", "EA", 0
    AddSeedRow "Guide Bushing", "DME", "5503", "Bushing", "EA", 0
    AddSeedRow "Top Plate", "DME", "1620-13-2", "Top plate", "EA", 0
    AddSeedRow "Bottom Plate", "DME", "1620-13-2", "Bottom plate", "EA", 0
    AddSeedRow "Guide Bushing", "DME", "B6", "Base / leader bushing family", "EA", 0
    AddSeedRow "Leader Pin", "DME", "P6", "Base / guide pin family", "EA", 0
    AddSeedRow "Retaining Ring", "DME", "MUD", "MUD retaining ring / circlip family", "EA", 0
    AddSeedRow "Support Pillar", "DME", "6143", "Support pillar", "EA", 0
    AddSeedRow "Sleeve Bearing", "Jaco", "HT200", "Sleeve bearing", "EA", 0
    AddSeedRow "Hardware", "McMaster-Carr", "6381K539", "Ejector bushing", "EA", 0
    AddSeedRow "Safety Strap", "PCS", "LSS-300", "Safety strap", "EA", 0
    AddSeedRow "Internal Retainer Ring", "McMaster-Carr", "99142A520", "Retaining ring", "EA", 0
    LogLine "Seeded built-in price list: " & PlCount & " components (prices 0)."
End Sub

Private Sub AddSeedRow(ByVal comp As String, ByVal ven As String, ByVal pn As String, _
                       ByVal descr As String, ByVal unit As String, ByVal price As Double)
    PlCount = PlCount + 1
    PlComp(PlCount) = comp
    PlVendor(PlCount) = ven
    PlPartNo(PlCount) = pn
    PlDescr(PlCount) = descr
    PlUnit(PlCount) = unit
    PlPrice(PlCount) = price
End Sub

' Distinctive search token for a price-list component (avoids false hits like
' matching "return pin" to a leader PIN).
Private Function PurchasedKeyword(ByVal comp As String) As String
    Dim c As String
    c = UCase(comp)
    If InStr(c, "LEADER") > 0 Or InStr(c, "GUIDE PIN") > 0 Then PurchasedKeyword = "PIN": Exit Function
    If InStr(c, "BUSHING") > 0 Then PurchasedKeyword = "BUSHING": Exit Function
    If InStr(c, "RETAIN") > 0 Or InStr(c, "RING") > 0 Or InStr(c, "CIRCLIP") > 0 Then PurchasedKeyword = "RETAIN": Exit Function
    If InStr(c, "STRAP") > 0 Then PurchasedKeyword = "STRAP": Exit Function
    If InStr(c, "INSULAT") > 0 Then PurchasedKeyword = "INSULAT": Exit Function
    If InStr(c, "SLEEVE") > 0 Or InStr(c, "BEARING") > 0 Then PurchasedKeyword = "BEARING": Exit Function
    PurchasedKeyword = ""
End Function

' Find the best price-list row for a BOM description (part number first, then keyword).
' matchKind reports how the row was found: PARTNO / DESCNO / PHRASE / KEYWORD / "".
' KEYWORD matches are weak evidence and must NOT silently price grouped
' assemblies (the J8420 "$25.46 for a latch-lock assembly" bug).
Private Function MatchPurchasedRow(ByVal desc As String, Optional ByVal partNo As String = "", _
                                   Optional ByRef matchKind As String) As Long
    Dim du As String, k As Long, tok As String, pu As String
    matchKind = ""
    du = " " & UCase(desc) & " "
    pu = UCase(Trim(partNo))
    ' 1) exact-ish match on the BOM's part number against the price list part numbers
    If pu <> "" Then
        For k = 1 To PlCount
            If PlPartNo(k) <> "" Then
                If InStr(pu, UCase(PlPartNo(k))) > 0 Or InStr(UCase(PlPartNo(k)), pu) > 0 Then matchKind = "PARTNO": MatchPurchasedRow = k: Exit Function
            End If
        Next k
    End If
    ' 2) price-list part number appearing inside the description text
    For k = 1 To PlCount
        If PlPartNo(k) <> "" Then
            If InStr(du, UCase(PlPartNo(k))) > 0 Then matchKind = "DESCNO": MatchPurchasedRow = k: Exit Function
        End If
    Next k
    ' 3) component phrase (Top Plate / Bottom Plate / exact named hardware).
    For k = 1 To PlCount
        If Trim(PlComp(k)) <> "" Then
            If InStr(du, " " & UCase(Trim(PlComp(k))) & " ") > 0 Then matchKind = "PHRASE": MatchPurchasedRow = k: Exit Function
        End If
    Next k
    ' 4) component keyword (LEADER / BUSHING / RETAIN / STRAP / INSULAT / BEARING).
    '    Guards: grouped assemblies (ASM/ASSEMBLY names) never keyword-match,
    '    and the PIN keyword requires LEADER/GUIDE in the description so
    '    return pins / dowel pins / PLC latch pins can't take the leader-pin price.
    If InStr(du, " ASM ") > 0 Or InStr(du, "_ASM") > 0 Or InStr(du, "-ASM") > 0 Or InStr(du, "ASSEMBLY") > 0 Then Exit Function
    For k = 1 To PlCount
        tok = PurchasedKeyword(PlComp(k))
        If tok <> "" Then
            If tok = "PIN" Then
                If InStr(du, "PIN") > 0 And (InStr(du, "LEADER") > 0 Or InStr(du, "GUIDE") > 0) Then matchKind = "KEYWORD": MatchPurchasedRow = k: Exit Function
            ElseIf InStr(du, tok) > 0 Then
                matchKind = "KEYWORD": MatchPurchasedRow = k: Exit Function
            End If
            If tok = "INSULAT" And InStr(du, "PYROPEL") > 0 Then matchKind = "KEYWORD": MatchPurchasedRow = k: Exit Function
        End If
    Next k
End Function

Private Function IsKnownPurchasedVendor(ByVal manuf As String) As Boolean
    Dim v As String
    v = UCase(Trim(manuf))
    If v = "" Then Exit Function
    If InStr(v, "DME") > 0 Then IsKnownPurchasedVendor = True: Exit Function
    If InStr(v, "PCS") > 0 Then IsKnownPurchasedVendor = True: Exit Function
    If InStr(v, "JACO") > 0 Then IsKnownPurchasedVendor = True: Exit Function
    If InStr(v, "MCMASTER") > 0 Then IsKnownPurchasedVendor = True: Exit Function
    If InStr(v, "MCMASTER-CARR") > 0 Then IsKnownPurchasedVendor = True
End Function

Private Function LooksLikePurchasedPartNo(ByVal partNo As String) As Boolean
    Dim p As String
    p = UCase(Trim(partNo))
    If p = "" Then Exit Function
    If InStr(p, "G1C-") > 0 Then Exit Function
    If InStr(p, "-") > 0 Then LooksLikePurchasedPartNo = True: Exit Function
    If p Like "*[A-Z]*" And p Like "*[0-9]*" Then LooksLikePurchasedPartNo = True
End Function

Private Sub CapturePurchased(ByVal desc As String, ByVal qty As Long, ByVal mat As String, _
                             Optional ByVal tThk As Double = 0, Optional ByVal wWid As Double = 0, _
                             Optional ByVal lLen As Double = 0, Optional ByVal partNo As String = "", _
                             Optional ByVal manuf As String = "", Optional ByVal detNo As String = "", _
                             Optional ByVal purchType As String = "")
    Dim isPurchase As Boolean
    isPurchase = (InStr(UCase(purchType), "PURCHASE") > 0)
    If isPurchase = False Then
        If IsKnownPurchasedVendor(manuf) And LooksLikePurchasedPartNo(partNo) Then isPurchase = True
    End If

    Dim k As Long
    Dim matchKind As String
    k = MatchPurchasedRow(desc, partNo, matchKind)

    ' Gate: if the BOM has an explicit TYPE = Purchase, capture it no matter what
    ' (so every purchased line shows up, even if it is not in the price list).
    ' If there is no TYPE column, fall back to the keyword/part# match as the gate
    ' so material rows are not captured.
    If isPurchase = False Then
        If k = 0 Then Exit Sub
    End If

    PpCount = PpCount + 1
    ReDim Preserve PpDesc(1 To PpCount): ReDim Preserve PpQty(1 To PpCount)
    ReDim Preserve PpComp(1 To PpCount): ReDim Preserve PpVendor(1 To PpCount)
    ReDim Preserve PpPartNo(1 To PpCount): ReDim Preserve PpPrice(1 To PpCount)
    ReDim Preserve PpW(1 To PpCount): ReDim Preserve PpL(1 To PpCount): ReDim Preserve PpT(1 To PpCount)
    ReDim Preserve PpDet(1 To PpCount)

    PpDesc(PpCount) = desc
    PpQty(PpCount) = IIf(qty < 1, 1, qty)
    ' Prefer the BOM's own manufacturer / part number (most accurate).
    PpVendor(PpCount) = IIf(Trim(manuf) <> "", Trim(manuf), IIf(k > 0, PlVendor(k), ""))
    PpPartNo(PpCount) = IIf(Trim(partNo) <> "", Trim(partNo), IIf(k > 0, PlPartNo(k), ""))
    PpComp(PpCount) = IIf(k > 0, PlComp(k), desc)
    PpW(PpCount) = wWid
    PpL(PpCount) = lLen
    PpT(PpCount) = tThk
    PpDet(PpCount) = Trim(detNo)

    ' Price: web lookup (off by default) -> direct part# match in the list ->
    ' the matched row -> 0. Logged so a $0 is easy to diagnose.
    ' KEYWORD matches with no part number are too weak to price: they are what
    ' put $25.46 (Leader Pin) on grouped assemblies in J8420. Those now stay
    ' $0 with a NEEDS PRICE warning instead of silently taking a wrong price.
    Dim p As Double
    p = GetOnlineUnitPrice(PpVendor(PpCount), PpPartNo(PpCount))
    If p <= 0 Then p = LookupListPriceByPartNo(PpPartNo(PpCount))
    If p <= 0 And k > 0 Then

        Dim isCadCapture As Boolean
        isCadCapture = (InStr(UCase$(purchType), "CAD") > 0)

        If matchKind = "KEYWORD" Or (isCadCapture And matchKind <> "PARTNO" And matchKind <> "DESCNO") Then

            LogLine "PRICE WARNING (" & desc & "): weak CAD/keyword price match suppressed. " & _
                    "partNo='" & partNo & "' matched list row '" & PlPartNo(k) & _
                    "' by " & matchKind & ". NEEDS PRICE."

            p = 0#

        Else
            p = PlPrice(k)
        End If

    End If
    If p <= 0 And ENABLE_PYTHON_PRICE_LOOKUP And InStr(UCase(PpVendor(PpCount)), "DME") > 0 Then
        p = LookupDmePriceWithPython(PpPartNo(PpCount))
        If p > 0 Then SavePriceToList PpVendor(PpCount), PpPartNo(PpCount), p
    End If
    PpPrice(PpCount) = p
    LogLine "PRICE " & PpPartNo(PpCount) & " (" & desc & "): listRow#=" & k & " match=" & matchKind & _
            " -> $" & FormatNumberForCsv(p) & "   [price list has " & PlCount & " row(s)]"
End Sub

' Direct price lookup by part number against the loaded price list.
Private Function LookupListPriceByPartNo(ByVal partNo As String) As Double
    LookupListPriceByPartNo = 0
    Dim pu As String, k As Long, plp As String
    pu = UCase(Trim(partNo))
    If pu = "" Then Exit Function
    For k = 1 To PlCount
        plp = UCase(Trim(PlPartNo(k)))
        If plp <> "" Then
            If plp = pu Or InStr(pu, plp) > 0 Or InStr(plp, pu) > 0 Then
                LookupListPriceByPartNo = PlPrice(k)
                Exit Function
            End If
        End If
    Next k
End Function

Private Function FindPriceLookupScript() As String
On Error Resume Next
    FindPriceLookupScript = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim cands(1 To 6) As String
    cands(1) = LOCAL_WORKSPACE_ROOT & "\cms_price_lookup.py"
    cands(2) = CurrentJobFolder & "\cms_price_lookup.py"
    cands(3) = DOWNLOADS_FOLDER & "\Updated\cms_price_lookup.py"
    cands(4) = DOWNLOADS_FOLDER & "\cms_price_lookup.py"
    cands(5) = DOWNLOADS_FOLDER & "\New folder (17)\cms_price_lookup.py"
    cands(6) = "C:\Users\lenovo\Downloads\Updated\cms_price_lookup.py"

    Dim i As Long
    For i = 1 To 6
        If cands(i) <> "" Then
            If fso.FileExists(cands(i)) Then
                FindPriceLookupScript = cands(i)
                Exit Function
            End If
        End If
    Next i

    FindPriceLookupScript = SearchForFileByName(DOWNLOADS_FOLDER, "cms_price_lookup.py", fso, 0)
End Function

Private Function CommandQuote(ByVal s As String) As String
    CommandQuote = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Private Function LookupDmePriceWithPython(ByVal partNo As String) As Double
On Error GoTo ErrHandler
    LookupDmePriceWithPython = 0
    partNo = Trim(partNo)
    If partNo = "" Then Exit Function

    Static cache As Object
    If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
    Dim key As String
    key = UCase(partNo)
    If cache.Exists(key) Then
        LookupDmePriceWithPython = CDbl(cache(key))
        Exit Function
    End If

    Dim scriptPath As String
    scriptPath = FindPriceLookupScript()
    If scriptPath = "" Then
        LogLine "DME Python price lookup skipped: cms_price_lookup.py not found."
        cache(key) = 0#
        Exit Function
    End If

    Dim sh As Object
    Dim ex As Object
    Dim cmd As String
    Set sh = CreateObject("WScript.Shell")
    cmd = CommandQuote(PYTHON_EXE) & " " & CommandQuote(scriptPath) & " --part " & CommandQuote(partNo)
    LogLine "DME Python price lookup: " & partNo
    Set ex = sh.Exec(cmd)

    Dim deadline As Date
    deadline = DateAdd("s", 60, Now)
    Do While ex.Status = 0 And Now < deadline
        WaitMilliseconds 250
    Loop
    If ex.Status = 0 Then
        On Error Resume Next
        ex.Terminate
        On Error GoTo ErrHandler
        LogLine "DME Python price lookup timed out: " & partNo
        cache(key) = 0#
        Exit Function
    End If

    Dim outText As String
    outText = ""
    On Error Resume Next
    outText = ex.StdOut.ReadAll & vbCrLf & ex.StdErr.ReadAll
    On Error GoTo ErrHandler

    Dim price As Double
    price = ExtractPriceFromHtml(outText)
    If price > 0 Then
        LogLine "DME Python price " & partNo & " = $" & FormatNumberForCsv(price)
    Else
        LogLine "DME Python price not found for " & partNo & ": " & Replace(Left(outText, 240), vbCrLf, " ")
    End If
    cache(key) = price
    LookupDmePriceWithPython = price
    Exit Function

ErrHandler:
    LogLine "LookupDmePriceWithPython error (" & partNo & "): " & Err.Description
    LookupDmePriceWithPython = 0
End Function

' ---- Live (dynamic) price lookup ------------------------------------------
' Looks the unit price up on the web for a vendor/part number every run. DME
' parts are looked up on the DME store, then a Bing search; everything else via
' a Bing search "<vendor> <partNo> price". Cached per part number for the run.
' Returns 0 if nothing parseable is found, so the caller falls back to the CSV.
Private Function GetOnlineUnitPrice(ByVal vendor As String, ByVal partNo As String) As Double
On Error GoTo eh
    GetOnlineUnitPrice = 0
    If ENABLE_ONLINE_PRICE_LOOKUP = False Then Exit Function
    Dim pn As String
    pn = Trim(partNo)
    If pn = "" Then Exit Function

    Static cache As Object
    If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
    Dim key As String
    key = UCase(pn)
    If cache.Exists(key) Then
        GetOnlineUnitPrice = cache(key)
        Exit Function
    End If

    Dim v As String, price As Double, html As String
    v = UCase(vendor)
    price = 0

    ' 1) DME parts -> DME store search page, then a Bing fallback.
    If InStr(v, "DME") > 0 Then
        html = HttpGetText("https://store.dme.net/search?keyword=" & UrlEncode(pn))
        price = ExtractPriceFromHtml(html)
        If price <= 0 Then
            html = HttpGetText("https://www.bing.com/search?q=" & UrlEncode("dme.net " & pn & " price"))
            price = ExtractPriceFromHtml(html)
        End If
    End If

    ' 2) Generic -> Bing search "<vendor> <partNo> price".
    If price <= 0 Then
        html = HttpGetText("https://www.bing.com/search?q=" & UrlEncode(Trim(vendor) & " " & pn & " price"))
        price = ExtractPriceFromHtml(html)
    End If

    ' 3) Still nothing -> open the page in the logged-in browser and ask once.
    If price <= 0 And ENABLE_ASSISTED_PRICE_PROMPT Then
        price = AskPriceViaBrowser(vendor, pn)
        If price > 0 Then SavePriceToList vendor, pn, price   ' persist so we never ask again
    End If

    cache(key) = price
    If price > 0 Then
        LogLine "Online price " & vendor & " " & pn & " = $" & FormatNumberForCsv(price)
    Else
        LogLine "Online price " & vendor & " " & pn & " : not found (falling back to list/0)"
    End If
    GetOnlineUnitPrice = price
    Exit Function
eh:
    LogLine "GetOnlineUnitPrice error (" & partNo & "): " & Err.Description
    GetOnlineUnitPrice = 0
End Function

Private Function HttpGetText(ByVal url As String) As String
On Error GoTo eh
    HttpGetText = ""
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHTTPRequest.5.1")
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 10000, 15000
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
    http.SetRequestHeader "Accept", "text/html"
    http.Send
    If http.Status = 200 Then HttpGetText = http.ResponseText
    Exit Function
eh:
    HttpGetText = ""
End Function

' Pull the first plausible price ($nn.nn or nn.nn USD) out of an HTML blob.
Private Function ExtractPriceFromHtml(ByVal html As String) As Double
On Error GoTo eh
    ExtractPriceFromHtml = 0
    If Len(html) = 0 Then Exit Function
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "\$\s*([0-9][0-9,]*\.[0-9]{2})|([0-9][0-9,]*\.[0-9]{2})\s*USD"
    Dim ms As Object, s As String, val As Double, i As Long
    Set ms = re.Execute(html)
    For i = 0 To ms.Count - 1
        s = ms(i).SubMatches(0)
        If s = "" Then s = ms(i).SubMatches(1)
        s = Replace(s, ",", "")
        If IsNumeric(s) Then
            val = CDbl(s)
            If val > 0 And val < 100000 Then
                ExtractPriceFromHtml = val
                Exit Function
            End If
        End If
    Next i
    Exit Function
eh:
    ExtractPriceFromHtml = 0
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, out As String
    out = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = AscW(c)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) _
           Or (code >= 97 And code <= 122) Or c = "-" Or c = "_" Or c = "." Then
            out = out & c
        ElseIf c = " " Then
            out = out & "+"
        Else
            out = out & "%" & Right("0" & Hex(code And 255), 2)
        End If
    Next i
    UrlEncode = out
End Function

' Open the part's page in the default (logged-in) browser and ask for the price.
' For DME parts it opens the DME store search; otherwise a Google search. The
' user reads the price off the page and types it in once.
Private Function AskPriceViaBrowser(ByVal vendor As String, ByVal partNo As String) As Double
On Error GoTo eh
    AskPriceViaBrowser = 0
    Dim url As String
    If InStr(UCase(vendor), "DME") > 0 Then
        url = "https://store.dme.net/search?keyword=" & UrlEncode(partNo)
    Else
        url = "https://www.google.com/search?q=" & UrlEncode(Trim(vendor) & " " & partNo & " price")
    End If

    ' Launch the page in the default browser.
    On Error Resume Next
    Shell "cmd /c start """" """ & url & """", vbNormalFocus
    On Error GoTo eh

    Dim ans As String, p As Double
    ans = InputBox("Look up " & vendor & " part " & partNo & " (the page just opened in your browser)." & vbCrLf & _
                   "Type the UNIT price and press OK." & vbCrLf & _
                   "(Leave blank / Cancel to skip and use 0.)", "CMS - price for " & partNo)
    ans = Trim(Replace(Replace(ans, "$", ""), ",", ""))
    If ans <> "" And IsNumeric(ans) Then
        p = CDbl(ans)
        If p > 0 And p < 1000000 Then AskPriceViaBrowser = p
    End If
    Exit Function
eh:
    AskPriceViaBrowser = 0
End Function

' Persist a captured price into the price-list CSV so it is reused next time.
' Updates the matching PartNumber row in place, or appends a new row.
Private Sub SavePriceToList(ByVal vendor As String, ByVal partNo As String, ByVal price As Double)
On Error GoTo eh
    ' Update in-memory list for the rest of this run.
    Dim i As Long, found As Boolean
    found = False
    For i = 1 To PlCount
        If UCase(Trim(PlPartNo(i))) = UCase(Trim(partNo)) Then
            PlPrice(i) = price: found = True: Exit For
        End If
    Next i
    If Not found Then
        PlCount = PlCount + 1
        If PlCount > UBound(PlComp) Then
            ReDim Preserve PlComp(1 To PlCount + 50): ReDim Preserve PlVendor(1 To PlCount + 50)
            ReDim Preserve PlPartNo(1 To PlCount + 50): ReDim Preserve PlDescr(1 To PlCount + 50)
            ReDim Preserve PlUnit(1 To PlCount + 50): ReDim Preserve PlPrice(1 To PlCount + 50)
        End If
        PlComp(PlCount) = partNo: PlVendor(PlCount) = vendor: PlPartNo(PlCount) = partNo
        PlDescr(PlCount) = "": PlUnit(PlCount) = "EA": PlPrice(PlCount) = price
    End If

    ' Write the change to the CSV file on disk (so it survives to the next run).
    If gPriceListPath = "" Or Dir(gPriceListPath) = "" Then Exit Sub
    Dim fIn As Integer, fOut As Integer, line As String, cols() As String
    Dim tmp As String, wroteRow As Boolean
    tmp = gPriceListPath & ".tmp"
    wroteRow = False
    fIn = FreeFile: Open gPriceListPath For Input As #fIn
    fOut = FreeFile: Open tmp For Output As #fOut
    Do While Not EOF(fIn)
        Line Input #fIn, line
        If Left(Trim(line), 1) = "#" Or Trim(line) = "" Then
            Print #fOut, line
        ElseIf InStr(1, line, "Component", vbTextCompare) > 0 And InStr(1, line, "Price", vbTextCompare) > 0 Then
            Print #fOut, line   ' header
        Else
            cols = Split(line, ",")
            If UBound(cols) >= 5 Then
                If UCase(Trim(cols(2))) = UCase(Trim(partNo)) Then
                    cols(5) = Format(price, "0.00")
                    Print #fOut, Join(cols, ",")
                    wroteRow = True
                Else
                    Print #fOut, line
                End If
            Else
                Print #fOut, line
            End If
        End If
    Loop
    If Not wroteRow Then
        Print #fOut, vendor & " part," & vendor & "," & partNo & ",,EA," & Format(price, "0.00") & ",captured from browser"
    End If
    Close #fIn: Close #fOut
    On Error Resume Next
    Kill gPriceListPath
    Name tmp As gPriceListPath
    On Error GoTo eh
    LogLine "Saved price to list: " & partNo & " = $" & Format(price, "0.00")
    Exit Sub
eh:
    On Error Resume Next
    If fIn > 0 Then Close #fIn
    If fOut > 0 Then Close #fOut
    LogLine "SavePriceToList error (" & partNo & "): " & Err.Description
End Sub

Private Sub WritePurchasedCategoryToSheet(ByVal ws As Object)
On Error GoTo eh
    If PpCount < 1 Or ws Is Nothing Then Exit Sub
    Dim r As Long, hr As Long, rr As Long, i As Long
    r = PURCHASED_QUOTE_START_ROW
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Merge
    ws.Cells(r, 1).value = "PURCHASED COMPONENTS  (DME / McMaster-Carr / Jaco)"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).HorizontalAlignment = -4108
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(84, 130, 53)
    hr = r + 1
    ws.Cells(hr, 1).value = "Component"
    ws.Cells(hr, 2).value = "Vendor"
    ws.Cells(hr, 3).value = "Part Number"
    ws.Cells(hr, 4).value = "Description"
    ws.Cells(hr, 6).value = "QTY"
    ws.Cells(hr, 7).value = "Unit $"
    ws.Cells(hr, 8).value = "Ext $"
    With ws.Range(ws.Cells(hr, 1), ws.Cells(hr, 8))
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
        .HorizontalAlignment = -4108
    End With
    rr = hr + 1
    For i = 1 To PpCount
        ws.Cells(rr, 1).value = PpComp(i)
        ws.Cells(rr, 2).value = PpVendor(i)
        ws.Cells(rr, 3).value = PpPartNo(i)
        ws.Cells(rr, 4).value = PpDesc(i)
        ws.Cells(rr, 6).value = PpQty(i)
        ws.Cells(rr, 7).value = PpPrice(i)
        ws.Cells(rr, 8).Formula = "=F" & rr & "*G" & rr
        ws.Cells(rr, 7).NumberFormat = "$#,##0.00"
        ws.Cells(rr, 8).NumberFormat = "$#,##0.00"
        rr = rr + 1
    Next i
    ws.Cells(rr, 1).value = "Total"
    ws.Cells(rr, 1).Font.Bold = True
    ws.Cells(rr, 8).Formula = "=SUM(H" & (hr + 1) & ":H" & (rr - 1) & ")"
    ws.Cells(rr, 8).NumberFormat = "$#,##0.00"
    ws.Cells(rr, 8).Font.Bold = True
    With ws.Range(ws.Cells(hr, 1), ws.Cells(rr, 8)).Borders
        .LineStyle = 1
        .Weight = 2
    End With
    LogLine "Purchased-components category written to Quote sheet starting row " & PURCHASED_QUOTE_START_ROW
    Exit Sub
eh:
    LogLine "WritePurchasedCategoryToSheet error: " & Err.Description
End Sub

Private Sub WritePurchasedPriceFile()
On Error GoTo eh
    If PpCount < 1 Then Exit Sub
    Dim s As String, i As Long, tot As Double
    s = "Component,Vendor,PartNumber,Description,QTY,UnitPrice,Extended" & vbCrLf
    For i = 1 To PpCount
        s = s & CsvText(PpComp(i)) & "," & CsvText(PpVendor(i)) & "," & CsvText(PpPartNo(i)) & "," & _
            CsvText(PpDesc(i)) & "," & PpQty(i) & "," & FormatNumberForCsv(PpPrice(i)) & "," & _
            FormatNumberForCsv(PpQty(i) * PpPrice(i)) & vbCrLf
        tot = tot + PpQty(i) * PpPrice(i)
    Next i
    s = s & "TOTAL,,,,,," & FormatNumberForCsv(tot) & vbCrLf
    Dim p As String, f As Integer
    p = GetWritableCsvPath(CurrentJobFolder & "\Purchased Components Quote.csv")
    f = FreeFile
    Open p For Output As #f
    Print #f, s;
    Close #f
    LogLine "Wrote purchased components CSV: " & p
    Exit Sub
eh:
    LogLine "WritePurchasedPriceFile error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub ComputePurchasedQuote()
    If PpCount < 1 Then
        If PlCount > 0 Then LogLine "Purchased components: none matched in this BOM."
        Exit Sub
    End If
    Dim i As Long, tot As Double, zero As Long
    For i = 1 To PpCount: tot = tot + PpQty(i) * PpPrice(i): Next i
    For i = 1 To PpCount
        If PpPrice(i) <= 0 Then zero = zero + 1
    Next i
    LogLine "PURCHASED COMPONENTS: " & PpCount & " line(s), extended total $" & FormatNumberForCsv(tot)
    If zero > 0 Then LogLine "  (" & zero & " line(s) priced at $0.00 - set their price in " & PURCHASED_PRICE_FILE & ")"
    WritePurchasedPriceFile
    WritePurchasedComponentsWorkbook
    SendProposalEmail tot
End Sub

' Email a short BOM-pricing summary back to the shop inbox after a quote is built.
' Subject:  BMS - <CustJob> <CNum> / PROPOSAL-BOM PRICING
' Body:     BOM: <priced purchased components>   $<total>   (total shown in red)
' The macro writes the details to a handoff file and lets the Python mailer send it
' (Python already holds the Gmail credentials).
' Pull the customer job number out of a folder name like
' "BMS-848200014-C18608 (Test)" -> "848200014" (first run of >=6 digits).
Private Function ExtractCustJobFromName(ByVal nm As String) As String
    ExtractCustJobFromName = ""
    If nm = "" Then Exit Function
    Dim i As Long, ch As String, run As String
    For i = 1 To Len(nm) + 1
        If i <= Len(nm) Then ch = Mid(nm, i, 1) Else ch = " "
        If ch >= "0" And ch <= "9" Then
            run = run & ch
        Else
            If Len(run) >= 6 Then ExtractCustJobFromName = run: Exit Function
            run = ""
        End If
    Next i
End Function

Private Function CleanSubjectToken(ByVal raw As String) As String
    Dim s As String
    s = Trim(raw)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    CleanSubjectToken = s
End Function

Private Function ProposalCustomerPrefix() As String
    Dim p As String
    p = CleanSubjectToken(CustomerPrefix)
    If p = "" Then p = CleanSubjectToken(CustomerDisplayName)
    If p = "" Then p = CleanSubjectToken(JobBaseName)
    If p = "" Then p = "QUOTE"
    ProposalCustomerPrefix = p
End Function

Private Function ProposalJobForSubject(ByVal prefix As String, ByVal custJob As String) As String
    Dim j As String
    j = CleanSubjectToken(custJob)
    If j = "" Then j = ExtractCustJobFromName(gExactJobFolderName)
    If j = "" Then j = ExtractCustJobFromName(CurrentJobFolder)
    If prefix <> "" Then
        If UCase(Left(j, Len(prefix) + 1)) = UCase(prefix & "-") Then
            j = Mid(j, Len(prefix) + 2)
        End If
    End If
    ProposalJobForSubject = j
End Function
Private Sub SendProposalEmail(ByVal total As Double)
On Error GoTo eh
    If PpCount < 1 Then
        gEmailStatus = "NOT sent - no purchased components matched in the BOM"
        LogLine "Proposal email skipped: no purchased components matched (PpCount=0)."
        Exit Sub
    End If

    ' Email gate: PROPOSAL_EMAIL_MODE = OFF | PROMPT | AUTO. This stops the
    ' unwanted automatic CDO proposal emails (J8420) unless explicitly allowed.
    If UCase(PROPOSAL_EMAIL_MODE) = "OFF" Then
        gEmailStatus = "NOT sent - PROPOSAL_EMAIL_MODE=OFF (preview written to cms_proposal.txt)"
        LogLine "Proposal email skipped: PROPOSAL_EMAIL_MODE=OFF."
        WriteProposalPreviewFile total, False
        Exit Sub
    ElseIf UCase(PROPOSAL_EMAIL_MODE) = "PROMPT" Then
        If MsgBox("Send the PROPOSAL-BOM PRICING email now?" & vbCrLf & vbCrLf & _
                  "Purchased lines: " & PpCount & "   Total: $" & FormatNumberForCsv(total), _
                  vbYesNo + vbQuestion, "CMS Proposal Email") <> vbYes Then
            gEmailStatus = "NOT sent - user declined at prompt (preview written to cms_proposal.txt)"
            LogLine "Proposal email skipped: user declined at PROMPT."
            WriteProposalPreviewFile total, False
            Exit Sub
        End If
    End If

    ' Build the BOM line: "Leader Pin (107), Guide Bushing (108), ..."
    Dim items As String, i As Long
    items = ""
    For i = 1 To PpCount
        If items <> "" Then items = items & ", "
        items = items & PpComp(i)
        If PpDet(i) <> "" Then items = items & " (" & PpDet(i) & ")"
    Next i

    ' Customer/job naming for the subject comes from the email/file handoff.
    ' Do not call it BMS unless the email/file actually said BMS.
    Dim prefix As String, cj As String
    prefix = ProposalCustomerPrefix()
    cj = ProposalJobForSubject(prefix, CustomerJobNumber)

    Dim cnum As String, subj As String, totalStr As String
    cnum = Replace(AssignedQuoteNumber, "-", "")          ' C18472 (no hyphen, like the sample)
    totalStr = FormatNumberForCsv(total)
    subj = prefix & IIf(cj <> "", " - " & cj, "") & " " & cnum & " / PROPOSAL-BOM PRICING"

    Dim html As String
    html = "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:14px'>" & _
           "<p>" & subj & "</p>" & _
           "<p>BOM: " & items & " &nbsp;&nbsp; " & _
           "<span style='color:#d00000'><b>$" & totalStr & "</b></span></p>" & _
           "</body></html>"

    ' Try to send DIRECTLY via CDO over Gmail SMTP (no Python / PATH dependency).
    Dim cdoOk As Boolean
    cdoOk = SendViaCdo(GMAIL_ADDRESS, subj, html)

    ' Write the handoff file WITH a Sent flag. If CDO failed (some shop PCs block
    ' CDO/SMTP), the picker reads this and sends it with Python after the job.
    Dim fso As Object, p As String, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(LOCAL_WORKSPACE_ROOT) Then fso.CreateFolder LOCAL_WORKSPACE_ROOT
    p = LOCAL_WORKSPACE_ROOT & "\cms_proposal.txt"
    Set ts = fso.CreateTextFile(p, True)
    ts.WriteLine "To=" & GMAIL_ADDRESS
    ts.WriteLine "CustomerPrefix=" & CustomerPrefix
    ts.WriteLine "CustomerName=" & CustomerDisplayName
    ts.WriteLine "CustJob=" & CustomerJobNumber
    ts.WriteLine "CNum=" & AssignedQuoteNumber
    ts.WriteLine "Items=" & items
    ts.WriteLine "Total=" & totalStr
    ts.WriteLine "Sent=" & IIf(cdoOk, "1", "0")
    ts.Close

    If cdoOk Then
        gEmailStatus = "SENT via CDO to " & GMAIL_ADDRESS
        LogLine "Proposal email SENT (CDO): " & subj & " | " & items & " | $" & totalStr
    Else
        gEmailStatus = "queued (CDO failed) - picker will send via Python"
        LogLine "CDO send failed; wrote " & p & " (Sent=0) for the picker to send via Python. " & subj
    End If
    Exit Sub
eh:
    LogLine "SendProposalEmail error: " & Err.Description
End Sub

' Write the proposal handoff/preview file without sending (used when the
' email gate is OFF or the user declines the prompt).
Private Sub WriteProposalPreviewFile(ByVal total As Double, ByVal wasSent As Boolean)
    On Error Resume Next
    Dim items As String, i As Long
    For i = 1 To PpCount
        If items <> "" Then items = items & ", "
        items = items & PpComp(i)
        If PpDet(i) <> "" Then items = items & " (" & PpDet(i) & ")"
    Next i
    Dim fso As Object, p As String, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(LOCAL_WORKSPACE_ROOT) Then fso.CreateFolder LOCAL_WORKSPACE_ROOT
    p = LOCAL_WORKSPACE_ROOT & "\cms_proposal.txt"
    Set ts = fso.CreateTextFile(p, True)
    ts.WriteLine "To=" & GMAIL_ADDRESS
    ts.WriteLine "CustomerPrefix=" & CustomerPrefix
    ts.WriteLine "CustomerName=" & CustomerDisplayName
    ts.WriteLine "CustJob=" & CustomerJobNumber
    ts.WriteLine "CNum=" & AssignedQuoteNumber
    ts.WriteLine "Items=" & items
    ts.WriteLine "Total=" & FormatNumberForCsv(total)
    ts.WriteLine "Sent=" & IIf(wasSent, "1", "0")
    ts.WriteLine "GatedBy=PROPOSAL_EMAIL_MODE:" & PROPOSAL_EMAIL_MODE
    ts.Close
End Sub

' Read the Gmail app password from the webapp Settings JSON (never hardcoded).
Private Function JsonStringField(ByVal jsonText As String, ByVal fieldName As String) As String
    JsonStringField = ""
    On Error Resume Next
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """" & fieldName & """\s*:\s*""([^""]*)"""
    Set matches = re.Execute(jsonText)
    If matches.Count > 0 Then JsonStringField = matches(0).SubMatches(0)
End Function

Private Function GmailAppPassword() As String
    GmailAppPassword = ""
    On Error Resume Next
    Dim fso As Object, ts As Object, jsonText As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(EMAIL_CREDENTIALS_FILE) Then Exit Function
    Set ts = fso.OpenTextFile(EMAIL_CREDENTIALS_FILE, 1)
    jsonText = ts.ReadAll
    ts.Close
    GmailAppPassword = JsonStringField(jsonText, "smtp_password")
    If GmailAppPassword = "" Then GmailAppPassword = JsonStringField(jsonText, "imap_password")
End Function

' Send an HTML email through Gmail SMTP using CDO. Tries SSL 465, then STARTTLS 587.
Private Function SendViaCdo(ByVal toAddr As String, ByVal subj As String, ByVal htmlBody As String) As Boolean
    SendViaCdo = False
    Dim appPw As String
    appPw = GmailAppPassword()
    If appPw = "" Then
        LogLine "CDO send skipped: no email credentials at " & EMAIL_CREDENTIALS_FILE & _
                " (open http://127.0.0.1:8000/settings and save your Gmail app password)."
        Exit Function
    End If

    Dim sch As String
    sch = "http://schemas.microsoft.com/cdo/configuration/"

    Dim ports As Variant, ssls As Variant, pi As Long
    ports = Array(465, 587)
    ssls = Array(True, False)   ' 465 implicit SSL; 587 STARTTLS

    For pi = 0 To 1
        On Error Resume Next
        Dim msg As Object
        Set msg = CreateObject("CDO.Message")
        With msg.Configuration.Fields
            .Item(sch & "sendusing") = 2
            .Item(sch & "smtpserver") = "smtp.gmail.com"
            .Item(sch & "smtpserverport") = ports(pi)
            .Item(sch & "smtpusessl") = ssls(pi)
            .Item(sch & "smtpauthenticate") = 1
            .Item(sch & "sendusername") = GMAIL_ADDRESS
            .Item(sch & "sendpassword") = appPw
            .Item(sch & "smtpconnectiontimeout") = 30
            .Update
        End With
        msg.To = toAddr
        msg.From = GMAIL_ADDRESS
        msg.Subject = subj
        msg.HTMLBody = htmlBody
        Err.Clear
        msg.Send
        If Err.Number = 0 Then
            SendViaCdo = True
            Set msg = Nothing
            On Error GoTo 0
            Exit Function
        End If
        Set msg = Nothing
        Err.Clear
        On Error GoTo 0
    Next pi
End Function

' Locate cms_gmail_search.py in the usual spots.
Private Function FindGmailScript() As String
    FindGmailScript = ""
    Dim fso As Object, cands(4) As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    cands(0) = LOCAL_WORKSPACE_ROOT & "\cms_gmail_search.py"
    cands(1) = TRUSTED_FOLDER & "\cms_gmail_search.py"
    cands(2) = DOWNLOADS_FOLDER & "\cms_gmail_search.py"
    cands(3) = "C:\Users\lenovo\Downloads\New folder (17)\cms_gmail_search.py"
    cands(4) = "C:\Users\lenovo\Downloads\cms_gmail_search.py"
    For i = 0 To 4
        If fso.FileExists(cands(i)) Then FindGmailScript = cands(i): Exit Function
    Next i
End Function

' Build a standalone Excel workbook listing the hardware we buy for this job
' (pulled from the BOM and priced from the CSV). Saved in the job folder and,
' when reachable, copied to the proposals folder next to the customer quote.
Private Sub WritePurchasedComponentsWorkbook()
On Error GoTo eh
    If PpCount < 1 Then Exit Sub

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    ws.Name = "Purchased Components"
    On Error GoTo eh

    ws.Cells(1, 1).value = "PURCHASED COMPONENTS"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).value = "Quote #:"
    ws.Cells(2, 2).value = AssignedQuoteNumber
    ws.Cells(3, 1).value = "Customer Job #:"
    ws.Cells(3, 2).value = CustomerJobNumber
    ws.Cells(4, 1).value = "Date:"
    ws.Cells(4, 2).value = Format(Now, "m/d/yyyy")

    Dim hdr As Long
    hdr = 6
    ws.Cells(hdr, 1).value = "Component"
    ws.Cells(hdr, 2).value = "Vendor"
    ws.Cells(hdr, 3).value = "Part Number"
    ws.Cells(hdr, 4).value = "Description"
    ws.Cells(hdr, 5).value = "QTY"
    ws.Cells(hdr, 6).value = "Unit Price"
    ws.Cells(hdr, 7).value = "Extended"
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 7)).Font.Bold = True

    Dim i As Long, rr As Long, tot As Double
    rr = hdr + 1
    For i = 1 To PpCount
        ws.Cells(rr, 1).value = PpComp(i)
        ws.Cells(rr, 2).value = PpVendor(i)
        ws.Cells(rr, 3).value = PpPartNo(i)
        ws.Cells(rr, 4).value = PpDesc(i)
        ws.Cells(rr, 5).value = PpQty(i)
        ws.Cells(rr, 6).value = PpPrice(i)
        ws.Cells(rr, 7).value = PpQty(i) * PpPrice(i)
        tot = tot + PpQty(i) * PpPrice(i)
        rr = rr + 1
    Next i
    ws.Cells(rr, 6).value = "TOTAL"
    ws.Cells(rr, 6).Font.Bold = True
    ws.Cells(rr, 7).value = tot
    ws.Cells(rr, 7).Font.Bold = True

    On Error Resume Next
    ws.Range(ws.Cells(hdr + 1, 6), ws.Cells(rr, 7)).NumberFormat = "$#,##0.00"
    ws.Columns("A:G").AutoFit
    On Error GoTo eh

    Dim baseName As String, outPath As String
    baseName = "Purchased Components " & IIf(JobBaseName <> "", JobBaseName, CurrentJobNumber)
    outPath = CurrentJobFolder & "\" & baseName & ".xlsx"

    On Error Resume Next
    wb.SaveAs outPath, 51                      ' xlOpenXMLWorkbook (.xlsx)
    If Err.Number <> 0 Then
        Err.Clear
        outPath = CurrentJobFolder & "\" & baseName & ".xls"
        wb.SaveAs outPath, 56                  ' xlExcel8 (.xls) fallback for old Excel
    End If
    On Error GoTo eh

    wb.Close True
    xl.Quit
    Set ws = Nothing: Set wb = Nothing: Set xl = Nothing
    LogLine "Wrote purchased components workbook: " & outPath
    Exit Sub
eh:
    LogLine "WritePurchasedComponentsWorkbook error: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
End Sub

' =====================================================================
'  ORIENTATION SUBSYSTEM ported verbatim from gemini1.bas
'  (top from TCP/BCP centers; front from holder long-side + pot/holder
'   depth so the pot blocks sit closer to the front than the holders).
' =====================================================================

Private Sub BuildFrontOrientIndexCollections(ByRef holderIndexes As Collection, _
                                             ByRef potIndexes As Collection)
On Error GoTo ErrHandler

    Set holderIndexes = New Collection
    Set potIndexes = New Collection

    ' First use geometry classification.
    ' This is essential for generic imported XT files where components are named
    ' Part-1, Part-2, etc. and BOM matching cannot find ID HOLDER / OD HOLDER.
    AddUniqueCadIndexToCollection holderIndexes, gIdxIDH
    AddUniqueCadIndexToCollection holderIndexes, gIdxODH
    AddUniqueCadIndexToCollection potIndexes, gIdxIDP
    AddUniqueCadIndexToCollection potIndexes, gIdxODP

    ' Then add BOM/name-based matches if available.
    AddUniqueCadIndexToCollection holderIndexes, _
        FindCadIndexForOrientationQuoteOrKeys("ID HOLDER", ID_HOLDER_KEYS)

    AddUniqueCadIndexToCollection holderIndexes, _
        FindCadIndexForOrientationQuoteOrKeys("OD HOLDER", OD_HOLDER_KEYS)

    AddUniqueCadIndexToCollection potIndexes, _
        FindCadIndexForOrientationQuoteOrKeys("ID POT BLOCK", _
            "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT")

    AddUniqueCadIndexToCollection potIndexes, _
        FindCadIndexForOrientationQuoteOrKeys("OD POT BLOCK", _
            "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT")

    LogLine "Front orientation index collections: holders=" & holderIndexes.Count & _
            " pots=" & potIndexes.Count

    Exit Sub

ErrHandler:
    LogLine "BuildFrontOrientIndexCollections error: " & Err.Description
    Set holderIndexes = New Collection
    Set potIndexes = New Collection
End Sub

Private Function OrientTcpTopFromCenters(ByVal model As Object, _
                                         ByVal tcpX As Double, _
                                         ByVal tcpY As Double, _
                                         ByVal tcpZ As Double, _
                                         ByVal bcpX As Double, _
                                         ByVal bcpY As Double, _
                                         ByVal bcpZ As Double, _
                                         ByVal sourceLabel As String) As Boolean
On Error GoTo ErrHandler

    OrientTcpTopFromCenters = False

    If model Is Nothing Then Exit Function

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = tcpX - bcpX
    dy = tcpY - bcpY
    dz = tcpZ - bcpZ

    Dim aX As Double
    Dim aY As Double
    Dim aZ As Double

    aX = Abs(dx)
    aY = Abs(dy)
    aZ = Abs(dz)

    LogLine sourceLabel & " TCP-BCP separation: " & _
            "X=" & FormatNumberForCsv(dx) & _
            " Y=" & FormatNumberForCsv(dy) & _
            " Z=" & FormatNumberForCsv(dz)

    Dim axisName As String
    Dim tcpHigh As Boolean

    If aY >= aX And aY >= aZ Then
        axisName = "Y"
        tcpHigh = (dy >= 0#)
    ElseIf aZ >= aX And aZ >= aY Then
        axisName = "Z"
        tcpHigh = (dz >= 0#)
    Else
        axisName = "X"
        tcpHigh = (dx >= 0#)
    End If

    Dim viewName As String
    Dim viewId As Long

    ' Correct SolidWorks default standard-view mapping:
    '   Y stack -> *Top / *Bottom
    '   Z stack -> *Front / *Back
    '   X stack -> *Right / *Left

    Select Case axisName

        Case "Y"
            If tcpHigh Then
                viewName = "*Top"
                viewId = 5
            Else
                viewName = "*Bottom"
                viewId = 6
            End If

        Case "Z"
            If tcpHigh Then
                viewName = "*Front"
                viewId = 1
            Else
                viewName = "*Back"
                viewId = 2
            End If

        Case "X"
            If tcpHigh Then
                viewName = "*Right"
                viewId = 4
            Else
                viewName = "*Left"
                viewId = 3
            End If

    End Select

    If viewName = "" Then
        LogLine sourceLabel & " orientation failed: no standard view selected."
        Exit Function
    End If

    model.ShowNamedView2 viewName, viewId

    LogLine sourceLabel & " selected " & viewName & _
            ". Stack axis=" & axisName & _
            ", TCP at high end=" & CStr(tcpHigh)

    RotateViewZSteps model, CMS_TOP_ROTATE_Z_STEPS
    StabilizeActiveView model, 100

    OrientTcpTopFromCenters = True
    Exit Function

ErrHandler:
    LogLine "OrientTcpTopFromCenters error: " & Err.Description
    OrientTcpTopFromCenters = False
End Function

Private Function PersistCurrentViewAsStandardTop(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    PersistCurrentViewAsStandardTop = False

    If model Is Nothing Then Exit Function

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    EnsureSwHidden

    Dim ok As Boolean
    ok = False

    On Error Resume Next

    Err.Clear
    ok = CBool(model.Extension.UpdateStandardViews("*Top", 5))
    If Err.Number <> 0 Then
        Err.Clear
        ok = False
    End If

    If ok = False Then
        Err.Clear
        ok = CBool(model.UpdateStandardViews("*Top", 5))
        If Err.Number <> 0 Then
            Err.Clear
            ok = False
        End If
    End If

    On Error GoTo ErrHandler

    If ok Then
        PersistCurrentViewAsStandardTop = True
        LogLine "Standard views REDEFINED: current TCP/top-side orientation assigned to *Top (ViewId 5)."
    Else
        LogLine "WARNING: UpdateStandardViews(*Top,5) did not succeed; orientation still carried by CMS_TOP named view."
    End If

    On Error Resume Next
    model.ForceRebuild3 False
    model.ViewZoomtofit2
    On Error GoTo ErrHandler

    Exit Function

ErrHandler:
    LogLine "PersistCurrentViewAsStandardTop error: " & Err.Description
    PersistCurrentViewAsStandardTop = False
End Function

Private Function DefineStandardFrontFromHolderAndPotCom(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    DefineStandardFrontFromHolderAndPotCom = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim holderIndexes As Collection
    Dim potIndexes As Collection

    BuildFrontOrientIndexCollections holderIndexes, potIndexes

    If holderIndexes Is Nothing Or holderIndexes.count = 0 Then
        LogLine "Front definition skipped: no holder CAD indexes found."
        Exit Function
    End If

    Dim holderIdx As Long
    holderIdx = PickLargestCadIndexFromCollection(holderIndexes)

    If holderIdx <= 0 Or holderIdx > PartCount Then
        LogLine "Front definition skipped: invalid holder index."
        Exit Function
    End If

    ' Step 1: after top is correct, start by looking at SolidWorks *Front.
    model.ShowNamedView2 "*Front", 1
    StabilizeActiveView model, 100

    Dim candidateViewName As String
    Dim candidateViewId As Long
    Dim oppositeViewName As String
    Dim oppositeViewId As Long
    Dim depthAxis As String

    candidateViewName = "*Front"
    candidateViewId = 1
    oppositeViewName = "*Back"
    oppositeViewId = 2
    depthAxis = "Z"

    LogLine "Front definition: starting from SolidWorks *Front, depth axis=Z."

    ' Step 2: if the holder long side is going into the screen,
    ' use the right face as the new front candidate.
    Dim holderLongIntoScreen As Boolean
    Dim gotLongTest As Boolean

    gotLongTest = IsHolderLongSideIntoCurrentView(model, holderIdx, holderLongIntoScreen)

    If gotLongTest Then

        If holderLongIntoScreen Then

            LogLine "Front definition: holder long side is perpendicular/into-screen from *Front. Trying *Right as front."

            model.ShowNamedView2 "*Right", 4
            StabilizeActiveView model, 100

            candidateViewName = "*Right"
            candidateViewId = 4
            oppositeViewName = "*Left"
            oppositeViewId = 3
            depthAxis = "X"

        Else

            LogLine "Front definition: holder long side is visible from *Front. Keeping *Front as front candidate."

        End If

    Else

        LogLine "Front definition: could not test holder long-side visibility. Keeping *Front as front candidate."

    End If

    ' Step 3:
    ' Force the pot blocks to be closer to the active front view than the holders.
    ' This uses the ACTIVE VIEW depth direction, not a guessed X/Z sign.
    Dim flippedForPots As Boolean
    flippedForPots = False

    If POT_BLOCKS_MUST_BE_FRONT_OF_HOLDERS Then

        If EnsurePotBlocksCloserThanHoldersInActiveView( _
                model, _
                holderIndexes, _
                potIndexes, _
                oppositeViewName, _
                oppositeViewId, _
                flippedForPots) Then

            If flippedForPots Then
                LogLine "Front definition: flipped to opposite face so pot blocks are closer to front."

                candidateViewName = oppositeViewName
                candidateViewId = oppositeViewId
            Else
                LogLine "Front definition: pot blocks are already closer to front."
            End If

        Else

            LogLine "WARNING: Could not verify pot blocks are closer to front than holders."

        End If

    End If

    ' Step 4: whatever view is active now becomes SolidWorks *Front.
    If PersistCurrentViewAsStandardFront(model) Then

        model.ShowNamedView2 "*Front", 1
        StabilizeActiveView model, 100

        ' Final safety check:
        ' After SolidWorks standard views are redefined, verify *Front still has
        ' the pot blocks closer than the holders. If not, flip *Back and save that
        ' as the new *Front.
        If POT_BLOCKS_MUST_BE_FRONT_OF_HOLDERS Then
            If EnforcePotBlocksCloserAfterFrontPersist(model, holderIndexes, potIndexes) = False Then
                LogLine "WARNING: Final *Front verification failed. Pot blocks may still be behind holders."
            End If
        End If

        ' Save a named corrected front view so DXF/JPG export can preserve it.
        model.ShowNamedView2 "*Front", 1
        StabilizeActiveView model, 50
        SaveCurrentViewAsNamed model, CMS_FRONT_VIEW_NAME
        LogLine "CMS_FRONT named view saved from corrected BMS *Front."

        LogLine "Front definition complete. Current orientation persisted as SolidWorks *Front."
        DefineStandardFrontFromHolderAndPotCom = True

    Else

        LogLine "Front definition failed: could not persist current view as SolidWorks *Front."
        DefineStandardFrontFromHolderAndPotCom = False

    End If

    Exit Function

ErrHandler:
    LogLine "DefineStandardFrontFromHolderAndPotCom error: " & Err.Description
    DefineStandardFrontFromHolderAndPotCom = False
End Function

Private Sub SaveCurrentViewAsNamed(ByVal model As Object, ByVal viewName As String)
On Error Resume Next

    If model Is Nothing Then Exit Sub
    If Trim(viewName) = "" Then Exit Sub

    model.DeleteNamedView viewName
    Err.Clear

    model.NameView viewName
    Err.Clear
End Sub

Private Function PersistCurrentViewAsStandardFront(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    PersistCurrentViewAsStandardFront = False

    If model Is Nothing Then Exit Function

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    EnsureSwHidden

    Dim ok As Boolean
    ok = False

    On Error Resume Next

    Err.Clear
    ok = CBool(model.Extension.UpdateStandardViews("*Front", 1))
    If Err.Number <> 0 Then
        Err.Clear
        ok = False
    End If

    If ok = False Then
        Err.Clear
        ok = CBool(model.UpdateStandardViews("*Front", 1))
        If Err.Number <> 0 Then
            Err.Clear
            ok = False
        End If
    End If

    On Error GoTo ErrHandler

    If ok Then
        PersistCurrentViewAsStandardFront = True
        LogLine "Standard views REDEFINED: current orientation assigned to *Front, ViewId 1."
    Else
        LogLine "WARNING: UpdateStandardViews(*Front,1) did not succeed."
    End If

    On Error Resume Next
    model.ForceRebuild3 False
    model.ViewZoomtofit2
    On Error GoTo ErrHandler

    Exit Function

ErrHandler:
    LogLine "PersistCurrentViewAsStandardFront error: " & Err.Description
    PersistCurrentViewAsStandardFront = False
End Function

Private Function PickLargestCadIndexFromCollection(ByVal col As Collection) As Long
On Error GoTo ErrHandler

    PickLargestCadIndexFromCollection = 0

    If col Is Nothing Then Exit Function
    If col.count = 0 Then Exit Function

    Dim i As Long
    Dim cadIdx As Long
    Dim bestIdx As Long
    Dim bestVol As Double

    bestIdx = 0
    bestVol = -1#

    For i = 1 To col.count

        cadIdx = CLng(col(i))

        If cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).BBoxVolume > bestVol Then
                bestVol = parts(cadIdx).BBoxVolume
                bestIdx = cadIdx
            End If
        End If

    Next i

    PickLargestCadIndexFromCollection = bestIdx
    Exit Function

ErrHandler:
    PickLargestCadIndexFromCollection = 0
End Function

Private Function IsHolderLongSideIntoCurrentView(ByVal model As Object, _
                                                 ByVal holderIdx As Long, _
                                                 ByRef longIntoScreen As Boolean) As Boolean
On Error GoTo ErrHandler

    IsHolderLongSideIntoCurrentView = False
    longIntoScreen = False

    If model Is Nothing Then Exit Function
    If holderIdx <= 0 Or holderIdx > PartCount Then Exit Function

    Dim swComp As Object
    Set swComp = FindAssemblyComponentByName(model, parts(holderIdx).componentName)

    If swComp Is Nothing Then
        LogLine "Holder long-side test failed: component not found: " & parts(holderIdx).componentName
        Exit Function
    End If

    Dim viewW As Double
    Dim viewH As Double

    If TryGetComponentViewWidthHeightInches(model, swComp, viewW, viewH) = False Then
        LogLine "Holder long-side test failed: could not get projected holder size."
        Exit Function
    End If

    Dim projectedLong As Double
    projectedLong = viewW
    If viewH > projectedLong Then projectedLong = viewH

    Dim actualLong As Double
    actualLong = parts(holderIdx).Length
    If actualLong <= 0# Then
        actualLong = parts(holderIdx).BoxDx
        If parts(holderIdx).BoxDy > actualLong Then actualLong = parts(holderIdx).BoxDy
        If parts(holderIdx).BoxDz > actualLong Then actualLong = parts(holderIdx).BoxDz
    End If

    If actualLong <= 0# Then Exit Function

    longIntoScreen = (projectedLong < actualLong * HOLDER_LONG_SIDE_VISIBLE_RATIO)

    LogLine "Holder long-side test:"
    LogLine "  holder=" & parts(holderIdx).componentName
    LogLine "  actual long=" & FormatNumberForCsv(actualLong)
    LogLine "  projected W/H=" & FormatNumberForCsv(viewW) & "/" & FormatNumberForCsv(viewH)
    LogLine "  projected long=" & FormatNumberForCsv(projectedLong)
    LogLine "  long side into screen=" & CStr(longIntoScreen)

    IsHolderLongSideIntoCurrentView = True
    Exit Function

ErrHandler:
    LogLine "IsHolderLongSideIntoCurrentView error: " & Err.Description
    IsHolderLongSideIntoCurrentView = False
End Function

Private Function TryGetComponentViewWidthHeightInches(ByVal model As Object, _
                                                      ByVal swComp As Object, _
                                                      ByRef viewWIn As Double, _
                                                      ByRef viewHIn As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentViewWidthHeightInches = False

    viewWIn = 0#
    viewHIn = 0#

    If model Is Nothing Then Exit Function
    If swComp Is Nothing Then Exit Function

    Dim vBox As Variant

    On Error Resume Next
    vBox = swComp.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    Dim mView As Variant
    mView = swView.Orientation3.ArrayData

    If IsEmpty(mView) Then Exit Function
    If IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    Dim xs(0 To 7) As Double
    Dim ys(0 To 7) As Double
    Dim zs(0 To 7) As Double

    xs(0) = CDbl(vBox(0)): ys(0) = CDbl(vBox(1)): zs(0) = CDbl(vBox(2))
    xs(1) = CDbl(vBox(3)): ys(1) = CDbl(vBox(1)): zs(1) = CDbl(vBox(2))
    xs(2) = CDbl(vBox(0)): ys(2) = CDbl(vBox(4)): zs(2) = CDbl(vBox(2))
    xs(3) = CDbl(vBox(3)): ys(3) = CDbl(vBox(4)): zs(3) = CDbl(vBox(2))

    xs(4) = CDbl(vBox(0)): ys(4) = CDbl(vBox(1)): zs(4) = CDbl(vBox(5))
    xs(5) = CDbl(vBox(3)): ys(5) = CDbl(vBox(1)): zs(5) = CDbl(vBox(5))
    xs(6) = CDbl(vBox(0)): ys(6) = CDbl(vBox(4)): zs(6) = CDbl(vBox(5))
    xs(7) = CDbl(vBox(3)): ys(7) = CDbl(vBox(4)): zs(7) = CDbl(vBox(5))

    Dim firstPoint As Boolean
    firstPoint = True

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double

    Dim i As Long

    For i = 0 To 7

        Dim vX As Double
        Dim vY As Double

        vX = (xs(i) * CDbl(mView(0))) + _
             (ys(i) * CDbl(mView(3))) + _
             (zs(i) * CDbl(mView(6)))

        vY = (xs(i) * CDbl(mView(1))) + _
             (ys(i) * CDbl(mView(4))) + _
             (zs(i) * CDbl(mView(7)))

        If firstPoint Then
            minX = vX
            maxX = vX
            minY = vY
            maxY = vY
            firstPoint = False
        Else
            If vX < minX Then minX = vX
            If vX > maxX Then maxX = vX
            If vY < minY Then minY = vY
            If vY > maxY Then maxY = vY
        End If

    Next i

    viewWIn = Abs(maxX - minX) * INCHES_PER_METER
    viewHIn = Abs(maxY - minY) * INCHES_PER_METER

    TryGetComponentViewWidthHeightInches = (viewWIn > 0# And viewHIn > 0#)
    Exit Function

ErrHandler:
    TryGetComponentViewWidthHeightInches = False
End Function

' Same as width/height, plus into-screen depth of the component AABB (inches).
Private Function TryGetComponentViewWidthHeightDepthInches(ByVal model As Object, _
                                                           ByVal swComp As Object, _
                                                           ByRef viewWIn As Double, _
                                                           ByRef viewHIn As Double, _
                                                           ByRef viewDIn As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentViewWidthHeightDepthInches = False
    viewWIn = 0#: viewHIn = 0#: viewDIn = 0#

    If model Is Nothing Then Exit Function
    If swComp Is Nothing Then Exit Function

    Dim vBox As Variant
    On Error Resume Next
    vBox = swComp.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView
    If swView Is Nothing Then Exit Function

    Dim mView As Variant
    mView = swView.Orientation3.ArrayData
    If IsEmpty(mView) Then Exit Function
    If IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    Dim xs(0 To 7) As Double, ys(0 To 7) As Double, zs(0 To 7) As Double
    xs(0) = CDbl(vBox(0)): ys(0) = CDbl(vBox(1)): zs(0) = CDbl(vBox(2))
    xs(1) = CDbl(vBox(3)): ys(1) = CDbl(vBox(1)): zs(1) = CDbl(vBox(2))
    xs(2) = CDbl(vBox(0)): ys(2) = CDbl(vBox(4)): zs(2) = CDbl(vBox(2))
    xs(3) = CDbl(vBox(3)): ys(3) = CDbl(vBox(4)): zs(3) = CDbl(vBox(2))
    xs(4) = CDbl(vBox(0)): ys(4) = CDbl(vBox(1)): zs(4) = CDbl(vBox(5))
    xs(5) = CDbl(vBox(3)): ys(5) = CDbl(vBox(1)): zs(5) = CDbl(vBox(5))
    xs(6) = CDbl(vBox(0)): ys(6) = CDbl(vBox(4)): zs(6) = CDbl(vBox(5))
    xs(7) = CDbl(vBox(3)): ys(7) = CDbl(vBox(4)): zs(7) = CDbl(vBox(5))

    Dim firstPoint As Boolean
    firstPoint = True
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim minZ As Double, maxZ As Double
    Dim i As Long, vX As Double, vY As Double, vZ As Double

    For i = 0 To 7
        vX = (xs(i) * CDbl(mView(0))) + (ys(i) * CDbl(mView(3))) + (zs(i) * CDbl(mView(6)))
        vY = (xs(i) * CDbl(mView(1))) + (ys(i) * CDbl(mView(4))) + (zs(i) * CDbl(mView(7)))
        vZ = (xs(i) * CDbl(mView(2))) + (ys(i) * CDbl(mView(5))) + (zs(i) * CDbl(mView(8)))

        If firstPoint Then
            minX = vX: maxX = vX: minY = vY: maxY = vY: minZ = vZ: maxZ = vZ
            firstPoint = False
        Else
            If vX < minX Then minX = vX
            If vX > maxX Then maxX = vX
            If vY < minY Then minY = vY
            If vY > maxY Then maxY = vY
            If vZ < minZ Then minZ = vZ
            If vZ > maxZ Then maxZ = vZ
        End If
    Next i

    viewWIn = Abs(maxX - minX) * INCHES_PER_METER
    viewHIn = Abs(maxY - minY) * INCHES_PER_METER
    viewDIn = Abs(maxZ - minZ) * INCHES_PER_METER

    TryGetComponentViewWidthHeightDepthInches = (viewWIn > 0# And viewHIn > 0# And viewDIn > 0#)
    Exit Function

ErrHandler:
    TryGetComponentViewWidthHeightDepthInches = False
End Function

Private Function EnsurePotBlocksCloserThanHoldersInActiveView( _
    ByVal model As Object, _
    ByVal holderIndexes As Collection, _
    ByVal potIndexes As Collection, _
    ByVal oppositeViewName As String, _
    ByVal oppositeViewId As Long, _
    ByRef flippedToOpposite As Boolean) As Boolean

On Error GoTo ErrHandler

    EnsurePotBlocksCloserThanHoldersInActiveView = False
    flippedToOpposite = False

    If model Is Nothing Then Exit Function
    If holderIndexes Is Nothing Then Exit Function
    If potIndexes Is Nothing Then Exit Function
    If holderIndexes.count = 0 Then Exit Function
    If potIndexes.count = 0 Then
        LogLine "Pot/front check skipped: no pot CAD indexes found."
        Exit Function
    End If

    Dim holderAvg As Double
    Dim holderMin As Double
    Dim holderMax As Double

    Dim potAvg As Double
    Dim potMin As Double
    Dim potMax As Double

    Dim currentDelta As Double

    If TryGetPotHolderActiveViewFrontDelta( _
            model, _
            holderIndexes, _
            potIndexes, _
            holderAvg, holderMin, holderMax, _
            potAvg, potMin, potMax, _
            currentDelta) = False Then

        LogLine "Pot/front check failed: could not calculate active-view depth."
        Exit Function

    End If

    LogLine "Pot/front active-view depth check BEFORE flip:"
    LogLine "  holder avg/min/max=" & _
            FormatNumberForCsv(holderAvg) & "/" & _
            FormatNumberForCsv(holderMin) & "/" & _
            FormatNumberForCsv(holderMax)

    LogLine "  pot    avg/min/max=" & _
            FormatNumberForCsv(potAvg) & "/" & _
            FormatNumberForCsv(potMin) & "/" & _
            FormatNumberForCsv(potMax)

    LogLine "  front delta=" & FormatNumberForCsv(currentDelta) & _
            "  requirement=" & IIf(POT_FRONT_REQUIRE_EVERY_POT_AHEAD_OF_EVERY_HOLDER, _
                                   "every pot ahead of every holder", _
                                   "average pot ahead of average holder")

    ' In active-view coordinates, larger depth = closer to the viewed/front face.
    If currentDelta > POT_FRONT_DEPTH_MIN_DELTA_IN Then
        EnsurePotBlocksCloserThanHoldersInActiveView = True
        Exit Function
    End If

    ' Current candidate has pots behind holders, so switch to the opposite face.
    LogLine "Pot/front check: pots are NOT closer than holders. Switching to opposite face: " & oppositeViewName

    model.ShowNamedView2 oppositeViewName, oppositeViewId
    StabilizeActiveView model, 100

    flippedToOpposite = True

    Dim newDelta As Double

    If TryGetPotHolderActiveViewFrontDelta( _
            model, _
            holderIndexes, _
            potIndexes, _
            holderAvg, holderMin, holderMax, _
            potAvg, potMin, potMax, _
            newDelta) = False Then

        LogLine "Pot/front check failed after flip: could not calculate active-view depth."
        Exit Function

    End If

    LogLine "Pot/front active-view depth check AFTER flip:"
    LogLine "  holder avg/min/max=" & _
            FormatNumberForCsv(holderAvg) & "/" & _
            FormatNumberForCsv(holderMin) & "/" & _
            FormatNumberForCsv(holderMax)

    LogLine "  pot    avg/min/max=" & _
            FormatNumberForCsv(potAvg) & "/" & _
            FormatNumberForCsv(potMin) & "/" & _
            FormatNumberForCsv(potMax)

    LogLine "  front delta after flip=" & FormatNumberForCsv(newDelta)

    If newDelta > POT_FRONT_DEPTH_MIN_DELTA_IN Then
        EnsurePotBlocksCloserThanHoldersInActiveView = True
    Else
        LogLine "WARNING: Opposite face still does not put pots clearly in front of holders."
        EnsurePotBlocksCloserThanHoldersInActiveView = False
    End If

    Exit Function

ErrHandler:
    LogLine "EnsurePotBlocksCloserThanHoldersInActiveView error: " & Err.Description
    EnsurePotBlocksCloserThanHoldersInActiveView = False
End Function

Private Function EnforcePotBlocksCloserAfterFrontPersist( _
    ByVal model As Object, _
    ByVal holderIndexes As Collection, _
    ByVal potIndexes As Collection) As Boolean

On Error GoTo ErrHandler

    EnforcePotBlocksCloserAfterFrontPersist = False

    If model Is Nothing Then Exit Function
    If holderIndexes Is Nothing Then Exit Function
    If potIndexes Is Nothing Then Exit Function
    If holderIndexes.count = 0 Then Exit Function
    If potIndexes.count = 0 Then Exit Function

    model.ShowNamedView2 "*Front", 1
    StabilizeActiveView model, 100

    Dim holderAvg As Double
    Dim holderMin As Double
    Dim holderMax As Double

    Dim potAvg As Double
    Dim potMin As Double
    Dim potMax As Double

    Dim delta As Double

    If TryGetPotHolderActiveViewFrontDelta( _
            model, _
            holderIndexes, _
            potIndexes, _
            holderAvg, holderMin, holderMax, _
            potAvg, potMin, potMax, _
            delta) = False Then

        LogLine "Final *Front pot verification failed: could not calculate depth."
        Exit Function

    End If

    LogLine "Final *Front pot verification:"
    LogLine "  holder avg/min/max=" & _
            FormatNumberForCsv(holderAvg) & "/" & _
            FormatNumberForCsv(holderMin) & "/" & _
            FormatNumberForCsv(holderMax)

    LogLine "  pot    avg/min/max=" & _
            FormatNumberForCsv(potAvg) & "/" & _
            FormatNumberForCsv(potMin) & "/" & _
            FormatNumberForCsv(potMax)

    LogLine "  final front delta=" & FormatNumberForCsv(delta)

    If delta > POT_FRONT_DEPTH_MIN_DELTA_IN Then
        LogLine "Final *Front verification OK: pot blocks are closer to front than holders."
        EnforcePotBlocksCloserAfterFrontPersist = True
        Exit Function
    End If

    ' If final *Front is still wrong, flip *Back and redefine that as *Front.
    LogLine "Final *Front verification failed. Flipping *Back and redefining that as *Front."

    model.ShowNamedView2 "*Back", 2
    StabilizeActiveView model, 100

    If PersistCurrentViewAsStandardFront(model) = False Then
        LogLine "WARNING: Could not persist flipped *Back as new *Front."
        Exit Function
    End If

    model.ShowNamedView2 "*Front", 1
    StabilizeActiveView model, 100

    Dim delta2 As Double

    If TryGetPotHolderActiveViewFrontDelta( _
            model, _
            holderIndexes, _
            potIndexes, _
            holderAvg, holderMin, holderMax, _
            potAvg, potMin, potMax, _
            delta2) = False Then

        LogLine "Final flipped *Front verification failed: could not calculate depth."
        Exit Function

    End If

    LogLine "Final flipped *Front verification:"
    LogLine "  holder avg/min/max=" & _
            FormatNumberForCsv(holderAvg) & "/" & _
            FormatNumberForCsv(holderMin) & "/" & _
            FormatNumberForCsv(holderMax)

    LogLine "  pot    avg/min/max=" & _
            FormatNumberForCsv(potAvg) & "/" & _
            FormatNumberForCsv(potMin) & "/" & _
            FormatNumberForCsv(potMax)

    LogLine "  final flipped front delta=" & FormatNumberForCsv(delta2)

    If delta2 > POT_FRONT_DEPTH_MIN_DELTA_IN Then
        LogLine "Final flipped *Front verification OK."
        EnforcePotBlocksCloserAfterFrontPersist = True
    Else
        LogLine "WARNING: Pot blocks are still not clearly in front after final flip."
        EnforcePotBlocksCloserAfterFrontPersist = False
    End If

    Exit Function

ErrHandler:
    LogLine "EnforcePotBlocksCloserAfterFrontPersist error: " & Err.Description
    EnforcePotBlocksCloserAfterFrontPersist = False
End Function

Private Function TryGetPotHolderActiveViewFrontDelta( _
    ByVal model As Object, _
    ByVal holderIndexes As Collection, _
    ByVal potIndexes As Collection, _
    ByRef holderAvg As Double, _
    ByRef holderMin As Double, _
    ByRef holderMax As Double, _
    ByRef potAvg As Double, _
    ByRef potMin As Double, _
    ByRef potMax As Double, _
    ByRef frontDelta As Double) As Boolean

On Error GoTo ErrHandler

    TryGetPotHolderActiveViewFrontDelta = False

    holderAvg = 0#
    holderMin = 0#
    holderMax = 0#

    potAvg = 0#
    potMin = 0#
    potMax = 0#

    frontDelta = 0#

    If TryGetActiveViewDepthStatsForCadIndexes(model, holderIndexes, True, holderAvg, holderMin, holderMax) = False Then
        Exit Function
    End If

    If TryGetActiveViewDepthStatsForCadIndexes(model, potIndexes, True, potAvg, potMin, potMax) = False Then
        Exit Function
    End If

    If POT_FRONT_REQUIRE_EVERY_POT_AHEAD_OF_EVERY_HOLDER Then
        ' Strict check:
        ' The farthest-back pot must still be ahead of the closest/front-most holder.
        frontDelta = potMin - holderMax
    Else
        ' Softer check:
        ' Average pot depth must be ahead of average holder depth.
        frontDelta = potAvg - holderAvg
    End If

    TryGetPotHolderActiveViewFrontDelta = True
    Exit Function

ErrHandler:
    TryGetPotHolderActiveViewFrontDelta = False
End Function

Private Function TryGetActiveViewDepthStatsForCadIndexes( _
    ByVal model As Object, _
    ByVal cadIndexes As Collection, _
    ByVal preferMassCenter As Boolean, _
    ByRef avgDepth As Double, _
    ByRef minDepth As Double, _
    ByRef maxDepth As Double) As Boolean

On Error GoTo ErrHandler

    TryGetActiveViewDepthStatsForCadIndexes = False

    avgDepth = 0#
    minDepth = 0#
    maxDepth = 0#

    If model Is Nothing Then Exit Function
    If cadIndexes Is Nothing Then Exit Function
    If cadIndexes.count = 0 Then Exit Function

    Dim total As Double
    Dim countVal As Long
    Dim firstVal As Boolean

    total = 0#
    countVal = 0
    firstVal = True

    Dim i As Long
    Dim cadIdx As Long

    For i = 1 To cadIndexes.count

        cadIdx = CLng(cadIndexes(i))

        Dim px As Double
        Dim py As Double
        Dim pz As Double

        If TryGetCadCenterPointForFrontCheck(cadIdx, preferMassCenter, px, py, pz) Then

            Dim depth As Double

            If TryProjectPointToActiveViewDepth(model, px, py, pz, depth) Then

                If firstVal Then
                    minDepth = depth
                    maxDepth = depth
                    firstVal = False
                Else
                    If depth < minDepth Then minDepth = depth
                    If depth > maxDepth Then maxDepth = depth
                End If

                total = total + depth
                countVal = countVal + 1

            End If

        End If

    Next i

    If countVal > 0 Then
        avgDepth = total / CDbl(countVal)
        TryGetActiveViewDepthStatsForCadIndexes = True
    End If

    Exit Function

ErrHandler:
    TryGetActiveViewDepthStatsForCadIndexes = False
End Function

Private Function TryProjectPointToActiveViewDepth( _
    ByVal model As Object, _
    ByVal px As Double, _
    ByVal py As Double, _
    ByVal pz As Double, _
    ByRef viewDepth As Double) As Boolean

On Error GoTo ErrHandler

    TryProjectPointToActiveViewDepth = False
    viewDepth = 0#

    If model Is Nothing Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    Dim mView As Variant
    mView = swView.Orientation3.ArrayData

    If IsEmpty(mView) Then Exit Function
    If IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    ' Same orientation convention used elsewhere in your macro:
    ' view X     = p dot [m0, m3, m6]
    ' view Y     = p dot [m1, m4, m7]
    ' view depth = p dot [m2, m5, m8]
    '
    ' In SolidWorks active-view coordinates, larger view depth is closer
    ' to the viewed/front face.
    viewDepth = (px * CDbl(mView(2))) + _
                (py * CDbl(mView(5))) + _
                (pz * CDbl(mView(8)))

    TryProjectPointToActiveViewDepth = True
    Exit Function

ErrHandler:
    TryProjectPointToActiveViewDepth = False
End Function

Private Function TryGetCadCenterPointForFrontCheck( _
    ByVal cadIdx As Long, _
    ByVal preferMassCenter As Boolean, _
    ByRef px As Double, _
    ByRef py As Double, _
    ByRef pz As Double) As Boolean

On Error GoTo ErrHandler

    TryGetCadCenterPointForFrontCheck = False

    px = 0#
    py = 0#
    pz = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    ' Prefer actual center of mass, because you specifically asked for center of mass.
    If preferMassCenter Then
        If parts(cadIdx).hasMassCenter Then
            px = parts(cadIdx).MassCenterX
            py = parts(cadIdx).MassCenterY
            pz = parts(cadIdx).MassCenterZ
            TryGetCadCenterPointForFrontCheck = True
            Exit Function
        End If
    End If

    ' Fallback to assembly bounding-box center.
    If parts(cadIdx).hasAsmCenter Then
        px = parts(cadIdx).AsmCenterX
        py = parts(cadIdx).AsmCenterY
        pz = parts(cadIdx).AsmCenterZ
        TryGetCadCenterPointForFrontCheck = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    TryGetCadCenterPointForFrontCheck = False
End Function

Private Function TryGetComponentViewY(ByVal model As Object, _
                                      ByVal swComp As Object, _
                                      ByRef screenY As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentViewY = False
    screenY = 0#

    If model Is Nothing Or swComp Is Nothing Then Exit Function

    Dim vBox As Variant
    On Error Resume Next
    vBox = swComp.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    cx = (CDbl(vBox(0)) + CDbl(vBox(3))) / 2#
    cy = (CDbl(vBox(1)) + CDbl(vBox(4))) / 2#
    cz = (CDbl(vBox(2)) + CDbl(vBox(5))) / 2#

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    Dim mView As Variant
    mView = swView.Orientation3.ArrayData

    If IsEmpty(mView) Then Exit Function
    If IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    screenY = (cx * CDbl(mView(1))) + (cy * CDbl(mView(4))) + (cz * CDbl(mView(7)))

    TryGetComponentViewY = True
    Exit Function

ErrHandler:
    TryGetComponentViewY = False
End Function

Private Function FindCadIndexForOrientationQuoteOrKeys(ByVal quoteName As String, _
                                                       ByVal fallbackKeys As String) As Long
On Error GoTo ErrHandler

    FindCadIndexForOrientationQuoteOrKeys = 0

    Dim cadIdx As Long

    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then
        FindCadIndexForOrientationQuoteOrKeys = cadIdx
    End If

    Exit Function

ErrHandler:
    FindCadIndexForOrientationQuoteOrKeys = 0
End Function

Private Function FindCadPartIndexByQuoteOrKeys(ByVal quoteName As String, ByVal pipeKeys As String) As Long
On Error GoTo ErrHandler

    FindCadPartIndexByQuoteOrKeys = 0

    Dim i As Long
    Dim hay As String

    For i = 1 To PartCount
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath
        If ContainsAnyPipeKey(hay, pipeKeys) Then
            FindCadPartIndexByQuoteOrKeys = i
            Exit Function
        End If
    Next i

    Dim k As String
    k = NormalizeKey(quoteName)

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k And ExportRows(i).HasCad Then
            FindCadPartIndexByQuoteOrKeys = ExportRows(i).CadPartIndex
            Exit Function
        End If
    Next i

    Exit Function

ErrHandler:
    FindCadPartIndexByQuoteOrKeys = 0
End Function

Private Sub AddUniqueCadIndexToCollection(ByVal col As Collection, ByVal cadIdx As Long)
On Error Resume Next

    If col Is Nothing Then Exit Sub
    If cadIdx <= 0 Then Exit Sub
    If cadIdx > PartCount Then Exit Sub

    Dim i As Long

    For i = 1 To col.count
        If CLng(col(i)) = cadIdx Then Exit Sub
    Next i

    col.Add cadIdx
End Sub

Private Sub UnsuppressAllAssemblyComponents(ByVal assyModel As Object)
On Error Resume Next

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    assyModel.ClearSelection2 True

    Dim i As Long
    Dim swComp As Object
    Dim selectedCount As Long

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed Then
                If swComp.Select4(True, Nothing, False) Then selectedCount = selectedCount + 1
            End If
        End If
    Next i

    If selectedCount > 0 Then assyModel.EditUnsuppress2

    assyModel.ClearSelection2 True
End Sub

Private Sub ShowAllAssemblyComponents(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    Dim vComps As Variant
    vComps = model.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long

    For i = 0 To UBound(vComps)
        If Not vComps(i) Is Nothing Then
            If Not vComps(i).IsSuppressed Then vComps(i).Visible = swComponentVisible
        End If
    Next i
End Sub

Private Function HideAllExceptComponentNamesOnce(ByVal assyModel As Object, _
                                                 ByVal keepNames As Collection, _
                                                 ByRef hiddenNames As Collection) As Boolean
On Error GoTo ErrHandler
    HideAllExceptComponentNamesOnce = False
    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function
    If keepNames Is Nothing Then Exit Function
    If keepNames.Count = 0 Then Exit Function
    If hiddenNames Is Nothing Then Set hiddenNames = New Collection

    Dim keepDict As Object
    Set keepDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To keepNames.Count
        keepDict(LCase(CStr(keepNames(i)))) = True
    Next i

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    assyModel.ClearSelection2 True

    Dim swComp As Object
    Dim selectedCount As Long
    Dim keepFoundCount As Long

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If keepDict.Exists(LCase(swComp.Name2)) Then
                    keepFoundCount = keepFoundCount + 1
                    swComp.Visible = swComponentVisible
                Else
                    If swComp.Select4(True, Nothing, False) Then
                        hiddenNames.Add swComp.Name2
                        selectedCount = selectedCount + 1
                    End If
                End If
            End If
        End If
    Next i

    If keepFoundCount = 0 Then
        LogLine "HideAllExceptComponentNamesOnce: none of the requested keep components were found."
        assyModel.ClearSelection2 True
        Exit Function
    End If

    If selectedCount > 0 Then
        LogLine "BASE DXF isolation: hiding non-selected components = " & selectedCount
        swAssembly.HideComponent2
    End If

    assyModel.ClearSelection2 True
    LogLine "BASE DXF isolation: selected components found = " & keepFoundCount
    HideAllExceptComponentNamesOnce = True
    Exit Function

ErrHandler:
    LogLine "HideAllExceptComponentNamesOnce error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    HideAllExceptComponentNamesOnce = False
End Function

Private Sub ShowNamedComponentsOnce(ByVal assyModel As Object, ByVal componentNames As Collection)
On Error GoTo ErrHandler
    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub
    If componentNames Is Nothing Then Exit Sub
    If componentNames.Count = 0 Then Exit Sub

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    assyModel.ClearSelection2 True

    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To componentNames.Count
        nameDict(LCase(CStr(componentNames(i)))) = True
    Next i

    Dim swComp As Object
    Dim selectedCount As Long
    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If nameDict.Exists(LCase(swComp.Name2)) Then
                    If swComp.Select4(True, Nothing, False) Then selectedCount = selectedCount + 1
                End If
            End If
        End If
    Next i

    If selectedCount > 0 Then
        swAssembly.ShowComponent2
    Else
        ShowAllAssemblyComponents assyModel
    End If

    assyModel.ClearSelection2 True
    Exit Sub

ErrHandler:
    LogLine "ShowNamedComponentsOnce error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    ShowAllAssemblyComponents assyModel
End Sub

Private Function FindComponentByKeys(ByVal assyModel As Object, ByVal pipeKeys As String) As Object
On Error GoTo ErrHandler

    Set FindComponentByKeys = Nothing

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function
    If IsArray(vComps) = False Then Exit Function

    Dim bestComp As Object
    Dim bestScore As Double
    bestScore = -1#

    Dim i As Long
    Dim swComp As Object
    Dim hay As String
    Dim score As Double

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                hay = swComp.Name2 & " " & swComp.GetPathName
                If ContainsAnyPipeKey(hay, pipeKeys) Then
                    score = Len(hay)
                    If score > bestScore Then
                        bestScore = score
                        Set bestComp = swComp
                    End If
                End If
            End If
        End If
    Next i

    Set FindComponentByKeys = bestComp
    Exit Function

ErrHandler:
    Set FindComponentByKeys = Nothing
End Function
Private Function FindAssemblyComponentByName(ByVal assyModel As Object, ByVal componentName As String) As Object
On Error GoTo ErrHandler

    Set FindAssemblyComponentByName = Nothing

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function
    If componentName = "" Then Exit Function

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    Dim i As Long
    Dim swComp As Object

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If LCase(swComp.Name2) = LCase(componentName) Then
                    Set FindAssemblyComponentByName = swComp
                    Exit Function
                End If
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    Set FindAssemblyComponentByName = Nothing
End Function










