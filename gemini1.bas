Option Explicit

' ============================================================
' CMS XT EXPORT MACRO - FIXED / SPEED + ROBUSTNESS UPDATE
' ------------------------------------------------------------
' UPDATED IN THIS VERSION:
'  1. Corrected TCP/top-side orientation axis mapping:
'       Y stack -> *Top / *Bottom
'       Z stack -> *Front / *Back
'       X stack -> *Right / *Left
'  2. Orientation now tries top/bottom HOLDER, POT, INS, then TCP/BCP.
'  3. CMS_TOP is saved from matched top-side orientation.
'  4. StabilizeActiveView restored.
'  5. DXF 1:1 setting remains forced for all DXF views/files.
'  6. Later parts include J BLOCK native-assembly DXF fix.
' ============================================================

' NOTE: No Windows API Declare is used. Waiting is handled by WaitMilliseconds.

' ============================================================
' USER SETTINGS
' ============================================================

' Job downloads root is resolved at runtime (company share vs local fallback).
Private Const PUBLIC_DOWNLOADS_PATH As String = "\\Mycloudex2ultra\mexico\Downloads"
Private Const LOCAL_DOWNLOADS_FALLBACK As String = "C:\Users\lenovo\Downloads"
Private Const EXTRACT_FOLDER_NAME As String = "_EXTRACTED_ZIP"
Private Const OUTPUT_FOLDER_SUFFIX As String = " PRINTS"

Private Const LOCAL_WORKSPACE_ROOT As String = "C:\CMS_Local_Workspace"

' --- Network-aware Elgin publish (no manual UNC paste needed) ---
Private Const COMPANY_WIFI_SSID As String = "NETGEAR"
Private Const PUBLIC_DATA_ROOT As String = "\\Mycloudex2ultra\mexico\Cameron's stuff\Matching software"
Private Const PRIVATE_DATA_ROOT As String = "C:\CMS_Local_Workspace\Matching"

' --- Match Studio publish/material/signature settings ---
Private Const ALWAYS_PUBLISH_TO_PUBLIC_MATCHING_SHARE As Boolean = True
Private Const MATCH_STUDIO_FORCE_CARBON_STEEL As Boolean = True
Private Const MATCH_STUDIO_APPLY_PHYSICAL_MATERIAL_TO_ALL_PARTS As Boolean = True
Private Const MATCH_STUDIO_USE_CARBON_STEEL_DENSITY_FOR_MASS As Boolean = False
Private Const MATCH_STUDIO_CARBON_STEEL_DENSITY_LB_PER_CUIN As Double = 0.283
Private Const MATCH_STUDIO_TURBO_ONLY As Boolean = False
Private Const MATCH_STUDIO_MATERIAL_DATABASE As String = "solidworks materials.sldmat"
Private Const MATCH_STUDIO_CARBON_STEEL_MATERIAL As String = "Plain Carbon Steel"

Private Const FORCE_LOCAL_PUBLISH As Boolean = True
Private Const PUBLISH_OUTPUTS As Boolean = True
Private Const AUTO_UPLOAD_JOB_SIGNATURE_TO_ELGIN As Boolean = True
Private Const ELGIN_API_BASE_URL As String = "http://localhost:2926"
Private Const ELGIN_SIGNATURE_API_KEY As String = "cms-signature-upload"
Private Const LIMIT_JOB_SEARCH_TO_CURRENT_AND_PREVIOUS_MONTH As Boolean = True
Private Const JOB_SEARCH_MONTH_WINDOW As Long = 4

Private Const PUSH_OUTPUTS_TO_NETWORK As Boolean = False
Private Const DELETE_EXTRACTED_ZIP_AFTER_FLATTEN As Boolean = True

Private Const USE_ACTIVE_SOLIDWORKS_DOC_FIRST As Boolean = False
Private Const SHOW_ERROR_MESSAGES As Boolean = True

Private Const READ_PDF_BOM_WITH_PDFTOTEXT As Boolean = True
Private Const PDFTOTEXT_EXE As String = "C:\Users\lenovo\Downloads\New folder (9)\poppler-26.02.0\Library\bin\pdftotext.exe"

Private Const ONLY_INCLUDE_ALLOWED_STEEL_BOM_ITEMS As Boolean = True
Private Const DEFAULT_STEEL_TYPE As String = "4140"

' Steel grades that count as normal base material and get exported. Add a grade
' here and it is allowed everywhere - the BOM filter, the PDF material reader
' and the material normaliser all read this one list.
Private Const ALLOWED_STEEL_TYPES As String = "4140|4130"

Private Const HIDE_QUARTER_INCH_THICKNESS As Boolean = True
Private Const PHYSICALLY_HIDE_250_BEFORE_SCAN As Boolean = False
Private Const QUARTER_INCH_THICKNESS As Double = 0.25
Private Const QUARTER_INCH_TOLERANCE As Double = 0.01

Private Const DIM_DECIMALS As Long = 3
Private Const DIM_OK_TOL As Double = 0.03
Private Const DIM_REVIEW_TOL As Double = 0.125
Private Const DIM_MAX_MATCH_TOTAL_DIFF As Double = 5#
Private Const MIN_STEEL_VOLUME_CUIN As Double = 1#

Private Const INCHES_PER_METER As Double = 39.3700787401575
Private Const CUIN_PER_CUBIC_METER As Double = 61023.7440947323

Private Const SAME_SIZE_PAIR_TOL As Double = 0.125

Private Const TURBO_READ_ONLY_BOM_SHEET As Boolean = True
Private Const TURBO_BOM_SHEET_NAME As String = "BOM"
Private Const BOM_HEADER_SEARCH_MAX_ROWS As Long = 150
Private Const STOP_BOM_READ_AFTER_BLANK_ROWS As Long = 20
Private Const VERBOSE_PROGRESS_LOG As Boolean = False

Private Const PROMPT_FOR_TOP_ORIENTATION As Boolean = False
Private Const CMS_TOP_VIEW_NAME As String = "CMS_TOP"

Private Const CMS_BASE_TOP_VIEW_NAME As String = "*Bottom"
Private Const CMS_BASE_TOP_VIEW_ID As Long = 6

' If the corrected top-side face is right but rotated on screen, adjust this:
' 0 = no rotation, 1 = rotate +Z once, 2 = 180, -1 = rotate -Z once.
Private Const CMS_TOP_ROTATE_Z_STEPS As Long = 0

Private Const AUTO_SELECT_TCP_TOP_ORIENTATION As Boolean = True

Private Const TCP_TOP_ORIENTATION_KEYS As String = "TCP|TOP CLAMPING|TOP CLAMPING PLATE|TOP SMED|TOP SMED PLATE|ID SMED"
Private Const BCP_BOTTOM_ORIENTATION_KEYS As String = "BCP|BOTTOM CLAMPING|BOTTOM CLAMPING PLATE|BOT CLAMPING|BOT CLAMPING PLATE|BOTTOM SMED|BOT SMED|OD SMED"

Private Const PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE As Boolean = True

Private Const AUTO_DEFINE_FRONT_FROM_HOLDER_POT_COM As Boolean = True

' If the holder long dimension is not visible in the current front view,
' the current front is looking down the holder length, so use *Right as front.
Private Const HOLDER_LONG_SIDE_VISIBLE_RATIO As Double = 0.8

' Small tolerance for pot-vs-holder front/back center comparison.
Private Const FRONT_COM_MIN_DELTA_IN As Double = 0.03

' SolidWorks normal convention is usually:
'   *Front sees from +Z
'   *Right sees from +X
' So larger Z/X is closer to that face.
' If your result is exactly backwards, change this to False.
Private Const FRONT_CLOSER_IS_LARGER_AXIS_VALUE As Boolean = True

Private Const POT_BLOCKS_MUST_BE_FRONT_OF_HOLDERS As Boolean = True

' Minimum depth difference required for pots to count as being in front of holders.
Private Const POT_FRONT_DEPTH_MIN_DELTA_IN As Double = 0.03

' True = every pot block must be in front of every holder.
' False = average pot center must be in front of average holder center.
'
' Set to False because True is geometrically unsatisfiable on a two-sided mold.
' The J8494 log proves it: holders span depth -0.969 to 6.968 and pots span
' -0.500 to 6.435, so the two groups interleave and potMin - holderMax is
' negative no matter which way round the view is. It failed, flipped, failed
' again, flipped back and failed again. Comparing group averages is the test
' that can actually pass.
Private Const POT_FRONT_REQUIRE_EVERY_POT_AHEAD_OF_EVERY_HOLDER As Boolean = False

' ============================================================
' TOP-CLAMPING-PLATE "NORMAL TO" ORIENTATION
' ------------------------------------------------------------
' This is what replaced physically rotating the assembly.
'
' Instead of moving geometry, it does what you would do by hand: pick the big
' flat face of the top clamping plate, hit Normal To, and save that as *Top.
' The parts never move, so nothing can shear and no component transform is
' touched - only the definition of the standard views changes, and the standard
' views are what every downstream step actually reads.
'
' Normal To on its own only fixes TILT (view perpendicular to the plate). It
' leaves the in-plane spin arbitrary, which would still leave the plate sitting
' at an angle on screen. So the view is also aligned to the plate's own longest
' straight edge, which squares the spin at the same time. That is the whole
' correction - tilt from the face normal, spin from the face edge.
'
' After this, the standard orientation process runs unchanged: it decides top
' vs bottom from the TCP/BCP centres, then sets *Front and puts the pot blocks
' closer to the front than the holders.
' ============================================================
Private Const ORIENT_TOP_BY_TCP_NORMAL_TO As Boolean = True
' --- PATCH: geometry-based top-view spin alignment after TCP Normal To ---
Private Const AUTO_ALIGN_TOP_SPIN_FROM_HOLDER_GEOMETRY As Boolean = True
Private Const TOP_SPIN_HOLDER_WIDTH_MARGIN_IN As Double = 0.05
' --- END PATCH ---

' Use a pot-block side face to set the front direction, which is what squares the
' spin. This is the important one. With it on, the TCP face decides which way is
' UP and the pot face decides which way is FRONT - two real faces off the model,
' no guessing. With it off the code falls back to the TCP's own longest edge,
' which is what left the base "close but rotated".
Private Const ORIENT_FRONT_BY_POT_NORMAL_TO As Boolean = True

' Component name keys for finding a pot block, when the matched quote misses.
Private Const POT_FRONT_ORIENTATION_KEYS As String = "ID POT BLOCK|OD POT BLOCK|TOP POT BLOCK|BOTTOM POT BLOCK|BOT POT BLOCK|ID POT|OD POT|POT BLOCK"

' A pot face counts as a SIDE face only if its normal is this close to
' perpendicular to the TCP normal. Keeps the pot's own top/bottom faces out.
Private Const POT_FRONT_MAX_TOP_DOT As Double = 0.25

' Really select the face and really press Normal To, instead of writing a view
' matrix. Writing the matrix produced a 90-deg-wrong view on J8494, because
' Orientation3's setter does not use the layout its getter reports. Selecting the
' face and running the command leaves all of that maths to SolidWorks.
Private Const TCP_USE_REAL_NORMAL_TO_COMMAND As Boolean = True

' --- PATCH: make Normal To actually run and force TCP to lighter clamping plate ---
Private Const USE_CTRL8_SENDKEYS_FOR_NORMAL_TO As Boolean = True
Private Const FORCE_TCP_TO_LIGHTER_CLAMPING_PLATE As Boolean = True
' --- END PATCH ---


' Candidate command IDs for View > Modify > Normal To. The ID moves between
' SolidWorks releases, so each is tried in turn and the view is measured after
' every attempt; the one that squares the face wins and gets logged. Once the log
' names the winner you can cut this list down to that single number.
Private Const NORMAL_TO_COMMAND_IDS As String = "1751|1338|1466|1465|1339|1752"

' Ray-pick radius in metres for selecting the face.
Private Const FACE_PICK_RAY_RADIUS_M As Double = 0.0015

' How far outside the part the pick ray starts, in metres (~20 in).
Private Const FACE_PICK_RAY_STANDOFF_M As Double = 0.5

' swSelectType_e.swSelFACES
Private Const swSelFACES As Long = 2

' A plate is square to the screen when its projected width and height match its
' real length and width to within this many inches. This is the truth test - it
' uses SolidWorks' own projection of real geometry, so it does not depend on any
' matrix convention being guessed right.
Private Const VIEW_SQUARE_TOL_IN As Double = 0.02

' Fallback only, used when ORIENT_FRONT_BY_POT_NORMAL_TO is False or the pot face
' cannot be found: align to the plate's longest straight edge.
Private Const TCP_NORMAL_ALIGN_TO_LONGEST_EDGE As Boolean = True

' A face smaller than this is not the plate's main face; ignore it.
Private Const TCP_NORMAL_MIN_FACE_AREA_SQIN As Double = 1#

' How square the face has to end up before the view is accepted and saved as
' *Top. Measured, not assumed - if the write is ignored this catches it.
Private Const TCP_NORMAL_MAX_OFF_DEG As Double = 0.5

' Named view saved at the Normal To orientation, for traceability in the tree.
Private Const TCP_NORMAL_VIEW_NAME As String = "CMS_TCP_NORMAL"

' ============================================================
' WHOLE-BASE STRAIGHTEN SETTINGS  (SUPERSEDED - OFF)
' ------------------------------------------------------------
' Kept only as a fallback. The Normal To method above replaced this because it
' does not move any geometry. Leave this False unless the Normal To path fails
' on a job and you want to compare; nothing else reads it.
' The imported SolidWorks base often arrives yawed a few degrees off the
' assembly X/Y/Z axes. Nothing downstream survives that: the CAD scan, the
' BOM dimension match, the TCP-top pick, the pot-vs-holder front test and
' every projected DXF all read assembly-space positions.
'
' This measures the WHOLE base the same way the pullcore straighten angle is
' measured - smallest bounding box - then FLOATS every component, rotates the
' whole assembly as ONE rigid body by the angle found, and FIXES everything
' again. Every part turns, by the same angle, in the same operation; no part
' type is singled out. Only after that does the standard orientation process
' run (TCP on top, then pots closer to *Front than the holders).
' ============================================================
Private Const STRAIGHTEN_WHOLE_BASE_BEFORE_ORIENT As Boolean = False

' Ignore skews smaller than this. A base that really is square still solves
' to a few hundredths of a degree of numeric noise.
Private Const STRAIGHTEN_BASE_MIN_DEG As Double = 0.25

' Never trust a solved angle larger than this. A real import skew is single
' digits to low double digits of degrees; a solve up near the 45 deg edge of
' the sweep means the box search locked onto the wrong axis. Hitting this limit
' is LOGGED and the base is left alone, so if a job legitimately arrives at a
' bigger yaw the log will say so and this number can be raised.
' Must stay below 45 or the guard can never fire.
Private Const STRAIGHTEN_BASE_MAX_DEG As Double = 30#

' One pass squares a single-axis yaw. A base tipped about two axes needs a
' second pass, re-measured after the first rotation is applied.
Private Const STRAIGHTEN_BASE_MAX_PASSES As Long = 3

' A pass must shrink the bounding-box volume by at least this fraction to be
' worth applying. 0.002 = 0.2%.
Private Const STRAIGHTEN_BASE_MIN_VOLUME_GAIN As Double = 0.002

' Angle sweep resolution, degrees. Coarse scans -45..+45 on each global axis,
' fine refines around the coarse winner.
Private Const STRAIGHTEN_BASE_COARSE_STEP_DEG As Double = 1#
Private Const STRAIGHTEN_BASE_FINE_STEP_DEG As Double = 0.05

' Mates are suppressed before the rotation so the solver cannot drag mated
' components back toward geometry that just moved. Leave them suppressed:
' every component is fixed afterwards and nothing downstream needs mates.
Private Const STRAIGHTEN_BASE_KEEP_MATES_SUPPRESSED As Boolean = True

Private Const USE_CAD_NAMING_LIBRARY As Boolean = True
Private Const LEARN_CAD_NAMING_LIBRARY_FROM_BOM As Boolean = True

Private Const CAD_NAMING_LIBRARY_FILE As String = "CMS_Block_Naming_Library.csv"

' Matching tolerances for no-BOM library fallback.
Private Const CAD_LIB_MAX_SINGLE_DIM_DIFF_IN As Double = 0.75
Private Const CAD_LIB_MAX_TOTAL_DIM_DIFF_IN As Double = 1.75
Private Const CAD_LIB_MAX_NORM_DISTANCE As Double = 0.4

' Score = dimension difference + normalized location distance.
Private Const CAD_LIB_SCORE_DIM_WEIGHT As Double = 1#
Private Const CAD_LIB_SCORE_LOC_WEIGHT As Double = 3#
Private Const CAD_LIB_MAX_SCORE As Double = 2.9

' One match per quote name by default: ID HOLDER once, OD HOLDER once, etc.
Private Const CAD_LIB_MAX_MATCHES_PER_QUOTE As Long = 1

' If two possible library names are close, still use best match but log warning.
Private Const CAD_LIB_AMBIGUOUS_SCORE_GAP As Double = 0.25

' ---- NO-BOM LIBRARY: PACKAGE + MATCHING-JOB FIXES (added) ----------------
' Reserve the J BLOCK and EJECTOR CAM with the same name-based finders the
' BOM flow uses BEFORE the size-only library matcher runs, so they keep their
' package DXFs instead of leaking into MISC DETAILS under another job's name.
Private Const CAD_LIB_RESERVE_PACKAGE_PARTS_FIRST As Boolean = True

' When a no-PDF job is a re-run of a previous job, most parts match one
' library source job almost exactly. Detect that job and prefer its names so
' the whole job is named/categorized like the original.
Private Const CAD_LIB_PREFER_DOMINANT_SOURCE_JOB As Boolean = True
Private Const CAD_LIB_DOMINANT_JOB_PERFECT_DIM_DIFF As Double = 0.06
Private Const CAD_LIB_DOMINANT_JOB_MIN_HITS As Long = 3
Private Const CAD_LIB_DOMINANT_JOB_SCORE_BONUS As Double = 0.75
' --------------------------------------------------------------------------


' ============================================================
' SPEED / GRAPHICS SETTINGS
' ============================================================

' These three are OFF on purpose while the TCP/pot orientation is being proved out.
' View manipulation is the least reliable corner of the SolidWorks API when the
' app is invisible, viewport graphics are disabled and there are no settle
' delays - a view write is accepted without error and simply does not happen.
' That is what made the first Normal To attempt log "36.000 deg off".
' Turn them back on for speed once the orientation is confirmed good on a few jobs.
Private Const DISABLE_STABILIZE_DELAYS As Boolean = False
Private Const DISABLE_MAIN_VIEWPORT_GRAPHICS As Boolean = False
Private Const RUN_SOLIDWORKS_INVISIBLE As Boolean = False
Private Const ENABLE_EXPORT_LOG As Boolean = True

' ============================================================
' TOP/BOT INSERT GEOMETRY FALLBACK
' ============================================================

Private Const ADD_TOP_BOT_INS_FROM_CAD_GEOMETRY As Boolean = True
Private Const INSERT_THICKNESS_TOL As Double = 0.035
Private Const INSERT_WIDTH_MATCH_TOL As Double = 0.25
Private Const INSERT_LENGTH_MATCH_TOL As Double = 0.75
Private Const INSERT_LENGTH_MATCH_TOL_LOOSE As Double = 1.5

' ============================================================
' DXF SETTINGS
' ============================================================

Private Const CREATE_INDIVIDUAL_DXFS As Boolean = False
Private Const CREATE_DXFS_DURING_XT_SAVE As Boolean = True
' --- Engraving DXF exports ---
Private Const CREATE_ENGRAVING_DXFS As Boolean = True
Private Const ENGRAVING_DXF_TOKEN_SUFFIX As String = " ENG"


' Keep fast batch behavior, but DO NOT skip DXFs.
Private Const FAST_BATCH_EXPORT As Boolean = True
Private Const FAST_SKIP_PRINTS_DXF As Boolean = False
Private Const FAST_SKIP_BASE_PACKAGE_DXF As Boolean = False
Private Const FAST_SKIP_HOLDERS_PACKAGE_DXF As Boolean = False

' Match Studio: no JPEG previews.
Private Const CREATE_COMPONENT_ISO_JPEGS As Boolean = False
Private Const CREATE_MATCH_STUDIO_HOLDERS_STL As Boolean = False
Private Const CREATE_MATCH_STUDIO_STL_PACKAGE As Boolean = False
Private Const CREATE_MATCH_STUDIO_STLS_DURING_EXISTING_EXPORTS As Boolean = True
Private Const MATCH_STUDIO_STL_FOLDER_SUFFIX As String = " STL"
Private Const MATCH_STUDIO_STL_MATCH_MAIN_BASE_ORIENTATION As Boolean = True

' Rotates exported STL mesh into the corrected SolidWorks *Front/*Top orientation.
' This is needed because STL does not store named views.
Private Const POST_ROTATE_STL_TO_CORRECTED_FRONT As Boolean = True

Private Const CREATE_COMPONENT_IGS_WITH_XT As Boolean = False
Private Const CREATE_COMPONENT_EASM_WITH_XT As Boolean = False

' UPDATED: this is now honored for BASE/HOLDERS too in later DXF code.
Private Const FORCE_ALL_DXF_VIEWS_1_TO_1 As Boolean = True

' BASE DXF spacing fix:
' 1:1 BASE views often cannot fit on 44 x 34 without overlap.
' This expands only the MAIN ASSEMBLY / BASE DXF sheet and spreads projected views.
Private Const BASE_DXF_AUTO_EXPAND_SHEET_TO_FIT_1_TO_1 As Boolean = True
Private Const BASE_DXF_EXTRA_SIDE_GAP_IN As Double = 8#
Private Const BASE_DXF_EXTRA_TOP_BOTTOM_GAP_IN As Double = 5#

Private Const BASE_DXF_MIN_SHEET_WIDTH_IN As Double = 60#
Private Const BASE_DXF_MIN_SHEET_HEIGHT_IN As Double = 44#

' TCP/BCP DXF in-plane rotation adjustment.
' Positive = counter-clockwise on the drawing sheet.
' Try 180 if upside down, 90 or -90 if sideways.
Private Const TCP_DXF_CENTER_ROTATION_DEG As Double = 0#
Private Const BCP_DXF_CENTER_ROTATION_DEG As Double = 0#

Private Const FREEZE_DXF_DRAWING_GRAPHICS As Boolean = True

Private Const FLIP_ID_HOLDER_CENTER_VIEW_180 As Boolean = True
Private Const FLIP_ID_HOLDER_CENTER_VIEW_180_FROM_ASSEMBLY As Boolean = False

Private Const OD_HOLDER_CENTER_ROTATION_DEG As Double = 0#

Private Const DIMENSION_J_BLOCK As Boolean = True
Private Const MULTIVIEW_FIT_SAFETY As Double = 0.9

Private Const SW_DRAWING_TEMPLATE_PATH As String = "C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Drawing.drwdot"

Private Const E_SHEET_WIDTH_IN As Double = 44#
Private Const E_SHEET_HEIGHT_IN As Double = 34#
Private Const DXF_MARGIN_IN As Double = 1#
Private Const DXF_MAX_SCALE As Double = 1#
Private Const DXF_PROJECTED_VIEW_GAP_IN As Double = 2.25
Private Const HOLDERS_SIDE_VIEW_EXTRA_GAP_IN As Double = 6#
Private Const PI_VALUE As Double = 3.14159265358979

' ============================================================
' PULL CORE DIMENSIONS EXCEL REPORT
' ============================================================

Private Const CREATE_PULLCORE_DIMENSIONS_EXCEL As Boolean = True
Private Const PULLCORE_DIMENSIONS_REPORT_FILE As String = "Pull Core Dimensions.xlsx"
Private Const JOB_SIGNATURE_REPORT_FILE As String = "XT_Export_Job_Signature.csv"

' ============================================================
' MAIN ASSEMBLY / HOLDERS PACKAGE SETTINGS
' ============================================================

Private Const CREATE_MAIN_ASSEMBLY_PACKAGE As Boolean = True
Private Const MAIN_ASSEMBLY_FILE_TOKEN As String = "BASE"
Private Const HOLDERS_ONLY_FILE_TOKEN As String = "HOLDERS"

' Whole base + every part as one STL, published into the J#### STL folder so
' Match Studio can overlay/compare the entire assembly between two jobs.
Private Const CREATE_FULL_ASSEMBLY_STL As Boolean = True
Private Const FULL_ASSEMBLY_STL_TOKEN As String = "FULL ASSEMBLY"

Private Const ID_HOLDER_KEYS As String = "ID HOLDER|TOP HOLDER|IDTE HOLDER|TOP CARRIER|ID CARRIER|ID MOLD BASE|ID MOLDBASE|IDTE MOLD BASE"
Private Const OD_HOLDER_KEYS As String = "OD HOLDER|BOTTOM HOLDER|BOT HOLDER|ODTE HOLDER|BOTTOM CARRIER|OD CARRIER|OD MOLD BASE|OD MOLDBASE|ODTE MOLD BASE"

' ============================================================
' J BLOCK PACKAGE SETTINGS
' ============================================================

Private Const CREATE_J_BLOCK_PACKAGE As Boolean = True
Private Const J_BLOCK_FOLDER_NAME As String = "J BLOCK"

Private Const J_BLOCK_NAME_KEYS As String = "EJECTOR J-BLOCK|EJECTOR J BLOCK|EJECTOR JBLOCK|EJ J-BLOCK|EJ J BLOCK|EJ JBLOCK|EJ. J-BLOCK|EJ. J BLOCK|EJ.J-BLOCK|EJ.J BLOCK|J-BLOCK|J BLOCK|JBLOCK|EJECTOR U-BLOCK|EJECTOR U BLOCK|EJECTOR UBLOCK|EJ U-BLOCK|EJ U BLOCK|EJ UBLOCK|EJ. U-BLOCK|EJ. U BLOCK|EJ.U-BLOCK|EJ.U BLOCK|U-BLOCK|U BLOCK|UBLOCK"

Private Const J_BLOCK_DIMENSION_DECIMALS As Long = 4

Private Const J_BLOCK_PARENT_VIEW_PRIMARY As String = "*Right"
Private Const J_BLOCK_PARENT_VIEW_FALLBACK As String = "*Right"

Private Const ROTATE_J_BLOCK_DXF_180 As Boolean = False

Private Const J_BLOCK_TARGET_THICKNESS As Double = 1#
Private Const J_BLOCK_TARGET_WIDTH As Double = 2.75
Private Const J_BLOCK_TARGET_LENGTH As Double = 3.838
Private Const J_BLOCK_DIM_MATCH_TOL As Double = 0.75

Private Const EJECTOR_CAM_NAME_KEYS As String = "EJECTOR CAM|EJ CAM|EJ. CAM|EJ.CAM|CAM"
Private Const EJECTOR_CAM_TARGET_THICKNESS As Double = 1#
Private Const EJECTOR_CAM_TARGET_WIDTH As Double = 1.28
Private Const EJECTOR_CAM_TARGET_LENGTH As Double = 1.63
Private Const EJECTOR_CAM_THICKNESS_TOL As Double = 0.35
Private Const EJECTOR_CAM_WIDTH_TOL As Double = 0.5
Private Const EJECTOR_CAM_LENGTH_TOL As Double = 0.6

Private Const J_BLOCK_BASE_FILE_TOKEN As String = "J BLOCK"
Private Const J_BLOCK_CAM_FILE_TOKEN As String = "CAM"

Private JBlockTgtL As Double
Private JBlockTgtW As Double
Private JBlockTgtT As Double
Private EjCamTgtL As Double
Private EjCamTgtW As Double
Private EjCamTgtT As Double

' ============================================================
' PULLCORE CAM AND KEY PACKAGE SETTINGS
' ============================================================

Private Const CREATE_PULLCORE_CAM_KEY_PACKAGE As Boolean = True
Private Const PULLCORE_CAM_KEY_FOLDER_NAME As String = "PULLCORE CAM AND KEY"

Private Const AUTO_LABEL_PULLCORE_ID_OD_BY_HEIGHT As Boolean = True
Private Const PULLCORE_ID_OD_HEIGHT_AXIS As String = "Y"
Private Const PULLCORE_ID_IS_HIGHER As Boolean = True

Private Const PULLCORE_T_TOL As Double = 0.175
Private Const PULLCORE_W_TOL As Double = 0.25
Private Const PULLCORE_L_TOL As Double = 0.35

Private Const USE_PULLCORE_BEST_FIT_BBOX As Boolean = True
Private Const PULLCORE_BEST_FIT_ALL_PARTS As Boolean = False

' ---- PULLCORE DETECTION / NAMING FIX (added) -----------------------------
' Respect the BOM quantity exactly. The old code forced every KEY row to be
' counted at least twice, which manufactured phantom rows whenever the model
' only contained one instance of each key -> those phantom rows became
' "NO CAD MATCH". Leave this True so cams and keys behave identically.
Private Const PULLCORE_RESPECT_BOM_QUANTITY As Boolean = True

' Identical instances of the same file are de-duplicated into a single part
' (Quantity is incremented). When True, that single part can satisfy as many
' BOM slots as it has instances, so symmetric / mirrored keys that share a
' file are no longer starved of matches.
Private Const PULLCORE_EXPAND_CANDIDATES_BY_QTY As Boolean = True

' Keep the descriptive BOM name (e.g. "OD TE Pullcore Cam") instead of
' overwriting it with a generic "ID/OD PULLCORE CAM/KEY" label. The generic
' relabeler assumed exactly one ID and one OD per class, dropped the TE/LE
' distinction and created duplicate names. When True, the descriptive name is
' preserved whenever the BOM already encodes a location (ID/OD/TE/LE); the
' generic labeler only runs for rows whose name is genuinely ambiguous.
Private Const PULLCORE_PRESERVE_DESCRIPTIVE_BOM_NAMES As Boolean = True

' Name each pullcore KEY after the CAM it physically sits inside, matched by
' centerline (assembly-center proximity). Keys of different locations are often
' the SAME size (e.g. an OD LE key and an OD TE key are both 3.375 x 1.25 x
' 0.75), so size matching alone cannot tell them apart and mislabels them. The
' cams ARE distinct, so the key inherits its nested cam's location (OD TE, ID
' TE, OD LE, ...). This is the authoritative source for key location/naming.
Private Const PULLCORE_NAME_KEYS_BY_NESTED_CAM As Boolean = True

' Optional height-based ID/OD nudge for the size matcher. Superseded by the
' nested-cam naming above (which uses real geometry instead of guessing a
' height axis), so it is OFF by default. Set True only to re-enable the older
' heuristic; it has no effect on key NAMES when nested-cam naming is on.
Private Const PULLCORE_USE_LOCATION_AWARE_MATCH As Boolean = True
Private Const PULLCORE_LOCATION_MATCH_WEIGHT As Double = 0.6

' DIMS DECIDE, LOCATION ONLY BREAKS TIES. Pair scores are bucketed by the
' fitted-dims distance: pairs whose dim scores fall in the same bucket are
' considered the same stock size and location picks between them; pairs in a
' better (smaller) bucket ALWAYS win regardless of location. This makes it
' impossible for the ID/OD side nudge to flip two different stock sizes.
Private Const PULLCORE_DIM_SCORE_BUCKET As Double = 0.15

' --------------------------------------------------------------------------

Private Const STRAIGHTEN_PULLCORE_DXF As Boolean = True
Private Const PULLCORE_STRAIGHTEN_SIGN As Double = -1#
Private Const PULLCORE_STRAIGHTEN_MIN_DEG As Double = 2#
Private Const PULLCORE_FALLBACK_ANGLE_DEG As Double = 32.5

' ---- PULLCORE DXF ORIENTATION (FIX) --------------------------------------
' The pullcore straighten angle is a rotation IN ONE PLANE. It is only a valid
' correction if the drawing's center view is projected in that same plane.
'
' When True the pullcore drawing source keeps the orientation it was exported
' with (assembly coordinates are baked into the X_T), so the temp part's *Front
' view is the assembly XY plane - the plane the pull tilt is measured in.
'
' When False the source is re-oriented by SetCmsTopOrientation first. On an
' isolated single pullcore body that routine has no TCP/BCP to key on, so it
' falls back to *Bottom + CMS_TOP_ROTATE_Z_STEPS: an arbitrary baseline that
' varies from part to part and invalidates the angle. That was the bug.
Private Const PULLCORE_DXF_KEEP_IMPORTED_ORIENTATION As Boolean = True

' Measure the straighten angle on the ACTUAL drawing source, in the ACTUAL
' center-view plane, rather than trusting the angle recorded during BOM matching
' (measured in whatever view the component's own part file happened to be saved
' with). The BOM-match angle is kept as a fallback.
Private Const PULLCORE_MEASURE_ANGLE_IN_VIEW_PLANE As Boolean = True

' The pullcore center/base view. Change these two together if the pullcore
' needs to be drawn off a different face; the angle is now measured in whatever
' view is named here, so the two can no longer drift apart.
Private Const PULLCORE_PARENT_VIEW_PRIMARY As String = "*Front"
Private Const PULLCORE_PARENT_VIEW_FALLBACK As String = "*Front"

' How close (deg) the in-view edge angle must be to the bounding-box magnitude
' before the edge value is trusted for magnitude. Within this band the edge's
' more precise value is used; outside it the robust box magnitude stands and
' the edge contributes only the sign (lean direction).
Private Const PULLCORE_ANGLE_AGREE_TOL As Double = 6#

' ---- ANGLE-AWARE PULLCORE CANDIDATE RECOVERY (added) ---------------------
' Pullcore cams/keys usually sit at an angle in the assembly, so their
' axis-aligned bounding box is much bigger than the BOM stock size and the
' direct L/W/T tolerance check rejects them (-> NO CAD MATCH). When this is
' True, a part is also accepted as a candidate if rotating the BOM block by
' a known pullcore angle reproduces the part's bounding box. Known angles =
' angles already detected on matched cams/keys in this job (so a matched CAM
' donates its angle to find the missing KEY and vice versa), then the
' configured fallback angle, then (optionally) a free angle search.
Private Const PULLCORE_ANGLE_AWARE_CANDIDATES As Boolean = True
Private Const PULLCORE_ANGLE_AABB_TOL As Double = 0.35
Private Const PULLCORE_ANGLE_INFER_ANY_ANGLE As Boolean = True
Private Const PULLCORE_ANGLE_VOL_RATIO_MIN As Double = 0.8
Private Const PULLCORE_ANGLE_VOL_RATIO_MAX As Double = 3.2

' A part may only become a pullcore candidate if its part file can be opened
' individually and a trusted FITTED bounding box obtained, and that fitted
' box must match one of the pullcore BOM stock sizes. Parts where only the
' regular (possibly rotated) box is available are dropped - names are never
' assigned from regular-box guesses.
Private Const PULLCORE_REQUIRE_FITTED_BBOX As Boolean = True
Private Const PULLCORE_FITTED_VERIFY_TOL_SCALE As Double = 1.5

' When fewer pullcore candidates than BOM rows survive the rough screen,
' sweep ALL remaining plausible parts: open each individually, compute its
' fitted box, and nominate it if the fitted box matches a stock size. This
' rescues parts rotated in ways the rough regular-box screen cannot predict.
Private Const PULLCORE_FITTED_SWEEP_FALLBACK As Boolean = True

' Regular (non-pullcore) BOM rows: when no part matches by its regular box
' (common in outsource layout assemblies where parts sit rotated), fall back
' to opening plausible parts individually and matching on the FITTED box.
Private Const REGULAR_MATCH_FITTED_BBOX_FALLBACK As Boolean = True
Private Const FITTED_FALLBACK_VOL_RATIO_MIN As Double = 0.6
Private Const FITTED_FALLBACK_VOL_RATIO_MAX As Double = 4.5

' --------------------------------------------------------------------------


' ============================================================
' PULLCORE STOP / FLIPPER CAM COVER PACKAGE SETTINGS
' ============================================================

Private Const CREATE_PULLCORE_STOP_PACKAGE As Boolean = True
Private Const PULLCORE_STOP_FOLDER_NAME As String = "PULLCORE STOP"

Private Const PULLCORE_STOP_NAME_KEYS As String = "PULLCORE STOP|PULL CORE STOP"
Private Const FLIPPER_CAM_COVER_NAME_KEYS As String = "FLIPPER CAM COVER PLATE|FLIPPER CAM COVER|CAM COVER PLATE"

Private Const FLIPPER_CAM_COVER_TARGET_THICKNESS As Double = 2.95
Private Const FLIPPER_CAM_COVER_TARGET_WIDTH As Double = 4.225
Private Const FLIPPER_CAM_COVER_TARGET_LENGTH As Double = 5.975
Private Const FLIPPER_CAM_COVER_DIM_TOL As Double = 0.35

' ============================================================
' PYROPEL SETTINGS
' ============================================================

Private Const ROUTE_PYROPEL_TO_SEPARATE_FOLDERS As Boolean = True
Private Const PYROPEL_POTS_FOLDER_SUFFIX As String = " PYROPEL POTS"
Private Const PYROPEL_HOLDERS_FOLDER_SUFFIX As String = " PYROPEL HOLDERS"

' ============================================================
' SOLIDWORKS / EXCEL CONSTANTS
' ============================================================

Private Const swDocNONE As Long = 0
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3

Private Const swOpenDocOptions_Silent As Long = 1
Private Const swOpenDocOptions_ReadOnly As Long = 2

Private Const swSaveAsCurrentVersion As Long = 0
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swSaveAsOptions_Copy As Long = 2

' swUserPreferenceToggle_e: "Save all components of an assembly in a single
' STL file" (vs one STL per component). Used only by the fallback path that
' exports the STL straight from the assembly.
Private Const swSTLComponentsIntoOneFile As Long = 248

' Saving an assembly as a .sldprt: include every component as a solid body.
' swUserPreferenceIntegerValue_e.swSaveAssemblyAsPartOptions = 201
' swSaveAsmAsPartOptions_e.swSaveAsmAsPart_AllComponents = 1
Private Const swSaveAssemblyAsPartOptions As Long = 201
Private Const swSaveAsmAsPart_AllComponents As Long = 1

' swBodyOperationType_e.SWBODYADD = 15903 (Combine -> Add / union of bodies).
Private Const SWBODYADD As Long = 15903

Private Const swSolidBody As Long = 0

Private Const swComponentHidden As Long = 0
Private Const swComponentVisible As Long = 1

Private Const xlCalculationManual As Long = -4135
Private Const xlOpenXMLWorkbook As Long = 51

' ============================================================
' TYPES
' ============================================================

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
    BBoxVolume As Double
    massValue As Double

    hasOriginalAsmBBox As Boolean
    OriginalAsmLength As Double
    OriginalAsmWidth As Double
    OriginalAsmThickness As Double

    hasAsmCenter As Boolean
    AsmCenterX As Double
    AsmCenterY As Double
    AsmCenterZ As Double

    hasMassCenter As Boolean
    MassCenterX As Double
    MassCenterY As Double
    MassCenterZ As Double

    hasTopRotation As Boolean
    topRotationRad As Double

    hasAsmPullAngle As Boolean
    asmPullAngleDeg As Double
    asmPullAngleSign As Double

    UsedForBomMatch As Boolean
    isBodyOnly As Boolean
End Type

Private Type BomInfo
    Description As String
    quoteName As String
    TypeField As String
    Quantity As Long
    material As String

    BomLength As Double
    BomWidth As Double
    BomThickness As Double
    hasDims As Boolean
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

Private Type PullcoreMatchInfo
    quoteName As String
    Description As String
    material As String
    Quantity As Long

    CadPartIndex As Long
    isCam As Boolean

    BomThickness As Double
    BomWidth As Double
    BomLength As Double

    CadThickness As Double
    CadWidth As Double
    CadLength As Double

    OriginalThickness As Double
    OriginalWidth As Double
    OriginalLength As Double

    FittedThickness As Double
    FittedWidth As Double
    FittedLength As Double

    DetectedAngleDeg As Double
    DetectedAngleSign As Double
    DxfRotationDeg As Double

    Status As String
End Type

Private Type CadNameLibEntry
    quoteName As String
    Length As Double
    Width As Double
    Thickness As Double

    centerX As Double
    centerY As Double
    centerZ As Double

    NormX As Double
    NormY As Double
    NormZ As Double

    sourceJob As String
    sourceComponent As String
End Type

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
Private swAssy As Object

Private parts() As PartInfo
Private PartCount As Long

Private BomRows() As BomInfo
Private BomCount As Long

Private ExportRows() As ExportInfo
Private ExportCount As Long

Private PullcoreMatches() As PullcoreMatchInfo
Private PullcoreMatchCount As Long

Private RunLogPath As String
Private StartupLogPath As String
Private CurrentJobFolder As String
Private CurrentJobNumber As String

Private NetworkJobFolder As String
Private LocalJobFolder As String
Private BaseNativeAssemblyPath As String

Private CustomerNumber As String
Private DateCode As String
Private NamingSourceText As String

Private ExportFilePaths As Object
Private SpecialBomCadMatches As Object
Private SpecialBomCadQuoteNames As Object
Private PullcoreBestFitDimCache As Object
Private PullcoreBestFitAngleCache As Object
Private MatchStudioStlExported As Object

Private MacroStartTime As Date
Private StepStartTime As Date
Private CurrentStepName As String

Private MainCadOpenedByMacro As Boolean
Private MainCadTitleForClose As String

Private DxfFreezeDoc As Object
Private MainViewportGraphicsDisabled As Boolean

Private LastJobFailReason As String

Private CurrentDxfForce1to1 As Boolean
Private CurrentIdHolderDxfFromAssembly As Boolean
Private CurrentDxfKeepImportedOrientation As Boolean
Private CurrentDxfStraightenAngleRad As Double

Private CurrentPullcoreStraightenCadIndex As Long

Private PullCoreDimensionsReportPath As String
Private PullcoreIdOdHeightAxisUsed As String
Private DetectedPullcoreAnglesCsv As String
Private CadLibDominantSourceJob As String
Private LastBestFitFromPlainBBox As Boolean

Private FinalStlCoordFrameReady As Boolean
Private FinalStlCoordM(0 To 8) As Double

' True once *Top and *Front have been set from the TCP and pot faces. When this
' is set, the old world-axis top pick must NOT run - it would call
' UpdateStandardViews on a global-axis view and wipe out the corrected frame.
Private TcpNormalTopApplied As Boolean

' ---- WHOLE-BASE STRAIGHTEN STATE ----
Private BaseStraightenApplied As Boolean
Private BaseStraightenTotalDeg As Double
Private BaseStraightenAxesUsed As String

' Assembly-space point cloud of every component's bounding-box corners.
' Rebuilt once per straighten pass; every angle in the sweep is then pure
' arithmetic over this cloud instead of another trip through the SW API.
Private BaseCloudX() As Double
Private BaseCloudY() As Double
Private BaseCloudZ() As Double
Private BaseCloudN As Long

' ============================================================
' MAIN - BATCH CONTROLLER
' ============================================================


' ============================================================
' POTS AND HOLDERS + PYROPEL SETTINGS
' Declared up here because Option Explicit requires module-level Const
' statements to live in the declarations section, above the first
' procedure. The procedures themselves are at the bottom of the module.
' ============================================================

Private Const CREATE_POTS_HOLDERS_DXF_PACKAGE As Boolean = True
Private Const POTS_HOLDERS_FOLDER_NAME As String = "POTS AND HOLDERS"
Private Const POTS_DXF_FILE_TOKEN As String = "POTS"
Private Const HOLDERS_DXF_FILE_TOKEN As String = "HOLDERS ONLY"
Private Const POT_BLOCK_COLLECT_KEYS As String = "ID POT BLOCK|OD POT BLOCK|TOP POT BLOCK|BOTTOM POT BLOCK|BOT POT BLOCK|ID POT|OD POT|POT BLOCK|POTBLOCK"
Private Const HOLDER_BLOCK_COLLECT_KEYS As String = "ID HOLDER|OD HOLDER|TOP HOLDER|BOTTOM HOLDER|BOT HOLDER|IDTE HOLDER|ODTE HOLDER|HOLDER BLOCK|TOP CARRIER|BOTTOM CARRIER|ID CARRIER|OD CARRIER"
' Dropped the bare words CAM, KEY and PIN from this list. They match as
' substrings, so "TOP SPIN HOLDER" was being excluded on the "PIN" inside
' "SPIN". The remaining tokens are long enough not to collide.
Private Const POTS_HOLDERS_EXCLUDE_KEYS As String = "PYROPEL|TCP|BCP|CLAMPING|SMED|J BLOCK|J-BLOCK|JBLOCK|PULLCORE|PULL CORE|EJECTOR CAM|SPACER|WASHER|SCREW|DOWEL"
Private Const POTS_HOLDERS_DIM_REPORT As String = "Pots and Holders Dimensions.csv"
Private Const RUN_PYRO_AFTER_JOB As Boolean = True
Private Const PYRO_SCRIPT_FOLDER As String = "C:\CMS_Local_Workspace"
Private Const PYRO_PYTHON_EXE As String = "python"
Private Const PYRO_OUTPUT_SUFFIX As String = " PYRO"
Private Const PYRO_JOBS_LABEL As String = ""
Private Const PYRO_WAIT_FOR_FINISH As Boolean = True
Private Const LABEL_PULLCORE_FROM_HOLDER_COM As Boolean = True

Sub main()
On Error GoTo ErrHandler

    Set swApp = Application.SldWorks

    If RUN_SOLIDWORKS_INVISIBLE Then
        On Error Resume Next
        swApp.Visible = False
        On Error GoTo ErrHandler
    End If

    MacroStartTime = Now
    StartupLogPath = Environ$("USERPROFILE") & "\Desktop\CMS_XT_Export_STARTUP_Log.txt"
    RunLogPath = StartupLogPath

    LogLine "========================================"
    LogLine "MACRO STARTED"
    LogLine "Root path: " & ResolveRootJobPath()
    LogLine "Matching publish root: " & ResolveMatchingRoot()
    LogLine "Local workspace root: " & LOCAL_WORKSPACE_ROOT
    LogLine "Push outputs to network: " & CStr(PUSH_OUTPUTS_TO_NETWORK)
    LogLine "Run SolidWorks invisible: " & CStr(RUN_SOLIDWORKS_INVISIBLE)
    LogLine "Create Pull Core Dimensions Excel: " & CStr(CREATE_PULLCORE_DIMENSIONS_EXCEL)
    LogLine "Auto-select TCP top orientation: " & CStr(AUTO_SELECT_TCP_TOP_ORIENTATION)
    LogLine "Auto-label pullcore ID/OD by height: " & CStr(AUTO_LABEL_PULLCORE_ID_OD_BY_HEIGHT)
    LogLine "OD holder center rotation deg: " & CStr(OD_HOLDER_CENTER_ROTATION_DEG)
    LogLine "FORCE_ALL_DXF_VIEWS_1_TO_1: " & CStr(FORCE_ALL_DXF_VIEWS_1_TO_1)
    LogLine "========================================"

    Dim jobInput As String
    jobInput = Trim(InputBox("Enter one or multiple job numbers/search texts." & vbCrLf & _
                             "Examples:" & vbCrLf & _
                             "J8410" & vbCrLf & _
                             "J8410, J8411, J8412" & vbCrLf & _
                             "833200084 J8413", _
                             "CMS XT Export Batch"))

    If jobInput = "" Then GoTo NormalEnd

    Dim jobs As Collection
    Set jobs = ParseJobInputList(jobInput)

    If jobs Is Nothing Or jobs.count = 0 Then
        MsgBox "No valid job numbers entered.", vbExclamation
        GoTo NormalEnd
    End If

    Dim completed As Collection
    Dim failed As Collection

    Set completed = New Collection
    Set failed = New Collection

    Dim i As Long
    Dim jobText As String
    Dim ok As Boolean

    For i = 1 To jobs.count

        jobText = UCase(Trim(CStr(jobs(i))))

        If jobText <> "" Then

            RunLogPath = StartupLogPath
            LogLine "========================================"
            LogLine "BATCH ITEM " & i & "/" & jobs.count & ": " & jobText
            LogLine "========================================"

            ok = ProcessOneJob(jobText)

            If ok Then
                completed.Add jobText
            Else
                failed.Add jobText & IIf(LastJobFailReason <> "", "  ->  " & LastJobFailReason, "")
            End If

            CloseAllDocumentsSafely
            Set swModel = Nothing
            Set swAssy = Nothing
            Set ExportFilePaths = Nothing
            Set SpecialBomCadMatches = Nothing
            Set SpecialBomCadQuoteNames = Nothing
            Set PullcoreBestFitDimCache = Nothing
            Set PullcoreBestFitAngleCache = Nothing
    Set PullcoreBestFitAngleCache = Nothing
            Set MatchStudioStlExported = Nothing
            ReleaseSolidWorksMemory "after batch job"

            Erase parts
            Erase BomRows
            Erase ExportRows
            Erase PullcoreMatches

            PartCount = 0
            BomCount = 0
            ExportCount = 0
            PullcoreMatchCount = 0
            CurrentPullcoreStraightenCadIndex = 0
            PullCoreDimensionsReportPath = ""
            PullcoreIdOdHeightAxisUsed = ""
            DetectedPullcoreAnglesCsv = ""

            BaseStraightenApplied = False
            BaseStraightenTotalDeg = 0#
            BaseStraightenAxesUsed = ""
            BaseCloudN = 0
            TcpNormalTopApplied = False

            DoEvents
            WaitMilliseconds 100
            DoEvents

        End If

    Next i

    Dim summary As String
    summary = BuildBatchSummary(completed, failed)

    LogLine "========================================"
    LogLine "BATCH DONE"
    LogLine Replace(summary, vbCrLf, " | ")
    LogLine "========================================"

    CloseAllDocumentsSafely

    On Error Resume Next
    If Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo ErrHandler

    MsgBox summary, IIf(failed.count > 0, vbExclamation, vbInformation)

NormalEnd:
    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing And Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo 0
    Exit Sub

ErrHandler:
    LogLine "FATAL BATCH ERROR"
    LogLine "Step: " & CurrentStepName
    LogLine "Err " & Err.Number & ": " & Err.Description

    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing And Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo 0

    MsgBox "Macro batch error at step: " & CurrentStepName & vbCrLf & _
           Err.Description & vbCrLf & RunLogPath, vbCritical
End Sub

Private Function ProcessOneJob(ByVal jobSearchText As String) As Boolean
On Error GoTo ErrHandler

    ProcessOneJob = False
    LastJobFailReason = ""

    Erase parts
    Erase BomRows
    Erase ExportRows
    Erase PullcoreMatches

    PartCount = 0
    BomCount = 0
    ExportCount = 0
    PullcoreMatchCount = 0
    CurrentPullcoreStraightenCadIndex = 0
    PullCoreDimensionsReportPath = ""
    PullcoreIdOdHeightAxisUsed = ""
    DetectedPullcoreAnglesCsv = ""
    FinalStlCoordFrameReady = False

    BaseStraightenApplied = False
    BaseStraightenTotalDeg = 0#
    BaseStraightenAxesUsed = ""
    BaseCloudN = 0
    TcpNormalTopApplied = False

    Dim stlCoordI As Long
    For stlCoordI = 0 To 8
        FinalStlCoordM(stlCoordI) = 0#
    Next stlCoordI

    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
    MainViewportGraphicsDisabled = False

    Set swModel = Nothing
    Set swAssy = Nothing
    Set SpecialBomCadMatches = CreateObject("Scripting.Dictionary")
    Set SpecialBomCadQuoteNames = CreateObject("Scripting.Dictionary")
    Set PullcoreBestFitDimCache = CreateObject("Scripting.Dictionary")
    Set PullcoreBestFitAngleCache = CreateObject("Scripting.Dictionary")
    Set MatchStudioStlExported = CreateObject("Scripting.Dictionary")

    CurrentJobNumber = UCase(Trim(jobSearchText))
    CurrentJobFolder = ""
    NetworkJobFolder = ""
    LocalJobFolder = ""
    BaseNativeAssemblyPath = ""

    CustomerNumber = ""
    DateCode = ""
    NamingSourceText = ""

    ReDim parts(1 To 1)
    ReDim BomRows(1 To 1)
    ReDim ExportRows(1 To 1)
    ReDim PullcoreMatches(1 To 1)

    LogStart "Find network job folder"

    NetworkJobFolder = FindJobFolderByText(ResolveRootJobPath(), CurrentJobNumber)
    LogLine "Network job folder result: " & NetworkJobFolder

    If NetworkJobFolder = "" Then
        LogErrorText "Could not find job folder for: " & CurrentJobNumber
        GoTo CleanExit
    End If

    LogDone "Find network job folder"

    LogStart "Prepare local job workspace"

    PrepareLocalJobWorkspace NetworkJobFolder, CurrentJobNumber, LocalJobFolder

    If LocalJobFolder = "" Then
        LogErrorText "Could not create local workspace for: " & CurrentJobNumber
        GoTo CleanExit
    End If

    CleanPreviousGeneratedOutputs LocalJobFolder, CurrentJobNumber

    CurrentJobFolder = LocalJobFolder
    RunLogPath = CurrentJobFolder & "\CMS_XT_Export_Log.txt"

    PullCoreDimensionsReportPath = CurrentJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE

    LogLine "========================================"
    LogLine "CMS XT export started"
    LogLine "Job/search text: " & CurrentJobNumber
    LogLine "Local job folder: " & CurrentJobFolder
    LogLine "Pull Core Dimensions report path: " & PullCoreDimensionsReportPath
    LogLine "========================================"

    LogDone "Prepare local job workspace"

    Dim extractFolder As String
    extractFolder = CurrentJobFolder & "\" & EXTRACT_FOLDER_NAME

    LogStart "Extract ZIP files"
    EnsureFolderDeep extractFolder
    ExtractAllZipFilesInJobFolder CurrentJobFolder, extractFolder
    FlattenExtractedZipContentsIntoJobFolder CurrentJobFolder, extractFolder

    If DELETE_EXTRACTED_ZIP_AFTER_FLATTEN Then
        LogLine "Deleting " & EXTRACT_FOLDER_NAME & " folder after flattening."
        DeleteFolderSafe extractFolder
    End If

    LogDone "Extract ZIP files"

    LogStart "Determine output naming info"
    DetermineOutputNamingInfo CurrentJobFolder
    LogLine "Customer/program number: " & CustomerNumber
    LogLine "Date code: " & DateCode
    LogLine "Naming source: " & NamingSourceText
    LogDone "Determine output naming info"

    LogStart "Find CAD file"

    If USE_ACTIVE_SOLIDWORKS_DOC_FIRST Then
        Set swModel = swApp.ActiveDoc
    End If

    Dim cadPath As String

    If swModel Is Nothing Then

        Dim cadCandidates As Collection
        Set cadCandidates = FindAllCadModelsRanked(CurrentJobFolder)
        AppendCadCandidates cadCandidates, FindAllCadModelsRanked(extractFolder)

        If cadCandidates.count = 0 Then
            LogErrorText "No CAD file found."
            GoTo CleanExit
        End If

        LogLine "CAD candidates found: " & cadCandidates.count

        Dim ci As Long
        For ci = 1 To cadCandidates.count
            LogLine "  candidate " & ci & ": " & CStr(cadCandidates(ci))
        Next ci

        LogDone "Find CAD file"
        LogStart "Open CAD"

        For ci = 1 To cadCandidates.count
            cadPath = CStr(cadCandidates(ci))
            LogLine "Trying to open CAD candidate " & ci & "/" & cadCandidates.count & ": " & cadPath
            Set swModel = OpenCadFile(cadPath)

            If Not swModel Is Nothing Then
                LogLine "CAD opened successfully: " & cadPath
                Exit For
            End If

            LogLine "Candidate failed to open, trying next."
        Next ci

        If swModel Is Nothing Then
            LogErrorText "Open CAD failed; no CAD candidate opened."
            GoTo CleanExit
        End If

        MainCadOpenedByMacro = True
        MainCadTitleForClose = swModel.GetTitle

        LogLine "Opened CAD title: " & swModel.GetTitle
        LogDone "Open CAD"

    Else

        MainCadOpenedByMacro = False
        MainCadTitleForClose = swModel.GetTitle

        LogLine "Using active document: " & swModel.GetTitle
        LogDone "Find CAD file"

    End If

    Dim errs As Long
    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    EnsureSwHidden

    MainCadTitleForClose = swModel.GetTitle

    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    EnsureSwHidden

    DisableMainViewportGraphics

    If MATCH_STUDIO_FORCE_CARBON_STEEL And MATCH_STUDIO_APPLY_PHYSICAL_MATERIAL_TO_ALL_PARTS Then
        LogStart "Apply Match Studio carbon steel material"
        ApplyMatchStudioCarbonSteelToDocument swModel
        LogDone "Apply Match Studio carbon steel material"
    End If

    If HIDE_QUARTER_INCH_THICKNESS And PHYSICALLY_HIDE_250_BEFORE_SCAN Then
        LogStart "Physically hide .250 items"
        HideQuarterInchThicknessItems swModel
        LogDone "Physically hide .250 items"
    ElseIf HIDE_QUARTER_INCH_THICKNESS Then
        LogLine ".250 physical hide skipped for speed. Filtering by BOM/CAD logic."
    End If

    StabilizeActiveView swModel, 300

    ' Superseded by the top-clamping-plate Normal To step, which runs after the
    ' BOM match. Off by default; see ORIENT_TOP_BY_TCP_NORMAL_TO.
    If STRAIGHTEN_WHOLE_BASE_BEFORE_ORIENT Then
        LogStart "Straighten whole base to assembly axes"
        StraightenWholeBaseToAssemblyAxes swModel
        LogDone "Straighten whole base to assembly axes"
        StabilizeActiveView swModel, 200
    End If

    LogStart "Scan CAD"

    ScanActiveSolidWorksDocument
    SortPartsByVolumeDescending

    LogLine "CAD PartCount=" & PartCount
    LogDone "Scan CAD"

    If PartCount = 0 Then
        LogErrorText "No measurable CAD parts found."
        GoTo CleanExit
    End If

    WritePartDimensionCsv CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"

    Dim usedCadNamingLibraryOnly As Boolean
    usedCadNamingLibraryOnly = False

    LogStart "Find BOM"

    Dim bomPath As String
    bomPath = FindCustomerBomFile(CurrentJobFolder)

    If bomPath = "" Then

        LogLine "No BOM found. Trying CAD naming library fallback."
        LogDone "Find BOM"

        If USE_CAD_NAMING_LIBRARY Then
            LogStart "No-BOM CAD naming library match"

            If TryBuildExportRowsFromCadNamingLibrary() Then
                usedCadNamingLibraryOnly = True
                LogLine "No-BOM CAD naming library fallback succeeded. ExportCount=" & ExportCount
            Else
                LogErrorText "No BOM found and CAD naming library could not identify parts."
                GoTo CleanExit
            End If

            LogDone "No-BOM CAD naming library match"
        Else
            LogErrorText "No BOM found."
            GoTo CleanExit
        End If

    Else

        LogLine "BOM selected: " & bomPath
        LogDone "Find BOM"

        If LCase(GetFileExtension(bomPath)) = "pdf" Then

            If READ_PDF_BOM_WITH_PDFTOTEXT Then
                LogStart "Read PDF BOM with Poppler"
                ReadCustomerBomPdfUsingPdfToText bomPath
                LogDone "Read PDF BOM with Poppler"
            Else
                LogErrorText "PDF BOM found but PDF reading disabled."
                GoTo CleanExit
            End If

        Else

            LogStart "Read Excel BOM TURBO"
            ReadCustomerBom bomPath
            LogDone "Read Excel BOM TURBO"

        End If

        LogLine "BomCount=" & BomCount

        If BomCount = 0 Then

            LogLine "BOM was found but no usable rows were parsed. Trying CAD naming library fallback."

            If USE_CAD_NAMING_LIBRARY Then
                LogStart "No-BOM CAD naming library match"

                If TryBuildExportRowsFromCadNamingLibrary() Then
                    usedCadNamingLibraryOnly = True
                    LogLine "CAD naming library fallback succeeded after empty BOM parse. ExportCount=" & ExportCount
                Else
                    LogErrorText "No usable BOM rows found and CAD naming library could not identify parts."
                    GoTo CleanExit
                End If

                LogDone "No-BOM CAD naming library match"
            Else
                LogErrorText "No usable BOM rows found."
                GoTo CleanExit
            End If

        End If

    End If

    LogStart "Match BOM to CAD"

    If usedCadNamingLibraryOnly Then

        LogLine "BOM matching skipped. Export rows were created from CAD naming library."

    Else

        BuildExportRowsFromBom

        If ADD_TOP_BOT_INS_FROM_CAD_GEOMETRY Then
            LogStart "Add missing TOP/BOT INS from CAD geometry"
            AddMissingTopBotInsFromCadGeometry
            LogDone "Add missing TOP/BOT INS from CAD geometry"
        End If

        If LEARN_CAD_NAMING_LIBRARY_FROM_BOM Then
            LogStart "Learn CAD naming library from matched BOM job"
            LearnCadNamingLibraryFromCurrentJob
            LogDone "Learn CAD naming library from matched BOM job"
        End If

    End If

    LogStart "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    ' STEP 1: square the base off two real faces on the model.
    '   the big flat face of the top clamping plate  -> which way is UP
    '   a side face of a pot block                   -> which way is FRONT
    ' No geometry moves; only the standard views are redefined, and the standard
    ' views are what everything downstream reads. The BOM match has already run,
    ' so both components can be found by their quote names.
    '
    ' TcpNormalTopApplied then tells the step below to leave the result alone.
    If ORIENT_TOP_BY_TCP_NORMAL_TO Then
        LogStart "Square *Top and *Front from the TCP and pot faces"
        TcpNormalTopApplied = TryOrientFromTcpTopAndPotFront(swModel)
        LogDone "Square *Top and *Front from the TCP and pot faces"
    End If

    ' STEP 2: standard orientation process, unchanged. *Top is now square to the
    ' clamping plate, so this only has to decide which way round it goes.
    '   EnsureCmsTopOrientationFromMatchedTcpBcp -> TCP on top, not bottom
    '     -> DefineStandardFrontFromHolderAndPotCom
    '        -> EnsurePotBlocksCloserThanHoldersInActiveView
    '           -> pots end up closer to *Front than the holders
    If BaseStraightenApplied Then
        LogLine "Standard orientation running on a base straightened by " & _
                Format(BaseStraightenTotalDeg, "0.000") & " deg (" & BaseStraightenAxesUsed & ")."
    End If

    EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE

    ' IMPORTANT:
    ' Capture the final redefined SolidWorks standard-view coordinate system.
    ' This is the coordinate system all STLs must use.
    CaptureFinalStandardViewsForStlCoordinateSystem swModel

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    SaveImportedBaseAssemblyToLocal swModel
    MainCadTitleForClose = swModel.GetTitle
    LogDone "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    WriteExportCheckCsv CurrentJobFolder & "\XT_Export_BOM_Match_Report.csv"

    Dim jobSignaturePath As String
    jobSignaturePath = CurrentJobFolder & "\" & JOB_SIGNATURE_REPORT_FILE

    WriteJobSignatureCsv jobSignaturePath
    UploadJobSignatureToElgin jobSignaturePath

    WriteExportLogBomSummary

    If CREATE_PULLCORE_DIMENSIONS_EXCEL And PullcoreMatchCount > 0 Then
        LogStart "Write Pull Core Dimensions Excel"
        WritePullCoreDimensionsExcel PullCoreDimensionsReportPath
        LogDone "Write Pull Core Dimensions Excel"
    End If

    LogLine "ExportCount=" & ExportCount
    LogLine "PullcoreMatchCount=" & PullcoreMatchCount
    LogDone "Match BOM to CAD"

    If ExportCount = 0 And PullcoreMatchCount = 0 And HasSpecialPackageCadMatches() = False Then
        LogErrorText "No matched export rows."
        GoTo CleanExit
    End If

    Dim outputFolder As String
    outputFolder = CurrentJobFolder & "\" & CurrentJobNumber & OUTPUT_FOLDER_SUFFIX

    EnsureFolderDeep outputFolder

    Set ExportFilePaths = CreateObject("Scripting.Dictionary")

    LogStart "Export XT/DXF files"
    ExportMatchedPartsAsXt outputFolder
    LogDone "Export XT/DXF files"

    If CREATE_INDIVIDUAL_DXFS Then
        LogStart "Export individual four-view DXFs"
        ExportIndividualHolderAndClampingDxfs outputFolder
        LogDone "Export individual four-view DXFs"
    End If

    LogStart "Restore visibility/state"

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ShowAllPartBodies swModel

    If HIDE_QUARTER_INCH_THICKNESS And PHYSICALLY_HIDE_250_BEFORE_SCAN Then
        HideQuarterInchThicknessItems swModel
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 200

    LogDone "Restore visibility/state"

    ' Move loose imported SLDPRT files into Base folder.
    ' Keep BASE and HOLDERS package files in main job folder.
    ' Collect PDFs/Excel into J#### pdfs; publish after DXFs exist.
    CloseAllDocumentsSafely
    Set swModel = Nothing
    Set swAssy = Nothing

    MoveLooseSolidWorksPartsToBaseFolder
    KeepBaseAndHoldersPackageFilesInMainFolder
    CollectPdfExcelFilesToJobPdfsFolder
    DeletePreviewImageFilesFromJob
    PublishJobOutputs

    If PUSH_OUTPUTS_TO_NETWORK Then
        LogStart "Push finished outputs back to network"
        PushFinishedOutputsBackToNetwork
        LogDone "Push finished outputs back to network"
    Else
        LogLine "Network push DISABLED. Outputs kept locally only at:"
        LogLine "  " & CurrentJobFolder
    End If

    LogLine "========================================"
    LogLine "DONE JOB. Total seconds=" & DateDiff("s", MacroStartTime, Now)
    LogLine "Export log file: " & RunLogPath
    LogLine "Local output folder: " & outputFolder
    LogLine "Pull Core Dimensions report: " & PullCoreDimensionsReportPath
    LogLine "========================================"

    ProcessOneJob = True

CleanExit:
    On Error Resume Next

    RestoreMainViewportGraphics

    If Not swApp Is Nothing Then
        If Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    End If

    CloseCurrentJobCadIfNeeded
    CloseAllDocumentsSafely

    Set swAssy = Nothing
    Set swModel = Nothing
    Set ExportFilePaths = Nothing
    Set SpecialBomCadMatches = Nothing
    Set SpecialBomCadQuoteNames = Nothing
    Set PullcoreBestFitDimCache = Nothing
    Set PullcoreBestFitAngleCache = Nothing
    Set MatchStudioStlExported = Nothing

    CurrentPullcoreStraightenCadIndex = 0
    PullcoreIdOdHeightAxisUsed = ""
    DetectedPullcoreAnglesCsv = ""

    Exit Function

ErrHandler:
    LogLine "FATAL JOB ERROR"
    LogLine "Job: " & CurrentJobNumber
    LogLine "Step: " & CurrentStepName
    LogLine "Err " & Err.Number & ": " & Err.Description
    LastJobFailReason = "Step '" & CurrentStepName & "' - Err " & Err.Number & ": " & Err.Description
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

    s = "Batch complete." & vbCrLf & vbCrLf
    s = s & "Completed: " & completed.count & vbCrLf

    If completed.count > 0 Then
        For i = 1 To completed.count
            s = s & "  - " & CStr(completed(i)) & vbCrLf
        Next i
    End If

    s = s & vbCrLf & "Failed: " & failed.count & vbCrLf

    If failed.count > 0 Then
        For i = 1 To failed.count
            s = s & "  - " & CStr(failed(i)) & vbCrLf
        Next i
    End If

    BuildBatchSummary = s
    Exit Function

ErrHandler:
    BuildBatchSummary = "Batch complete."
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
        If MainCadTitleForClose <> "" Then
            LogLine "Closing main CAD document for finished job: " & MainCadTitleForClose
            swApp.CloseDoc MainCadTitleForClose
        End If
    Else
        If MainCadTitleForClose <> "" Then
            LogLine "Main CAD was not opened by macro. Leaving open: " & MainCadTitleForClose
        End If
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
        swApp.CloseDoc swDoc.GetTitle
        Set swDoc = nextDoc
    Loop

    Set swDoc = Nothing
    Set nextDoc = Nothing

    ReleaseSolidWorksMemory "CloseAllDocumentsSafely"
End Sub

Private Sub ReleaseSolidWorksMemory(Optional ByVal reason As String = "")
On Error Resume Next

    If Not swModel Is Nothing Then swModel.ClearSelection2 True

    If Not swApp Is Nothing Then
        swApp.CommandInProgress = False
        swApp.UserControl = False
    End If

    Set DxfFreezeDoc = Nothing

    DoEvents
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

Private Sub CleanPreviousGeneratedOutputs(ByVal jobFolder As String, ByVal jobNumber As String)
On Error GoTo ErrHandler

    If jobFolder = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(jobFolder) = False Then Exit Sub

    LogLine "Cleaning previous generated outputs for rerun: " & jobFolder

    DeleteFolderSafe jobFolder & "\" & jobNumber & OUTPUT_FOLDER_SUFFIX
    DeleteFolderSafe jobFolder & "\" & jobNumber & " " & J_BLOCK_FOLDER_NAME
    DeleteFolderSafe jobFolder & "\" & jobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME
    DeleteFolderSafe jobFolder & "\" & jobNumber & " " & PULLCORE_STOP_FOLDER_NAME
    DeleteFolderSafe jobFolder & "\" & jobNumber & " " & POTS_HOLDERS_FOLDER_NAME
    DeleteFolderSafe jobFolder & "\" & jobNumber & PYRO_OUTPUT_SUFFIX
    DeleteFolderSafe jobFolder & "\" & jobNumber & " MISC DETAILS"
    DeleteFolderSafe jobFolder & "\" & jobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX

    DeleteOldBaseHolderRootFiles jobFolder, jobNumber

    Exit Sub

ErrHandler:
    LogLine "CleanPreviousGeneratedOutputs error: " & Err.Description
End Sub

Private Sub DeleteOldBaseHolderRootFiles(ByVal jobFolder As String, ByVal jobNumber As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(jobFolder) = False Then Exit Sub

    Dim folder As Object
    Set folder = fso.GetFolder(jobFolder)

    Dim fileList As Collection
    Set fileList = New Collection

    Dim f As Object
    Dim n As String
    Dim ext As String

    For Each f In folder.Files

        n = UCase(f.name)
        ext = LCase(fso.GetExtensionName(f.path))

        If Left(n, Len(UCase(jobNumber))) = UCase(jobNumber) Then

            If InStr(n, "_BASE_") > 0 Or InStr(n, "_HOLDERS_") > 0 Then

                Select Case ext
                    Case "x_t", "igs", "iges", "easm", "dxf", "stl"
                        fileList.Add f.path
                End Select

            End If

        End If

    Next f

    Dim i As Long

    For i = 1 To fileList.count
        fso.DeleteFile CStr(fileList(i)), True
        LogLine "Deleted old generated root package file: " & CStr(fileList(i))
    Next i

    Exit Sub

ErrHandler:
    LogLine "DeleteOldBaseHolderRootFiles error: " & Err.Description
End Sub

' ============================================================
' LOCAL STAGING / NETWORK PUSH-BACK
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
        LogLine "Clearing old local workspace: " & localFolderOut
        On Error Resume Next
        fso.DeleteFolder localFolderOut, True
        On Error GoTo ErrHandler
    End If

    EnsureFolderDeep localFolderOut

    If fso.FolderExists(localFolderOut) = False Then
        LogLine "Could not create local workspace folder: " & localFolderOut
        localFolderOut = ""
        Exit Sub
    End If

    LogLine "Copying network job folder to local workspace."
    LogLine "  FROM: " & sourceNetworkFolder
    LogLine "  TO  : " & localFolderOut

    If CopyFolderWithRobocopy(sourceNetworkFolder, localFolderOut) Then
        LogLine "Local copy complete (robocopy)."
    Else
        LogLine "Robocopy unavailable/failed. Using VBA copy fallback."
        CopyFolderContentsFiltered sourceNetworkFolder, localFolderOut
        LogLine "Local copy complete (VBA fallback)."
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

Private Sub CopyFolderContentsFiltered(ByVal sourceFolder As String, ByVal destFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(sourceFolder) Then Exit Sub

    EnsureFolderDeep destFolder

    Dim src As Object
    Set src = fso.GetFolder(sourceFolder)

    Dim file As Object
    For Each file In src.Files
        fso.CopyFile file.path, destFolder & "\" & file.name, True
    Next file

    Dim subFolder As Object
    Dim folderNameUpper As String

    For Each subFolder In src.SubFolders

        folderNameUpper = UCase(subFolder.name)

        If folderNameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextFolder
        If InStr(folderNameUpper, " PRINTS") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "J BLOCK") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "PULLCORE") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "PULL CORE") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "PYROPEL") > 0 Then GoTo NextFolder
        If folderNameUpper = "BASE" Then GoTo NextFolder
        If InStr(folderNameUpper, " BASE") > 0 Then GoTo NextFolder

        EnsureFolderDeep destFolder & "\" & subFolder.name
        CopyFolderContentsFiltered subFolder.path, destFolder & "\" & subFolder.name

NextFolder:
    Next subFolder

    Exit Sub

ErrHandler:
    LogLine "CopyFolderContentsFiltered error: " & Err.Description
End Sub

Private Sub PushFinishedOutputsBackToNetwork()
On Error GoTo ErrHandler

    If PUSH_OUTPUTS_TO_NETWORK = False Then
        LogLine "PushFinishedOutputsBackToNetwork called but push is disabled. Skipping."
        Exit Sub
    End If

    If NetworkJobFolder = "" Then
        LogLine "Push back skipped: NetworkJobFolder is blank."
        Exit Sub
    End If

    If CurrentJobFolder = "" Then
        LogLine "Push back skipped: CurrentJobFolder is blank."
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(CurrentJobFolder) = False Then
        LogLine "Push back skipped: local CurrentJobFolder does not exist."
        Exit Sub
    End If

    EnsureFolderDeep NetworkJobFolder

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & OUTPUT_FOLDER_SUFFIX, _
                       NetworkJobFolder & "\" & CurrentJobNumber & OUTPUT_FOLDER_SUFFIX

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & J_BLOCK_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & J_BLOCK_FOLDER_NAME

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_STOP_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_STOP_FOLDER_NAME

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & POTS_HOLDERS_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & POTS_HOLDERS_FOLDER_NAME

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & PYRO_OUTPUT_SUFFIX, _
                       NetworkJobFolder & "\" & CurrentJobNumber & PYRO_OUTPUT_SUFFIX

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX, _
                       NetworkJobFolder & "\" & CurrentJobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX

    CopyFolderIfExists LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber) & "\" & CurrentJobNumber & " Base", _
                       NetworkJobFolder & "\" & CurrentJobNumber & " Base"

    CopyRootOutputFilesToNetwork CurrentJobFolder, NetworkJobFolder

    CopyFileIfExists CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv", _
                     NetworkJobFolder & "\XT_Export_CAD_Dimensions.csv"

    CopyFileIfExists CurrentJobFolder & "\XT_Export_BOM_Match_Report.csv", _
                     NetworkJobFolder & "\XT_Export_BOM_Match_Report.csv"

    CopyFileIfExists CurrentJobFolder & "\XT_Export_BOM_PDF_Text.txt", _
                     NetworkJobFolder & "\XT_Export_BOM_PDF_Text.txt"

    CopyFileIfExists CurrentJobFolder & "\" & JOB_SIGNATURE_REPORT_FILE, _
                     NetworkJobFolder & "\" & JOB_SIGNATURE_REPORT_FILE

    CopyFileIfExists CurrentJobFolder & "\CMS_XT_Export_Log.txt", _
                     NetworkJobFolder & "\CMS_XT_Export_Log.txt"

    CopyFileIfExists CurrentJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE, _
                     NetworkJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE

    LogLine "Push back to network complete."
    Exit Sub

ErrHandler:
    LogLine "PushFinishedOutputsBackToNetwork error: " & Err.Description
End Sub

Private Sub CopyFolderIfExists(ByVal sourceFolder As String, ByVal destFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(sourceFolder) = False Then
        LogLine "CopyFolderIfExists skipped missing folder: " & sourceFolder
        Exit Sub
    End If

    If fso.FolderExists(destFolder) Then
        LogLine "Deleting old network output folder: " & destFolder
        fso.DeleteFolder destFolder, True
    End If

    EnsureFolderDeep destFolder
    CopyFolderContents sourceFolder, destFolder

    LogLine "Copied folder to network:"
    LogLine "  FROM: " & sourceFolder
    LogLine "  TO  : " & destFolder
    Exit Sub

ErrHandler:
    LogLine "CopyFolderIfExists error: " & Err.Description
End Sub

Private Sub CopyFileIfExists(ByVal sourceFile As String, ByVal destFile As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(sourceFile) = False Then Exit Sub

    EnsureFolderDeep fso.GetParentFolderName(destFile)
    fso.CopyFile sourceFile, destFile, True
    LogLine "Copied file to network: " & destFile
    Exit Sub

ErrHandler:
    LogLine "CopyFileIfExists error: " & Err.Description
End Sub

Private Sub CopyRootOutputFilesToNetwork(ByVal localRoot As String, ByVal networkRoot As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(localRoot) = False Then Exit Sub

    Dim folder As Object
    Set folder = fso.GetFolder(localRoot)

    Dim file As Object
    Dim ext As String
    Dim nameUpper As String

    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.path))
        nameUpper = UCase(file.name)

        If Left(nameUpper, Len(UCase(CurrentJobNumber))) = UCase(CurrentJobNumber) Then
            Select Case ext
                Case "x_t", "igs", "iges", "easm", "dxf", "stl", "sldasm"
                    fso.CopyFile file.path, networkRoot & "\" & file.name, True
                    LogLine "Copied root output file to network: " & file.name
            End Select
        End If

    Next file

    Exit Sub

ErrHandler:
    LogLine "CopyRootOutputFilesToNetwork error: " & Err.Description
End Sub

' ============================================================
' OPEN CAD / BASE SAVE / GRAPHICS / ORIENTATION
' ============================================================

Private Function OpenCadFile(ByVal cadPath As String) As Object
On Error GoTo ErrHandler

    LogLine "OpenCadFile: " & cadPath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(cadPath) Then
        LogLine "OpenCadFile: file does not exist."
        Set OpenCadFile = Nothing
        Exit Function
    End If

    LogLine "CAD size bytes: " & fso.GetFile(cadPath).Size

    Dim ext As String
    ext = LCase(fso.GetExtensionName(cadPath))

    Dim errs As Long
    Dim warns As Long
    Dim importErrors As Long
    Dim m As Object

    Set m = Nothing

    If ext = "sldasm" Then

        errs = 0
        warns = 0
        Set m = swApp.OpenDoc6(cadPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
        LogLine "OpenDoc6(ASSEMBLY) errs=" & errs & " warns=" & warns & " ok=" & CStr(Not m Is Nothing)

    ElseIf ext = "sldprt" Then

        errs = 0
        warns = 0
        Set m = swApp.OpenDoc6(cadPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
        LogLine "OpenDoc6(PART) errs=" & errs & " warns=" & warns & " ok=" & CStr(Not m Is Nothing)

    ElseIf ext = "slddrw" Then

        LogLine "OpenCadFile: drawing file skipped."
        Set m = Nothing

    Else

        importErrors = 0
        Set m = swApp.LoadFile4(cadPath, "r", Nothing, importErrors)
        LogLine "LoadFile4 importErrors=" & importErrors & " ok=" & CStr(Not m Is Nothing)

        If m Is Nothing Then
            importErrors = 0
            Set m = swApp.LoadFile4(cadPath, "", Nothing, importErrors)
            LogLine "LoadFile4(no opts) importErrors=" & importErrors & " ok=" & CStr(Not m Is Nothing)
        End If

        If m Is Nothing Then
            errs = 0
            warns = 0
            Set m = swApp.OpenDoc6(cadPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
            LogLine "Fallback OpenDoc6(ASSEMBLY) errs=" & errs & " ok=" & CStr(Not m Is Nothing)
        End If

        If m Is Nothing Then
            errs = 0
            warns = 0
            Set m = swApp.OpenDoc6(cadPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
            LogLine "Fallback OpenDoc6(PART) errs=" & errs & " ok=" & CStr(Not m Is Nothing)
        End If

    End If

    If m Is Nothing Then
        LogLine "OpenCadFile: ALL open methods failed for ext '" & ext & "'."
    End If

    Set OpenCadFile = m
    Exit Function

ErrHandler:
    LogLine "OpenCadFile error: " & Err.Description
    Set OpenCadFile = Nothing
End Function

Private Sub SaveImportedBaseAssemblyToLocal(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub

    If model.GetType <> swDocASSEMBLY Then
        LogLine "SaveImportedBaseAssemblyToLocal skipped: model is not an assembly."
        Exit Sub
    End If

    Dim baseFolder As String
    baseFolder = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber) & "\" & CurrentJobNumber & " Base"

    EnsureFolderDeep baseFolder

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    BaseNativeAssemblyPath = baseFolder & "\" & _
                             CurrentJobNumber & "_BASE_" & _
                             custToken & "_" & dateToken & ".sldasm"

    LogLine "Saving imported BASE assembly locally as native SLDASM:"
    LogLine "  " & BaseNativeAssemblyPath

    On Error Resume Next
    model.ResolveAllLightWeightComponents True
    On Error GoTo ErrHandler

    UnsuppressAllAssemblyComponents model
    ShowAllAssemblyComponents model

    ApplyCmsTopView model
    StabilizeActiveView model, 50

    Dim errs As Long
    Dim warns As Long

    model.Extension.SaveAs3 BaseNativeAssemblyPath, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent, _
                            Nothing, Nothing, errs, warns

    LogLine "BASE native SLDASM save done. Errors=" & errs & " Warnings=" & warns

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(BaseNativeAssemblyPath) = False Then
        LogLine "WARNING: BASE native SLDASM was not created."
    Else
        LogLine "BASE native SLDASM created successfully."
    End If

    Exit Sub

ErrHandler:
    LogLine "SaveImportedBaseAssemblyToLocal error: " & Err.Description
End Sub

Private Sub DisableMainViewportGraphics()
On Error Resume Next

    If DISABLE_MAIN_VIEWPORT_GRAPHICS = False Then Exit Sub
    If swModel Is Nothing Then Exit Sub

    Dim swView As Object
    Set swView = swModel.ActiveView

    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = False
        MainViewportGraphicsDisabled = True
        LogLine "Main 3D viewport graphics updates disabled."
    End If
End Sub

Private Sub RestoreMainViewportGraphics()
On Error Resume Next

    If swModel Is Nothing Then Exit Sub

    Dim swView As Object
    Set swView = swModel.ActiveView

    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = True
    End If

    If MainViewportGraphicsDisabled Then
        LogLine "Main 3D viewport graphics updates restored."
    End If

    MainViewportGraphicsDisabled = False
End Sub

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

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = tcpX - bcpX
    dY = tcpY - bcpY
    dZ = tcpZ - bcpZ

    Dim ax As Double
    Dim ay As Double
    Dim az As Double

    ax = Abs(dX)
    ay = Abs(dY)
    az = Abs(dZ)

    LogLine sourceLabel & " TCP-BCP separation: " & _
            "X=" & FormatNumberForCsv(dX) & _
            " Y=" & FormatNumberForCsv(dY) & _
            " Z=" & FormatNumberForCsv(dZ)

    Dim axisName As String
    Dim tcpHigh As Boolean

    If ay >= ax And ay >= az Then
        axisName = "Y"
        tcpHigh = (dY >= 0#)
    ElseIf az >= ax And az >= ay Then
        axisName = "Z"
        tcpHigh = (dZ >= 0#)
    Else
        axisName = "X"
        tcpHigh = (dX >= 0#)
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

    ' *Top and *Front were already built from the TCP and pot faces. Do NOT run
    ' the world-axis pick below: it calls ShowNamedView2 on a global-axis view and
    ' then UpdateStandardViews, which would overwrite the corrected frame with an
    ' axis-aligned one. That is what put the crooked *Top back last run.
    '
    ' Still worth doing is the cheap sanity pass: make sure the TCP side is the
    ' one facing up rather than the BCP side, then rebuild CMS_TOP.
    If TcpNormalTopApplied Then

        LogLine "*Top/*Front already squared from the TCP and pot faces. " & _
                "Skipping the world-axis top pick so it cannot overwrite them."

        EnsureNormalTopLooksAtTcpSide model


        If AUTO_ALIGN_TOP_SPIN_FROM_HOLDER_GEOMETRY Then
            If AlignTopSpinFromHolderGeometry(model) Then
                LogLine "TCP Normal To path: top-view spin aligned from holder geometry."
            Else
                LogLine "WARNING: TCP Normal To path could not align top-view spin from holder geometry."
            End If
        End If



        ' PATCH:
        ' TCP Normal To already fixed *Top. Now use the old reliable front logic:
        ' choose a standard front and flip if needed so the pot blocks are in front
        ' of the holders. Do not Normal To the pot face.
        If AUTO_DEFINE_FRONT_FROM_HOLDER_POT_COM Then
            If DefineStandardFrontFromHolderAndPotCom(model) Then
                LogLine "TCP Normal To path: Standard *Front defined from holder/pot COM; pots forced in front."
            Else
                LogLine "WARNING: TCP Normal To path could not define *Front from holder/pot COM."
            End If
        End If

        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 100

        On Error Resume Next
        model.DeleteNamedView CMS_TOP_VIEW_NAME
        Err.Clear
        model.NameView CMS_TOP_VIEW_NAME
        Err.Clear
        On Error GoTo ErrHandler

        LogLine "CMS_TOP named view saved from the TCP/pot corrected *Top."

        model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
        StabilizeActiveView model, 100

        Exit Sub

    End If

    Dim oriented As Boolean
    oriented = False

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
        LogLine "Matched top-side orientation failed from holder/pot/ins/TCP pairs."
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

Private Sub BuildFrontOrientIndexCollections(ByRef holderIndexes As Collection, _
                                             ByRef potIndexes As Collection)
On Error GoTo ErrHandler

    Set holderIndexes = New Collection
    Set potIndexes = New Collection

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

    Exit Sub

ErrHandler:
    LogLine "BuildFrontOrientIndexCollections error: " & Err.Description
    Set holderIndexes = New Collection
    Set potIndexes = New Collection
End Sub

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

Private Function TryAverageAxisValueForCadIndexes(ByVal cadIndexes As Collection, _
                                                  ByVal axisName As String, _
                                                  ByVal preferMassCenter As Boolean, _
                                                  ByRef avgVal As Double) As Boolean
On Error GoTo ErrHandler

    TryAverageAxisValueForCadIndexes = False
    avgVal = 0#

    If cadIndexes Is Nothing Then Exit Function
    If cadIndexes.count = 0 Then Exit Function

    Dim total As Double
    Dim countVal As Long

    total = 0#
    countVal = 0

    Dim i As Long
    Dim cadIdx As Long

    For i = 1 To cadIndexes.count

        cadIdx = CLng(cadIndexes(i))

        Dim px As Double
        Dim py As Double
        Dim pz As Double

        If TryGetCadCenterPointInches(cadIdx, preferMassCenter, px, py, pz) Then

            Select Case UCase(Trim(axisName))

                Case "X"
                    total = total + px
                    countVal = countVal + 1

                Case "Y"
                    total = total + py
                    countVal = countVal + 1

                Case "Z"
                    total = total + pz
                    countVal = countVal + 1

            End Select

        End If

    Next i

    If countVal > 0 Then
        avgVal = total / CDbl(countVal)
        TryAverageAxisValueForCadIndexes = True
    End If

    Exit Function

ErrHandler:
    TryAverageAxisValueForCadIndexes = False
End Function

Private Function TryGetCadCenterPointInches(ByVal cadIdx As Long, _
                                            ByVal preferMassCenter As Boolean, _
                                            ByRef px As Double, _
                                            ByRef py As Double, _
                                            ByRef pz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetCadCenterPointInches = False

    px = 0#
    py = 0#
    pz = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    If preferMassCenter Then
        If parts(cadIdx).hasMassCenter Then
            px = parts(cadIdx).MassCenterX
            py = parts(cadIdx).MassCenterY
            pz = parts(cadIdx).MassCenterZ
            TryGetCadCenterPointInches = True
            Exit Function
        End If
    End If

    If parts(cadIdx).hasAsmCenter Then
        px = parts(cadIdx).AsmCenterX
        py = parts(cadIdx).AsmCenterY
        pz = parts(cadIdx).AsmCenterZ
        TryGetCadCenterPointInches = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    TryGetCadCenterPointInches = False
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

        Dim vx As Double
        Dim vy As Double

        vx = (xs(i) * CDbl(mView(0))) + _
             (ys(i) * CDbl(mView(3))) + _
             (zs(i) * CDbl(mView(6)))

        vy = (xs(i) * CDbl(mView(1))) + _
             (ys(i) * CDbl(mView(4))) + _
             (zs(i) * CDbl(mView(7)))

        If firstPoint Then
            minX = vx
            maxX = vx
            minY = vy
            maxY = vy
            firstPoint = False
        Else
            If vx < minX Then minX = vx
            If vx > maxX Then maxX = vx
            If vy < minY Then minY = vy
            If vy > maxY Then maxY = vy
        End If

    Next i

    viewWIn = Abs(maxX - minX) * INCHES_PER_METER
    viewHIn = Abs(maxY - minY) * INCHES_PER_METER

    TryGetComponentViewWidthHeightInches = (viewWIn > 0# And viewHIn > 0#)
    Exit Function

ErrHandler:
    TryGetComponentViewWidthHeightInches = False
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

' ============================================================
' PATCH: GEOMETRY-BASED TOP SPIN ALIGNMENT
' ------------------------------------------------------------
' TCP Normal To fixes the plate tilt, but Normal To does not define the
' in-plane roll/spin. The pot/front check only proves front-vs-back depth;
' it cannot distinguish a correct view from a 90-degree-spun top view.
'
' This routine fixes that missing roll geometrically:
'   1. Show corrected SolidWorks *Top.
'   2. Try top-view rotations 0, 90, 180, 270.
'   3. Project the largest holder into the active view.
'   4. Pick the spin where the holder's projected width is largest.
'      That makes the holder long direction run left-to-right on the screen.
'   5. Persist that orientation as SolidWorks *Top.
'
' This is not a hardcoded 90. It chooses the required quarter-turn from the
' live holder projection after TCP Normal To.
' ============================================================
' ============================================================
' PATCH: EXACT ACTIVE-VIEW AXIS SETTER
' ------------------------------------------------------------
' SolidWorks ViewRotateplusz is NOT a 90-degree quarter turn in this setup.
' It rotates by the user's small view-increment angle. These helpers set the
' active view axes directly so the spin candidates are true 0/90/180/270 turns.
' ============================================================

Private Function TrySetActiveViewAxesExact(ByVal model As Object, _
                                           ByVal xx As Double, ByVal xy As Double, ByVal xz As Double, _
                                           ByVal yx As Double, ByVal yy As Double, ByVal yz As Double, _
                                           ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double) As Boolean
On Error GoTo ErrHandler

    TrySetActiveViewAxesExact = False

    If model Is Nothing Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then
        LogLine "TrySetActiveViewAxesExact: no active view."
        Exit Function
    End If

    Dim mathUtil As Object
    Set mathUtil = swApp.GetMathUtility

    If mathUtil Is Nothing Then
        LogLine "TrySetActiveViewAxesExact: MathUtility unavailable."
        Exit Function
    End If

    On Error Resume Next
    swView.EnableGraphicsUpdate = True
    Err.Clear
    On Error GoTo ErrHandler

    If TrySetActiveViewAxesExactOneLayout(model, mathUtil, _
                                          xx, xy, xz, yx, yy, yz, dX, dY, dZ, False) Then
        TrySetActiveViewAxesExact = True
        Exit Function
    End If

    If TrySetActiveViewAxesExactOneLayout(model, mathUtil, _
                                          xx, xy, xz, yx, yy, yz, dX, dY, dZ, True) Then
        TrySetActiveViewAxesExact = True
        Exit Function
    End If

    LogLine "TrySetActiveViewAxesExact: both matrix layouts failed verification."
    Exit Function

ErrHandler:
    LogLine "TrySetActiveViewAxesExact error: " & Err.Description
    TrySetActiveViewAxesExact = False
End Function

Private Function TrySetActiveViewAxesExactOneLayout(ByVal model As Object, _
                                                    ByVal mathUtil As Object, _
                                                    ByVal xx As Double, ByVal xy As Double, ByVal xz As Double, _
                                                    ByVal yx As Double, ByVal yy As Double, ByVal yz As Double, _
                                                    ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, _
                                                    ByVal transposedLayout As Boolean) As Boolean
On Error GoTo ErrHandler

    TrySetActiveViewAxesExactOneLayout = False

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    Dim arr(0 To 15) As Double

    If transposedLayout Then
        arr(0) = xx: arr(1) = xy: arr(2) = xz
        arr(3) = yx: arr(4) = yy: arr(5) = yz
        arr(6) = dX: arr(7) = dY: arr(8) = dZ
    Else
        arr(0) = xx: arr(1) = yx: arr(2) = dX
        arr(3) = xy: arr(4) = yy: arr(5) = dY
        arr(6) = xz: arr(7) = yz: arr(8) = dZ
    End If

    arr(9) = 0#
    arr(10) = 0#
    arr(11) = 0#
    arr(12) = 1#
    arr(13) = 0#
    arr(14) = 0#
    arr(15) = 0#

    Dim xf As Object
    Set xf = mathUtil.CreateTransform(arr)

    If xf Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set swView.Orientation3 = xf
    If Err.Number <> 0 Then
        Err.Clear
        swView.Orientation3 = xf
    End If
    Err.Clear
    model.GraphicsRedraw2
    DoEvents
    WaitMilliseconds 100
    model.GraphicsRedraw2
    DoEvents
    On Error GoTo ErrHandler

    Dim v As Variant
    v = swView.Orientation3.ArrayData

    If IsArray(v) = False Then Exit Function
    If UBound(v) < 8 Then Exit Function

    Dim rxX As Double, rxY As Double, rxZ As Double
    Dim ryX As Double, ryY As Double, ryZ As Double
    Dim rdX As Double, rdY As Double, rdZ As Double

    rxX = CDbl(v(0)): rxY = CDbl(v(3)): rxZ = CDbl(v(6))
    ryX = CDbl(v(1)): ryY = CDbl(v(4)): ryZ = CDbl(v(7))
    rdX = CDbl(v(2)): rdY = CDbl(v(5)): rdZ = CDbl(v(8))

    NormalizeStraightenVector rxX, rxY, rxZ
    NormalizeStraightenVector ryX, ryY, ryZ
    NormalizeStraightenVector rdX, rdY, rdZ

    Dim dotX As Double
    Dim dotY As Double
    Dim dotD As Double

    dotX = rxX * xx + rxY * xy + rxZ * xz
    dotY = ryX * yx + ryY * yy + ryZ * yz
    dotD = rdX * dX + rdY * dY + rdZ * dZ

    If dotX > 0.999 And dotY > 0.999 And dotD > 0.999 Then
        LogLine "TrySetActiveViewAxesExact: view write verified" & _
                IIf(transposedLayout, " using transposed layout.", " using normal layout.")
        TrySetActiveViewAxesExactOneLayout = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    TrySetActiveViewAxesExactOneLayout = False
End Function
' ============================================================
' PATCH HELPER: MEASURE VIEW ROTATE Z INCREMENT
' ------------------------------------------------------------
' Measures how many degrees SolidWorks rotates for one ViewRotateplusz call.
' This avoids assuming ViewRotateplusz is 90 degrees.
' ============================================================
Private Function TryMeasureViewRotateZIncrementDeg(ByVal model As Object, _
                                                   ByRef stepDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryMeasureViewRotateZIncrementDeg = False
    stepDegOut = 0#

    If model Is Nothing Then Exit Function

    Dim swView As Object
    Dim m0 As Variant
    Dim m1 As Variant

    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 150

    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    m0 = swView.Orientation3.ArrayData

    If IsArray(m0) = False Then Exit Function
    If UBound(m0) < 8 Then Exit Function

    Dim x0x As Double
    Dim x0y As Double
    Dim x0z As Double

    x0x = CDbl(m0(0))
    x0y = CDbl(m0(3))
    x0z = CDbl(m0(6))

    If NormalizeStraightenVector(x0x, x0y, x0z) = False Then Exit Function

    RotateViewZSteps model, 1
    StabilizeActiveView model, 150

    Set swView = model.ActiveView
    If swView Is Nothing Then Exit Function

    m1 = swView.Orientation3.ArrayData

    If IsArray(m1) = False Then Exit Function
    If UBound(m1) < 8 Then Exit Function

    Dim x1x As Double
    Dim x1y As Double
    Dim x1z As Double

    x1x = CDbl(m1(0))
    x1y = CDbl(m1(3))
    x1z = CDbl(m1(6))

    If NormalizeStraightenVector(x1x, x1y, x1z) = False Then Exit Function

    Dim dotVal As Double
    dotVal = (x0x * x1x) + (x0y * x1y) + (x0z * x1z)

    If dotVal > 1# Then dotVal = 1#
    If dotVal < -1# Then dotVal = -1#

    stepDegOut = AcosDegSafe(dotVal)

    ' Return to clean *Top before caller starts testing candidates.
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 100

    LogLine "Measured ViewRotateplusz one-step angle = " & Format(stepDegOut, "0.000") & " deg."

    TryMeasureViewRotateZIncrementDeg = (stepDegOut > 0.5)

    Exit Function

ErrHandler:
    LogLine "TryMeasureViewRotateZIncrementDeg error: " & Err.Description

    On Error Resume Next
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 100

    TryMeasureViewRotateZIncrementDeg = False
End Function

Private Function AcosDegSafe(ByVal x As Double) As Double
On Error Resume Next

    If x > 1# Then x = 1#
    If x < -1# Then x = -1#

    If x >= 0.999999999 Then
        AcosDegSafe = 0#
    ElseIf x <= -0.999999999 Then
        AcosDegSafe = 180#
    Else
        AcosDegSafe = (Atn(-x / Sqr(-x * x + 1#)) + 2# * Atn(1#)) * 180# / PI_VALUE
    End If
End Function

Private Function AlignTopSpinFromHolderGeometry(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    AlignTopSpinFromHolderGeometry = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim holderIndexes As Collection
    Dim potIndexes As Collection

    BuildFrontOrientIndexCollections holderIndexes, potIndexes

    If holderIndexes Is Nothing Then
        LogLine "Top spin calibrated alignment skipped: holder collection is Nothing."
        Exit Function
    End If

    If holderIndexes.count = 0 Then
        LogLine "Top spin calibrated alignment skipped: no holder CAD indexes found."
        Exit Function
    End If

    Dim holderIdx As Long
    holderIdx = PickLargestCadIndexFromCollection(holderIndexes)

    If holderIdx <= 0 Or holderIdx > PartCount Then
        LogLine "Top spin calibrated alignment skipped: invalid largest holder index."
        Exit Function
    End If

    Dim holderComp As Object
    Set holderComp = FindAssemblyComponentByName(model, parts(holderIdx).componentName)

    If holderComp Is Nothing Then
        LogLine "Top spin calibrated alignment skipped: holder component not found: " & _
                parts(holderIdx).componentName
        Exit Function
    End If

    LogLine "Top spin calibrated alignment: using holder " & parts(holderIdx).componentName & _
            " L/W/T=" & FormatNumberForCsv(parts(holderIdx).Length) & "/" & _
            FormatNumberForCsv(parts(holderIdx).Width) & "/" & _
            FormatNumberForCsv(parts(holderIdx).Thickness)

    Dim errs As Long
    On Error Resume Next
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    Err.Clear
    On Error GoTo ErrHandler

    ' Measure the actual view-rotate increment. In your log, one RotateViewZSteps
    ' call was clearly NOT 90 degrees. It was about 15 degrees, so 3 calls landed
    ' near 45 degrees. This calibration converts quarter turns into the correct
    ' number of SolidWorks view-rotate increments.
    Dim oneStepDeg As Double
    Dim stepsPerQuarter As Long

    If TryMeasureViewRotateZIncrementDeg(model, oneStepDeg) = False Then
        LogLine "Top spin calibrated alignment failed: could not measure ViewRotateplusz increment."
        Exit Function
    End If

    If oneStepDeg <= 0.5 Then
        LogLine "Top spin calibrated alignment failed: measured ViewRotateplusz increment too small: " & _
                Format(oneStepDeg, "0.000") & " deg."
        Exit Function
    End If

    stepsPerQuarter = CLng((90# / oneStepDeg) + 0.5)

    If stepsPerQuarter < 1 Then stepsPerQuarter = 1
    If stepsPerQuarter > 200 Then
        LogLine "Top spin calibrated alignment failed: unreasonable stepsPerQuarter=" & CStr(stepsPerQuarter)
        Exit Function
    End If

    LogLine "Top spin calibrated alignment: ViewRotateplusz increment=" & _
            Format(oneStepDeg, "0.000") & " deg; steps per 90 deg=" & CStr(stepsPerQuarter)

    Dim bestQuarter As Long
    Dim bestScore As Double
    Dim bestW As Double
    Dim bestH As Double
    Dim gotAny As Boolean

    bestQuarter = 0
    bestScore = -1E+99
    bestW = 0#
    bestH = 0#
    gotAny = False

    Dim q As Long
    Dim rotateSteps As Long
    Dim viewW As Double
    Dim viewH As Double
    Dim score As Double

    For q = 0 To 3

        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 100

        rotateSteps = q * stepsPerQuarter

        If rotateSteps > 0 Then
            RotateViewZSteps model, rotateSteps
            StabilizeActiveView model, 100
        End If

        If TryGetComponentViewWidthHeightInches(model, holderComp, viewW, viewH) Then

            ' Positive score means the holder long direction is more left-right on screen.
            score = viewW - viewH

            LogLine "Top spin calibrated candidate +" & CStr(q * 90) & _
                    " deg (" & CStr(rotateSteps) & " view steps): projected holder W/H=" & _
                    FormatNumberForCsv(viewW) & "/" & FormatNumberForCsv(viewH) & _
                    " score=" & FormatNumberForCsv(score)

            If gotAny = False Or score > bestScore Then
                gotAny = True
                bestScore = score
                bestQuarter = q
                bestW = viewW
                bestH = viewH
            End If

        Else

            LogLine "Top spin calibrated candidate +" & CStr(q * 90) & _
                    " deg: could not measure projected holder W/H."

        End If

    Next q

    If gotAny = False Then
        LogLine "Top spin calibrated alignment failed: no candidate measurements."
        Exit Function
    End If

    ' Re-apply the selected candidate from clean *Top.
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 100

    rotateSteps = bestQuarter * stepsPerQuarter

    If rotateSteps > 0 Then
        RotateViewZSteps model, rotateSteps
        StabilizeActiveView model, 150
    End If

    LogLine "Top spin calibrated alignment selected +" & CStr(bestQuarter * 90) & _
            " deg (" & CStr(rotateSteps) & " view steps). Final holder W/H=" & _
            FormatNumberForCsv(bestW) & "/" & FormatNumberForCsv(bestH)

    If bestW + TOP_SPIN_HOLDER_WIDTH_MARGIN_IN < bestH Then
        LogLine "WARNING: Top spin calibrated alignment selected best candidate, but holder is still taller than wide."
    End If

    If PersistCurrentViewAsStandardTop(model) Then

        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 150

        LogLine "Top spin calibrated alignment persisted selected spin as SolidWorks *Top."
        AlignTopSpinFromHolderGeometry = True
        Exit Function

    Else

        LogLine "WARNING: Top spin calibrated alignment could not persist selected spin as *Top."
        AlignTopSpinFromHolderGeometry = False
        Exit Function

    End If

    Exit Function

ErrHandler:
    LogLine "AlignTopSpinFromHolderGeometry calibrated error: " & Err.Description
    AlignTopSpinFromHolderGeometry = False
End Function

Private Sub ApplyCmsTopView(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
End Sub

' Show a view by name, resolving the SolidWorks standard views to their view IDs
' (the reliable form - see how CMS_BASE_TOP_VIEW_NAME/_ID are paired). Anything
' that is not a standard name is passed through as a user-named view.
Private Sub ShowNamedOrStandardView(ByVal model As Object, ByVal viewName As String)
On Error Resume Next

    If model Is Nothing Then Exit Sub
    If Trim(viewName) = "" Then Exit Sub

    Dim k As String
    Dim viewId As Long

    ' NormalizeKey does not strip "*", so drop the standard-view prefix first.
    k = NormalizeKey(Replace(viewName, "*", ""))
    viewId = -1

    Select Case k
        Case "FRONT":      viewId = 1
        Case "BACK":       viewId = 2
        Case "LEFT":       viewId = 3
        Case "RIGHT":      viewId = 4
        Case "TOP":        viewId = 5
        Case "BOTTOM":     viewId = 6
        Case "ISOMETRIC":  viewId = 7
        Case "TRIMETRIC":  viewId = 8
        Case "DIMETRIC":   viewId = 9
    End Select

    model.ShowNamedView2 viewName, viewId
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

Private Function TryFindBestComponentCenterByKeys(ByVal assyModel As Object, _
                                                  ByVal pipeKeys As String, _
                                                  ByRef cx As Double, _
                                                  ByRef cy As Double, _
                                                  ByRef cz As Double, _
                                                  ByRef foundName As String) As Boolean
On Error GoTo ErrHandler

    TryFindBestComponentCenterByKeys = False

    cx = 0#
    cy = 0#
    cz = 0#
    foundName = ""

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function
    If IsArray(vComps) = False Then Exit Function

    Dim bestScore As Double
    bestScore = -1#

    Dim i As Long
    Dim swComp As Object
    Dim hay As String
    Dim score As Double

    Dim tx As Double
    Dim ty As Double
    Dim tz As Double

    For i = 0 To UBound(vComps)

        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                hay = swComp.Name2 & " " & swComp.GetPathName

                If ContainsAnyPipeKey(hay, pipeKeys) Then

                    score = Len(hay)

                    If TryGetComponentCenterInches(swComp, tx, ty, tz) Then

                        If score > bestScore Then
                            bestScore = score
                            cx = tx
                            cy = ty
                            cz = tz
                            foundName = swComp.Name2
                            TryFindBestComponentCenterByKeys = True
                        End If

                    End If

                End If

            End If
        End If

    Next i

    Exit Function

ErrHandler:
    TryFindBestComponentCenterByKeys = False
End Function

' ============================================================
' END OF PART 1
' Paste Part 2 immediately after this.
' ============================================================
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

    LogLine "ZIP count=" & zips.count

    Dim i As Long
    For i = 1 To zips.count
        LogLine "Extracting ZIP " & i & "/" & zips.count & ": " & CStr(zips(i))
        If UnzipOneZipRobust(CStr(zips(i)), extractRoot) Then
            LogLine "ZIP extracted OK."
        Else
            LogLine "ZIP extract failed."
        End If
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
    folderName = UCase(folder.name)

    If folderName = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub
    If InStr(folderName, "PRINT") > 0 Then Exit Sub

    Dim file As Object
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.path)) = "zip" Then
            zips.Add file.path
        End If
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

    If Not ok Then
        ok = ExtractZipUsingShell(localZip, tempOut)
    End If

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
    WaitMilliseconds 3000

    ExtractZipUsingShell = True
    Exit Function

ErrHandler:
    ExtractZipUsingShell = False
End Function

Private Sub FlattenExtractedZipContentsIntoJobFolder(ByVal jobFolder As String, ByVal extractRoot As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If jobFolder = "" Then Exit Sub
    If extractRoot = "" Then Exit Sub
    If fso.FolderExists(jobFolder) = False Then Exit Sub
    If fso.FolderExists(extractRoot) = False Then Exit Sub

    LogLine "Flattening extracted ZIP contents into main job folder."
    LogLine "  FROM: " & extractRoot
    LogLine "  TO  : " & jobFolder

    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(extractRoot)

    Dim subFolder As Object
    Dim file As Object

    For Each file In rootFolder.Files
        fso.CopyFile file.path, jobFolder & "\" & file.name, True
        LogLine "Flatten copied file: " & file.name
    Next file

    For Each subFolder In rootFolder.SubFolders
        CopyExtractedFolderContentsToMain subFolder.path, jobFolder
    Next subFolder

    LogLine "Flatten extracted ZIP contents complete."
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
    Dim destSub As String

    For Each file In src.Files
        fso.CopyFile file.path, mainJobFolder & "\" & file.name, True
        LogLine "Flatten copied file: " & file.name
    Next file

    For Each subFolder In src.SubFolders
        destSub = mainJobFolder & "\" & subFolder.name
        EnsureFolderDeep destSub
        CopyFolderContents subFolder.path, destSub
        LogLine "Flatten copied folder: " & subFolder.name
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
        fso.CopyFile file.path, destFolder & "\" & file.name, True
    Next file

    Dim subFolder As Object
    For Each subFolder In src.SubFolders
        EnsureFolderDeep destFolder & "\" & subFolder.name
        CopyFolderContents subFolder.path, destFolder & "\" & subFolder.name
    Next subFolder
End Sub

' ============================================================
' CAD SCAN
' ============================================================

Private Sub ScanActiveSolidWorksDocument()
On Error GoTo ErrHandler

    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub

    If swModel.GetType = swDocASSEMBLY Then

        Set swAssy = swModel

        On Error Resume Next
        swAssy.ResolveAllLightWeightComponents True
        On Error GoTo ErrHandler

        Dim vComps As Variant
        vComps = swAssy.GetComponents(False)
        If IsEmpty(vComps) Then Exit Sub

        Dim i As Long
        For i = 0 To UBound(vComps)
            If VERBOSE_PROGRESS_LOG And i Mod 10 = 0 Then
                LogProgress "Scanning comp " & (i + 1) & "/" & (UBound(vComps) + 1)
            End If

            ProcessAssemblyComponent vComps(i)
            DoEvents
        Next i

    ElseIf swModel.GetType = swDocPART Then

        ScanPartBodies swModel

    End If

    Exit Sub

ErrHandler:
    LogLine "ScanActiveSolidWorksDocument error: " & Err.Description
End Sub

Private Sub ProcessAssemblyComponent(ByVal swComp As Object)
On Error GoTo SafeExit

    If swComp Is Nothing Then Exit Sub
    If swComp.IsSuppressed Then Exit Sub

    If HIDE_QUARTER_INCH_THICKNESS And PHYSICALLY_HIDE_250_BEFORE_SCAN Then
        If IsComponentHidden(swComp) Then Exit Sub
    End If

    Dim compPath As String
    Dim compName As String
    Dim configName As String

    compPath = swComp.GetPathName
    compName = swComp.Name2
    configName = swComp.ReferencedConfiguration

    Dim swCompModel As Object
    Set swCompModel = swComp.GetModelDoc2

    If swCompModel Is Nothing And compPath <> "" Then
        Dim errs As Long
        Dim warns As Long

        Set swCompModel = swApp.OpenDoc6(compPath, _
                                         swDocPART, _
                                         swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, _
                                         configName, errs, warns)
    End If

    If swCompModel Is Nothing Then Exit Sub
    If swCompModel.GetType <> swDocPART Then Exit Sub

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim hasCenter As Boolean

    hasCenter = TryGetComponentCenterInches(swComp, cx, cy, cz)

    Dim existingIndex As Long
    existingIndex = FindExistingPart(compPath, configName, compName, cy, hasCenter)

    If existingIndex > 0 Then
        parts(existingIndex).Quantity = parts(existingIndex).Quantity + 1
        Exit Sub
    End If

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    If Not GetPartBoundingBoxInches(swCompModel, dX, dY, dZ) Then Exit Sub

    Dim massValue As Double
    massValue = GetModelMassOrVolumeValue(swCompModel)

    Dim massCx As Double
    Dim massCy As Double
    Dim massCz As Double
    Dim massFromCom As Double
    Dim hasMassCenter As Boolean

    hasMassCenter = TryGetComponentMassCenterInches(swComp, massCx, massCy, massCz, massFromCom)

    If hasMassCenter And massFromCom > 0# Then
        massValue = massFromCom
    End If

    Dim asmBoxL As Double
    Dim asmBoxW As Double
    Dim asmBoxT As Double
    Dim hasAsmBox As Boolean

    hasAsmBox = TryGetComponentAssemblyBoundingBoxInches(swComp, asmBoxL, asmBoxW, asmBoxT)

    ' How far this component is tilted in the ASSEMBLY (from its placement
    ' transform), measured against the assembly axes. This is per-part even
    ' when imported part files store their bodies at a common orientation,
    ' so it is the reliable source for the pull-core rotation.
    Dim asmPullDeg As Double
    Dim asmPullSign As Double
    Dim hasAsmPull As Boolean
    asmPullSign = 1#
    hasAsmPull = TryGetComponentAssemblyPullAngleDeg(swComp, asmPullDeg, asmPullSign)

    If hasAsmBox Then
        If VERBOSE_PROGRESS_LOG Then
            LogLine "Assembly/world bbox for " & compName & " = " & _
                    FormatNumberForCsv(asmBoxL) & "/" & _
                    FormatNumberForCsv(asmBoxW) & "/" & _
                    FormatNumberForCsv(asmBoxT)
        End If
    Else
        If VERBOSE_PROGRESS_LOG Then
            LogLine "Assembly/world bbox unavailable for " & compName & "; using part-local bbox."
        End If
    End If

    AddCadPart compName, _
               CleanDisplayName(compName, compPath), _
               compPath, configName, "", 1, _
               dX, dY, dZ, massValue, False, _
               hasCenter, cx, cy, cz, _
               False, 0#, _
               hasAsmBox, asmBoxL, asmBoxW, asmBoxT, _
               hasMassCenter, massCx, massCy, massCz, _
               hasAsmPull, asmPullDeg, asmPullSign

SafeExit:
End Sub

Private Sub ScanPartBodies(ByVal swPartModel As Object)
On Error GoTo ErrHandler

    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Sub

    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant

    For i = 0 To UBound(vBodies)

        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then

            vBox = swBody.GetBodyBox

            If Not IsEmpty(vBox) Then

                Dim dX As Double
                Dim dY As Double
                Dim dZ As Double

                dX = Abs(CDbl(vBox(3)) - CDbl(vBox(0))) * INCHES_PER_METER
                dY = Abs(CDbl(vBox(4)) - CDbl(vBox(1))) * INCHES_PER_METER
                dZ = Abs(CDbl(vBox(5)) - CDbl(vBox(2))) * INCHES_PER_METER

                Dim massValue As Double
                massValue = GetBodyMassOrVolumeValue(swBody)
                If massValue <= 0 Then massValue = dX * dY * dZ

                AddCadPart swBody.name, swBody.name, _
                           swPartModel.GetPathName, "", swBody.name, 1, _
                           dX, dY, dZ, massValue, True
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "ScanPartBodies error: " & Err.Description
End Sub

Private Sub AddCadPart(ByVal componentName As String, ByVal cleanName As String, _
                       ByVal filePath As String, ByVal configName As String, _
                       ByVal bodyName As String, ByVal qty As Long, _
                       ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, _
                       ByVal massValue As Double, ByVal isBodyOnly As Boolean, _
                       Optional ByVal hasAsmCenter As Boolean = False, _
                       Optional ByVal asmCx As Double = 0#, Optional ByVal asmCy As Double = 0#, _
                       Optional ByVal asmCz As Double = 0#, _
                       Optional ByVal hasTopRotation As Boolean = False, _
                       Optional ByVal topRotationRad As Double = 0#, _
                       Optional ByVal hasOriginalAsmBBox As Boolean = False, _
                       Optional ByVal originalAsmL As Double = 0#, Optional ByVal originalAsmW As Double = 0#, _
                       Optional ByVal originalAsmT As Double = 0#, _
                       Optional ByVal hasMassCenter As Boolean = False, _
                       Optional ByVal massCx As Double = 0#, Optional ByVal massCy As Double = 0#, _
                       Optional ByVal massCz As Double = 0#, _
                       Optional ByVal hasAsmPullAngle As Boolean = False, _
                       Optional ByVal asmPullAngleDeg As Double = 0#, _
                       Optional ByVal asmPullAngleSign As Double = 1#)

    Dim L As Double
    Dim W As Double
    Dim T As Double

    SortThreeDimensions dX, dY, dZ, L, W, T

    L = Round(L, DIM_DECIMALS)
    W = Round(W, DIM_DECIMALS)
    T = Round(T, DIM_DECIMALS)

    If L * W * T < MIN_STEEL_VOLUME_CUIN Then Exit Sub

    PartCount = PartCount + 1
    ReDim Preserve parts(1 To PartCount)

    parts(PartCount).componentName = componentName
    parts(PartCount).cleanName = cleanName
    parts(PartCount).filePath = filePath
    parts(PartCount).configName = configName
    parts(PartCount).bodyName = bodyName
    parts(PartCount).Quantity = qty

    parts(PartCount).Length = L
    parts(PartCount).Width = W
    parts(PartCount).Thickness = T
    parts(PartCount).BBoxVolume = L * W * T

    parts(PartCount).hasOriginalAsmBBox = hasOriginalAsmBBox

    If hasOriginalAsmBBox Then
        parts(PartCount).OriginalAsmLength = originalAsmL
        parts(PartCount).OriginalAsmWidth = originalAsmW
        parts(PartCount).OriginalAsmThickness = originalAsmT
    Else
        parts(PartCount).OriginalAsmLength = L
        parts(PartCount).OriginalAsmWidth = W
        parts(PartCount).OriginalAsmThickness = T
    End If

    If massValue > 0 Then
        parts(PartCount).massValue = massValue
    Else
        parts(PartCount).massValue = parts(PartCount).BBoxVolume
    End If

    parts(PartCount).hasAsmCenter = hasAsmCenter
    parts(PartCount).AsmCenterX = asmCx
    parts(PartCount).AsmCenterY = asmCy
    parts(PartCount).AsmCenterZ = asmCz

    parts(PartCount).hasMassCenter = hasMassCenter
    parts(PartCount).MassCenterX = massCx
    parts(PartCount).MassCenterY = massCy
    parts(PartCount).MassCenterZ = massCz

    parts(PartCount).hasTopRotation = hasTopRotation
    parts(PartCount).topRotationRad = topRotationRad

    parts(PartCount).hasAsmPullAngle = hasAsmPullAngle
    parts(PartCount).asmPullAngleDeg = asmPullAngleDeg
    parts(PartCount).asmPullAngleSign = asmPullAngleSign

    parts(PartCount).UsedForBomMatch = False
    parts(PartCount).isBodyOnly = isBodyOnly
End Sub

Private Function FindExistingPart(ByVal filePath As String, _
                                  ByVal configName As String, _
                                  ByVal compName As String, _
                                  Optional ByVal newCenterY As Double = 0#, _
                                  Optional ByVal newHasCenter As Boolean = False) As Long

    Dim i As Long

    For i = 1 To PartCount

        If filePath <> "" Then

            If LCase(parts(i).filePath) = LCase(filePath) And _
               LCase(parts(i).configName) = LCase(configName) Then

                ' Different assembly instances must stay separate.
                ' Part-1 and Part-2 may use the same file but be different physical parts.
                If compName <> "" Then
                    If LCase(parts(i).componentName) <> LCase(compName) Then GoTo NextPart
                End If

                ' Same file/config at a different Y center should stay separate.
                If newHasCenter And parts(i).hasAsmCenter Then
                    If Abs(parts(i).AsmCenterY - newCenterY) > 0.05 Then GoTo NextPart
                End If

                FindExistingPart = i
                Exit Function

            End If

        Else

            If LCase(parts(i).componentName) = LCase(compName) Then
                FindExistingPart = i
                Exit Function
            End If

        End If

NextPart:
    Next i

    FindExistingPart = 0
End Function

Private Function GetPartBoundingBoxInches(ByVal swPartModel As Object, _
                                          ByRef dxIn As Double, _
                                          ByRef dyIn As Double, _
                                          ByRef dzIn As Double) As Boolean
On Error GoTo ErrHandler

    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then
        GetPartBoundingBoxInches = False
        Exit Function
    End If

    Dim firstBody As Boolean
    firstBody = True

    Dim xmin As Double
    Dim yMin As Double
    Dim zMin As Double
    Dim xmax As Double
    Dim yMax As Double
    Dim zMax As Double

    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant

    For i = 0 To UBound(vBodies)

        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then

            vBox = swBody.GetBodyBox

            If Not IsEmpty(vBox) Then

                If firstBody Then
                    xmin = CDbl(vBox(0))
                    yMin = CDbl(vBox(1))
                    zMin = CDbl(vBox(2))
                    xmax = CDbl(vBox(3))
                    yMax = CDbl(vBox(4))
                    zMax = CDbl(vBox(5))
                    firstBody = False
                Else
                    If CDbl(vBox(0)) < xmin Then xmin = CDbl(vBox(0))
                    If CDbl(vBox(1)) < yMin Then yMin = CDbl(vBox(1))
                    If CDbl(vBox(2)) < zMin Then zMin = CDbl(vBox(2))
                    If CDbl(vBox(3)) > xmax Then xmax = CDbl(vBox(3))
                    If CDbl(vBox(4)) > yMax Then yMax = CDbl(vBox(4))
                    If CDbl(vBox(5)) > zMax Then zMax = CDbl(vBox(5))
                End If

            End If

        End If

    Next i

    If firstBody Then
        GetPartBoundingBoxInches = False
        Exit Function
    End If

    dxIn = Abs(xmax - xmin) * INCHES_PER_METER
    dyIn = Abs(yMax - yMin) * INCHES_PER_METER
    dzIn = Abs(zMax - zMin) * INCHES_PER_METER

    GetPartBoundingBoxInches = True
    Exit Function

ErrHandler:
    GetPartBoundingBoxInches = False
End Function

Private Function TryGetComponentAssemblyBoundingBoxInches(ByVal swComp As Object, _
                                                          ByRef outL As Double, _
                                                          ByRef outW As Double, _
                                                          ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentAssemblyBoundingBoxInches = False

    outL = 0#
    outW = 0#
    outT = 0#

    If swComp Is Nothing Then Exit Function

    Dim swPartModel As Object
    Set swPartModel = swComp.GetModelDoc2

    If swPartModel Is Nothing Then Exit Function
    If swPartModel.GetType <> swDocPART Then Exit Function

    Dim xmin As Double
    Dim yMin As Double
    Dim zMin As Double
    Dim xmax As Double
    Dim yMax As Double
    Dim zMax As Double

    If TryGetPartRawBoxMeters(swPartModel, xmin, yMin, zMin, xmax, yMax, zMax) = False Then Exit Function

    Dim xform As Object
    Set xform = swComp.Transform2

    If xform Is Nothing Then Exit Function

    Dim m As Variant
    m = xform.ArrayData

    If IsEmpty(m) Then Exit Function
    If IsArray(m) = False Then Exit Function
    If UBound(m) < 12 Then Exit Function

    Dim scaleVal As Double
    scaleVal = CDbl(m(12))
    If Abs(scaleVal) < 0.0000001 Then scaleVal = 1#

    Dim xs(0 To 7) As Double
    Dim ys(0 To 7) As Double
    Dim zs(0 To 7) As Double

    xs(0) = xmin: ys(0) = yMin: zs(0) = zMin
    xs(1) = xmax: ys(1) = yMin: zs(1) = zMin
    xs(2) = xmin: ys(2) = yMax: zs(2) = zMin
    xs(3) = xmax: ys(3) = yMax: zs(3) = zMin
    xs(4) = xmin: ys(4) = yMin: zs(4) = zMax
    xs(5) = xmax: ys(5) = yMin: zs(5) = zMax
    xs(6) = xmin: ys(6) = yMax: zs(6) = zMax
    xs(7) = xmax: ys(7) = yMax: zs(7) = zMax

    Dim firstPoint As Boolean
    firstPoint = True

    Dim axmin As Double
    Dim aymin As Double
    Dim azmin As Double
    Dim axmax As Double
    Dim aymax As Double
    Dim azmax As Double

    Dim i As Long

    For i = 0 To 7

        Dim tx As Double
        Dim ty As Double
        Dim tz As Double

        tx = scaleVal * ((xs(i) * CDbl(m(0))) + (ys(i) * CDbl(m(3))) + (zs(i) * CDbl(m(6)))) + CDbl(m(9))
        ty = scaleVal * ((xs(i) * CDbl(m(1))) + (ys(i) * CDbl(m(4))) + (zs(i) * CDbl(m(7)))) + CDbl(m(10))
        tz = scaleVal * ((xs(i) * CDbl(m(2))) + (ys(i) * CDbl(m(5))) + (zs(i) * CDbl(m(8)))) + CDbl(m(11))

        If firstPoint Then
            axmin = tx
            aymin = ty
            azmin = tz
            axmax = tx
            aymax = ty
            azmax = tz
            firstPoint = False
        Else
            If tx < axmin Then axmin = tx
            If ty < aymin Then aymin = ty
            If tz < azmin Then azmin = tz

            If tx > axmax Then axmax = tx
            If ty > aymax Then aymax = ty
            If tz > azmax Then azmax = tz
        End If

    Next i

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = Abs(axmax - axmin) * INCHES_PER_METER
    dY = Abs(aymax - aymin) * INCHES_PER_METER
    dZ = Abs(azmax - azmin) * INCHES_PER_METER

    SortThreeDimensions dX, dY, dZ, outL, outW, outT

    outL = Round(outL, DIM_DECIMALS)
    outW = Round(outW, DIM_DECIMALS)
    outT = Round(outT, DIM_DECIMALS)

    TryGetComponentAssemblyBoundingBoxInches = (outL > 0# And outW > 0# And outT > 0#)

    Exit Function

ErrHandler:
    TryGetComponentAssemblyBoundingBoxInches = False
End Function

Private Function TryGetPartRawBoxMeters(ByVal swPartModel As Object, _
                                        ByRef xmin As Double, _
                                        ByRef yMin As Double, _
                                        ByRef zMin As Double, _
                                        ByRef xmax As Double, _
                                        ByRef yMax As Double, _
                                        ByRef zMax As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPartRawBoxMeters = False

    If swPartModel Is Nothing Then Exit Function
    If swPartModel.GetType <> swDocPART Then Exit Function

    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function

    Dim firstBody As Boolean
    firstBody = True

    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant

    For i = 0 To UBound(vBodies)

        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then

            vBox = swBody.GetBodyBox

            If IsEmpty(vBox) = False Then
                If IsArray(vBox) Then
                    If UBound(vBox) >= 5 Then

                        If firstBody Then
                            xmin = CDbl(vBox(0))
                            yMin = CDbl(vBox(1))
                            zMin = CDbl(vBox(2))
                            xmax = CDbl(vBox(3))
                            yMax = CDbl(vBox(4))
                            zMax = CDbl(vBox(5))
                            firstBody = False
                        Else
                            If CDbl(vBox(0)) < xmin Then xmin = CDbl(vBox(0))
                            If CDbl(vBox(1)) < yMin Then yMin = CDbl(vBox(1))
                            If CDbl(vBox(2)) < zMin Then zMin = CDbl(vBox(2))

                            If CDbl(vBox(3)) > xmax Then xmax = CDbl(vBox(3))
                            If CDbl(vBox(4)) > yMax Then yMax = CDbl(vBox(4))
                            If CDbl(vBox(5)) > zMax Then zMax = CDbl(vBox(5))
                        End If

                    End If
                End If
            End If

        End If

    Next i

    TryGetPartRawBoxMeters = Not firstBody

    Exit Function

ErrHandler:
    TryGetPartRawBoxMeters = False
End Function

Private Function TryGetComponentCenterInches(ByVal swComp As Object, _
                                             ByRef cx As Double, _
                                             ByRef cy As Double, _
                                             ByRef cz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentCenterInches = False

    cx = 0#
    cy = 0#
    cz = 0#

    If swComp Is Nothing Then Exit Function

    Dim vBox As Variant

    On Error Resume Next
    vBox = swComp.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    cx = ((CDbl(vBox(0)) + CDbl(vBox(3))) / 2#) * INCHES_PER_METER
    cy = ((CDbl(vBox(1)) + CDbl(vBox(4))) / 2#) * INCHES_PER_METER
    cz = ((CDbl(vBox(2)) + CDbl(vBox(5))) / 2#) * INCHES_PER_METER

    TryGetComponentCenterInches = True
    Exit Function

ErrHandler:
    TryGetComponentCenterInches = False
End Function

Private Function GetModelMassOrVolumeValue(ByVal model As Object) As Double
On Error GoTo ErrHandler

    If model Is Nothing Then
        GetModelMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim mp As Object
    Set mp = model.Extension.CreateMassProperty

    If mp Is Nothing Then
        GetModelMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim massLb As Double
    Dim volCuIn As Double

    massLb = 0#
    volCuIn = 0#

    On Error Resume Next
    massLb = CDbl(mp.Mass)
    volCuIn = CDbl(mp.Volume) * CUIN_PER_CUBIC_METER
    On Error GoTo ErrHandler

    ' Fast Match Studio mode:
    ' Treat every part as solid/plain carbon steel without physically assigning
    ' material to hundreds of imported components.
    If MATCH_STUDIO_USE_CARBON_STEEL_DENSITY_FOR_MASS Then
        If volCuIn > 0# Then
            GetModelMassOrVolumeValue = volCuIn * MATCH_STUDIO_CARBON_STEEL_DENSITY_LB_PER_CUIN
            Exit Function
        End If
    End If

    If massLb > 0# Then
        GetModelMassOrVolumeValue = massLb
        Exit Function
    End If

    If volCuIn > 0# Then
        GetModelMassOrVolumeValue = volCuIn
        Exit Function
    End If

    GetModelMassOrVolumeValue = 0#
    Exit Function

ErrHandler:
    GetModelMassOrVolumeValue = 0#
End Function

Private Function GetBodyMassOrVolumeValue(ByVal swBody As Object) As Double
On Error GoTo ErrHandler

    If swBody Is Nothing Then
        GetBodyMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim vProps As Variant

    On Error Resume Next
    vProps = swBody.GetMassProperties(1#)
    On Error GoTo ErrHandler

    If IsEmpty(vProps) Then
        GetBodyMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim v As Double
    v = 0#

    On Error Resume Next
    v = CDbl(vProps(3)) * CUIN_PER_CUBIC_METER
    On Error GoTo ErrHandler

    GetBodyMassOrVolumeValue = v
    Exit Function

ErrHandler:
    GetBodyMassOrVolumeValue = 0#
End Function

Private Sub SortThreeDimensions(ByVal a As Double, _
                                ByVal b As Double, _
                                ByVal c As Double, _
                                ByRef L As Double, _
                                ByRef W As Double, _
                                ByRef T As Double)

    Dim arr(1 To 3) As Double
    Dim i As Long
    Dim j As Long
    Dim tmp As Double

    arr(1) = a
    arr(2) = b
    arr(3) = c

    For i = 1 To 2
        For j = i + 1 To 3
            If arr(j) > arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    L = arr(1)
    W = arr(2)
    T = arr(3)
End Sub

Private Sub SortPartsByVolumeDescending()

    Dim i As Long
    Dim j As Long
    Dim tmp As PartInfo

    For i = 1 To PartCount - 1
        For j = i + 1 To PartCount
            If parts(j).BBoxVolume > parts(i).BBoxVolume Then
                tmp = parts(i)
                parts(i) = parts(j)
                parts(j) = tmp
            End If
        Next j
    Next i
End Sub

' ============================================================
' EXPORT X_T / DXF
' ============================================================

Private Sub ExportMatchedPartsAsXt(ByVal outputFolder As String)
On Error GoTo ErrHandler

    EnsureFolderDeep outputFolder
    EnsureSwHidden

    Dim i As Long

    For i = 1 To ExportCount
        LogLine "Exporting item " & i & "/" & ExportCount & ": " & ExportRows(i).quoteName
        ExportOneMatchedPartAsXt ExportRows(i), outputFolder, i
        ReleaseSolidWorksMemory "after export item"
        DoEvents
    Next i

    If CREATE_MAIN_ASSEMBLY_PACKAGE Then
        LogStart "Create MAIN ASSEMBLY / HOLDERS package"
        ExportMainAssemblyAndHoldersPackage CurrentJobFolder
        LogDone "Create MAIN ASSEMBLY / HOLDERS package"
    End If

    ' Slow second-pass STL package disabled.
    ' STLs are now saved during existing isolated export states.
    'If CREATE_MATCH_STUDIO_STL_PACKAGE Then
    '    LogStart "Create MATCH STUDIO STL package"
    '    ExportMatchStudioStlPackage CurrentJobFolder
    '    LogDone "Create MATCH STUDIO STL package"
    'End If

    If CREATE_J_BLOCK_PACKAGE And MATCH_STUDIO_TURBO_ONLY = False Then
        LogStart "Create J BLOCK package"
        ExportJBlockPackage CurrentJobFolder
        LogDone "Create J BLOCK package"
    End If

    If CREATE_PULLCORE_CAM_KEY_PACKAGE And MATCH_STUDIO_TURBO_ONLY = False Then
        LogStart "Create PULLCORE CAM AND KEY package"
        ExportPullcoreCamKeyPackage CurrentJobFolder
        LogDone "Create PULLCORE CAM AND KEY package"
    End If

    If CREATE_PULLCORE_STOP_PACKAGE And MATCH_STUDIO_TURBO_ONLY = False Then
        LogStart "Create PULLCORE STOP package"
        ExportPullcoreStopPackage CurrentJobFolder
        LogDone "Create PULLCORE STOP package"
    End If

    If CREATE_POTS_HOLDERS_DXF_PACKAGE And MATCH_STUDIO_TURBO_ONLY = False Then
        LogStart "Create POTS AND HOLDERS DXF package"
        ExportPotsAndHoldersDxfPackage CurrentJobFolder
        LogDone "Create POTS AND HOLDERS DXF package"
    End If

    If RUN_PYRO_AFTER_JOB And MATCH_STUDIO_TURBO_ONLY = False Then
        LogStart "Run PYROPEL face script"
        RunPyropelFaceScriptsAfterJob CurrentJobFolder
        LogDone "Run PYROPEL face script"
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportMatchedPartsAsXt error: " & Err.Description
End Sub

Private Sub ExportOneMatchedPartAsXt(ByRef q As ExportInfo, ByVal outputFolder As String, ByVal itemNumber As Long)
On Error GoTo ErrHandler

    If q.CadPartIndex <= 0 Or q.CadPartIndex > PartCount Then
        LogLine "  Skipped export (no CAD geometry for this BOM row): " & q.quoteName
        Exit Sub
    End If

    Dim p As PartInfo
    p = parts(q.CadPartIndex)

    Dim quoteToken As String
    quoteToken = CleanQuoteTokenForFile(q.quoteName)
    If quoteToken = "" Then quoteToken = "ITEM"

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim targetFolder As String
    Dim isPyropel As Boolean

    isPyropel = IsPyropelExportRow(q, p)

    If isPyropel And ROUTE_PYROPEL_TO_SEPARATE_FOLDERS Then

        If InStr(UCase(p.componentName), "POT") > 0 _
           Or InStr(UCase(p.cleanName), "POT") > 0 _
           Or InStr(UCase(q.quoteName), "POT") > 0 Then

            targetFolder = outputFolder & "\" & CurrentJobNumber & PYROPEL_POTS_FOLDER_SUFFIX
        Else
            targetFolder = outputFolder & "\" & CurrentJobNumber & PYROPEL_HOLDERS_FOLDER_SUFFIX
        End If

        LogLine "Pyropel routing -> " & targetFolder

    Else

        targetFolder = GetOutputFolderForRegularExport(q.quoteName, outputFolder)

    End If

    EnsureFolderDeep targetFolder

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = targetFolder & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".x_t"
    xtPath = GetUniqueFilePath(xtPath)

    dxfPath = targetFolder & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".dxf"
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Export file name: " & xtPath
    LogLine "  Quote name: " & q.quoteName
    LogLine "  Material: " & q.material
    LogLine "  CAD component: " & p.componentName
    LogLine "  Target folder: " & targetFolder

    Dim makeDxf As Boolean
    makeDxf = ShouldCreateStandardPrintDxf(q.quoteName)
    If FAST_BATCH_EXPORT And FAST_SKIP_PRINTS_DXF Then makeDxf = False

    If MATCH_STUDIO_TURBO_ONLY Then
        makeDxf = False
        LogLine "MATCH STUDIO TURBO: skipping print DXF for " & q.quoteName
    End If

    If isPyropel Then
        makeDxf = True
        LogLine "Pyropel export: creating X_T + DXF for routing."
    End If

    If swModel.GetType = swDocASSEMBLY Then

        ExportAssemblyOneComponentHideRestWithOptionalDxf swModel, p, xtPath, dxfPath, q.quoteName, makeDxf

    ElseIf swModel.GetType = swDocPART Then

        If p.isBodyOnly Then

            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath

            If makeDxf And CREATE_DXFS_DURING_XT_SAVE Then
                CreateStandardPrintDxfFromXtPath xtPath, dxfPath, q.quoteName
            End If

            ShowAllPartBodies swModel

        Else

            SaveModelAs swModel, xtPath

            If makeDxf And CREATE_DXFS_DURING_XT_SAVE Then
                CreateStandardPrintDxfFromXtPath xtPath, dxfPath, q.quoteName
            End If

        End If

    End If

    If Not ExportFilePaths Is Nothing Then
        ExportFilePaths(NormalizeKey(q.quoteName)) = xtPath
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportOneMatchedPartAsXt error: " & Err.Description
End Sub

Private Function IsPyropelExportRow(ByRef q As ExportInfo, ByRef p As PartInfo) As Boolean
On Error Resume Next

    Dim hay As String
    hay = UCase(q.material & " " & q.quoteName & " " & p.componentName & " " & p.cleanName & " " & p.filePath)

    IsPyropelExportRow = (InStr(hay, "PYROPEL") > 0)
End Function

Private Function ShouldCreateStandardPrintDxf(ByVal quoteName As String) As Boolean

    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k

        Case "IDHOLDER", _
             "ODHOLDER", _
             "TCP", _
             "BCP", _
             "TOPSMED", _
             "BOTTOMSMED", _
             "BOTSMED", _
             "TOPSMEDPLATE", _
             "BOTTOMSMEDPLATE", _
             "BOTSMEDPLATE", _
             "TOPCLAMPINGPLATE", _
             "BOTTOMCLAMPINGPLATE", _
             "BOTCLAMPINGPLATE"

            ShouldCreateStandardPrintDxf = True
            Exit Function

    End Select

    ShouldCreateStandardPrintDxf = False

End Function

Private Function IsHolderQuoteForMatchStudio(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)
    IsHolderQuoteForMatchStudio = (k = "IDHOLDER" Or k = "ODHOLDER" Or _
                                   k = "IDMOLDBASE" Or k = "ODMOLDBASE")
End Function

Private Function ShouldExportComponentArtifacts(ByVal quoteName As String) As Boolean
    ShouldExportComponentArtifacts = IsHolderQuoteForMatchStudio(quoteName)
End Function

Private Sub ExportComponentMatchStudioArtifacts(ByVal assyModel As Object, _
                                                ByVal quoteName As String, _
                                                ByVal xtPath As String)
On Error GoTo ErrHandler
    If assyModel Is Nothing Then Exit Sub
    If xtPath = "" Then Exit Sub
    If IsHolderQuoteForMatchStudio(quoteName) = False Then Exit Sub

    Dim folder As String
    folder = GetParentFolderPath(xtPath)
    EnsureFolderDeep folder

    If CREATE_COMPONENT_EASM_WITH_XT Then
        If assyModel.GetType = swDocASSEMBLY Then
            Dim easmPath As String
            easmPath = ReplaceExtension(xtPath, "easm")
            SaveModelAs assyModel, easmPath
        End If
    End If

    If CREATE_COMPONENT_ISO_JPEGS Then
        ExportComponentIsoJpeg assyModel, quoteName, folder
    End If
    Exit Sub
ErrHandler:
    LogLine "ExportComponentMatchStudioArtifacts error: " & Err.Description
End Sub

Private Function ReplaceExtension(ByVal path As String, ByVal newExt As String) As String
    Dim dot As Long
    dot = InStrRev(path, ".")
    If dot > 0 Then
        ReplaceExtension = Left(path, dot - 1) & "." & newExt
    Else
        ReplaceExtension = path & "." & newExt
    End If
End Function

Private Function GetParentFolderPath(ByVal path As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolderPath = fso.GetParentFolderName(path)
End Function

Private Sub ExportComponentIsoJpeg(ByVal assyModel As Object, _
                                   ByVal quoteName As String, _
                                   ByVal outputFolder As String)
On Error GoTo ErrHandler
    If assyModel Is Nothing Then Exit Sub
    If outputFolder = "" Then Exit Sub

    EnsureFolderDeep outputFolder
    assyModel.ShowNamedView2 "*Isometric", 7
    assyModel.ViewZoomtofit2
    assyModel.GraphicsRedraw2

    Dim jpgPath As String
    jpgPath = GetUniqueFilePath(outputFolder & "\" & CurrentJobNumber & "_" & quoteName & " ISO.jpg")
    SaveViewAsImage assyModel, jpgPath
    LogLine "Saved component ISO preview: " & jpgPath
    Exit Sub
ErrHandler:
    LogLine "ExportComponentIsoJpeg error: " & Err.Description
End Sub

Private Sub SaveViewAsImage(ByVal model As Object, ByVal imagePath As String)
On Error GoTo ErrHandler
    Dim errs As Long
    Dim warns As Long
    model.Extension.SaveAs3 imagePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns
    Exit Sub
ErrHandler:
    LogLine "SaveViewAsImage error: " & Err.Description
End Sub

' ============================================================
' SUPPRESS-ISOLATE EXPORT ROUTINES
' ============================================================

Private Sub ExportAssemblyOneComponentHideRestWithOptionalDxf(ByVal assyModel As Object, _
                                                              ByRef p As PartInfo, _
                                                              ByVal xtPath As String, _
                                                              ByVal dxfPath As String, _
                                                              ByVal quoteName As String, _
                                                              ByVal makeDxf As Boolean)
On Error GoTo ErrHandler

    If assyModel Is Nothing Then Exit Sub

    LogLine "Suppress-isolate export. Target component: " & p.componentName

    If ExportAssemblyComponentBySuppressRestWithOptionalDxf(assyModel, p, xtPath, dxfPath, quoteName, makeDxf) Then
        LogLine "Suppress-isolate export succeeded."
    Else
        LogLine "WARNING: Suppress-isolate export failed for: " & p.componentName
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportAssemblyOneComponentHideRestWithOptionalDxf error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents assyModel
    ShowAllAssemblyComponents assyModel
End Sub

Private Function TryExportAssemblyComponentDirectWithOptionalDxf(ByVal assyModel As Object, _
                                                                 ByRef p As PartInfo, _
                                                                 ByVal xtPath As String, _
                                                                 ByVal dxfPath As String, _
                                                                 ByVal quoteName As String, _
                                                                 ByVal makeDxf As Boolean) As Boolean
On Error GoTo ErrHandler

    TryExportAssemblyComponentDirectWithOptionalDxf = _
        ExportAssemblyComponentBySuppressRestWithOptionalDxf(assyModel, p, xtPath, dxfPath, quoteName, makeDxf)
    Exit Function

ErrHandler:
    LogLine "TryExportAssemblyComponentDirectWithOptionalDxf error: " & Err.Description
    TryExportAssemblyComponentDirectWithOptionalDxf = False
End Function

Private Function ExportAssemblyComponentBySuppressRestWithOptionalDxf(ByVal assyModel As Object, _
                                                                      ByRef p As PartInfo, _
                                                                      ByVal xtPath As String, _
                                                                      ByVal dxfPath As String, _
                                                                      ByVal quoteName As String, _
                                                                      ByVal makeDxf As Boolean) As Boolean
On Error GoTo ErrHandler

    ExportAssemblyComponentBySuppressRestWithOptionalDxf = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    Dim i As Long
    Dim swComp As Object
    Dim foundTarget As Boolean
    Dim suppressCount As Long

    Dim suppressedObjects As Collection
    Set suppressedObjects = New Collection

    assyModel.ClearSelection2 True

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If LCase(swComp.Name2) = LCase(p.componentName) Then
                    If swComp.Select4(False, Nothing, False) Then foundTarget = True
                    Exit For
                End If
            End If
        End If
    Next i

    If foundTarget = False Then
        LogLine "Suppress-isolate failed: target component not found: " & p.componentName
        assyModel.ClearSelection2 True
        Exit Function
    End If

    assyModel.ClearSelection2 True

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If LCase(swComp.Name2) <> LCase(p.componentName) Then
                    If swComp.Select4(True, Nothing, False) Then
                        suppressCount = suppressCount + 1
                        suppressedObjects.Add swComp
                    End If
                End If
            End If
        End If
    Next i

    If suppressCount > 0 Then
        On Error Resume Next
        assyModel.EditSuppress2
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrHandler
    End If

    assyModel.ClearSelection2 True

    SaveModelAs assyModel, xtPath

    If CREATE_MATCH_STUDIO_STLS_DURING_EXISTING_EXPORTS Then
        Dim msStlLabel As String
        msStlLabel = MatchStudioStlLabelFromQuote(quoteName)

        If msStlLabel <> "" Then
            SaveMatchStudioVisibleComponentStl assyModel, msStlLabel, p.componentName
        End If
    End If

    If fso.FileExists(xtPath) = False Then
        LogLine "Suppress-isolate failed: XT was not created."
        GoTo CleanExit
    End If

    If ShouldExportComponentArtifacts(quoteName) Then
        ExportComponentMatchStudioArtifacts assyModel, quoteName, xtPath
    End If

    If makeDxf And CREATE_DXFS_DURING_XT_SAVE Then

        Dim qk As String
        qk = NormalizeKey(quoteName)

        Select Case qk

            Case "IDHOLDER", "ODHOLDER"
                If CreateHolderDxfFromAssemblyBottomView(assyModel, dxfPath, quoteName) = False Then
                    LogLine "WARNING: " & quoteName & " assembly holder-face DXF failed; using part-based fallback."
                    CreateStandardPrintDxfFromXtPath xtPath, dxfPath, quoteName
                End If

            Case "TCP", "BCP"
                If CreateClampingPlateDxfFromAssemblyTopView(assyModel, dxfPath, quoteName) = False Then
                    LogLine "WARNING: " & quoteName & " assembly-top DXF failed; using part-based fallback."
                    CreateStandardPrintDxfFromXtPath xtPath, dxfPath, quoteName
                End If

            Case Else
                CreateStandardPrintDxfFromXtPath xtPath, dxfPath, quoteName

        End Select
    End If

    If CREATE_ENGRAVING_DXFS Then
        MaybeCreateEngravingDxfForIsolatedComponent assyModel, quoteName, dxfPath
    End If

    ExportAssemblyComponentBySuppressRestWithOptionalDxf = True

CleanExit:
    On Error Resume Next

    assyModel.ClearSelection2 True

    Dim obj As Object

    If Not suppressedObjects Is Nothing Then
        For i = 1 To suppressedObjects.count
            Set obj = suppressedObjects(i)
            If Not obj Is Nothing Then obj.Select4 True, Nothing, False
        Next i

        If suppressedObjects.count > 0 Then assyModel.EditUnsuppress2
    End If

    assyModel.ClearSelection2 True
    Exit Function

ErrHandler:
    LogLine "ExportAssemblyComponentBySuppressRestWithOptionalDxf error: " & Err.Description
    Resume CleanExit
End Function

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

' ============================================================
' ASSEMBLY/PART VISIBILITY HELPERS
' ============================================================

Private Function HideAllExceptTwoTargetComponentsOnce(ByVal assyModel As Object, _
                                                      ByVal targetComponentName1 As String, _
                                                      ByVal targetComponentName2 As String, _
                                                      ByRef hiddenNames As Collection) As Boolean
On Error GoTo ErrHandler

    HideAllExceptTwoTargetComponentsOnce = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    assyModel.ClearSelection2 True

    Dim i As Long
    Dim swComp As Object
    Dim foundTarget1 As Boolean
    Dim foundTarget2 As Boolean
    Dim selectedCount As Long

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                If LCase(swComp.Name2) = LCase(targetComponentName1) Then
                    foundTarget1 = True
                ElseIf LCase(swComp.Name2) = LCase(targetComponentName2) Then
                    foundTarget2 = True
                Else
                    If swComp.Select4(True, Nothing, False) Then
                        hiddenNames.Add swComp.Name2
                        selectedCount = selectedCount + 1
                    End If
                End If

            End If
        End If
    Next i

    If foundTarget1 = False Or foundTarget2 = False Then
        LogLine "WARNING: Did not find both holder target components."
        assyModel.ClearSelection2 True
        Exit Function
    End If

    If selectedCount > 0 Then
        LogLine "Selected non-holder components to hide: " & selectedCount
        swAssembly.HideComponent2
    End If

    assyModel.ClearSelection2 True
    HideAllExceptTwoTargetComponentsOnce = True
    Exit Function

ErrHandler:
    LogLine "HideAllExceptTwoTargetComponentsOnce error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    HideAllExceptTwoTargetComponentsOnce = False
End Function

Private Function HideAllExceptComponentNamesOnce(ByVal assyModel As Object, _
                                                 ByVal keepNames As Collection, _
                                                 ByRef hiddenNames As Collection) As Boolean
On Error GoTo ErrHandler

    HideAllExceptComponentNamesOnce = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function
    If keepNames Is Nothing Then Exit Function
    If keepNames.count = 0 Then Exit Function

    If hiddenNames Is Nothing Then Set hiddenNames = New Collection

    Dim keepDict As Object
    Set keepDict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To keepNames.count
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

    selectedCount = 0
    keepFoundCount = 0

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
    If componentNames.count = 0 Then Exit Sub

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    assyModel.ClearSelection2 True

    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To componentNames.count
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

Private Sub ShowOnlyTwoAssemblyComponents(ByVal assyModel As Object, _
                                          ByVal targetComponentName1 As String, _
                                          ByVal targetComponentName2 As String)
On Error Resume Next

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long
    Dim swComp As Object

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If Not swComp.IsSuppressed Then
                If LCase(swComp.Name2) = LCase(targetComponentName1) Or _
                   LCase(swComp.Name2) = LCase(targetComponentName2) Then
                    swComp.Visible = swComponentVisible
                Else
                    swComp.Visible = swComponentHidden
                End If
            End If
        End If
    Next i
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

Private Sub ShowOnlyPartBody(ByVal partModel As Object, ByVal targetBodyName As String)
On Error Resume Next

    If partModel Is Nothing Then Exit Sub
    If partModel.GetType <> swDocPART Then Exit Sub

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Sub

    Dim i As Long

    For i = 0 To UBound(vBodies)
        If LCase(vBodies(i).name) = LCase(targetBodyName) Then
            vBodies(i).Hide2 False
        Else
            vBodies(i).Hide2 True
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

    Dim i As Long

    For i = 0 To UBound(vBodies)
        vBodies(i).Hide2 False
    Next i
End Sub

Private Sub SaveModelAs(ByVal model As Object, ByVal fullPath As String)
On Error GoTo ErrHandler

    Dim errs As Long
    Dim warns As Long

    LogLine "Saving: " & fullPath
    EnsureSwHidden

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
    If fullPath = "" Then Exit Function

    Dim errs As Long
    Dim warns As Long

    LogLine "Saving copy: " & fullPath
    EnsureSwHidden

    ' Use Copy so temporary native drawing sources do not rename/repath the
    ' live assembly document. Without this, the next export can keep using a
    ' stale COM object after the temp DXF source is closed/deleted.
    model.Extension.SaveAs3 fullPath, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent + swSaveAsOptions_Copy, _
                            Nothing, Nothing, errs, warns

    LogLine "Save copy done. Errors=" & errs & " Warnings=" & warns

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    SaveModelCopyAs = fso.FileExists(fullPath)
    Exit Function

ErrHandler:
    LogLine "SaveModelCopyAs error: " & Err.Description
    SaveModelCopyAs = False
End Function

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

Private Function ReorientBinaryStlFileToMatrix(ByVal stlPath As String, _
                                               ByRef m() As Double) As Boolean
On Error GoTo ErrHandler

    ReorientBinaryStlFileToMatrix = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If stlPath = "" Then Exit Function
    If fso.FileExists(stlPath) = False Then Exit Function

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
        LogLine "STL reorient skipped: file does not look like binary STL. Size=" & _
                CStr(LOF(f)) & " expected=" & CStr(expectedLen)
        Exit Function
    End If

    Dim tri As BinaryStlTriangle

    If Len(tri) <> 50 Then
        Close #f
        LogLine "STL reorient skipped: BinaryStlTriangle size is " & CStr(Len(tri)) & ", expected 50."
        Exit Function
    End If

    Dim i As Long
    Dim triPos As Long

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

Private Function ReorientAsciiStlFileToMatrix(ByVal stlPath As String, _
                                              ByRef m() As Double) As Boolean
On Error GoTo ErrHandler

    ReorientAsciiStlFileToMatrix = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If stlPath = "" Then Exit Function
    If fso.FileExists(stlPath) = False Then Exit Function

    Dim txt As String
    txt = ReadAllTextFile(stlPath)

    If Trim(txt) = "" Then Exit Function

    ' Quick ASCII STL sanity checks.
    If InStr(1, Left$(txt, 512), "solid", vbTextCompare) = 0 Then Exit Function
    If InStr(1, txt, "vertex", vbTextCompare) = 0 Then Exit Function
    If InStr(1, txt, "facet normal", vbTextCompare) = 0 Then Exit Function

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Dim lines() As String
    lines = Split(txt, vbLf)

    Dim outLines() As String
    ReDim outLines(LBound(lines) To UBound(lines))

    Dim i As Long
    Dim rawLine As String
    Dim T As String
    Dim indent As String
    Dim toks() As String

    Dim x As Double
    Dim y As Double
    Dim z As Double

    For i = LBound(lines) To UBound(lines)

        rawLine = lines(i)
        T = Trim(rawLine)
        indent = LeadingWhitespace(rawLine)

        If T <> "" Then

            toks = Split(NormalizeSpaces(T), " ")

            If UBound(toks) >= 4 Then

                If LCase$(toks(0)) = "facet" And LCase$(toks(1)) = "normal" Then

                    x = ParseStlNumber(toks(2))
                    y = ParseStlNumber(toks(3))
                    z = ParseStlNumber(toks(4))

                    TransformDoubleVectorByMatrix x, y, z, m
                    NormalizeDoubleVector x, y, z

                    outLines(i) = indent & "facet normal " & _
                                  StlNumber(x) & " " & _
                                  StlNumber(y) & " " & _
                                  StlNumber(z)

                ElseIf LCase$(toks(0)) = "vertex" And UBound(toks) >= 3 Then

                    x = ParseStlNumber(toks(1))
                    y = ParseStlNumber(toks(2))
                    z = ParseStlNumber(toks(3))

                    TransformDoubleVectorByMatrix x, y, z, m

                    outLines(i) = indent & "vertex " & _
                                  StlNumber(x) & " " & _
                                  StlNumber(y) & " " & _
                                  StlNumber(z)

                Else

                    outLines(i) = rawLine

                End If

            ElseIf UBound(toks) >= 3 Then

                If LCase$(toks(0)) = "vertex" Then

                    x = ParseStlNumber(toks(1))
                    y = ParseStlNumber(toks(2))
                    z = ParseStlNumber(toks(3))

                    TransformDoubleVectorByMatrix x, y, z, m

                    outLines(i) = indent & "vertex " & _
                                  StlNumber(x) & " " & _
                                  StlNumber(y) & " " & _
                                  StlNumber(z)

                Else

                    outLines(i) = rawLine

                End If

            Else

                outLines(i) = rawLine

            End If

        Else

            outLines(i) = rawLine

        End If

    Next i

    Dim f As Integer
    f = FreeFile

    Open stlPath For Output As #f
    Print #f, Join(outLines, vbCrLf)
    Close #f

    ReorientAsciiStlFileToMatrix = True
    Exit Function

ErrHandler:
    LogLine "ReorientAsciiStlFileToMatrix error: " & Err.Description

    On Error Resume Next
    Close #f

    ReorientAsciiStlFileToMatrix = False
End Function

Private Function LeadingWhitespace(ByVal s As String) As String
On Error Resume Next

    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)
        ch = mid$(s, i, 1)

        If ch <> " " And ch <> vbTab Then
            LeadingWhitespace = Left$(s, i - 1)
            Exit Function
        End If
    Next i

    LeadingWhitespace = s
End Function

Private Function ParseStlNumber(ByVal s As String) As Double
On Error Resume Next

    s = Trim$(s)
    s = Replace(s, "D", "E")
    s = Replace(s, "d", "E")

    ' Val handles STL-style decimal points regardless of Windows locale.
    ParseStlNumber = val(s)
End Function

Private Function StlNumber(ByVal v As Double) As String
On Error Resume Next

    If Abs(v) < 0.000000000001 Then v = 0#

    StlNumber = Replace(Format$(v, "0.#########"), ",", ".")
End Function

Private Sub TransformDoubleVectorByMatrix(ByRef x As Double, _
                                          ByRef y As Double, _
                                          ByRef z As Double, _
                                          ByRef m() As Double)
On Error Resume Next

    Dim ox As Double
    Dim oy As Double
    Dim oz As Double

    ox = x
    oy = y
    oz = z

    ' Same convention as the binary STL transform and drawing-view projection:
    ' corrected X = model dot [m0, m3, m6]
    ' corrected Y = model dot [m1, m4, m7]
    ' corrected Z = model dot [m2, m5, m8]
    x = (ox * m(0)) + (oy * m(3)) + (oz * m(6))
    y = (ox * m(1)) + (oy * m(4)) + (oz * m(7))
    z = (ox * m(2)) + (oy * m(5)) + (oz * m(8))
End Sub

Private Sub NormalizeDoubleVector(ByRef x As Double, _
                                  ByRef y As Double, _
                                  ByRef z As Double)
On Error Resume Next

    Dim L As Double

    L = Sqr((x * x) + (y * y) + (z * z))

    If L <= 0.0000001 Then Exit Sub

    x = x / L
    y = y / L
    z = z / L
End Sub

Private Sub ExportIndividualHolderAndClampingDxfs(ByVal outputFolder As String)
On Error Resume Next
    LogLine "ExportIndividualHolderAndClampingDxfs skipped. DXFs created during XT export."
End Sub

' ============================================================
' END OF PART 2
' Paste Part 3 immediately after this.
' ============================================================
' ============================================================
' MAIN ASSEMBLY / HOLDERS PACKAGE
' ============================================================

Private Sub ExportMainAssemblyAndHoldersPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "MAIN ASSEMBLY package skipped: active model is not an assembly."
        Exit Sub
    End If

    EnsureFolderDeep outputFolder

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 200

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim mainToken As String
    mainToken = CleanQuoteTokenForFile(MAIN_ASSEMBLY_FILE_TOKEN)
    If mainToken = "" Then mainToken = "BASE"

    Dim mainEasmPath As String
    Dim mainXtPath As String
    Dim mainIgsPath As String
    Dim mainDxfPath As String

    mainEasmPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".easm"
    mainXtPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".x_t"
    mainIgsPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".igs"
    mainDxfPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".dxf"

    mainEasmPath = GetUniqueFilePath(mainEasmPath)
    mainXtPath = GetUniqueFilePath(mainXtPath)
    mainIgsPath = GetUniqueFilePath(mainIgsPath)
    mainDxfPath = GetUniqueFilePath(mainDxfPath)

    LogLine "Exporting MAIN ASSEMBLY BASE package."

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    SaveModelAs swModel, mainEasmPath
    SaveModelAs swModel, mainXtPath
    SaveModelAs swModel, mainIgsPath

    ' Whole base + all parts as a single STL for Match Studio assembly overlay.
    ' Done here while every component is unsuppressed and shown.
    ExportFullAssemblyStl

    If CREATE_COMPONENT_ISO_JPEGS Then
        ExportComponentIsoJpeg swModel, "BASE ISO", outputFolder
    End If

    CreateBaseDxfFromSelectedComponents mainXtPath, mainDxfPath

    Dim idHolderIdx As Long
    Dim odHolderIdx As Long

    idHolderIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odHolderIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    If idHolderIdx <= 0 Or odHolderIdx <= 0 Then
        LogLine "HOLDERS package skipped: could not find both holders."
    Else
        ExportHoldersOnlyMainFolderPackage CurrentJobFolder, idHolderIdx, odHolderIdx
    End If

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    Exit Sub

ErrHandler:
    LogLine "ExportMainAssemblyAndHoldersPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

' ============================================================
' FULL ASSEMBLY STL (whole base + every part)
' Published into the J#### STL folder so Match Studio can overlay the entire
' assembly between two jobs, alongside the individual component STLs.
' ============================================================
Private Sub ExportFullAssemblyStl()
On Error GoTo ErrHandler

    If CREATE_FULL_ASSEMBLY_STL = False Then Exit Sub
    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "FULL ASSEMBLY STL skipped: active model is not an assembly."
        Exit Sub
    End If

    ' Everything visible so the STL contains the whole base and all parts.
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    Dim stlFolder As String
    stlFolder = GetMatchStudioStlFolder()
    If stlFolder = "" Then Exit Sub

    Dim custToken As String
    Dim dateToken As String
    Dim token As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)
    token = CleanQuoteTokenForFile(FULL_ASSEMBLY_STL_TOKEN)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")
    If token = "" Then token = "FULL ASSEMBLY"

    Dim stlPath As String
    stlPath = stlFolder & "\" & _
              CurrentJobNumber & "_" & token & "_" & custToken & "_" & dateToken & ".stl"

    stlPath = GetUniqueFilePath(stlPath)

    LogLine "Saving FULL ASSEMBLY STL (whole base + all parts):"
    LogLine "  " & stlPath

    ' Build the STL from a throwaway merged part so every component ends up in
    ' one combined mesh (one STL file), then delete that temp part.
    SaveAssemblyAsMergedPartStl swModel, stlPath

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel

    Exit Sub

ErrHandler:
    LogLine "ExportFullAssemblyStl error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

' ------------------------------------------------------------
' Save the whole assembly as a temporary multibody part, combine every part
' into a single body, export the STL from it, then delete the temp part.
' This guarantees one merged STL instead of one mesh per component. If any
' step fails it falls back to exporting the STL straight from the assembly.
' ------------------------------------------------------------
Private Sub SaveAssemblyAsMergedPartStl(ByVal assyModel As Object, _
                                        ByVal stlPath As String)
On Error GoTo ErrHandler

    If assyModel Is Nothing Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Throwaway part in the system temp area (never published).
    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_FULL_ASM_MERGE"
    EnsureFolderDeep tempFolder

    Dim partPath As String
    partPath = tempFolder & "\FULL_ASSEMBLY_MERGE_" & _
               Format(Now, "yyyymmdd_hhnnss") & ".sldprt"
    partPath = GetUniqueFilePath(partPath)

    ' Save the assembly as a part with ALL components, so each part becomes a
    ' solid body in the new part. Restore the prior preference afterward.
    Dim asmAsPartSet As Boolean
    Dim priorAsmAsPart As Long
    asmAsPartSet = False
    If Not swApp Is Nothing Then
        priorAsmAsPart = swApp.GetUserPreferenceIntegerValue(swSaveAssemblyAsPartOptions)
        swApp.SetUserPreferenceIntegerValue swSaveAssemblyAsPartOptions, swSaveAsmAsPart_AllComponents
        asmAsPartSet = True
    End If

    Dim errs As Long
    Dim warns As Long

    LogLine "FULL ASSEMBLY: saving assembly as temp multibody part:"
    LogLine "  " & partPath

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

    ' Merge all solid bodies into one (best effort; STL stays one file regardless).
    MergeAllPartBodies partModel

    ' Same corrected Top/Front coordinate frame as the rest of the STL package.
    SaveStlWithMainBaseOrientation partModel, stlPath, "FULL ASSEMBLY"

    On Error Resume Next
    swApp.CloseDoc partModel.GetTitle
    Set partModel = Nothing
    ReleaseSolidWorksMemory "after FULL ASSEMBLY merge part"

    If fso.FileExists(partPath) Then
        fso.DeleteFile partPath, True
        LogLine "FULL ASSEMBLY: deleted temp merge part."
    End If

    ' Reactivate the assembly so the rest of the export continues normally.
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
    swApp.ActivateDoc3 assyModel.GetTitle, False, 0, errs
End Sub

' Fallback: export the STL directly from the assembly, forcing the single-file
' STL option so all components still land in one file.
Private Sub SaveFullAssemblyStlFromAssembly(ByVal assyModel As Object, ByVal stlPath As String)
On Error GoTo ErrHandler

    Dim priorOneFile As Boolean
    Dim oneFileSet As Boolean
    oneFileSet = False
    If Not swApp Is Nothing Then
        priorOneFile = swApp.GetUserPreferenceToggle(swSTLComponentsIntoOneFile)
        swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, True
        oneFileSet = True
    End If

    SaveStlWithMainBaseOrientation assyModel, stlPath, "FULL ASSEMBLY"

    If oneFileSet Then
        swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, priorOneFile
    End If
    Exit Sub

ErrHandler:
    LogLine "SaveFullAssemblyStlFromAssembly error: " & Err.Description
    On Error Resume Next
    If oneFileSet Then swApp.SetUserPreferenceToggle swSTLComponentsIntoOneFile, priorOneFile
End Sub

' Select every solid body in the part and Combine -> Add them into one body.
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
        If Not vBodies(i) Is Nothing Then
            vBodies(i).Select2 True, Nothing
        End If
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

Private Sub CreateBaseDxfFromSelectedComponents(ByVal fullBaseXtPath As String, _
                                                ByVal mainDxfPath As String)
On Error GoTo ErrHandler

    If FAST_BATCH_EXPORT And FAST_SKIP_BASE_PACKAGE_DXF Then
        LogLine "FAST: skipping MAIN ASSEMBLY BASE DXF."
        Exit Sub
    End If

    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "BASE selected-components DXF skipped: active model is not an assembly."

        If BaseNativeAssemblyPath <> "" Then
            CreateProjectedDxfFromNativePath BaseNativeAssemblyPath, mainDxfPath, "MAIN ASSEMBLY", _
                                             CMS_TOP_VIEW_NAME, "*Top", _
                                             False, False, False, True
        Else
            CreateProjectedDxfFromXtPath fullBaseXtPath, mainDxfPath, "MAIN ASSEMBLY", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, False, False, True
        End If

        Exit Sub
    End If

    Dim keepNames As Collection
    Set keepNames = BuildBaseDxfKeepComponentNames()

    If keepNames Is Nothing Or keepNames.count = 0 Then
        LogLine "WARNING: BASE DXF selected component list is empty. Falling back to full BASE DXF."

        If BaseNativeAssemblyPath <> "" Then
            CreateProjectedDxfFromNativePath BaseNativeAssemblyPath, mainDxfPath, "MAIN ASSEMBLY", _
                                             CMS_TOP_VIEW_NAME, "*Top", _
                                             False, False, False, True
        Else
            CreateProjectedDxfFromXtPath fullBaseXtPath, mainDxfPath, "MAIN ASSEMBLY", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, False, False, True
        End If

        Exit Sub
    End If

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_BASE_DXF_SELECTED_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    ' IMPORTANT:
    ' Use native SLDASM for DXF creation.
    ' X_T does not preserve CMS_TOP / redefined standard views.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_BASE_DXF_SELECTED_TEMP.sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    LogLine "Creating BASE DXF from selected components only. Count=" & keepNames.count

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "WARNING: Could not isolate selected BASE DXF components. Falling back to full BASE DXF."

        If BaseNativeAssemblyPath <> "" Then
            CreateProjectedDxfFromNativePath BaseNativeAssemblyPath, mainDxfPath, "MAIN ASSEMBLY", _
                                             CMS_TOP_VIEW_NAME, "*Top", _
                                             False, False, False, True
        Else
            CreateProjectedDxfFromXtPath fullBaseXtPath, mainDxfPath, "MAIN ASSEMBLY", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, False, False, True
        End If

        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    If MATCH_STUDIO_TURBO_ONLY Then
        LogLine "MATCH STUDIO TURBO: skipping BASE DXF creation."
        GoTo CleanExit
    End If

    LogLine "BASE DXF: saving selected-component native temp SLDASM:"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(swModel, tempNativePath) = False Then GoTo CleanExit

    CreateProjectedDxfFromNativePath tempNativePath, mainDxfPath, "MAIN ASSEMBLY", _
                                     CMS_TOP_VIEW_NAME, "*Top", _
                                     False, False, False, True

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.count > 0 Then
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

    If Not fso Is Nothing Then
        If tempFolder <> "" Then
            If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
        End If
    End If

    Exit Sub

ErrHandler:
    LogLine "CreateBaseDxfFromSelectedComponents error: " & Err.Description
    Resume CleanExit
End Sub

Private Function BuildBaseDxfKeepComponentNames() As Collection
On Error GoTo ErrHandler

    Dim keepNames As New Collection

    AddBaseDxfKeepComponentFromQuote keepNames, "TCP", "TCP", _
        "TCP|TOP SMED|TOP SMED PLATE|TOP CLAMPING PLATE|TOP CLAMPING|ID SMED"

    AddBaseDxfKeepComponentFromQuote keepNames, "BCP", "BCP", _
        "BCP|BOTTOM SMED|BOT SMED|BOTTOM SMED PLATE|BOT SMED PLATE|BOTTOM CLAMPING PLATE|BOT CLAMPING PLATE|BOTTOM CLAMPING|BOT CLAMPING|OD SMED"

    AddBaseDxfKeepComponentFromQuote keepNames, "ID HOLDER", "ID HOLDER", _
        ID_HOLDER_KEYS

    AddBaseDxfKeepComponentFromQuote keepNames, "OD HOLDER", "OD HOLDER", _
        OD_HOLDER_KEYS

    AddBaseDxfKeepComponentFromQuote keepNames, "ID POT", "ID POT BLOCK", _
        "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT"

    AddBaseDxfKeepComponentFromQuote keepNames, "OD POT", "OD POT BLOCK", _
        "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT"

    Set BuildBaseDxfKeepComponentNames = keepNames
    Exit Function

ErrHandler:
    LogLine "BuildBaseDxfKeepComponentNames error: " & Err.Description
    Set BuildBaseDxfKeepComponentNames = New Collection
End Function

Private Sub AddBaseDxfKeepComponentFromQuote(ByVal keepNames As Collection, _
                                             ByVal label As String, _
                                             ByVal quoteName As String, _
                                             ByVal fallbackKeys As String)
On Error GoTo ErrHandler

    If keepNames Is Nothing Then Exit Sub

    Dim cadIdx As Long

    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then

        AddUniqueComponentName keepNames, parts(cadIdx).componentName

        LogLine "BASE DXF include " & label & ": CAD '" & _
                parts(cadIdx).componentName & "'  L/W/T=" & _
                FormatNumberForCsv(parts(cadIdx).Length) & "/" & _
                FormatNumberForCsv(parts(cadIdx).Width) & "/" & _
                FormatNumberForCsv(parts(cadIdx).Thickness)
    Else
        LogLine "WARNING: BASE DXF could not find component for " & label & _
                " using quote '" & quoteName & "'."
    End If

    Exit Sub

ErrHandler:
    LogLine "AddBaseDxfKeepComponentFromQuote error (" & label & "): " & Err.Description
End Sub

Private Function FindCadIndexFromExportQuote(ByVal quoteName As String) As Long
On Error GoTo ErrHandler

    FindCadIndexFromExportQuote = 0

    Dim k As String
    k = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            If ExportRows(i).HasCad Then
                FindCadIndexFromExportQuote = ExportRows(i).CadPartIndex
                Exit Function
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    FindCadIndexFromExportQuote = 0
End Function

Private Sub AddUniqueComponentName(ByVal names As Collection, ByVal componentName As String)
On Error Resume Next

    If names Is Nothing Then Exit Sub

    componentName = Trim(componentName)
    If componentName = "" Then Exit Sub

    Dim i As Long

    For i = 1 To names.count
        If LCase(CStr(names(i))) = LCase(componentName) Then Exit Sub
    Next i

    names.Add componentName
End Sub

Private Sub ExportHoldersOnlyMainFolderPackage(ByVal outputFolder As String, _
                                               ByVal idHolderIdx As Long, _
                                               ByVal odHolderIdx As Long)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocASSEMBLY Then Exit Sub
    If idHolderIdx <= 0 Or idHolderIdx > PartCount Then Exit Sub
    If odHolderIdx <= 0 Or odHolderIdx > PartCount Then Exit Sub

    Dim idPart As PartInfo
    Dim odPart As PartInfo

    idPart = parts(idHolderIdx)
    odPart = parts(odHolderIdx)

    Dim custToken As String
    Dim dateToken As String
    Dim holderToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)
    holderToken = CleanQuoteTokenForFile(HOLDERS_ONLY_FILE_TOKEN)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")
    If holderToken = "" Then holderToken = "HOLDERS"

    Dim holdersIgsPath As String
    Dim holdersStlPath As String
    Dim holdersDxfPath As String

    holdersIgsPath = outputFolder & "\" & CurrentJobNumber & "_" & holderToken & "_" & custToken & "_" & dateToken & ".igs"
    holdersStlPath = outputFolder & "\" & CurrentJobNumber & "_" & holderToken & "_" & custToken & "_" & dateToken & ".stl"
    holdersDxfPath = outputFolder & "\" & CurrentJobNumber & "_" & holderToken & "_" & custToken & "_" & dateToken & ".dxf"

    holdersIgsPath = GetUniqueFilePath(holdersIgsPath)
    holdersStlPath = GetUniqueFilePath(holdersStlPath)
    holdersDxfPath = GetUniqueFilePath(holdersDxfPath)

    Dim tempFolder As String
    Dim holdersTempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_HOLDERS_ONLY_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    ' IMPORTANT:
    ' Native SLDASM for DXF so CMS_TOP / standard orientation survives.
    holdersTempNativePath = tempFolder & "\" & CurrentJobNumber & "_HOLDERS_TEMP.sldasm"

    LogLine "Exporting HOLDERS package."

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    If HideAllExceptTwoTargetComponentsOnce(swModel, idPart.componentName, odPart.componentName, hiddenNames) = False Then
        ShowOnlyTwoAssemblyComponents swModel, idPart.componentName, odPart.componentName
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 200

    SaveModelAs swModel, holdersIgsPath

    If CREATE_MATCH_STUDIO_HOLDERS_STL Then
        LogLine "Exporting HOLDERS package STL for Match Studio overlay."
        SaveStlWithMainBaseOrientation swModel, holdersStlPath, "HOLDERS"
    End If

    If FAST_BATCH_EXPORT And FAST_SKIP_HOLDERS_PACKAGE_DXF Then
        LogLine "FAST: skipping HOLDERS DXF (IGS still saved)."
    Else
        LogLine "HOLDERS DXF: saving native temp SLDASM:"
        LogLine "  " & holdersTempNativePath

        If SaveModelCopyAs(swModel, holdersTempNativePath) = False Then GoTo CleanExit

        CreateProjectedDxfFromNativePath holdersTempNativePath, holdersDxfPath, "HOLDERS", _
                                         CMS_TOP_VIEW_NAME, "*Top", _
                                         False, False, False, True
    End If

    If hiddenNames.count > 0 Then
        ShowNamedComponentsOnce swModel, hiddenNames
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

CleanExit:
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso Is Nothing Then
        If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportHoldersOnlyMainFolderPackage error: " & Err.Description
    On Error Resume Next
    ShowAllAssemblyComponents swModel
    Resume CleanExit
End Sub

' ============================================================
' FAST MATCH STUDIO STL EXPORT HELPERS
' These save STLs during already-isolated export states.
' No extra suppress/hide pass.
' ============================================================

Private Function GetMatchStudioStlFolder() As String
On Error Resume Next

    Dim folderPath As String

    folderPath = CurrentJobFolder & "\" & CurrentJobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX

    EnsureFolderDeep folderPath

    GetMatchStudioStlFolder = folderPath
End Function

Private Function BuildMatchStudioStlPath(ByVal label As String) As String
On Error GoTo ErrHandler

    Dim stlFolder As String
    Dim custToken As String
    Dim dateToken As String
    Dim token As String

    stlFolder = GetMatchStudioStlFolder()

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)
    token = CleanQuoteTokenForFile(label)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")
    If token = "" Then token = "ITEM"

    BuildMatchStudioStlPath = stlFolder & "\" & _
                              CurrentJobNumber & "_" & token & "_" & _
                              custToken & "_" & dateToken & ".stl"

    Exit Function

ErrHandler:
    BuildMatchStudioStlPath = ""
End Function

Private Function MatchStudioStlLabelFromQuote(ByVal quoteName As String) As String
On Error GoTo ErrHandler

    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k

        Case "TCP", _
             "TOPSMED", _
             "TOPSMEDPLATE", _
             "TOPCLAMPING", _
             "TOPCLAMPINGPLATE"
            MatchStudioStlLabelFromQuote = "TCP"
            Exit Function

        Case "BCP", _
             "BOTTOMSMED", _
             "BOTSMED", _
             "BOTTOMSMEDPLATE", _
             "BOTSMEDPLATE", _
             "BOTTOMCLAMPING", _
             "BOTCLAMPING", _
             "BOTTOMCLAMPINGPLATE", _
             "BOTCLAMPINGPLATE"
            MatchStudioStlLabelFromQuote = "BCP"
            Exit Function

        Case "IDHOLDER", "TOPHOLDER", "IDTEHOLDER"
            MatchStudioStlLabelFromQuote = "ID HOLDER"
            Exit Function

        Case "ODHOLDER", "BOTTOMHOLDER", "BOTHOLDER", "ODTEHOLDER"
            MatchStudioStlLabelFromQuote = "OD HOLDER"
            Exit Function

        Case "IDPOT", "IDPOTBLOCK", "TOPPOT", "TOPPOTBLOCK"
            MatchStudioStlLabelFromQuote = "ID POT"
            Exit Function

        Case "ODPOT", "ODPOTBLOCK", "BOTTOMPOT", "BOTTOMPOTBLOCK", "BOTPOT", "BOTPOTBLOCK"
            MatchStudioStlLabelFromQuote = "OD POT"
            Exit Function

    End Select

    MatchStudioStlLabelFromQuote = ""
    Exit Function

ErrHandler:
    MatchStudioStlLabelFromQuote = ""
End Function

Private Sub SaveMatchStudioVisibleComponentStl(ByVal model As Object, _
                                               ByVal label As String, _
                                               Optional ByVal sourceComponentName As String = "")
On Error GoTo ErrHandler

    If CREATE_MATCH_STUDIO_STLS_DURING_EXISTING_EXPORTS = False Then Exit Sub
    If model Is Nothing Then Exit Sub
    If label = "" Then Exit Sub

    If MatchStudioStlExported Is Nothing Then
        Set MatchStudioStlExported = CreateObject("Scripting.Dictionary")
    End If

    Dim key As String
    key = NormalizeKey(label)

    If MatchStudioStlExported.Exists(key) Then Exit Sub

    Dim stlPath As String
    stlPath = BuildMatchStudioStlPath(label)

    If stlPath = "" Then Exit Sub

    LogLine "Fast Match Studio STL save while already isolated:"
    LogLine "  " & label & " -> " & stlPath

    SaveStlWithMainBaseOrientation model, stlPath, label, sourceComponentName

    MatchStudioStlExported(key) = stlPath

    Exit Sub

ErrHandler:
    LogLine "SaveMatchStudioVisibleComponentStl error (" & label & "): " & Err.Description
End Sub

' ============================================================
' MATCH STUDIO STL PACKAGE (slow fallback ? not used by default)
' Exports individual STL files and one combined six-component STL:
'   ID POT, OD POT, BCP, TCP, ID HOLDER, OD HOLDER
' ============================================================

Private Sub ExportMatchStudioStlPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "MATCH STUDIO STL package skipped: active model is not an assembly."
        Exit Sub
    End If

    Dim stlFolder As String
    stlFolder = outputFolder & "\" & CurrentJobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX

    EnsureFolderDeep stlFolder

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    LogLine "Creating Match Studio STL folder:"
    LogLine "  " & stlFolder

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel

    ExportMatchStudioOneComponentStl stlFolder, "ID POT", "ID POT BLOCK", _
        "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT", _
        custToken, dateToken

    ExportMatchStudioOneComponentStl stlFolder, "OD POT", "OD POT BLOCK", _
        "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT", _
        custToken, dateToken

    ExportMatchStudioOneComponentStl stlFolder, "BCP", "BCP", _
        "BCP|BOTTOM SMED|BOT SMED|BOTTOM SMED PLATE|BOT SMED PLATE|BOTTOM CLAMPING PLATE|BOT CLAMPING PLATE|BOTTOM CLAMPING|BOT CLAMPING|OD SMED", _
        custToken, dateToken

    ExportMatchStudioOneComponentStl stlFolder, "TCP", "TCP", _
        "TCP|TOP SMED|TOP SMED PLATE|TOP CLAMPING PLATE|TOP CLAMPING|ID SMED", _
        custToken, dateToken

    ExportMatchStudioOneComponentStl stlFolder, "ID HOLDER", "ID HOLDER", _
        ID_HOLDER_KEYS, _
        custToken, dateToken

    ExportMatchStudioOneComponentStl stlFolder, "OD HOLDER", "OD HOLDER", _
        OD_HOLDER_KEYS, _
        custToken, dateToken

    ExportMatchStudioCombinedSixStl stlFolder, custToken, dateToken

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel

    LogLine "Match Studio STL package complete."

    Exit Sub

ErrHandler:
    LogLine "ExportMatchStudioStlPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Sub ExportMatchStudioOneComponentStl(ByVal stlFolder As String, _
                                             ByVal label As String, _
                                             ByVal quoteName As String, _
                                             ByVal fallbackKeys As String, _
                                             ByVal custToken As String, _
                                             ByVal dateToken As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim cadIdx As Long
    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    If cadIdx <= 0 Or cadIdx > PartCount Then
        LogLine "MATCH STUDIO STL skipped missing component: " & label
        Exit Sub
    End If

    Dim keepNames As Collection
    Set keepNames = New Collection

    AddUniqueComponentName keepNames, parts(cadIdx).componentName

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "MATCH STUDIO STL failed to isolate: " & label
        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 50

    Dim token As String
    token = CleanQuoteTokenForFile(label)
    If token = "" Then token = "ITEM"

    Dim stlPath As String
    stlPath = stlFolder & "\" & _
              CurrentJobNumber & "_" & token & "_" & custToken & "_" & dateToken & ".stl"

    stlPath = GetUniqueFilePath(stlPath)

    LogLine "Saving Match Studio component STL:"
    LogLine "  " & label & " -> " & stlPath

    SaveStlWithMainBaseOrientation swModel, stlPath, label, parts(cadIdx).componentName

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel
    Exit Sub

ErrHandler:
    LogLine "ExportMatchStudioOneComponentStl error (" & label & "): " & Err.Description
    Resume CleanExit
End Sub

Private Sub ExportMatchStudioCombinedSixStl(ByVal stlFolder As String, _
                                            ByVal custToken As String, _
                                            ByVal dateToken As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim keepNames As Collection
    Set keepNames = BuildMatchStudioSixComponentNames()

    If keepNames Is Nothing Or keepNames.count = 0 Then
        LogLine "MATCH STUDIO combined STL skipped: no six-pack components found."
        Exit Sub
    End If

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "MATCH STUDIO combined STL failed: could not isolate six components."
        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 50

    Dim token As String
    token = CleanQuoteTokenForFile("MATCH SET")
    If token = "" Then token = "MATCH SET"

    Dim stlPath As String
    stlPath = stlFolder & "\" & _
              CurrentJobNumber & "_" & token & "_" & custToken & "_" & dateToken & ".stl"

    stlPath = GetUniqueFilePath(stlPath)

    LogLine "Saving Match Studio combined six-component STL:"
    LogLine "  Components kept=" & CStr(keepNames.count)
    LogLine "  " & stlPath

    SaveStlWithMainBaseOrientation swModel, stlPath, "MATCH SET"

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel
    Exit Sub

ErrHandler:
    LogLine "ExportMatchStudioCombinedSixStl error: " & Err.Description
    Resume CleanExit
End Sub

Private Function BuildMatchStudioSixComponentNames() As Collection
On Error GoTo ErrHandler

    Dim keepNames As New Collection

    AddMatchStudioStlKeepComponent keepNames, "TCP", "TCP", _
        "TCP|TOP SMED|TOP SMED PLATE|TOP CLAMPING PLATE|TOP CLAMPING|ID SMED"

    AddMatchStudioStlKeepComponent keepNames, "BCP", "BCP", _
        "BCP|BOTTOM SMED|BOT SMED|BOTTOM SMED PLATE|BOT SMED PLATE|BOTTOM CLAMPING PLATE|BOT CLAMPING PLATE|BOTTOM CLAMPING|BOT CLAMPING|OD SMED"

    AddMatchStudioStlKeepComponent keepNames, "ID HOLDER", "ID HOLDER", _
        ID_HOLDER_KEYS

    AddMatchStudioStlKeepComponent keepNames, "OD HOLDER", "OD HOLDER", _
        OD_HOLDER_KEYS

    AddMatchStudioStlKeepComponent keepNames, "ID POT", "ID POT BLOCK", _
        "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT"

    AddMatchStudioStlKeepComponent keepNames, "OD POT", "OD POT BLOCK", _
        "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT"

    Set BuildMatchStudioSixComponentNames = keepNames
    Exit Function

ErrHandler:
    LogLine "BuildMatchStudioSixComponentNames error: " & Err.Description
    Set BuildMatchStudioSixComponentNames = New Collection
End Function

Private Sub AddMatchStudioStlKeepComponent(ByVal keepNames As Collection, _
                                           ByVal label As String, _
                                           ByVal quoteName As String, _
                                           ByVal fallbackKeys As String)
On Error GoTo ErrHandler

    If keepNames Is Nothing Then Exit Sub

    Dim cadIdx As Long
    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then
        AddUniqueComponentName keepNames, parts(cadIdx).componentName

        LogLine "MATCH STUDIO combined STL include " & label & ": " & _
                parts(cadIdx).componentName
    Else
        LogLine "WARNING: MATCH STUDIO combined STL missing " & label
    End If

    Exit Sub

ErrHandler:
    LogLine "AddMatchStudioStlKeepComponent error (" & label & "): " & Err.Description
End Sub

Private Function FindHolderCadIndexForMainPackage(ByVal holderQuoteName As String) As Long
On Error GoTo ErrHandler

    Dim k As String
    k = NormalizeKey(holderQuoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            If ExportRows(i).HasCad Then
                FindHolderCadIndexForMainPackage = ExportRows(i).CadPartIndex
                Exit Function
            End If
        End If
    Next i

    If k = "IDHOLDER" Then
        FindHolderCadIndexForMainPackage = FindCadPartIndexByQuoteOrKeys("ID HOLDER", ID_HOLDER_KEYS)
        Exit Function
    ElseIf k = "ODHOLDER" Then
        FindHolderCadIndexForMainPackage = FindCadPartIndexByQuoteOrKeys("OD HOLDER", OD_HOLDER_KEYS)
        Exit Function
    End If

    FindHolderCadIndexForMainPackage = 0
    Exit Function

ErrHandler:
    LogLine "FindHolderCadIndexForMainPackage error for " & holderQuoteName & ": " & Err.Description
    FindHolderCadIndexForMainPackage = 0
End Function

' ============================================================

' ============================================================
' ENGRAVING DXF HELPERS - ADDED BY PATCHER
' Creates one-view, solid engraving DXFs for:
'   ID HOLDER ENG
'   OD HOLDER ENG
'   TCP ENG
' ============================================================

Private Function IsEngravingDxfQuote(ByVal quoteName As String, _
                                     ByRef canonicalQuoteName As String, _
                                     ByRef engravingLabel As String, _
                                     ByRef parentViewName As String, _
                                     ByRef parentViewId As Long) As Boolean
On Error GoTo ErrHandler

    IsEngravingDxfQuote = False

    canonicalQuoteName = ""
    engravingLabel = ""
    parentViewName = "*Top"
    parentViewId = 5

    Dim qKey As String
    qKey = NormalizeKey(quoteName)

    Select Case qKey

        Case "IDHOLDER", "TOPHOLDER", "IDTEHOLDER", "IDMOLDBASE", "IDMOLDBASE"
            canonicalQuoteName = "ID HOLDER"
            engravingLabel = "ID HOLDER" & ENGRAVING_DXF_TOKEN_SUFFIX

            ' ID HOLDER engraving uses actual SolidWorks *Bottom view directly.
            parentViewName = "*Bottom"
            parentViewId = 6

            IsEngravingDxfQuote = True
            Exit Function

        Case "ODHOLDER", "BOTTOMHOLDER", "BOTHOLDER", "ODTEHOLDER", "ODMOLDBASE", "ODMOLDBASE"
            canonicalQuoteName = "OD HOLDER"
            engravingLabel = "OD HOLDER" & ENGRAVING_DXF_TOKEN_SUFFIX

            ' OD HOLDER engraving stays on corrected SolidWorks *Top view.
            parentViewName = "*Top"
            parentViewId = 5

            IsEngravingDxfQuote = True
            Exit Function

        Case "TCP", _
             "TOPSMED", _
             "TOPSMEDPLATE", _
             "TOPCLAMPING", _
             "TOPCLAMPINGPLATE"
            canonicalQuoteName = "TCP"
            engravingLabel = "TCP" & ENGRAVING_DXF_TOKEN_SUFFIX
            IsEngravingDxfQuote = True
            Exit Function

    End Select

    Exit Function

ErrHandler:
    IsEngravingDxfQuote = False
End Function

Private Sub MaybeCreateEngravingDxfForIsolatedComponent(ByVal assyModel As Object, _
                                                        ByVal quoteName As String, _
                                                        ByVal normalDxfPath As String)
On Error GoTo ErrHandler

    If CREATE_ENGRAVING_DXFS = False Then Exit Sub
    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim canonicalQuoteName As String
    Dim engravingLabel As String
    Dim parentViewName As String
    Dim parentViewId As Long

    If IsEngravingDxfQuote(quoteName, canonicalQuoteName, engravingLabel, parentViewName, parentViewId) = False Then
        Exit Sub
    End If

    Dim engravingDxfPath As String
    engravingDxfPath = BuildEngravingDxfPath(normalDxfPath, engravingLabel)

    If engravingDxfPath = "" Then Exit Sub

    LogLine "Creating engraving DXF:"
    LogLine "  Quote: " & quoteName
    LogLine "  Label: " & engravingLabel
    LogLine "  Path : " & engravingDxfPath

    If CreateEngravingDxfFromIsolatedAssemblyComponent( _
            assyModel, _
            engravingDxfPath, _
            canonicalQuoteName, _
            engravingLabel, _
            parentViewName, _
            parentViewId) Then

        LogLine "Engraving DXF created: " & engravingDxfPath

    Else

        LogLine "WARNING: Engraving DXF was not created: " & engravingDxfPath

    End If

    Exit Sub

ErrHandler:
    LogLine "MaybeCreateEngravingDxfForIsolatedComponent error: " & Err.Description
End Sub

Private Function BuildEngravingDxfPath(ByVal normalDxfPath As String, _
                                       ByVal engravingLabel As String) As String
On Error GoTo ErrHandler

    BuildEngravingDxfPath = ""

    Dim folderPath As String
    folderPath = GetParentFolderPath(normalDxfPath)

    If folderPath = "" Then folderPath = CurrentJobFolder
    If folderPath = "" Then Exit Function

    EnsureFolderDeep folderPath

    Dim custToken As String
    Dim dateToken As String
    Dim quoteToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)
    quoteToken = CleanQuoteTokenForFile(engravingLabel)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")
    If quoteToken = "" Then quoteToken = "ENG"

    BuildEngravingDxfPath = GetUniqueFilePath( _
        folderPath & "\" & _
        CurrentJobNumber & "_" & _
        quoteToken & "_" & _
        custToken & "_" & _
        dateToken & ".dxf")

    Exit Function

ErrHandler:
    BuildEngravingDxfPath = ""
End Function

Private Function CreateEngravingDxfFromIsolatedAssemblyComponent(ByVal assyModel As Object, _
                                                                 ByVal dxfPath As String, _
                                                                 ByVal canonicalQuoteName As String, _
                                                                 ByVal engravingLabel As String, _
                                                                 ByVal parentViewName As String, _
                                                                 ByVal parentViewId As Long) As Boolean
On Error GoTo ErrHandler

    CreateEngravingDxfFromIsolatedAssemblyComponent = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function
    If dxfPath = "" Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_ENG_DXF_" & _
                 Format(Now, "yyyymmdd_hhnnss") & "_" & _
                 NormalizeKey(engravingLabel)

    EnsureFolderDeep tempFolder

    Dim tempNativePath As String
    tempNativePath = tempFolder & "\" & _
                     CurrentJobNumber & "_" & _
                     NormalizeKey(engravingLabel) & _
                     "_TEMP.sldasm"

    Dim oldIdHolderDxfFromAssembly As Boolean
    oldIdHolderDxfFromAssembly = CurrentIdHolderDxfFromAssembly

    If NormalizeKey(canonicalQuoteName) = "IDHOLDER" Then
        CurrentIdHolderDxfFromAssembly = True
    End If

    assyModel.ShowNamedView2 parentViewName, parentViewId
    StabilizeActiveView assyModel, 100

    LogLine "Engraving DXF: saving isolated native SLDASM:"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(assyModel, tempNativePath) = False Then GoTo CleanExit

    If fso.FileExists(tempNativePath) = False Then
        LogLine "Engraving DXF skipped: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    CreateSingleViewSolidDxfFromNativePath tempNativePath, _
                                           dxfPath, _
                                           canonicalQuoteName, _
                                           parentViewName, _
                                           CMS_TOP_VIEW_NAME

    CreateEngravingDxfFromIsolatedAssemblyComponent = fso.FileExists(dxfPath)

CleanExit:
    On Error Resume Next

    CurrentIdHolderDxfFromAssembly = oldIdHolderDxfFromAssembly

    ApplyCmsTopView assyModel

    If Not fso Is Nothing Then
        If tempFolder <> "" Then
            If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateEngravingDxfFromIsolatedAssemblyComponent error: " & Err.Description
    Resume CleanExit
End Function

Private Sub CreateSingleViewSolidDxfFromNativePath(ByVal nativePath As String, _
                                                   ByVal dxfPath As String, _
                                                   ByVal quoteName As String, _
                                                   ByVal parentPrimaryView As String, _
                                                   ByVal parentFallbackView As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim drawTitle As String
    drawTitle = ""

    Dim freezeApplied As Boolean
    freezeApplied = False

    If fso.FileExists(SW_DRAWING_TEMPLATE_PATH) = False Then
        LogLine "Engraving DXF skipped. Drawing template not found: " & SW_DRAWING_TEMPLATE_PATH
        Exit Sub
    End If

    If fso.FileExists(nativePath) = False Then
        LogLine "Engraving DXF skipped. Native source missing: " & nativePath
        Exit Sub
    End If

    Dim partL As Double
    Dim partW As Double
    Dim partT As Double

    partL = 0#
    partW = 0#
    partT = 0#

    GetBestCurrentOutputDimensions quoteName, partL, partW, partT

    Dim liveL As Double
    Dim liveW As Double
    Dim liveT As Double

    If TryGetNativeModelDimsInches(nativePath, liveL, liveW, liveT) Then
        If liveL > 0# Then partL = liveL
        If liveW > 0# Then partW = liveW
        If liveT > 0# Then partT = liveT
    End If

    If partL <= 0# Then partL = 10#
    If partW <= 0# Then partW = 10#
    If partT <= 0# Then partT = 2#

    CurrentDxfForce1to1 = FORCE_ALL_DXF_VIEWS_1_TO_1

    Dim scaleVal As Double

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Then
        scaleVal = 1#
        LogLine "Engraving DXF scale forced 1:1 for " & quoteName
    Else
        scaleVal = CalculateSingleViewDxfScale(partL, partW)
        If scaleVal <= 0# Then scaleVal = 1#
        LogLine "Engraving DXF auto-fit scale for " & quoteName & " = " & Format(scaleVal, "0.0000")
    End If

    Dim sheetWIn As Double
    Dim sheetHIn As Double

    sheetWIn = E_SHEET_WIDTH_IN
    sheetHIn = E_SHEET_HEIGHT_IN

    Dim swDraw As Object
    Set swDraw = swApp.NewDocument(SW_DRAWING_TEMPLATE_PATH, 0, _
                                   sheetWIn / INCHES_PER_METER, _
                                   sheetHIn / INCHES_PER_METER)

    If swDraw Is Nothing Then
        LogLine "Engraving DXF skipped. Could not create drawing."
        GoTo CleanExit
    End If

    drawTitle = swDraw.GetTitle

    Dim errs As Long
    swApp.ActivateDoc3 drawTitle, False, 0, errs
    EnsureSwHidden

    SetupDrawingAsESize swDraw, sheetWIn, sheetHIn

    If FREEZE_DXF_DRAWING_GRAPHICS Then
        FreezeDxfDrawingGraphics swDraw
        freezeApplied = True
    End If

    Dim centerX As Double
    Dim centerY As Double

    centerX = sheetWIn / 2#
    centerY = sheetHIn / 2#

    Dim parentView As Object

    Set parentView = CreateParentDrawingView(swDraw, _
                                             nativePath, _
                                             parentPrimaryView, _
                                             parentFallbackView, _
                                             centerX, _
                                             centerY, _
                                             scaleVal)

    If parentView Is Nothing Then
        LogLine "Engraving DXF skipped. Could not create center drawing view."
        GoTo CleanExit
    End If

    ApplyEngravingCenterViewRotation parentView, quoteName

    SetDrawingViewSolid parentView

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Then
        ForceAllDxfScales1To1 swDraw, sheetWIn, sheetHIn
    End If

    SetDrawingViewSolid parentView

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    Dim saveErrs As Long
    Dim saveWarns As Long

    LogLine "Saving engraving DXF: " & dxfPath
    EnsureSwHidden

    swDraw.Extension.SaveAs3 dxfPath, _
                             swSaveAsCurrentVersion, _
                             swSaveAsOptions_Silent, _
                             Nothing, Nothing, saveErrs, saveWarns

    LogLine "Engraving DXF save done. Errors=" & saveErrs & " Warnings=" & saveWarns

CleanExit:
    On Error Resume Next

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    If drawTitle <> "" Then swApp.CloseDoc drawTitle

    Set parentView = Nothing
    Set swDraw = Nothing
    Set fso = Nothing

    ReleaseSolidWorksMemory "after engraving DXF save"
    Exit Sub

ErrHandler:
    LogLine "CreateSingleViewSolidDxfFromNativePath error for " & quoteName & ": " & Err.Description
    Resume CleanExit
End Sub

Private Sub ApplyEngravingCenterViewRotation(ByVal swView As Object, _
                                             ByVal quoteName As String)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    Dim qKey As String
    qKey = NormalizeKey(quoteName)

    Dim angleRad As Double
    angleRad = 0#

    Select Case qKey

        Case "TCP", _
             "TOPSMED", _
             "TOPSMEDPLATE", _
             "TOPCLAMPING", _
             "TOPCLAMPINGPLATE"

            If Abs(TCP_DXF_CENTER_ROTATION_DEG) > 0.000001 Then
                angleRad = angleRad + DegToRad(TCP_DXF_CENTER_ROTATION_DEG)

                LogLine "TCP engraving DXF: rotating center view " & _
                        Format(TCP_DXF_CENTER_ROTATION_DEG, "0.00") & " deg."
            End If

        Case "IDHOLDER"

            ' ID HOLDER engraving uses *Bottom directly.
            ' No extra in-plane rotation.
            angleRad = 0#
            LogLine "ID HOLDER engraving DXF: using *Bottom view directly with no rotation."
        Case "ODHOLDER"

            ' OD HOLDER engraving stays on *Top and keeps the normal OD holder rotation setting.
            If Abs(OD_HOLDER_CENTER_ROTATION_DEG) > 0.000001 Then
                angleRad = angleRad + DegToRad(OD_HOLDER_CENTER_ROTATION_DEG)

                LogLine "OD HOLDER engraving DXF: rotating center view " & _
                        Format(OD_HOLDER_CENTER_ROTATION_DEG, "0.00") & " deg."
            Else
                LogLine "OD HOLDER engraving DXF: using *Top view with no extra rotation."
            End If
    End Select

    If Abs(angleRad) > 0.000001 Then
        RotateDrawingViewByAngle swView, angleRad
    End If
End Sub

Private Function CalculateSingleViewDxfScale(ByVal partL As Double, _
                                             ByVal partW As Double) As Double
On Error GoTo ErrHandler

    Dim usableW As Double
    Dim usableH As Double

    usableW = E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)
    usableH = E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)

    Dim sW As Double
    Dim sh As Double
    Dim scaleVal As Double

    scaleVal = DXF_MAX_SCALE

    If partL > 0# Then
        sW = usableW / partL
        If sW < scaleVal Then scaleVal = sW
    End If

    If partW > 0# Then
        sh = usableH / partW
        If sh < scaleVal Then scaleVal = sh
    End If

    If scaleVal <= 0# Then scaleVal = 1#
    If scaleVal > DXF_MAX_SCALE Then scaleVal = DXF_MAX_SCALE

    CalculateSingleViewDxfScale = scaleVal
    Exit Function

ErrHandler:
    CalculateSingleViewDxfScale = 1#
End Function

' ============================================================
' END ENGRAVING DXF HELPERS - ADDED BY PATCHER
' ============================================================

' STANDARD DXF FROM SAVED XT
' ============================================================

Private Sub CreateStandardPrintDxfFromXtPath(ByVal xtPath As String, _
                                             ByVal dxfPath As String, _
                                             ByVal quoteName As String)
On Error GoTo ErrHandler

    If xtPath = "" Then Exit Sub
    If dxfPath = "" Then Exit Sub

    Dim parentPrimary As String
    Dim parentFallback As String

    Select Case NormalizeKey(quoteName)

        Case "IDHOLDER", "ODHOLDER", _
             "TCP", "BCP", _
             "TOPSMED", "BOTTOMSMED", "BOTSMED", _
             "TOPSMEDPLATE", "BOTTOMSMEDPLATE", "BOTSMEDPLATE", _
             "TOPCLAMPINGPLATE", "BOTTOMCLAMPINGPLATE", "BOTCLAMPINGPLATE"

            parentPrimary = "*Top"
            parentFallback = CMS_TOP_VIEW_NAME

        Case Else
            parentPrimary = CMS_TOP_VIEW_NAME
            parentFallback = "*Top"

    End Select

    CreateProjectedDxfFromXtPath xtPath, dxfPath, quoteName, _
                                 parentPrimary, parentFallback, True, False, False

    Exit Sub

ErrHandler:
    LogLine "CreateStandardPrintDxfFromXtPath error for " & quoteName & ": " & Err.Description
End Sub

Private Function CreateHolderDxfFromAssemblyBottomView(ByVal assyModel As Object, _
                                                       ByVal dxfPath As String, _
                                                       ByVal quoteName As String) As Boolean
On Error GoTo ErrHandler

    CreateHolderDxfFromAssemblyBottomView = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_HOLDER_DXF_" & Format(Now, "yyyymmdd_hhnnss") & "_" & NormalizeKey(quoteName)
    EnsureFolderDeep tempFolder

    Dim holderViewName As String
    Dim holderViewId As Long
    Dim holderViewToken As String

    If NormalizeKey(quoteName) = "IDHOLDER" Then

        ' ID HOLDER: use the actual SolidWorks *Bottom view directly.
        ' No manual 180-degree flip is applied here.
        holderViewName = "*Bottom"
        holderViewId = 6
        holderViewToken = "BOTTOMVIEW"

        LogLine "ID HOLDER DXF: using SolidWorks *Bottom view directly."

    Else

        ' OD HOLDER stays on the corrected SolidWorks *Top view.
        holderViewName = "*Top"
        holderViewId = 5
        holderViewToken = "TOPVIEW"

        LogLine quoteName & " holder DXF: using SolidWorks *Top view."

    End If

    ' Native assembly, not X_T.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_" & NormalizeKey(quoteName) & "_" & holderViewToken & "_TEMP.sldasm"

    ApplyCmsTopView assyModel
    StabilizeActiveView assyModel, 100

    assyModel.ShowNamedView2 holderViewName, holderViewId
    StabilizeActiveView assyModel, 100

    LogLine quoteName & " holder DXF: saving native temp SLDASM from " & holderViewName & ":"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(assyModel, tempNativePath) = False Then GoTo CleanExit

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "Holder assembly " & holderViewName & " DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    CurrentIdHolderDxfFromAssembly = (NormalizeKey(quoteName) = "IDHOLDER")

    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, quoteName, _
                                     holderViewName, CMS_TOP_VIEW_NAME, _
                                     True, False, False, False, False

    CurrentIdHolderDxfFromAssembly = False

    CreateHolderDxfFromAssemblyBottomView = fso.FileExists(dxfPath)

CleanExit:
    On Error Resume Next

    CurrentIdHolderDxfFromAssembly = False

    ApplyCmsTopView assyModel

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If Not fso2 Is Nothing Then
        If tempFolder <> "" Then
            If fso2.FolderExists(tempFolder) Then fso2.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateHolderDxfFromAssemblyBottomView error: " & Err.Description
    CreateHolderDxfFromAssemblyBottomView = False
    Resume CleanExit
End Function

Private Function CreateClampingPlateDxfFromAssemblyTopView(ByVal assyModel As Object, _
                                                          ByVal dxfPath As String, _
                                                          ByVal quoteName As String) As Boolean
On Error GoTo ErrHandler

    CreateClampingPlateDxfFromAssemblyTopView = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_CLAMP_DXF_" & Format(Now, "yyyymmdd_hhnnss") & "_" & NormalizeKey(quoteName)
    EnsureFolderDeep tempFolder

    ' Native assembly, not X_T.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_" & NormalizeKey(quoteName) & "_TOPVIEW_TEMP.sldasm"

    ' Force the actual SolidWorks standard TOP view before saving the temp native source.
    assyModel.ShowNamedView2 "*Top", 5
    StabilizeActiveView assyModel, 100

    LogLine quoteName & " clamping plate DXF: saving native temp SLDASM from SolidWorks *Top:"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(assyModel, tempNativePath) = False Then GoTo CleanExit

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "Clamping-plate assembly-top DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, quoteName, _
                                     "*Top", CMS_TOP_VIEW_NAME, _
                                     True, False, False, False, False

    CreateClampingPlateDxfFromAssemblyTopView = fso.FileExists(dxfPath)

CleanExit:
    On Error Resume Next

    ApplyCmsTopView assyModel

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If Not fso2 Is Nothing Then
        If tempFolder <> "" Then
            If fso2.FolderExists(tempFolder) Then fso2.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateClampingPlateDxfFromAssemblyTopView error: " & Err.Description
    CreateClampingPlateDxfFromAssemblyTopView = False
    Resume CleanExit
End Function

Private Sub CreateProjectedDxfFromXtPath(ByVal xtPath As String, _
                                         ByVal dxfPath As String, _
                                         ByVal quoteName As String, _
                                         ByVal parentPrimaryView As String, _
                                         ByVal parentFallbackView As String, _
                                         ByVal allWireframe As Boolean, _
                                         ByVal jBlockMixedDisplay As Boolean, _
                                         ByVal dimensionedJBlock As Boolean, _
                                         Optional ByVal allSolid As Boolean = False, _
                                         Optional ByVal pullcoreAux As Boolean = False)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(xtPath) = False Then
        LogLine "DXF skipped. XT source missing: " & xtPath
        Exit Sub
    End If

    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_DXF_FROM_XT_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CleanFileName(quoteName)

    EnsureFolderDeep tempFolder

    Dim nativePath As String
    nativePath = OpenXtAndSaveNativeForDrawing_Fixed(xtPath, tempFolder, NormalizeKey(quoteName))

    If nativePath = "" Then
        LogLine "DXF skipped. Could not open XT/native source for: " & xtPath
        GoTo CleanExit
    End If

    CreateProjectedDxfFromNativePath nativePath, dxfPath, quoteName, _
                                     parentPrimaryView, parentFallbackView, _
                                     allWireframe, jBlockMixedDisplay, dimensionedJBlock, allSolid, pullcoreAux

CleanExit:
    On Error Resume Next

    If tempFolder <> "" Then fso.DeleteFolder tempFolder, True
    Exit Sub

ErrHandler:
    LogLine "CreateProjectedDxfFromXtPath error for " & quoteName & ": " & Err.Description
    Resume CleanExit
End Sub

Private Function OpenXtAndSaveNativeForDrawing_Fixed(ByVal xtPath As String, _
                                                     ByVal tempFolder As String, _
                                                     ByVal baseName As String) As String
On Error GoTo ErrHandler

    OpenXtAndSaveNativeForDrawing_Fixed = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(xtPath) = False Then Exit Function

    EnsureFolderDeep tempFolder

    Dim importErrors As Long
    Dim errs As Long
    Dim warns As Long

    Dim mdl As Object
    Set mdl = swApp.LoadFile4(xtPath, "", Nothing, importErrors)

    If mdl Is Nothing Then
        Set mdl = swApp.OpenDoc6(xtPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then
        Set mdl = swApp.OpenDoc6(xtPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then
        LogLine "OpenXtAndSaveNativeForDrawing_Fixed failed to open: " & xtPath
        Exit Function
    End If

    Dim nativePath As String

    If mdl.GetType = swDocASSEMBLY Then
        nativePath = tempFolder & "\" & CleanFileName(baseName) & ".sldasm"
    Else
        nativePath = tempFolder & "\" & CleanFileName(baseName) & ".sldprt"
    End If

    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs
    EnsureSwHidden

    If CurrentDxfKeepImportedOrientation Then
        LogLine "Native drawing source kept in imported orientation (no re-orient)."
    Else
        SetCmsTopOrientation mdl
        ApplyCmsTopView mdl
        StabilizeActiveView mdl, 150
    End If

    mdl.Extension.SaveAs3 nativePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns

    If fso.FileExists(nativePath) Then
        OpenXtAndSaveNativeForDrawing_Fixed = nativePath
    Else
        LogLine "Native temp save failed: " & nativePath
    End If

    On Error Resume Next
    swApp.CloseDoc mdl.GetTitle
    Set mdl = Nothing
    Set fso = Nothing
    ReleaseSolidWorksMemory "after XT native conversion"
    Exit Function

ErrHandler:
    LogLine "OpenXtAndSaveNativeForDrawing_Fixed error: " & Err.Description
    OpenXtAndSaveNativeForDrawing_Fixed = ""
End Function

Private Sub CreateProjectedDxfFromNativePath(ByVal nativePath As String, _
                                             ByVal dxfPath As String, _
                                             ByVal quoteName As String, _
                                             ByVal parentPrimaryView As String, _
                                             ByVal parentFallbackView As String, _
                                             ByVal allWireframe As Boolean, _
                                             ByVal jBlockMixedDisplay As Boolean, _
                                             ByVal dimensionedJBlock As Boolean, _
                                             Optional ByVal allSolid As Boolean = False, _
                                             Optional ByVal pullcoreAux As Boolean = False)
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

    Dim partL As Double
    Dim partW As Double
    Dim partT As Double

    partL = 0#
    partW = 0#
    partT = 0#

    GetBestCurrentOutputDimensions quoteName, partL, partW, partT

    Dim liveL As Double
    Dim liveW As Double
    Dim liveT As Double

    If TryGetNativeModelDimsInches(nativePath, liveL, liveW, liveT) Then
        If liveL > 0# Then partL = liveL
        If liveW > 0# Then partW = liveW
        If liveT > 0# Then partT = liveT
    End If

    If IsHoldersDxfQuote(quoteName) Then
        If partL <= 0# Then partL = GetCombinedHoldersLengthForDxf()
        If partW <= 0# Then partW = GetCombinedHoldersWidthForDxf()
        If partT <= 0# Then partT = GetCombinedHoldersThicknessForDxf()

        If partL <= 0# Then partL = 42#
        If partW <= 0# Then partW = 28#
        If partT <= 0# Then partT = 10#
    End If

    If IsMainBaseDxfQuote(quoteName) Then
        If partL <= 0# Then partL = 42#
        If partW <= 0# Then partW = 30#
        If partT <= 0# Then partT = 10#
    End If

    If IsJBlockDxfQuote(quoteName) Then
        If partL <= 0# Then partL = JBlockTgtL
        If partW <= 0# Then partW = JBlockTgtW
        If partT <= 0# Then partT = JBlockTgtT
    End If

    If partL <= 0# Then partL = 10#
    If partW <= 0# Then partW = 10#
    If partT <= 0# Then partT = 2#

    ' UPDATED:
    ' Force 1:1 for ALL DXF views/files when enabled.
    ' No BASE/HOLDERS override.
    CurrentDxfForce1to1 = FORCE_ALL_DXF_VIEWS_1_TO_1

    Dim scaleVal As Double

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Then

        CurrentDxfForce1to1 = True
        scaleVal = 1#

        LogLine "DXF scale forced 1:1 for " & quoteName

    Else

        CurrentDxfForce1to1 = False

        scaleVal = CalculateProjectedFourViewDxfScale(partL, partW, partT)
        scaleVal = scaleVal * MULTIVIEW_FIT_SAFETY

        If scaleVal <= 0# Then scaleVal = 0.1

        LogLine "DXF auto-fit scale for " & quoteName & " = " & Format(scaleVal, "0.0000")

    End If

    If scaleVal = 1# Then
        Dim layoutWCheck As Double
        Dim layoutHCheck As Double

        layoutWCheck = partL + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
        layoutHCheck = partW + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)

        If layoutWCheck > (E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)) Or _
           layoutHCheck > (E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)) Then

            LogLine "WARNING: 1:1 DXF layout may not fit E-size sheet for " & quoteName & _
                    ". Required W/H=" & FormatNumberForCsv(layoutWCheck) & "/" & _
                    FormatNumberForCsv(layoutHCheck) & _
                    ", usable W/H=" & FormatNumberForCsv(E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)) & "/" & _
                    FormatNumberForCsv(E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN))
        End If
    End If

    If IsDxfInPrintsFolder(dxfPath) Then

        allWireframe = True
        allSolid = False
        jBlockMixedDisplay = False
        dimensionedJBlock = False

        If IsClampingPlateDxfQuote(quoteName) Then
            LogLine "TCP/BCP DXF detected. Starting hidden-lines-visible; left/bottom will be forced solid."
        Else
            LogLine "PRINTS folder DXF detected. Forcing HIDDEN LINES VISIBLE: " & dxfPath
        End If

    End If

    LogLine "Creating DXF: quote=" & quoteName & " native=" & nativePath
    LogLine "  L/W/T=" & FormatNumberForCsv(partL) & "/" & FormatNumberForCsv(partW) & "/" & FormatNumberForCsv(partT) & _
            "  scale=" & Format(scaleVal, "0.0000")

    CurrentDxfStraightenAngleRad = 0#

    If STRAIGHTEN_PULLCORE_DXF And IsPullcoreDxfQuote(quoteName) Then

        Dim pcDetectedDeg As Double
        Dim pcGotAngle As Boolean
        Dim pcSign As Double
        Dim pcAsmDeg As Double

        pcDetectedDeg = 0#
        pcGotAngle = False
        pcSign = GetPullcoreTiltSign(CurrentPullcoreStraightenCadIndex)

        ' Cross-check line. parts().asmPullAngleDeg is the tilt read straight off
        ' the component's assembly placement transform. It was being computed and
        ' stored but never read by anything; logging it lets the measured angle
        ' below be compared against it without changing behaviour.
        If CurrentPullcoreStraightenCadIndex > 0 And CurrentPullcoreStraightenCadIndex <= PartCount Then
            If parts(CurrentPullcoreStraightenCadIndex).hasAsmPullAngle Then
                pcAsmDeg = parts(CurrentPullcoreStraightenCadIndex).asmPullAngleSign * _
                           Abs(parts(CurrentPullcoreStraightenCadIndex).asmPullAngleDeg)

                LogLine "PULLCORE cross-check: assembly transform tilt = " & _
                        Format(pcAsmDeg, "0.00") & " deg"
            End If
        End If

        ' FIX - SOURCE ORDER.
        ' Measure on the ACTUAL drawing source, in the ACTUAL center-view plane,
        ' first. Previously the angle recorded during BOM matching won by
        ' short-circuit, and that value was measured in whatever view the
        ' component's own part file happened to be saved with - not the plane the
        ' drawing view is projected in. The DXF was therefore rotated about the
        ' wrong axis, which is why it moved but never squared up.
        ' PRIMARY - smallest bounding rectangle of the geometry that is actually
        ' being drawn, measured in the plane it is actually being drawn in.
        ' See the header on TryComputePullcoreStraightenAngleFromNativeGeometry
        ' for why the previous sources could only ever return 0 here.
        If PULLCORE_MEASURE_ANGLE_IN_VIEW_PLANE Then
            pcGotAngle = TryComputePullcoreStraightenAngleFromNativeGeometry( _
                             nativePath, parentPrimaryView, pcDetectedDeg)
        End If

        If pcGotAngle = False And PULLCORE_MEASURE_ANGLE_IN_VIEW_PLANE Then
            pcGotAngle = TryComputePullcoreBestFitStraightenAngleFromNativePath( _
                             nativePath, parentPrimaryView, pcDetectedDeg)

            If pcGotAngle Then
                LogLine "PULLCORE angle from longest-edge fallback in the " & _
                        parentPrimaryView & " plane: " & _
                        Format(pcDetectedDeg, "0.00") & " deg"
            End If
        End If

        If pcGotAngle = False And CurrentPullcoreStraightenCadIndex > 0 Then
            pcGotAngle = TryComputePullcoreStraightenAngleFromMatchedCad(CurrentPullcoreStraightenCadIndex, pcDetectedDeg)
        End If

        If pcGotAngle = False And PULLCORE_MEASURE_ANGLE_IN_VIEW_PLANE = False Then
            pcGotAngle = TryComputePullcoreBestFitStraightenAngleFromNativePath( _
                             nativePath, parentPrimaryView, pcDetectedDeg)
        End If

        If pcGotAngle = False And PULLCORE_FALLBACK_ANGLE_DEG <> 0# Then
            ' FIX - PULLCORE_FALLBACK_ANGLE_DEG is a MAGNITUDE. Applied unsigned
            ' it rotated every pullcore the same way, which is backwards for one
            ' half of a mirrored ID/OD pair. Carry the part's own lean direction.
            pcDetectedDeg = pcSign * Abs(PULLCORE_FALLBACK_ANGLE_DEG)
            pcGotAngle = True
            LogLine "PULLCORE: using fixed fallback angle " & _
                    Format(pcDetectedDeg, "0.00") & " deg (sign " & _
                    Format(pcSign, "0") & ") because geometry detection failed."
        End If

        If pcGotAngle Then
            If Abs(pcDetectedDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
                CurrentDxfStraightenAngleRad = PULLCORE_STRAIGHTEN_SIGN * pcDetectedDeg * PI_VALUE / 180#

                LogLine "PULLCORE geometry straighten angle detected: " & _
                        Format(pcDetectedDeg, "0.00") & " deg"

                LogLine "PULLCORE DXF correction angle: " & _
                        Format(RadToDeg(CurrentDxfStraightenAngleRad), "0.00") & " deg"
            Else
                CurrentDxfStraightenAngleRad = 0#
                LogLine "PULLCORE geometry angle " & Format(pcDetectedDeg, "0.00") & _
                        " deg is below threshold; DXF left unrotated."
            End If
        Else
            CurrentDxfStraightenAngleRad = 0#
            LogLine "PULLCORE: no straightening angle available; DXF left unrotated."
        End If

    End If

    Dim sheetWIn As Double
    Dim sheetHIn As Double

    sheetWIn = E_SHEET_WIDTH_IN
    sheetHIn = E_SHEET_HEIGHT_IN

    Dim baseExtraX As Double
    Dim baseExtraY As Double

    baseExtraX = 0#
    baseExtraY = 0#

    If IsMainBaseDxfQuote(quoteName) Then

        baseExtraX = BASE_DXF_EXTRA_SIDE_GAP_IN
        baseExtraY = BASE_DXF_EXTRA_TOP_BOTTOM_GAP_IN

        If BASE_DXF_AUTO_EXPAND_SHEET_TO_FIT_1_TO_1 Then

            Dim neededBaseW As Double
            Dim neededBaseH As Double

            neededBaseW = ((partL + (2# * partT) + _
                           (2# * (DXF_PROJECTED_VIEW_GAP_IN + baseExtraX))) * scaleVal) + _
                           (2# * DXF_MARGIN_IN)

            neededBaseH = ((partW + (2# * partT) + _
                           (2# * (DXF_PROJECTED_VIEW_GAP_IN + baseExtraY))) * scaleVal) + _
                           (2# * DXF_MARGIN_IN)

            If sheetWIn < neededBaseW Then sheetWIn = neededBaseW
            If sheetHIn < neededBaseH Then sheetHIn = neededBaseH

            If sheetWIn < BASE_DXF_MIN_SHEET_WIDTH_IN Then sheetWIn = BASE_DXF_MIN_SHEET_WIDTH_IN
            If sheetHIn < BASE_DXF_MIN_SHEET_HEIGHT_IN Then sheetHIn = BASE_DXF_MIN_SHEET_HEIGHT_IN

            LogLine "BASE DXF sheet expanded/spread to prevent overlap:"
            LogLine "  sheet W/H=" & FormatNumberForCsv(sheetWIn) & "/" & FormatNumberForCsv(sheetHIn)
            LogLine "  extra side/top gap=" & FormatNumberForCsv(baseExtraX) & "/" & FormatNumberForCsv(baseExtraY)

        End If

    End If

    Dim swDraw As Object
    Set swDraw = swApp.NewDocument(SW_DRAWING_TEMPLATE_PATH, 0, _
                                   sheetWIn / INCHES_PER_METER, _
                                   sheetHIn / INCHES_PER_METER)

    If swDraw Is Nothing Then
        LogLine "DXF skipped. Could not create drawing."
        GoTo CleanExit
    End If

    drawTitle = swDraw.GetTitle

    Dim errs As Long
    swApp.ActivateDoc3 drawTitle, False, 0, errs
    EnsureSwHidden

    SetupDrawingAsESize swDraw, sheetWIn, sheetHIn

    If FREEZE_DXF_DRAWING_GRAPHICS Then
        FreezeDxfDrawingGraphics swDraw
        freezeApplied = True
    End If

    Dim centerX As Double
    Dim centerY As Double

    centerX = sheetWIn / 2#
    centerY = sheetHIn / 2#

    Dim projectedXOffset As Double
    Dim projectedYOffset As Double

    projectedXOffset = ((partL / 2#) + DXF_PROJECTED_VIEW_GAP_IN + baseExtraX + (partT / 2#)) * scaleVal
    projectedYOffset = ((partW / 2#) + DXF_PROJECTED_VIEW_GAP_IN + baseExtraY + (partT / 2#)) * scaleVal

    If IsHoldersDxfQuote(quoteName) Then
        projectedXOffset = projectedXOffset + (HOLDERS_SIDE_VIEW_EXTRA_GAP_IN * scaleVal)
        LogLine "HOLDERS DXF: spreading left/right views by extra " & _
                FormatNumberForCsv(HOLDERS_SIDE_VIEW_EXTRA_GAP_IN) & " in."
    End If

    Dim xLeft As Double
    Dim yLeft As Double
    Dim xRight As Double
    Dim yRight As Double
    Dim xTop As Double
    Dim yTop As Double
    Dim xBottom As Double
    Dim yBottom As Double

    xLeft = centerX - projectedXOffset
    yLeft = centerY

    xRight = centerX + projectedXOffset
    yRight = centerY

    xTop = centerX
    yTop = centerY + projectedYOffset

    xBottom = centerX
    yBottom = centerY - projectedYOffset

    If xLeft < DXF_MARGIN_IN Then xLeft = DXF_MARGIN_IN
    If xRight > sheetWIn - DXF_MARGIN_IN Then xRight = sheetWIn - DXF_MARGIN_IN
    If yTop > sheetHIn - DXF_MARGIN_IN Then yTop = sheetHIn - DXF_MARGIN_IN
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

    Dim rotateJBlockParentBeforeProjection As Boolean

    ' FIX: the J BLOCK ejector cam DXF is exported under the bare quote name
    ' "CAM", so IsJBlockDxfQuote was False for it and the cam did NOT follow the
    ' block when ROTATE_J_BLOCK_DXF_180 was enabled. Include it explicitly.
    rotateJBlockParentBeforeProjection = (ROTATE_J_BLOCK_DXF_180 And _
                                          (IsJBlockDxfQuote(quoteName) Or _
                                           IsJBlockCamDxfQuote(quoteName)))

    Dim parentPreProjectionAngleRad As Double
    Dim qKey As String

    parentPreProjectionAngleRad = 0#
    qKey = NormalizeKey(quoteName)

    Select Case qKey

        Case "TCP", _
             "TOPSMED", _
             "TOPSMEDPLATE", _
             "TOPCLAMPING", _
             "TOPCLAMPINGPLATE"

            If Abs(TCP_DXF_CENTER_ROTATION_DEG) > 0.000001 Then
                parentPreProjectionAngleRad = parentPreProjectionAngleRad + DegToRad(TCP_DXF_CENTER_ROTATION_DEG)

                LogLine "TCP DXF: rotating center/base view " & _
                        Format(TCP_DXF_CENTER_ROTATION_DEG, "0.00") & _
                        " deg BEFORE projected views are created."
            End If

        Case "BCP", _
             "BOTTOMSMED", _
             "BOTSMED", _
             "BOTTOMSMEDPLATE", _
             "BOTSMEDPLATE", _
             "BOTTOMCLAMPING", _
             "BOTCLAMPING", _
             "BOTTOMCLAMPINGPLATE", _
             "BOTCLAMPINGPLATE"

            If Abs(BCP_DXF_CENTER_ROTATION_DEG) > 0.000001 Then
                parentPreProjectionAngleRad = parentPreProjectionAngleRad + DegToRad(BCP_DXF_CENTER_ROTATION_DEG)

                LogLine "BCP DXF: rotating center/base view " & _
                        Format(BCP_DXF_CENTER_ROTATION_DEG, "0.00") & _
                        " deg BEFORE projected views are created."
            End If

    End Select

    If rotateJBlockParentBeforeProjection Then
        parentPreProjectionAngleRad = parentPreProjectionAngleRad + PI_VALUE
        LogLine "J BLOCK: rotating center/base view 180 BEFORE projected views are created."
    End If

    If qKey = "IDHOLDER" Then
        Dim idHolderFlip As Boolean
        If CurrentIdHolderDxfFromAssembly Then
            idHolderFlip = FLIP_ID_HOLDER_CENTER_VIEW_180_FROM_ASSEMBLY
        Else
            idHolderFlip = FLIP_ID_HOLDER_CENTER_VIEW_180
        End If

        If idHolderFlip Then
            parentPreProjectionAngleRad = parentPreProjectionAngleRad + PI_VALUE
            LogLine "ID HOLDER: rotating center/base view 180 BEFORE projected views are created. (fromAssembly=" & CStr(CurrentIdHolderDxfFromAssembly) & ")"
        Else
            LogLine "ID HOLDER: center/base view used as-is BEFORE projected views are created. (fromAssembly=" & CStr(CurrentIdHolderDxfFromAssembly) & ")"
        End If
    End If

    If qKey = "ODHOLDER" Then
        If Abs(OD_HOLDER_CENTER_ROTATION_DEG) > 0.000001 Then
            parentPreProjectionAngleRad = parentPreProjectionAngleRad + DegToRad(OD_HOLDER_CENTER_ROTATION_DEG)

            LogLine "OD HOLDER: rotating center/base view " & _
                    Format(OD_HOLDER_CENTER_ROTATION_DEG, "0.00") & _
                    " deg BEFORE projected views are created."
        Else
            LogLine "OD HOLDER: center/base view used as-is BEFORE projected views are created."
        End If
    End If

    If Abs(CurrentDxfStraightenAngleRad) > 0.000001 Then
        parentPreProjectionAngleRad = parentPreProjectionAngleRad + CurrentDxfStraightenAngleRad

        LogLine "PULLCORE: applying center/base view straightening BEFORE projected views are created: " & _
                Format(RadToDeg(CurrentDxfStraightenAngleRad), "0.00") & " deg"
    ElseIf IsPullcoreDxfQuote(quoteName) Then
        LogLine "PULLCORE: straighten angle is ZERO for " & quoteName & _
                " - the center view will be left exactly as the model sits."
    End If

    LogLine "DXF total pre-projection center-view rotation for " & quoteName & " = " & _
            Format(RadToDeg(parentPreProjectionAngleRad), "0.00") & " deg"

    If Abs(parentPreProjectionAngleRad) > 0.000001 Then
        RotateDrawingViewByAngle parentView, parentPreProjectionAngleRad

        swDraw.ClearSelection2 True
        swDraw.ForceRebuild3 False
        swDraw.GraphicsRedraw2
    End If

    Set viewLeft = Nothing
    Set viewRight = Nothing
    Set viewTop = Nothing
    Set viewBottom = Nothing

    If pullcoreAux Then
        LogLine "PULLCORE: building 4 edge auxiliary views around the already-oriented front view."
        AddFourEdgeAuxViews swDraw, parentView, scaleVal
    Else
        Set viewLeft = CreateProjectedDrawingView(swDraw, parentView, xLeft, yLeft, scaleVal, "PROJECTED LEFT")
        Set viewRight = CreateProjectedDrawingView(swDraw, parentView, xRight, yRight, scaleVal, "PROJECTED RIGHT")
        Set viewTop = CreateProjectedDrawingView(swDraw, parentView, xTop, yTop, scaleVal, "PROJECTED TOP")
        Set viewBottom = CreateProjectedDrawingView(swDraw, parentView, xBottom, yBottom, scaleVal, "PROJECTED BOTTOM")
    End If

    If allSolid Then
        ForceAllDrawingViewsSolid swDraw
    ElseIf allWireframe Then
        ForceAllDrawingViewsWireframe swDraw
    ElseIf jBlockMixedDisplay Then
        SetDrawingViewSolid parentView
        SetDrawingViewSolid viewLeft
        SetDrawingViewSolid viewRight
        SetDrawingViewWireframe viewTop
        SetDrawingViewWireframe viewBottom
    Else
        ForceAllDrawingViewsWireframe swDraw
    End If

    If dimensionedJBlock Then
        AddBestEffortJBlockDimensions swDraw, parentView, viewTop, viewBottom, viewLeft, viewRight, partL, partW, partT, scaleVal
    End If

    If allSolid Then ForceAllDrawingViewsSolid swDraw
    If allWireframe Then ForceAllDrawingViewsWireframe swDraw

    If IsDxfInPrintsFolder(dxfPath) Then

        If IsClampingPlateDxfQuote(quoteName) Then
            ForceTcpBcpClampingPlateViewModes parentView, viewLeft, viewRight, viewTop, viewBottom
            LogLine "TCP/BCP DXF: center/right/top forced wireframe; left/bottom forced solid."
        Else
            ForceAllDrawingViewsWireframe swDraw
        End If

    End If

    If IsMainBaseDxfQuote(quoteName) Then
        ForceAllDrawingViewsSolid swDraw
    End If

    If IsHoldersDxfQuote(quoteName) Then
        ForceAllDrawingViewsSolid swDraw
        LogLine "HOLDERS DXF: all views forced SOLID."
    End If

    If IsJBlockDxfQuote(quoteName) Then
        ForceJBlockLeftRightViewsSolid viewLeft, viewRight
        LogLine "J BLOCK DXF: left/right projected views forced SOLID."
    End If

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    Dim saveErrs As Long
    Dim saveWarns As Long

    ForceAllDxfScales1To1 swDraw, sheetWIn, sheetHIn

    LogLine "Saving DXF: " & dxfPath
    EnsureSwHidden

    swDraw.Extension.SaveAs3 dxfPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, saveErrs, saveWarns

    LogLine "DXF save done. Errors=" & saveErrs & " Warnings=" & saveWarns

CleanExit:
    On Error Resume Next

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    If drawTitle <> "" Then swApp.CloseDoc drawTitle

    Set viewBottom = Nothing
    Set viewTop = Nothing
    Set viewRight = Nothing
    Set viewLeft = Nothing
    Set parentView = Nothing
    Set swDraw = Nothing
    Set fso = Nothing

    ReleaseSolidWorksMemory "after DXF save"
    Exit Sub

ErrHandler:
    LogLine "CreateProjectedDxfFromNativePath error for " & quoteName & ": " & Err.Description
    Resume CleanExit
End Sub

' ============================================================
' END OF PART 3
' Paste Part 4 immediately after this.
' ============================================================
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

Private Function CalculateProjectedFourViewDxfScale(ByVal partL As Double, _
                                                    ByVal partW As Double, _
                                                    ByVal partT As Double) As Double

    Dim scaleVal As Double
    scaleVal = DXF_MAX_SCALE

    Dim usableW As Double
    Dim usableH As Double

    usableW = E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)
    usableH = E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)

    Dim layoutW As Double
    Dim layoutH As Double

    layoutW = partL + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
    layoutH = partW + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)

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

Private Function CreateParentDrawingView(ByVal swDraw As Object, _
                                         ByVal nativePath As String, _
                                         ByVal primaryViewName As String, _
                                         ByVal fallbackViewName As String, _
                                         ByVal xIn As Double, _
                                         ByVal yIn As Double, _
                                         ByVal scaleVal As Double) As Object
On Error GoTo ErrHandler

    Set CreateParentDrawingView = Nothing

    If swDraw Is Nothing Then Exit Function

    Dim swView As Object
    Set swView = Nothing

    On Error Resume Next

    Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, primaryViewName, _
                                                      xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)

    If swView Is Nothing Then
        Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, fallbackViewName, _
                                                          xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    End If

    If swView Is Nothing Then
        Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, "*Top", _
                                                          xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    End If

    On Error GoTo ErrHandler

    If swView Is Nothing Then
        LogLine "Could not create parent drawing view."
        Exit Function
    End If

    SetDrawingViewScale swView, scaleVal

    Set CreateParentDrawingView = swView
    Exit Function

ErrHandler:
    LogLine "CreateParentDrawingView error: " & Err.Description
    Set CreateParentDrawingView = Nothing
End Function

Private Function CreateProjectedDrawingView(ByVal swDraw As Object, _
                                            ByVal parentView As Object, _
                                            ByVal xIn As Double, _
                                            ByVal yIn As Double, _
                                            ByVal scaleVal As Double, _
                                            ByVal labelText As String) As Object
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

Private Sub AddFourEdgeAuxViews(ByVal swDraw As Object, ByVal parentView As Object, ByVal scaleVal As Double)
On Error GoTo ErrHandler

    If swDraw Is Nothing Or parentView Is Nothing Then Exit Sub

    Dim vOut As Variant
    vOut = parentView.GetOutline

    If Not IsArray(vOut) Then Exit Sub
    If UBound(vOut) < 3 Then Exit Sub

    Dim xmin As Double
    Dim yMin As Double
    Dim xmax As Double
    Dim yMax As Double

    xmin = CDbl(vOut(0))
    yMin = CDbl(vOut(1))
    xmax = CDbl(vOut(2))
    yMax = CDbl(vOut(3))

    Dim mx As Double
    Dim my As Double
    Dim W As Double
    Dim h As Double

    mx = (xmin + xmax) / 2#
    my = (yMin + yMax) / 2#
    W = xmax - xmin
    h = yMax - yMin

    Dim dX As Double
    Dim dY As Double

    dX = W * 1.4 + (1.5 / INCHES_PER_METER)
    dY = h * 1.4 + (1.5 / INCHES_PER_METER)

    TryMakeEdgeAuxView swDraw, parentView, scaleVal, xmin, my, True, mx - dX, my + dY, "VIEW B"
    TryMakeEdgeAuxView swDraw, parentView, scaleVal, xmax, my, True, mx + dX, my + dY, "VIEW C"
    TryMakeEdgeAuxView swDraw, parentView, scaleVal, mx, yMin, False, mx - dX, my - dY, "VIEW A"
    TryMakeEdgeAuxView swDraw, parentView, scaleVal, mx, yMax, False, mx + dX, my - dY, "VIEW D"

    Exit Sub

ErrHandler:
    LogLine "AddFourEdgeAuxViews error: " & Err.Description
End Sub

Private Sub TryMakeEdgeAuxView(ByVal swDraw As Object, ByVal parentView As Object, ByVal scaleVal As Double, _
                               ByVal selX As Double, ByVal selY As Double, ByVal edgeVertical As Boolean, _
                               ByVal placeX As Double, ByVal placeY As Double, ByVal label As String)
On Error GoTo ErrHandler

    swDraw.ClearSelection2 True

    If TrySelectViewSilhouetteEdge(swDraw, selX, selY, edgeVertical, False) = False Then
        LogLine "Aux " & label & ": could not grab an edge to project from."
        swDraw.ClearSelection2 True
        Exit Sub
    End If

    Dim auxV As Object
    Set auxV = Nothing

    On Error Resume Next
    Set auxV = swDraw.CreateAuxiliaryViewAt2(placeX, placeY, 0#, label)
    On Error GoTo ErrHandler

    swDraw.ClearSelection2 True

    If auxV Is Nothing Then
        LogLine "Aux " & label & ": CreateAuxiliaryViewAt2 returned nothing."
    Else
        SetDrawingViewScale auxV, scaleVal
        LogLine "Aux " & label & " created."
    End If

    Exit Sub

ErrHandler:
    LogLine "TryMakeEdgeAuxView error (" & label & "): " & Err.Description
    On Error Resume Next
    swDraw.ClearSelection2 True
End Sub

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

Private Sub RotateDrawingViewByAngle(ByVal swView As Object, ByVal angleRad As Double)
On Error Resume Next

    If swView Is Nothing Then Exit Sub
    If Abs(angleRad) < 0.000001 Then Exit Sub

    ' This whole Sub used to run under On Error Resume Next with no read-back,
    ' so if IView::Angle refused the write it failed SILENTLY and the DXF just
    ' came out crooked with nothing in the log. Verify and report.
    Dim beforeRad As Double
    Dim afterRad As Double
    Dim wantRad As Double

    beforeRad = CDbl(swView.Angle)
    wantRad = beforeRad + angleRad

    Err.Clear
    swView.Angle = wantRad

    afterRad = CDbl(swView.Angle)

    LogLine "VIEW ROTATE: before=" & Format(RadToDeg(beforeRad), "0.00") & _
            " request=" & Format(RadToDeg(angleRad), "0.00") & _
            " want=" & Format(RadToDeg(wantRad), "0.00") & _
            " after=" & Format(RadToDeg(afterRad), "0.00") & " deg"

    If Abs(afterRad - wantRad) > 0.0017 Then
        LogLine "VIEW ROTATE FAILED: IView.Angle did not take. The drawing view" & _
                " will stay crooked. Geometry-level rotation is needed instead."
    End If
End Sub

Private Sub RotateDrawingView180(ByVal swView As Object)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    Dim a As Double
    a = CDbl(swView.Angle)

    swView.Angle = a + PI_VALUE
End Sub

Private Sub RotateJBlockProjectedViewSet180(ByVal swDraw As Object, _
                                            ByVal parentView As Object, _
                                            ByVal viewLeft As Object, _
                                            ByVal viewRight As Object, _
                                            ByVal viewTop As Object, _
                                            ByVal viewBottom As Object)
On Error Resume Next

    RotateDrawingView180 parentView
    RotateDrawingView180 viewLeft
    RotateDrawingView180 viewRight
    RotateDrawingView180 viewTop
    RotateDrawingView180 viewBottom

    SwapDrawingViewPositions viewLeft, viewRight
    SwapDrawingViewPositions viewTop, viewBottom

    If Not swDraw Is Nothing Then
        swDraw.ForceRebuild3 False
        swDraw.GraphicsRedraw2
    End If
End Sub

Private Sub SwapDrawingViewPositions(ByVal viewA As Object, ByVal viewB As Object)
On Error Resume Next

    If viewA Is Nothing Then Exit Sub
    If viewB Is Nothing Then Exit Sub

    Dim posA As Variant
    Dim posB As Variant

    posA = viewA.Position
    posB = viewB.Position

    If IsArray(posA) = False Then Exit Sub
    If IsArray(posB) = False Then Exit Sub
    If UBound(posA) < 1 Then Exit Sub
    If UBound(posB) < 1 Then Exit Sub

    Dim newA As Variant
    Dim newB As Variant

    newA = Array(CDbl(posB(0)), CDbl(posB(1)))
    newB = Array(CDbl(posA(0)), CDbl(posA(1)))

    viewA.Position = newA
    viewB.Position = newB
End Sub

' ============================================================
' ANGLE / DRAWING DISPLAY HELPERS
' ============================================================

Private Function Atan2Safe(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        Atan2Safe = Atn(y / x)
    ElseIf x < 0# And y >= 0# Then
        Atan2Safe = Atn(y / x) + PI_VALUE
    ElseIf x < 0# And y < 0# Then
        Atan2Safe = Atn(y / x) - PI_VALUE
    ElseIf x = 0# And y > 0# Then
        Atan2Safe = PI_VALUE / 2#
    ElseIf x = 0# And y < 0# Then
        Atan2Safe = -PI_VALUE / 2#
    Else
        Atan2Safe = 0#
    End If
End Function

Private Function RadToDeg(ByVal rad As Double) As Double
    RadToDeg = rad * 180# / PI_VALUE
End Function

Private Function DegToRad(ByVal deg As Double) As Double
    DegToRad = deg * PI_VALUE / 180#
End Function

Private Function NormalizeAngleToPlusMinus90(ByVal angleRad As Double) As Double
    Do While angleRad > PI_VALUE / 2#
        angleRad = angleRad - PI_VALUE
    Loop

    Do While angleRad < -PI_VALUE / 2#
        angleRad = angleRad + PI_VALUE
    Loop

    NormalizeAngleToPlusMinus90 = angleRad
End Function

' Pull-core tilt of a component in the ASSEMBLY, in degrees (0..45), taken
' from its placement transform. Transform2.ArrayData holds a 3x3 rotation in
' the first 9 entries (column-major: cols are the part's local X,Y,Z axes
' expressed in assembly coordinates). The tilt is the largest angle any of
' those three axes makes with its nearest assembly axis, folded to 0..45.
' This is distinct per component even when imported part files store their
' bodies at a shared orientation, so it is the correct pull-angle source.
Private Function TryGetComponentAssemblyPullAngleDeg(ByVal swComp As Object, _
                                                     ByRef angleDeg As Double, _
                                                     Optional ByRef outSign As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentAssemblyPullAngleDeg = False
    angleDeg = 0#
    outSign = 1#

    If swComp Is Nothing Then Exit Function

    Dim xform As Object
    Set xform = swComp.Transform2
    If xform Is Nothing Then Exit Function

    Dim v As Variant
    v = xform.ArrayData
    If IsEmpty(v) Then Exit Function
    If IsArray(v) = False Then Exit Function
    If UBound(v) < 8 Then Exit Function

    ' Local X axis = v(0..2), local Y = v(3..5), local Z = v(6..8).
    ' The pull-core de-rotation happens in the DXF (X/Y) plane, so the angle
    ' that matters is the in-plane rotation of the part's local X axis about
    ' assembly Z: signed Atan2(localX.y, localX.x). This carries DIRECTION
    ' (clockwise vs counter-clockwise) so the DXF is straightened the correct
    ' way for both ID and OD parts, which are mirrored about the centerline.
    Dim lxX As Double, lxY As Double
    lxX = CDbl(v(0))
    lxY = CDbl(v(1))

    Dim signedRad As Double
    signedRad = Atan2Safe(lxY, lxX)

    ' Fold into -45..+45 about the box's 90-degree symmetry, preserving sign.
    Dim signedDeg As Double
    signedDeg = signedRad * 180# / PI_VALUE

    Do While signedDeg > 45#
        signedDeg = signedDeg - 90#
    Loop
    Do While signedDeg < -45#
        signedDeg = signedDeg + 90#
    Loop

    angleDeg = Abs(signedDeg)

    If signedDeg >= 0# Then
        outSign = 1#
    Else
        outSign = -1#
    End If

    TryGetComponentAssemblyPullAngleDeg = True

    Exit Function

ErrHandler:
    TryGetComponentAssemblyPullAngleDeg = False
End Function

Private Function TryGetComponentTopRotationRad(ByVal swComp As Object, ByRef angleRad As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentTopRotationRad = False
    angleRad = 0#

    If swComp Is Nothing Then Exit Function

    Dim xform As Object
    Set xform = swComp.Transform2

    If xform Is Nothing Then Exit Function

    Dim v As Variant
    v = xform.ArrayData

    If IsEmpty(v) Then Exit Function
    If IsArray(v) = False Then Exit Function
    If UBound(v) < 8 Then Exit Function

    Dim xAxisX As Double
    Dim xAxisY As Double

    xAxisX = CDbl(v(0))
    xAxisY = CDbl(v(1))

    angleRad = Atan2Safe(xAxisY, xAxisX)
    angleRad = NormalizeAngleToPlusMinus90(angleRad)

    TryGetComponentTopRotationRad = True
    Exit Function

ErrHandler:
    TryGetComponentTopRotationRad = False
End Function

Private Sub RotateAllModelDrawingViews(ByVal swDraw As Object, ByVal angleRad As Double)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub
    If Abs(angleRad) < 0.000001 Then Exit Sub

    Dim v As Object
    Set v = swDraw.GetFirstView

    If Not v Is Nothing Then Set v = v.GetNextView

    Do While Not v Is Nothing
        v.Angle = CDbl(v.Angle) + angleRad
        Set v = v.GetNextView
    Loop

    swDraw.ForceRebuild3 False
    swDraw.GraphicsRedraw2
End Sub

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

Private Sub ForceTcpBcpClampingPlateViewModes(ByVal parentView As Object, _
                                              ByVal viewLeft As Object, _
                                              ByVal viewRight As Object, _
                                              ByVal viewTop As Object, _
                                              ByVal viewBottom As Object)
On Error Resume Next

    SetDrawingViewWireframe parentView
    SetDrawingViewWireframe viewRight
    SetDrawingViewWireframe viewTop

    SetDrawingViewSolid viewLeft
    SetDrawingViewSolid viewBottom
End Sub

Private Sub ForceJBlockLeftRightViewsSolid(ByVal viewLeft As Object, _
                                           ByVal viewRight As Object)
On Error Resume Next

    SetDrawingViewSolid viewLeft
    SetDrawingViewSolid viewRight
End Sub

Private Sub SetupDrawingAsESize(ByVal swDraw As Object, _
                                Optional ByVal sheetWIn As Double = 0#, _
                                Optional ByVal sheetHIn As Double = 0#)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    If sheetWIn <= 0# Then sheetWIn = E_SHEET_WIDTH_IN
    If sheetHIn <= 0# Then sheetHIn = E_SHEET_HEIGHT_IN

    Dim swSheet As Object
    Set swSheet = swDraw.GetCurrentSheet

    If Not swSheet Is Nothing Then
        swSheet.SetSize 12, sheetWIn / INCHES_PER_METER, sheetHIn / INCHES_PER_METER
    End If

    ForceDrawingSheetScale1To1 swDraw, sheetWIn, sheetHIn

    swDraw.GraphicsRedraw2
End Sub

Private Sub ForceDrawingSheetScale1To1(ByVal swDraw As Object, _
                                       Optional ByVal sheetWIn As Double = 0#, _
                                       Optional ByVal sheetHIn As Double = 0#)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    If sheetWIn <= 0# Then sheetWIn = E_SHEET_WIDTH_IN
    If sheetHIn <= 0# Then sheetHIn = E_SHEET_HEIGHT_IN

    Dim swSheet As Object
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then Exit Sub

    ' SolidWorks templates and view insertion can reset the sheet to 1:2.
    ' Force the sheet itself to 1:1 using multiple late-bound API paths.
    Err.Clear
    swSheet.SetScale 1#, 1#, False, False
    Err.Clear
    swSheet.SetScale 1#, 1#, True, True
    Err.Clear

    Dim sheetName As String
    sheetName = ""
    sheetName = CStr(swSheet.GetName)

    If sheetName <> "" Then

        Err.Clear
        swDraw.SetupSheet5 sheetName, 12, 12, 1#, 1#, False, "", _
                           sheetWIn / INCHES_PER_METER, _
                           sheetHIn / INCHES_PER_METER, "", False
        Err.Clear

        Err.Clear
        swDraw.SetupSheet4 sheetName, 12, 12, 1#, 1#, False, "", _
                           sheetWIn / INCHES_PER_METER, _
                           sheetHIn / INCHES_PER_METER, ""
        Err.Clear

    End If
End Sub

Private Sub ForceAllDxfScales1To1(ByVal swDraw As Object, _
                                  Optional ByVal sheetWIn As Double = 0#, _
                                  Optional ByVal sheetHIn As Double = 0#)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    ForceDrawingSheetScale1To1 swDraw, sheetWIn, sheetHIn

    Dim v As Object
    Set v = swDraw.GetFirstView

    If Not v Is Nothing Then Set v = v.GetNextView

    Do While Not v Is Nothing
        SetDrawingViewScale v, 1#
        Set v = v.GetNextView
    Loop

    ForceDrawingSheetScale1To1 swDraw, sheetWIn, sheetHIn

    swDraw.ForceRebuild3 False
    swDraw.GraphicsRedraw2
End Sub

Private Sub SetDrawingViewWireframe(ByVal swView As Object)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    swView.UseParentStyle = False
    swView.SetDisplayMode3 False, 1, False, True
    swView.DisplayMode = 1
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

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Or CurrentDxfForce1to1 Then
        scaleVal = 1#
    End If

    If scaleVal <= 0# Then scaleVal = 1#

    swView.UseSheetScale = False
    swView.ScaleDecimal = scaleVal

    If Abs(scaleVal - 1#) < 0.000001 Then
        Dim scaleRatio(0 To 1) As Double
        scaleRatio(0) = 1#
        scaleRatio(1) = 1#

        Err.Clear
        swView.scaleRatio = scaleRatio
        Err.Clear
        swView.scaleRatio = "1:1"
        Err.Clear
    End If

    ' Set this again last. Some SolidWorks view operations flip back to sheet scale.
    swView.UseSheetScale = False
    swView.ScaleDecimal = scaleVal
End Sub

' ============================================================
' J BLOCK OVERALL DIMENSIONS
' ============================================================

Private Sub AddBestEffortJBlockDimensions(ByVal swDraw As Object, _
                                          ByVal parentView As Object, _
                                          ByVal viewTop As Object, _
                                          ByVal viewBottom As Object, _
                                          ByVal viewLeft As Object, _
                                          ByVal viewRight As Object, _
                                          ByVal partL As Double, _
                                          ByVal partW As Double, _
                                          ByVal partT As Double, _
                                          ByVal scaleVal As Double)
On Error GoTo ErrHandler

    If DIMENSION_J_BLOCK = False Then
        LogLine "J BLOCK dimensioning disabled by setting."
        Exit Sub
    End If

    If swDraw Is Nothing Then Exit Sub

    LogLine "Adding overall J BLOCK dimensions."

    AddViewOverallDimension swDraw, parentView, False
    AddViewOverallDimension swDraw, parentView, True

    If Not viewRight Is Nothing Then
        AddViewOverallDimension swDraw, viewRight, True
    ElseIf Not viewLeft Is Nothing Then
        AddViewOverallDimension swDraw, viewLeft, True
    ElseIf Not viewTop Is Nothing Then
        AddViewOverallDimension swDraw, viewTop, False
    End If

    AddJBlockDimensionNotes swDraw, parentView, partL, partW, partT

    swDraw.GraphicsRedraw2
    Exit Sub

ErrHandler:
    LogLine "AddBestEffortJBlockDimensions error: " & Err.Description
End Sub

Private Sub AddViewOverallDimension(ByVal swDraw As Object, _
                                    ByVal swView As Object, _
                                    ByVal horizontal As Boolean)
On Error GoTo ErrHandler

    If swDraw Is Nothing Then Exit Sub
    If swView Is Nothing Then Exit Sub

    Dim vOut As Variant
    vOut = swView.GetOutline

    If IsEmpty(vOut) Then Exit Sub
    If IsArray(vOut) = False Then Exit Sub
    If UBound(vOut) < 3 Then Exit Sub

    Dim xmin As Double
    Dim yMin As Double
    Dim xmax As Double
    Dim yMax As Double
    Dim midx As Double
    Dim midy As Double
    Dim gap As Double

    xmin = CDbl(vOut(0))
    yMin = CDbl(vOut(1))
    xmax = CDbl(vOut(2))
    yMax = CDbl(vOut(3))

    midx = (xmin + xmax) / 2#
    midy = (yMin + yMax) / 2#

    gap = 0.625 / INCHES_PER_METER

    swDraw.ClearSelection2 True

    Dim ok1 As Boolean
    Dim ok2 As Boolean
    Dim dispDim As Object

    If horizontal Then

        ok1 = TrySelectViewSilhouetteEdge(swDraw, xmin, midy, True, False)
        ok2 = TrySelectViewSilhouetteEdge(swDraw, xmax, midy, True, True)

        If ok1 And ok2 Then
            Set dispDim = swDraw.AddHorizontalDimension2(midx, yMin - gap, 0#)
        End If

    Else

        ok1 = TrySelectViewSilhouetteEdge(swDraw, midx, yMin, False, False)
        ok2 = TrySelectViewSilhouetteEdge(swDraw, midx, yMax, False, True)

        If ok1 And ok2 Then
            Set dispDim = swDraw.AddVerticalDimension2(xmax + gap, midy, 0#)
        End If

    End If

    swDraw.ClearSelection2 True
    Exit Sub

ErrHandler:
    On Error Resume Next
    swDraw.ClearSelection2 True
    LogLine "AddViewOverallDimension error: " & Err.Description
End Sub

Private Function TrySelectViewSilhouetteEdge(ByVal swDraw As Object, _
                                             ByVal x As Double, _
                                             ByVal y As Double, _
                                             ByVal edgeIsVertical As Boolean, _
                                             ByVal appendToSelection As Boolean) As Boolean
On Error GoTo ErrHandler

    TrySelectViewSilhouetteEdge = False

    Dim nudges(0 To 4) As Double

    nudges(0) = 0#
    nudges(1) = 0.015 / INCHES_PER_METER
    nudges(2) = -0.015 / INCHES_PER_METER
    nudges(3) = 0.04 / INCHES_PER_METER
    nudges(4) = -0.04 / INCHES_PER_METER

    Dim i As Long
    Dim tx As Double
    Dim ty As Double
    Dim okSel As Boolean

    For i = 0 To 4

        tx = x
        ty = y

        If edgeIsVertical Then
            tx = x + nudges(i)
        Else
            ty = y + nudges(i)
        End If

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

Private Sub TryInsertModelAnnotations(ByVal swView As Object)
    ' Reserved. Imported solids usually carry no model dimensions.
End Sub

Private Sub AddJBlockDimensionNotes(ByVal swDraw As Object, ByVal swView As Object, _
                                    ByVal partL As Double, ByVal partW As Double, ByVal partT As Double)
On Error GoTo ErrHandler

    If swDraw Is Nothing Then Exit Sub

    Dim xmin As Double
    Dim yMin As Double
    Dim xmax As Double
    Dim yMax As Double

    xmin = E_SHEET_WIDTH_IN / 2# / INCHES_PER_METER
    yMin = E_SHEET_HEIGHT_IN / 2# / INCHES_PER_METER
    xmax = xmin
    yMax = yMin

    If Not swView Is Nothing Then

        Dim vOut As Variant
        vOut = swView.GetOutline

        If IsArray(vOut) Then
            If UBound(vOut) >= 3 Then
                xmin = CDbl(vOut(0))
                yMin = CDbl(vOut(1))
                xmax = CDbl(vOut(2))
                yMax = CDbl(vOut(3))
            End If
        End If

    End If

    Dim midx As Double
    Dim midy As Double
    Dim gap As Double

    midx = (xmin + xmax) / 2#
    midy = (yMin + yMax) / 2#
    gap = 0.6 / INCHES_PER_METER

    PlaceDrawingNote swDraw, midx - gap, yMin - gap, "OVERALL LENGTH (L) = " & FormatJBlockDim(partL)
    PlaceDrawingNote swDraw, xmin - (2.4 / INCHES_PER_METER), midy, "OVERALL HEIGHT (W) = " & FormatJBlockDim(partW)
    PlaceDrawingNote swDraw, xmax + gap, yMax + gap, "THICKNESS (T) = " & FormatJBlockDim(partT)

    swDraw.ClearSelection2 True
    Exit Sub

ErrHandler:
    LogLine "AddJBlockDimensionNotes error: " & Err.Description
End Sub

Private Sub PlaceDrawingNote(ByVal swDraw As Object, ByVal xMeters As Double, _
                             ByVal yMeters As Double, ByVal text As String)
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

Private Function FormatJBlockDim(ByVal v As Double) As String
    FormatJBlockDim = Format(v, "0.000")
End Function

' ============================================================
' DXF DIMENSION / QUOTE HELPERS
' ============================================================

Private Sub GetBestCurrentOutputDimensions(ByVal quoteName As String, _
                                           ByRef L As Double, _
                                           ByRef W As Double, _
                                           ByRef T As Double)
On Error Resume Next

    Dim k As String
    k = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            If ExportRows(i).HasCad Then
                L = parts(ExportRows(i).CadPartIndex).Length
                W = parts(ExportRows(i).CadPartIndex).Width
                T = parts(ExportRows(i).CadPartIndex).Thickness
                Exit Sub
            End If
        End If
    Next i

    For i = 1 To PullcoreMatchCount
        If NormalizeKey(PullcoreMatches(i).quoteName) = k Then
            If PullcoreMatches(i).CadPartIndex > 0 Then
                L = PullcoreMatches(i).FittedLength
                W = PullcoreMatches(i).FittedWidth
                T = PullcoreMatches(i).FittedThickness

                If L <= 0 Then L = PullcoreMatches(i).CadLength
                If W <= 0 Then W = PullcoreMatches(i).CadWidth
                If T <= 0 Then T = PullcoreMatches(i).CadThickness

                If L <= 0 Then L = parts(PullcoreMatches(i).CadPartIndex).Length
                If W <= 0 Then W = parts(PullcoreMatches(i).CadPartIndex).Width
                If T <= 0 Then T = parts(PullcoreMatches(i).CadPartIndex).Thickness

                Exit Sub
            End If
        End If
    Next i
End Sub

Private Function GetCombinedHoldersLengthForDxf() As Double
    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    Dim best As Double
    best = 0#

    If idIdx > 0 Then If parts(idIdx).Length > best Then best = parts(idIdx).Length
    If odIdx > 0 Then If parts(odIdx).Length > best Then best = parts(odIdx).Length

    GetCombinedHoldersLengthForDxf = best
End Function

Private Function GetCombinedHoldersWidthForDxf() As Double
    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    Dim total As Double
    total = 0#

    If idIdx > 0 Then total = total + parts(idIdx).Width
    If odIdx > 0 Then total = total + parts(odIdx).Width

    GetCombinedHoldersWidthForDxf = total
End Function

Private Function GetCombinedHoldersThicknessForDxf() As Double
    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    Dim best As Double
    best = 0#

    If idIdx > 0 Then If parts(idIdx).Thickness > best Then best = parts(idIdx).Thickness
    If odIdx > 0 Then If parts(odIdx).Thickness > best Then best = parts(odIdx).Thickness

    GetCombinedHoldersThicknessForDxf = best
End Function

Private Function IsDxfInPrintsFolder(ByVal dxfPath As String) As Boolean
    Dim p As String
    p = UCase(dxfPath)

    If InStr(p, "\" & UCase(CurrentJobNumber & OUTPUT_FOLDER_SUFFIX) & "\") > 0 Then
        If InStr(p, "PYROPEL") > 0 Then
            IsDxfInPrintsFolder = False
            Exit Function
        End If

        IsDxfInPrintsFolder = True
        Exit Function
    End If

    If InStr(p, "\PRINTS\") > 0 Then IsDxfInPrintsFolder = True
End Function

Private Function IsHoldersDxfQuote(ByVal quoteName As String) As Boolean
    IsHoldersDxfQuote = (UCase(Trim(quoteName)) = "HOLDERS")
End Function

Private Function IsMainBaseDxfQuote(ByVal quoteName As String) As Boolean
    IsMainBaseDxfQuote = (UCase(Trim(quoteName)) = "MAIN ASSEMBLY")
End Function

Private Function IsJBlockDxfQuote(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    IsJBlockDxfQuote = (InStr(s, "J BLOCK") > 0 Or InStr(s, "J-BLOCK") > 0 Or InStr(s, "JBLOCK") > 0 _
                        Or InStr(s, "U BLOCK") > 0 Or InStr(s, "U-BLOCK") > 0 Or InStr(s, "UBLOCK") > 0)
End Function

' The J BLOCK ejector cam is exported with the bare quote name "CAM"
' (J_BLOCK_CAM_FILE_TOKEN). It has to follow the J BLOCK's view orientation, but
' it must NOT inherit the J BLOCK's dimension fallbacks, so it is deliberately
' kept separate from IsJBlockDxfQuote.
Private Function IsJBlockCamDxfQuote(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    IsJBlockCamDxfQuote = (s = "CAM" Or s = "EJECTOR CAM" Or s = "EJ CAM")
End Function

Private Function IsClampingPlateDxfQuote(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "TCP", "BCP", _
             "TOPCLAMPINGPLATE", "BOTTOMCLAMPINGPLATE", _
             "TOPSMEDPLATE", "BOTTOMSMEDPLATE", _
             "TOPSMED", "BOTTOMSMED"
            IsClampingPlateDxfQuote = True
    End Select
End Function

Private Function IsPullcoreDxfQuote(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    If s = "PULLCORE" Then
        IsPullcoreDxfQuote = True
        Exit Function
    End If

    ' FIX: the PULLCORE STOP is a flat plate, not a cam or key. It does not pull
    ' at an angle so it must not be de-rotated. Before this it matched on the
    ' word "PULLCORE", found no cam angle to use, and fell through to the blanket
    ' PULLCORE_FALLBACK_ANGLE_DEG rotation.
    If InStr(s, "PULLCORE STOP") > 0 Or InStr(s, "PULL CORE STOP") > 0 Then
        IsPullcoreDxfQuote = False
        Exit Function
    End If

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then
        IsPullcoreDxfQuote = True
    End If
End Function

Private Function TryGetNativeModelDimsInches(ByVal nativePath As String, _
                                             ByRef L As Double, _
                                             ByRef W As Double, _
                                             ByRef T As Double) As Boolean
On Error GoTo ErrHandler

    TryGetNativeModelDimsInches = False

    L = 0#
    W = 0#
    T = 0#

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

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    Dim gotBox As Boolean

    gotBox = TryGetModelDocBoxDimsInches(mdl, dX, dY, dZ)

    On Error Resume Next
    swApp.CloseDoc mdl.GetTitle
    On Error GoTo ErrHandler

    If gotBox = False Then Exit Function

    SortThreeDimensions dX, dY, dZ, L, W, T

    L = Round(L, DIM_DECIMALS)
    W = Round(W, DIM_DECIMALS)
    T = Round(T, DIM_DECIMALS)

    TryGetNativeModelDimsInches = (L > 0 And W > 0 And T > 0)
    Exit Function

ErrHandler:
    TryGetNativeModelDimsInches = False
End Function

Private Function TryGetModelDocBoxDimsInches(ByVal mdl As Object, _
                                             ByRef dX As Double, _
                                             ByRef dY As Double, _
                                             ByRef dZ As Double) As Boolean
On Error GoTo ErrHandler

    TryGetModelDocBoxDimsInches = False

    If mdl Is Nothing Then Exit Function

    If mdl.GetType = swDocPART Then
        TryGetModelDocBoxDimsInches = GetPartBoundingBoxInches(mdl, dX, dY, dZ)
        Exit Function
    End If

    Dim vBox As Variant

    On Error Resume Next
    vBox = mdl.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsValidBoxArray(vBox) = False Then Exit Function

    dX = Abs(CDbl(vBox(3)) - CDbl(vBox(0))) * INCHES_PER_METER
    dY = Abs(CDbl(vBox(4)) - CDbl(vBox(1))) * INCHES_PER_METER
    dZ = Abs(CDbl(vBox(5)) - CDbl(vBox(2))) * INCHES_PER_METER

    TryGetModelDocBoxDimsInches = True
    Exit Function

ErrHandler:
    TryGetModelDocBoxDimsInches = False
End Function

Private Function IsValidBoxArray(ByVal vBox As Variant) As Boolean
    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    IsValidBoxArray = True
End Function

' ============================================================
' END OF PART 4
' Paste Part 5 immediately after this.
' ============================================================
' ============================================================
' BOM FIND / READ (EXCEL)
' ============================================================

Private Function FindCustomerBomFile(ByVal jobFolder As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(jobFolder) Then Exit Function

    Dim bestPath As String
    Dim bestScore As Long

    bestPath = ""
    bestScore = -1

    SearchBomFilesRecursive fso.GetFolder(jobFolder), bestPath, bestScore

    If bestPath <> "" Then
        LogLine "BOM file chosen (score " & bestScore & "): " & bestPath
    Else
        LogLine "No BOM/Excel/PDF file found anywhere under: " & jobFolder
        LogLine "  Files present at the job folder root:"

        Dim ff As Object
        For Each ff In fso.GetFolder(jobFolder).Files
            LogLine "    - " & ff.name
        Next ff
    End If

    FindCustomerBomFile = bestPath
    Exit Function

ErrHandler:
    LogLine "FindCustomerBomFile error: " & Err.Description
    FindCustomerBomFile = ""
End Function

Private Sub SearchBomFilesRecursive(ByVal folder As Object, ByRef bestPath As String, ByRef bestScore As Long)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderName As String
    folderName = UCase(folder.name)

    If folderName = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub

    Dim file As Object
    Dim ext As String
    Dim nameUpper As String
    Dim score As Long

    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.path))
        nameUpper = UCase(file.name)

        If ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "pdf" Then

            score = 0

            If InStr(nameUpper, "BOM") > 0 Then score = score + 100
            If InStr(nameUpper, "BILL OF MATERIAL") > 0 Then score = score + 100
            If InStr(nameUpper, "QUOTE") > 0 Then score = score + 40
            If InStr(nameUpper, "MATERIAL") > 0 Then score = score + 30
            If InStr(nameUpper, CurrentJobNumber) > 0 Then score = score + 20

            If ext = "xlsx" Then score = score + 8
            If ext = "xlsm" Then score = score + 7
            If ext = "xls" Then score = score + 6
            If ext = "pdf" Then score = score + 5

            If score > bestScore Then
                bestScore = score
                bestPath = file.path
            End If

        End If

    Next file

    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        SearchBomFilesRecursive subFolder, bestPath, bestScore
    Next subFolder
End Sub

Private Sub ReadCustomerBom(ByVal bomPath As String)
On Error GoTo ErrHandler

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False

    Set xlWb = xlApp.Workbooks.Open(bomPath, False, True)

    On Error Resume Next
    xlApp.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    Dim usedThisSheet As Boolean
    usedThisSheet = False

    If TURBO_READ_ONLY_BOM_SHEET Then
        On Error Resume Next
        Set xlWs = xlWb.Worksheets(TURBO_BOM_SHEET_NAME)
        On Error GoTo ErrHandler

        If Not xlWs Is Nothing Then
            LogLine "TURBO: reading only sheet '" & TURBO_BOM_SHEET_NAME & "'"
            ReadBomWorksheetFastArray xlWs
            usedThisSheet = True
        End If
    End If

    If usedThisSheet = False Then

        Dim ws As Object

        For Each ws In xlWb.Worksheets
            If IsLikelyBomWorksheet(ws) And ShouldSkipBomWorksheet(ws) = False Then
                LogLine "Reading worksheet: " & ws.name
                ReadBomWorksheetFastArray ws
                If BomCount > 0 Then Exit For
            End If
        Next ws

    End If

CleanExit:
    On Error Resume Next

    If Not xlWb Is Nothing Then xlWb.Close False

    If Not xlApp Is Nothing Then
        xlApp.Calculation = -4105
        xlApp.EnableEvents = True
        xlApp.ScreenUpdating = True
        xlApp.Quit
    End If

    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrHandler:
    LogLine "ReadCustomerBom error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub ReadBomWorksheetFastArray(ByVal ws As Object)
On Error GoTo ErrHandler

    Dim usedRange As Object
    Set usedRange = ws.usedRange

    Dim rowCount As Long
    Dim colCount As Long

    rowCount = usedRange.rows.count
    colCount = usedRange.Columns.count

    If rowCount = 0 Or colCount = 0 Then Exit Sub

    Dim data As Variant
    data = usedRange.value

    If IsEmpty(data) Then Exit Sub
    If IsArray(data) = False Then Exit Sub

    Dim headerRow As Long
    Dim descCol As Long
    Dim qtyCol As Long
    Dim matCol As Long
    Dim lenCol As Long
    Dim widCol As Long
    Dim thkCol As Long
    Dim typeCol As Long

    headerRow = 0
    descCol = 0
    qtyCol = 0
    matCol = 0
    lenCol = 0
    widCol = 0
    thkCol = 0
    typeCol = 0

    Dim r As Long
    Dim maxHeaderRow As Long

    maxHeaderRow = rowCount
    If maxHeaderRow > BOM_HEADER_SEARCH_MAX_ROWS Then maxHeaderRow = BOM_HEADER_SEARCH_MAX_ROWS

    For r = 1 To maxHeaderRow
        descCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("DESCRIPTION", "DESC", "PART NAME", "ITEM DESCRIPTION", "COMPONENT"))
        If descCol > 0 Then
            headerRow = r
            Exit For
        End If
    Next r

    If headerRow = 0 Then
        LogLine "No header row found by description."
        Exit Sub
    End If

    qtyCol = FindBomQtyColumnInArrayRow(data, headerRow, colCount)
    matCol = FindBomMaterialColumnInArrayRow(data, headerRow, colCount)
    typeCol = FindBomHeaderLikeInArrayRow(data, headerRow, colCount, Array("TYPE", "SOURCE", "MAKE/BUY", "MAKE BUY", "PROCESS"))

    FindBomDimensionColumnsInArrayRow data, headerRow, colCount, lenCol, widCol, thkCol

    LogLine "BOM header row=" & headerRow & " descCol=" & descCol & " qtyCol=" & qtyCol & _
            " matCol=" & matCol & " typeCol=" & typeCol & _
            " L=" & lenCol & " W=" & widCol & " T=" & thkCol

    Dim blanks As Long
    blanks = 0

    For r = headerRow + 1 To rowCount

        Dim desc As String
        Dim typeText As String
        Dim matText As String
        Dim qtyText As String

        desc = GetArrayValue(data, r, descCol)
        typeText = GetArrayValue(data, r, typeCol)
        matText = GetArrayValue(data, r, matCol)
        qtyText = GetArrayValue(data, r, qtyCol)

        If Trim(desc) = "" Then

            blanks = blanks + 1
            If blanks >= STOP_BOM_READ_AFTER_BLANK_ROWS Then Exit For

        Else

            blanks = 0

            Dim lenVal As Double
            Dim widVal As Double
            Dim thkVal As Double
            Dim hasDims As Boolean

            lenVal = 0#
            widVal = 0#
            thkVal = 0#

            If lenCol > 0 Then lenVal = val(GetArrayValue(data, r, lenCol))
            If widCol > 0 Then widVal = val(GetArrayValue(data, r, widCol))
            If thkCol > 0 Then thkVal = val(GetArrayValue(data, r, thkCol))

            hasDims = (lenVal > 0 And widVal > 0 And thkVal > 0)

            If ShouldUseBomItem(desc, typeText, matText, hasDims) Then

                Dim qtyVal As Long
                qtyVal = CLng(val(qtyText))

                If qtyVal <= 0 Then qtyVal = 1

                AddBomRow desc, typeText, matText, qtyVal, lenVal, widVal, thkVal, hasDims

            End If

        End If

    Next r

    Exit Sub

ErrHandler:
    LogLine "ReadBomWorksheetFastArray error: " & Err.Description
End Sub

Private Function ShouldUseBomItem(ByVal desc As String, _
                                  ByVal typeText As String, _
                                  ByVal materialText As String, _
                                  ByVal hasDims As Boolean) As Boolean

    Dim d As String
    d = NormalizeText(desc)

    If d = "" Then Exit Function
    If IsHardwareName(d) Then Exit Function

    If InStr(d, "PYROPEL") > 0 Or InStr(NormalizeText(materialText), "PYROPEL") > 0 Then
        ShouldUseBomItem = True
        Exit Function
    End If

    If InStr(d, "EJ J-BLOCK") > 0 Or InStr(d, "EJECTOR J-BLOCK") > 0 Or _
       InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Then
        ShouldUseBomItem = True
        Exit Function
    End If

    If InStr(d, "PULLCORE") > 0 Or InStr(d, "PULL CORE") > 0 Then
        ShouldUseBomItem = True
        Exit Function
    End If

    If InStr(d, "HOLDER") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "SMED") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "TCP") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "BCP") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "INSERT") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "INS") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "CAM") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "KEY") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "PLATE") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "STOP") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "COVER") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "POT") > 0 Then ShouldUseBomItem = True: Exit Function

    If hasDims Then ShouldUseBomItem = True: Exit Function

    ShouldUseBomItem = False
End Function

Private Sub AddBomRow(ByVal desc As String, _
                      ByVal typeText As String, _
                      ByVal matText As String, _
                      ByVal qty As Long, _
                      ByVal lenVal As Double, _
                      ByVal widVal As Double, _
                      ByVal thkVal As Double, _
                      ByVal hasDims As Boolean)

    Dim L As Double
    Dim W As Double
    Dim T As Double

    If hasDims Then
        SortThreeDimensions lenVal, widVal, thkVal, L, W, T
    Else
        L = 0#
        W = 0#
        T = 0#
    End If

    Dim isPyropel As Boolean
    isPyropel = (InStr(NormalizeText(desc), "PYROPEL") > 0 Or InStr(NormalizeText(matText), "PYROPEL") > 0)

    If HIDE_QUARTER_INCH_THICKNESS And hasDims And isPyropel = False Then
        If IsQuarterInchThickness(T) Then
            If IsInsertQuoteName(StandardPlateName(desc)) = False Then
                LogLine "BOM .250 item hidden: " & desc
                Exit Sub
            End If
        End If
    End If

    BomCount = BomCount + 1
    ReDim Preserve BomRows(1 To BomCount)

    BomRows(BomCount).Description = Trim(desc)
    BomRows(BomCount).quoteName = StandardPlateName(desc)
    BomRows(BomCount).TypeField = Trim(typeText)
    BomRows(BomCount).Quantity = qty
    BomRows(BomCount).material = NormalizeSteelType(matText)

    BomRows(BomCount).BomLength = Round(L, DIM_DECIMALS)
    BomRows(BomCount).BomWidth = Round(W, DIM_DECIMALS)
    BomRows(BomCount).BomThickness = Round(T, DIM_DECIMALS)
    BomRows(BomCount).hasDims = hasDims
End Sub

' ============================================================
' BOM READ (PDF via Poppler pdftotext)
' ============================================================

Private Sub ReadCustomerBomPdfUsingPdfToText(ByVal pdfPath As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(PDFTOTEXT_EXE) = False Then
        LogErrorText "pdftotext.exe not found at: " & PDFTOTEXT_EXE
        Exit Sub
    End If

    If fso.FileExists(pdfPath) = False Then
        LogErrorText "PDF not found: " & pdfPath
        Exit Sub
    End If

    Dim txtPath As String
    txtPath = CurrentJobFolder & "\XT_Export_BOM_PDF_Text.txt"

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = Chr(34) & PDFTOTEXT_EXE & Chr(34) & " -layout " & _
          Chr(34) & pdfPath & Chr(34) & " " & Chr(34) & txtPath & Chr(34)

    LogLine "Running pdftotext: " & cmd
    sh.Run cmd, 0, True

    If fso.FileExists(txtPath) = False Then
        LogErrorText "pdftotext produced no output file."
        Exit Sub
    End If

    Dim allText As String
    allText = ReadAllTextFile(txtPath)

    If Trim(allText) = "" Then
        LogErrorText "pdftotext output is empty."
        Exit Sub
    End If

    ParseBomTextFromPdf allText
    Exit Sub

ErrHandler:
    LogLine "ReadCustomerBomPdfUsingPdfToText error: " & Err.Description
End Sub

Private Sub ParseBomTextFromPdf(ByVal allText As String)
On Error GoTo ErrHandler

    Dim lines() As String

    allText = Replace(allText, vbCr, vbLf)

    Do While InStr(allText, vbLf & vbLf) > 0
        allText = Replace(allText, vbLf & vbLf, vbLf)
    Loop

    lines = Split(allText, vbLf)

    Dim i As Long

    For i = LBound(lines) To UBound(lines)
        ParseBomPdfTextLine lines(i)
    Next i

    LogLine "=== Parsed " & BomCount & " BOM rows from customer BOM ==="

    For i = 1 To BomCount
        LogLine "  BOMROW: '" & BomRows(i).Description & "' mat=" & BomRows(i).material & _
                " L/W/T=" & FormatNumberForCsv(BomRows(i).BomLength) & "/" & _
                FormatNumberForCsv(BomRows(i).BomWidth) & "/" & _
                FormatNumberForCsv(BomRows(i).BomThickness) & _
                " pullcore=" & CStr(IsPullcoreBomRow(BomRows(i)))
    Next i

    Exit Sub

ErrHandler:
    LogLine "ParseBomTextFromPdf error: " & Err.Description
End Sub

Private Sub ParseBomPdfTextLine(ByVal rawLine As String)
On Error GoTo ErrHandler

    Dim line As String
    line = Trim(rawLine)

    If line = "" Then Exit Sub

    Dim upperLine As String
    upperLine = UCase(line)

    If InStr(upperLine, "LTH") > 0 And InStr(upperLine, "WTH") > 0 Then Exit Sub
    If InStr(upperLine, "DESCRIPTION") > 0 And InStr(upperLine, "QTY") > 0 Then Exit Sub

    If TryParseTempcraftBasePdfMaterialLine(line) = False Then Exit Sub

    Exit Sub

ErrHandler:
    LogLine "ParseBomPdfTextLine error: " & Err.Description
End Sub

Private Function TryParseTempcraftBasePdfMaterialLine(ByVal line As String) As Boolean
On Error GoTo ErrHandler

    TryParseTempcraftBasePdfMaterialLine = False

    Dim upperLine As String
    upperLine = UCase(line)

    Dim desc As String
    Dim qty As Long
    Dim mat As String

    desc = ExtractTempcraftPdfDescription(line)
    If desc = "" Then Exit Function

    qty = ExtractTempcraftPdfQty(upperLine)
    mat = ExtractTempcraftPdfMaterial(upperLine)

    Dim nums() As Double
    Dim numCount As Long

    ExtractDecimalNumbers line, nums, numCount

    Dim lenVal As Double
    Dim widVal As Double
    Dim thkVal As Double
    Dim hasDims As Boolean

    lenVal = 0#
    widVal = 0#
    thkVal = 0#
    hasDims = False

    If numCount >= 3 Then
        lenVal = nums(0)
        widVal = nums(1)
        thkVal = nums(2)
        hasDims = (lenVal > 0 And widVal > 0 And thkVal > 0)
    End If

    If qty <= 0 Then qty = 1
    If mat = "" Then mat = DEFAULT_STEEL_TYPE

    If ShouldUseBomItem(desc, "Outsource", mat, hasDims) Then
        AddBomRow desc, "Outsource", mat, qty, lenVal, widVal, thkVal, hasDims
        TryParseTempcraftBasePdfMaterialLine = True
    End If

    Exit Function

ErrHandler:
    LogLine "TryParseTempcraftBasePdfMaterialLine error: " & Err.Description
    TryParseTempcraftBasePdfMaterialLine = False
End Function

Private Function ExtractTempcraftPdfDescription(ByVal line As String) As String
On Error GoTo ErrHandler

    Dim work As String
    work = RemoveLeadingItemNumber(line)

    Dim cutPos As Long
    cutPos = 0

    Dim upperWork As String
    upperWork = UCase(work)

    Dim kw As Variant
    Dim p As Long

    For Each kw In Array("OUTSOURCE", "PURCHASE", "MAKE", "BUY", "STOCK")
        p = InStr(upperWork, CStr(kw))
        If p > 0 Then
            If cutPos = 0 Or p < cutPos Then cutPos = p
        End If
    Next kw

    If cutPos > 1 Then
        ExtractTempcraftPdfDescription = Trim(Left(work, cutPos - 1))
    Else
        Dim i As Long

        For i = 1 To Len(work)
            If IsNumeric(mid(work, i, 1)) Then
                cutPos = i
                Exit For
            End If
        Next i

        If cutPos > 1 Then
            ExtractTempcraftPdfDescription = Trim(Left(work, cutPos - 1))
        Else
            ExtractTempcraftPdfDescription = Trim(work)
        End If
    End If

    Exit Function

ErrHandler:
    ExtractTempcraftPdfDescription = ""
End Function

Private Function ExtractTempcraftPdfQty(ByVal upperLine As String) As Long
On Error GoTo ErrHandler
    ExtractTempcraftPdfQty = ExtractQtyAfterOutsource(upperLine)
    Exit Function
ErrHandler:
    ExtractTempcraftPdfQty = 1
End Function

Private Function ExtractTempcraftPdfMaterial(ByVal upperLine As String) As String

    If InStr(upperLine, "4140") > 0 Then ExtractTempcraftPdfMaterial = "4140": Exit Function
    If InStr(upperLine, "4130") > 0 Then ExtractTempcraftPdfMaterial = "4130": Exit Function
    If InStr(upperLine, "A-2") > 0 Or InStr(upperLine, "A2") > 0 Then ExtractTempcraftPdfMaterial = "A-2": Exit Function
    If InStr(upperLine, "H-13") > 0 Or InStr(upperLine, "H13") > 0 Then ExtractTempcraftPdfMaterial = "H13": Exit Function
    If InStr(upperLine, "D-2") > 0 Or InStr(upperLine, "D2") > 0 Then ExtractTempcraftPdfMaterial = "D-2": Exit Function
    If InStr(upperLine, "DRILL ROD") > 0 Then ExtractTempcraftPdfMaterial = "Drill Rod": Exit Function
    If InStr(upperLine, "PYROPEL") > 0 Then ExtractTempcraftPdfMaterial = "Pyropel": Exit Function

    ExtractTempcraftPdfMaterial = ""
End Function

' ============================================================
' MATCH BOM TO CAD
' ============================================================

Private Sub BuildExportRowsFromBom()
On Error GoTo ErrHandler

    Dim i As Long

    Dim pullcoreRowCount As Long
    pullcoreRowCount = 0

    For i = 1 To BomCount
        If IsPullcoreBomRow(BomRows(i)) Then pullcoreRowCount = pullcoreRowCount + 1
    Next i

    LogLine "BOM rows total=" & BomCount & "  pullcore rows=" & pullcoreRowCount

    SetPackageTargetsFromBom

    MatchPullcoreRowsByClass

    For i = 1 To BomCount

        If IsPullcoreBomRow(BomRows(i)) = False Then

            Dim b As BomInfo
            b = BomRows(i)

            Dim cadIdx As Long
            cadIdx = FindBestCadMatchForBom(b)

            AddExportRow b, cadIdx

        End If

    Next i

    LogLine "After matching: ExportRows=" & ExportCount & "  PullcoreMatches=" & PullcoreMatchCount

    Exit Sub

ErrHandler:
    LogLine "BuildExportRowsFromBom error: " & Err.Description
End Sub

Private Sub SetPackageTargetsFromBom()
On Error GoTo ErrHandler

    JBlockTgtL = J_BLOCK_TARGET_LENGTH
    JBlockTgtW = J_BLOCK_TARGET_WIDTH
    JBlockTgtT = J_BLOCK_TARGET_THICKNESS

    EjCamTgtL = EJECTOR_CAM_TARGET_LENGTH
    EjCamTgtW = EJECTOR_CAM_TARGET_WIDTH
    EjCamTgtT = EJECTOR_CAM_TARGET_THICKNESS

    Dim i As Long
    Dim d As String

    For i = 1 To BomCount

        d = NormalizeText(BomRows(i).Description)

        If (InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Or InStr(d, "JBLOCK") > 0 _
            Or InStr(d, "U-BLOCK") > 0 Or InStr(d, "U BLOCK") > 0 Or InStr(d, "UBLOCK") > 0) Then
            If BomRows(i).BomLength > 0 Then
                JBlockTgtL = BomRows(i).BomLength
                JBlockTgtW = BomRows(i).BomWidth
                JBlockTgtT = BomRows(i).BomThickness

                LogLine "J BLOCK target from BOM: L=" & FormatNumberForCsv(JBlockTgtL) & _
                        " W=" & FormatNumberForCsv(JBlockTgtW) & _
                        " T=" & FormatNumberForCsv(JBlockTgtT)
            End If
        End If

        If (InStr(d, "CAM") > 0) And (InStr(d, "EJ") > 0 Or InStr(d, "EJECTOR") > 0) _
           And InStr(d, "COVER") = 0 And InStr(d, "J-BLOCK") = 0 And InStr(d, "J BLOCK") = 0 Then

            If BomRows(i).BomLength > 0 Then
                EjCamTgtL = BomRows(i).BomLength
                EjCamTgtW = BomRows(i).BomWidth
                EjCamTgtT = BomRows(i).BomThickness

                LogLine "EJECTOR CAM target from BOM: L=" & FormatNumberForCsv(EjCamTgtL) & _
                        " W=" & FormatNumberForCsv(EjCamTgtW) & _
                        " T=" & FormatNumberForCsv(EjCamTgtT)
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "SetPackageTargetsFromBom error: " & Err.Description
End Sub

Private Function IsPullcoreBomRow(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)

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
        IsPullcoreBomRow = True
        Exit Function
    End If

    If HasPullcoreLocationToken(d) Then
        IsPullcoreBomRow = True
        Exit Function
    End If

    ' BOM qty-2 rows that only list material (H-13 cam / A-2 key) without "Pullcore" text.
    If b.hasDims Then
        If IsPullcoreCamBomRow(b) And LooksLikePullcoreCamBomDims(b) Then
            IsPullcoreBomRow = True
            Exit Function
        End If
        If IsPullcoreKeyBomRow(b) And LooksLikePullcoreKeyBomDims(b) Then
            IsPullcoreBomRow = True
            Exit Function
        End If
    End If

    If InStr(d, "CAM") = 0 And InStr(d, "KEY") = 0 Then Exit Function

    If InStr(d, "PULLCORE") > 0 Or InStr(d, "PULL CORE") > 0 Then
        IsPullcoreBomRow = True
    End If
End Function

Private Function IsPullcoreCamBomRow(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)

    If InStr(d, "KEY") > 0 And InStr(d, "CAM") = 0 Then Exit Function
    If InStr(d, "CAM") > 0 Then IsPullcoreCamBomRow = True: Exit Function
    If NormalizeSteelType(b.material) = "H13" Then IsPullcoreCamBomRow = True
End Function

Private Function IsPullcoreKeyBomRow(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)

    If InStr(d, "CAM") > 0 And InStr(d, "KEY") = 0 Then Exit Function
    If InStr(d, "KEY") > 0 Then IsPullcoreKeyBomRow = True: Exit Function

    If NormalizeSteelType(b.material) = "A-2" Then
        If InStr(d, "EJECTOR") = 0 And InStr(d, "J-BLOCK") = 0 And InStr(d, "J BLOCK") = 0 Then
            IsPullcoreKeyBomRow = True
        End If
    End If
End Function

Private Function LooksLikePullcoreCamBomDims(ByRef b As BomInfo) As Boolean
    If b.hasDims = False Then Exit Function

    If b.BomLength < 3.4 Or b.BomLength > 4.6 Then Exit Function
    If b.BomWidth < 2# Or b.BomWidth > 3.2 Then Exit Function
    If b.BomThickness < 0.75 Or b.BomThickness > 1.6 Then Exit Function

    LooksLikePullcoreCamBomDims = True
End Function

Private Function LooksLikePullcoreKeyBomDims(ByRef b As BomInfo) As Boolean
    If b.hasDims = False Then Exit Function

    If b.BomLength < 2.5 Or b.BomLength > 3.6 Then Exit Function
    If b.BomWidth < 0.9 Or b.BomWidth > 1.7 Then Exit Function
    If b.BomThickness < 0.55 Or b.BomThickness > 1.1 Then Exit Function

    LooksLikePullcoreKeyBomDims = True
End Function

Private Function HasPullcoreLocationToken(ByVal d As String) As Boolean
    If GetPullcoreLocationCode(d) <> "" Then
        HasPullcoreLocationToken = True
        Exit Function
    End If

    Dim toks() As String
    toks = Split(d, " ")

    Dim i As Long
    For i = LBound(toks) To UBound(toks)
        Select Case toks(i)
            Case "TE", "LE", "ID", "OD", "IDTE", "IDLE", "ODTE", "ODLE"
                HasPullcoreLocationToken = True
                Exit Function
        End Select
    Next i
End Function

Private Sub AddPullcoreMatchRow(ByRef b As BomInfo, ByVal cadIdx As Long, _
                                ByVal cadL As Double, ByVal cadW As Double, ByVal cadT As Double)

    PullcoreMatchCount = PullcoreMatchCount + 1
    ReDim Preserve PullcoreMatches(1 To PullcoreMatchCount)

    Dim displayName As String
    displayName = CleanPullcoreDisplayName(b.Description)

    If displayName = "" Then displayName = b.quoteName

    PullcoreMatches(PullcoreMatchCount).quoteName = displayName
    PullcoreMatches(PullcoreMatchCount).Description = b.Description
    PullcoreMatches(PullcoreMatchCount).material = b.material
    PullcoreMatches(PullcoreMatchCount).Quantity = b.Quantity

    PullcoreMatches(PullcoreMatchCount).CadPartIndex = cadIdx
    PullcoreMatches(PullcoreMatchCount).isCam = IsPullcoreCamBomRow(b)

    PullcoreMatches(PullcoreMatchCount).BomThickness = b.BomThickness
    PullcoreMatches(PullcoreMatchCount).BomWidth = b.BomWidth
    PullcoreMatches(PullcoreMatchCount).BomLength = b.BomLength

    If cadIdx > 0 Then

        PullcoreMatches(PullcoreMatchCount).CadThickness = cadT
        PullcoreMatches(PullcoreMatchCount).CadWidth = cadW
        PullcoreMatches(PullcoreMatchCount).CadLength = cadL

        PullcoreMatches(PullcoreMatchCount).FittedThickness = cadT
        PullcoreMatches(PullcoreMatchCount).FittedWidth = cadW
        PullcoreMatches(PullcoreMatchCount).FittedLength = cadL

        If parts(cadIdx).hasOriginalAsmBBox Then
            PullcoreMatches(PullcoreMatchCount).OriginalLength = parts(cadIdx).OriginalAsmLength
            PullcoreMatches(PullcoreMatchCount).OriginalWidth = parts(cadIdx).OriginalAsmWidth
            PullcoreMatches(PullcoreMatchCount).OriginalThickness = parts(cadIdx).OriginalAsmThickness
        Else
            PullcoreMatches(PullcoreMatchCount).OriginalLength = parts(cadIdx).Length
            PullcoreMatches(PullcoreMatchCount).OriginalWidth = parts(cadIdx).Width
            PullcoreMatches(PullcoreMatchCount).OriginalThickness = parts(cadIdx).Thickness
        End If

        PullcoreMatches(PullcoreMatchCount).Status = ComparePullcoreDimsToBomStatus(b, cadL, cadW, cadT)

        ' PULL ANGLE - hybrid of the two measurements, taking each one's
        ' strength:
        '   * BOX magnitude (original vs fitted AABB) is robust, unambiguous and
        '     identical for a cam and its key (they pull together). It is the
        '     trustworthy anchor for HOW MUCH the part is tipped.
        '   * EDGE angle (longest cross edge in the active-view plane) gives the
        '     SIGN (which way it leans in the front view) and, when it AGREES
        '     with the box magnitude, a more precise value.
        ' The edge value is only adopted for magnitude when it is within a few
        ' degrees of the box value - this rejects the edge method's known
        ' failure modes (a body oriented so its cross edge sits near an axis can
        ' fold to the wrong value). Otherwise the box magnitude stands and the
        ' edge contributes sign only.
        Dim edgeDeg As Double
        Dim haveEdge As Boolean
        Dim boxDeg As Double
        Dim pcModel As Object

        haveEdge = False
        edgeDeg = 0#

        Set pcModel = GetModelDocForCadIndex(cadIdx)
        If Not pcModel Is Nothing Then
            If pcModel.GetType = swDocPART Then
                If TryComputePullcoreLongestModelEdgeAngleFromModel(pcModel, edgeDeg) Then
                    haveEdge = True
                End If
            End If
        End If

        boxDeg = EstimatePullcoreAngleFromFittedAndOriginalDeg( _
                     cadL, cadW, cadT, _
                     PullcoreMatches(PullcoreMatchCount).OriginalLength, _
                     PullcoreMatches(PullcoreMatchCount).OriginalWidth, _
                     PullcoreMatches(PullcoreMatchCount).OriginalThickness)

        Dim finalDeg As Double
        Dim angleSign As Double

        finalDeg = Abs(boxDeg)
        angleSign = 1#

        ' Sign from the edge if we have it, else from the assembly transform.
        If haveEdge And Abs(edgeDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            If edgeDeg < 0# Then angleSign = -1# Else angleSign = 1#
        ElseIf cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).hasAsmPullAngle Then angleSign = parts(cadIdx).asmPullAngleSign
        End If

        ' Adopt the (more precise) edge magnitude only when it corroborates the
        ' box magnitude, i.e. they agree to within PULLCORE_ANGLE_AGREE_TOL deg.
        If haveEdge And Abs(boxDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            If Abs(Abs(edgeDeg) - Abs(boxDeg)) <= PULLCORE_ANGLE_AGREE_TOL Then
                finalDeg = Abs(edgeDeg)
                LogLine "  Pullcore angle: edge " & Format(edgeDeg, "0.00") & _
                        " agrees with box " & Format(boxDeg, "0.00") & " - using edge."
            Else
                LogLine "  Pullcore angle: edge " & Format(edgeDeg, "0.00") & _
                        " disagrees with box " & Format(boxDeg, "0.00") & _
                        " - using box magnitude (edge gives sign only)."
            End If
        End If

        If finalDeg >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg = finalDeg
            PullcoreMatches(PullcoreMatchCount).DetectedAngleSign = angleSign
            LogLine "  Pullcore pull angle = " & Format(finalDeg, "0.00") & _
                    " deg (sign " & Format(angleSign, "0") & ")"
        Else
            PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg = 0#
            PullcoreMatches(PullcoreMatchCount).DetectedAngleSign = 1#
            LogLine "  Pullcore angle below threshold - part is square in front view."
        End If

        If Abs(PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            ' Rotate by the OPPOSITE of the part's actual tilt direction to
            ' square it. angleSign encodes which way it is tipped.
            PullcoreMatches(PullcoreMatchCount).DxfRotationDeg = _
                PULLCORE_STRAIGHTEN_SIGN * angleSign * _
                PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg
        Else
            PullcoreMatches(PullcoreMatchCount).DxfRotationDeg = 0#
        End If

        ' Donate this angle to the angle-aware candidate filter so a sibling
        ' part (e.g. the KEY of a matched CAM) at the same angle can be found.
        RememberDetectedPullcoreAngle PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg


    Else
        PullcoreMatches(PullcoreMatchCount).Status = "NO CAD MATCH"
    End If

    Dim cadNameForLog As String

    If cadIdx > 0 Then
        cadNameForLog = parts(cadIdx).componentName
    Else
        cadNameForLog = "(none)"
    End If

    LogLine "PULLCORE matched: '" & displayName & "' -> CAD '" & cadNameForLog & "'"

    If cadIdx > 0 Then
        LogLine "  Original assembly/world bbox L/W/T=" & _
                FormatNumberForCsv(PullcoreMatches(PullcoreMatchCount).OriginalLength) & "/" & _
                FormatNumberForCsv(PullcoreMatches(PullcoreMatchCount).OriginalWidth) & "/" & _
                FormatNumberForCsv(PullcoreMatches(PullcoreMatchCount).OriginalThickness)

        LogLine "  Fitted bbox L/W/T=" & _
                FormatNumberForCsv(cadL) & "/" & _
                FormatNumberForCsv(cadW) & "/" & _
                FormatNumberForCsv(cadT)

        LogLine "  Pullcore angle=" & _
                Format(PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg, "0.00") & _
                " deg, DXF rotation=" & _
                Format(PullcoreMatches(PullcoreMatchCount).DxfRotationDeg, "0.00") & " deg"

        If parts(cadIdx).hasAsmCenter Then
            LogLine "  Pullcore assembly center X/Y/Z=" & _
                    FormatNumberForCsv(parts(cadIdx).AsmCenterX) & "/" & _
                    FormatNumberForCsv(parts(cadIdx).AsmCenterY) & "/" & _
                    FormatNumberForCsv(parts(cadIdx).AsmCenterZ)
        End If
    End If
End Sub

Private Function ComparePullcoreDimsToBomStatus(ByRef b As BomInfo, _
                                                ByVal cadL As Double, _
                                                ByVal cadW As Double, _
                                                ByVal cadT As Double) As String
On Error GoTo ErrHandler

    If b.hasDims = False Then
        ComparePullcoreDimsToBomStatus = "MATCH (no BOM dims)"
        Exit Function
    End If

    Dim worst As Double

    worst = Abs(cadL - b.BomLength)

    If Abs(cadW - b.BomWidth) > worst Then worst = Abs(cadW - b.BomWidth)
    If Abs(cadT - b.BomThickness) > worst Then worst = Abs(cadT - b.BomThickness)

    If worst <= DIM_OK_TOL Then
        ComparePullcoreDimsToBomStatus = "OK"
    ElseIf worst <= DIM_REVIEW_TOL Then
        ComparePullcoreDimsToBomStatus = "REVIEW (" & FormatNumberForCsv(worst) & ")"
    Else
        ComparePullcoreDimsToBomStatus = "CHECK (" & FormatNumberForCsv(worst) & ")"
    End If

    Exit Function

ErrHandler:
    ComparePullcoreDimsToBomStatus = "ERROR"
End Function

Private Function EstimatePullcoreAngleFromFittedAndOriginalDeg( _
    ByVal fitL As Double, ByVal fitW As Double, ByVal fitT As Double, _
    ByVal origL As Double, ByVal origW As Double, ByVal origT As Double) As Double
On Error GoTo ErrHandler

    EstimatePullcoreAngleFromFittedAndOriginalDeg = 0#

    If fitL <= 0 Or fitW <= 0 Or fitT <= 0 Then Exit Function
    If origL <= 0 Or origW <= 0 Or origT <= 0 Then Exit Function

    Dim b(1 To 3) As Double
    Dim n(1 To 3) As Double

    b(1) = fitL
    b(2) = fitW
    b(3) = fitT

    n(1) = origL
    n(2) = origW
    n(3) = origT

    SortArray3Descending b
    SortArray3Descending n

    If Abs(b(1) - n(1)) <= 0.001 And _
       Abs(b(2) - n(2)) <= 0.001 And _
       Abs(b(3) - n(3)) <= 0.001 Then
        EstimatePullcoreAngleFromFittedAndOriginalDeg = 0#
        Exit Function
    End If

    Dim i As Long
    Dim unchangedIdx As Long
    Dim bestDiff As Double
    Dim d As Double

    unchangedIdx = 1
    bestDiff = 1E+99

    For i = 1 To 3
        d = Abs(b(i) - n(i))
        If d < bestDiff Then
            bestDiff = d
            unchangedIdx = i
        End If
    Next i

    Dim ba As Double
    Dim bB As Double
    Dim nA As Double
    Dim nB As Double
    Dim c As Long

    c = 0

    For i = 1 To 3
        If i <> unchangedIdx Then
            c = c + 1

            If c = 1 Then
                ba = b(i)
                nA = n(i)
            ElseIf c = 2 Then
                bB = b(i)
                nB = n(i)
            End If
        End If
    Next i

    If bB > ba Then
        SwapDoubleValues ba, bB
        SwapDoubleValues nA, nB
    End If

    EstimatePullcoreAngleFromFittedAndOriginalDeg = SearchBoxRotationDeg(ba, bB, nA, nB)

    Exit Function

ErrHandler:
    EstimatePullcoreAngleFromFittedAndOriginalDeg = 0#
End Function

Private Sub SortArray3Descending(ByRef a() As Double)
    If a(2) > a(1) Then SwapDoubleValues a(1), a(2)
    If a(3) > a(2) Then SwapDoubleValues a(2), a(3)
    If a(2) > a(1) Then SwapDoubleValues a(1), a(2)
End Sub

Private Sub SwapDoubleValues(ByRef a As Double, ByRef b As Double)
    Dim T As Double
    T = a
    a = b
    b = T
End Sub

' ============================================================
' END OF PART 5
' Paste Part 6 immediately after this.
' ============================================================
' ============================================================
' PULLCORE MATCHING
' ============================================================

Private Sub MatchPullcoreRowsByClass()
On Error GoTo ErrHandler

    Dim camRows() As Long
    Dim keyRows() As Long
    Dim camN As Long
    Dim keyN As Long

    camN = 0
    keyN = 0

    ReDim camRows(1 To 1)
    ReDim keyRows(1 To 1)

    Dim i As Long

    For i = 1 To BomCount

        If IsPullcoreBomRow(BomRows(i)) Then

            Dim reps As Long
            reps = BomRows(i).Quantity

            If reps < 1 Then reps = 1
            If reps > 8 Then reps = 8

            Dim isCamRow As Boolean
            isCamRow = IsPullcoreCamBomRow(BomRows(i))

            ' Old behavior forced every key row to count at least twice, which
            ' invented phantom "NO CAD MATCH" rows when the model held only one
            ' instance of each key. Honor the real BOM quantity instead.
            If Not PULLCORE_RESPECT_BOM_QUANTITY Then
                If (Not isCamRow) And reps < 2 Then reps = 2
            End If

            Dim r As Long
            For r = 1 To reps
                If isCamRow Then
                    camN = camN + 1
                    ReDim Preserve camRows(1 To camN)
                    camRows(camN) = i
                Else
                    keyN = keyN + 1
                    ReDim Preserve keyRows(1 To keyN)
                    keyRows(keyN) = i
                End If
            Next r

        End If

    Next i

    LogLine "Pullcore BOM rows (qty-expanded): cams=" & camN & " keys=" & keyN

    If camN > 0 Then AssignPullcoreClass camRows, camN, True
    If keyN > 0 Then AssignPullcoreClass keyRows, keyN, False

    ' Make absolutely sure a CAD pullcore component is not reused twice.
    BlockDuplicatePullcoreCadMatches

    ' PRIMARY: ID/OD from the holder centres of mass, with the Z direction
    ' standardised so the ID holder is always on the positive side. This is
    ' immune to a flipped assembly, which the fixed-axis method below is not.
    Dim labelledByHolder As Boolean
    labelledByHolder = LabelPullcoreByHolderComZ()

    If labelledByHolder = False Then

        LogLine "Pullcore ID/OD: holder-COM method unavailable; using the height axis."

        If AUTO_LABEL_PULLCORE_ID_OD_BY_HEIGHT Then
            If LabelPullcoreCamsAndKeysByHeight() = False Then
                LabelPullcoreKeysByProximity
            End If
        Else
            LabelPullcoreKeysByProximity
        End If

    End If

    ' Final authority for generic qty-2 pullcore rows:
    ' highest Y = ID, lowest Y = OD.
    ' Skipped when the holder-COM method already labelled them, otherwise this
    ' fixed-Y rule would immediately undo the holder-referenced result.
    If labelledByHolder Then
        LogLine "Pullcore ID/OD: keeping holder-COM labels (fixed-Y finaliser skipped)."
    Else
        FinalizePullcoreCamIdOdNames
    End If

    ' For keys, keep them tied to the nearest/nested cam side.
    RelabelPullcoreKeysByNestedCam

    ' When one BOM row (qty 2+) expands to several physical parts and the BOM
    ' name has no ID/OD yet, split them ID/OD by height instead of " (2)".
    DisambiguatePullcoreDescriptiveGroupsByIdOd

    ' Must run last so final names are unique.
    DisambiguatePullcoreQuoteNames

    On Error Resume Next
    If Not swModel Is Nothing Then
        Dim errs As Long
        swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    End If
    EnsureSwHidden
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    LogLine "MatchPullcoreRowsByClass error: " & Err.Description
End Sub

Private Function LabelPullcoreCamsAndKeysByHeight() As Boolean
On Error GoTo ErrHandler

    LabelPullcoreCamsAndKeysByHeight = False
    PullcoreIdOdHeightAxisUsed = ""

    If PullcoreMatchCount <= 0 Then Exit Function

    Dim camMatches As Collection
    Dim keyMatches As Collection

    Set camMatches = New Collection
    Set keyMatches = New Collection

    Dim i As Long
    Dim cadIdx As Long

    For i = 1 To PullcoreMatchCount

        cadIdx = PullcoreMatches(i).CadPartIndex

        If cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).hasAsmCenter Then

                If PullcoreMatches(i).isCam Then
                    camMatches.Add i
                Else
                    keyMatches.Add i
                End If

            End If
        End If

    Next i

    If camMatches.count < 2 Then
        LogLine "Pullcore ID/OD height labeling skipped: need at least 2 CAM matches with assembly centers."
        Exit Function
    End If

    Dim axisName As String
    axisName = DeterminePullcoreIdOdHeightAxis(camMatches, keyMatches)

    If axisName = "" Then axisName = "Y"

    PullcoreIdOdHeightAxisUsed = axisName

    LogLine "Pullcore ID/OD height labeling axis selected: " & axisName

    Dim highCamMatch As Long
    Dim lowCamMatch As Long

    FindHighLowPullcoreMatchesByAxis camMatches, axisName, highCamMatch, lowCamMatch

    If highCamMatch <= 0 Or lowCamMatch <= 0 Then
        LogLine "Pullcore ID/OD height labeling skipped: could not determine high/low CAM."
        Exit Function
    End If

    Dim idCamMatch As Long
    Dim odCamMatch As Long

    If PULLCORE_ID_IS_HIGHER Then
        idCamMatch = highCamMatch
        odCamMatch = lowCamMatch
    Else
        idCamMatch = lowCamMatch
        odCamMatch = highCamMatch
    End If

    SetPullcoreMatchLabel idCamMatch, "ID PULLCORE CAM", axisName
    SetPullcoreMatchLabel odCamMatch, "OD PULLCORE CAM", axisName

    If keyMatches.count > 0 Then
        LabelPullcoreKeysToNearestLabeledCams keyMatches, idCamMatch, odCamMatch, axisName
    End If

    LabelPullcoreCamsAndKeysByHeight = True

    Exit Function

ErrHandler:
    LogLine "LabelPullcoreCamsAndKeysByHeight error: " & Err.Description
    LabelPullcoreCamsAndKeysByHeight = False
End Function

Private Function DeterminePullcoreIdOdHeightAxis(ByVal camMatches As Collection, _
                                                 ByVal keyMatches As Collection) As String
On Error GoTo ErrHandler

    Dim setting As String
    setting = UCase(Trim(PULLCORE_ID_OD_HEIGHT_AXIS))

    If setting = "Y" Or setting = "Z" Then
        DeterminePullcoreIdOdHeightAxis = setting
        Exit Function
    End If

    Dim yMin As Double
    Dim yMax As Double
    Dim zMin As Double
    Dim zMax As Double
    Dim firstVal As Boolean

    firstVal = True

    Dim i As Long
    Dim mi As Long
    Dim cadIdx As Long

    For i = 1 To camMatches.count

        mi = CLng(camMatches(i))
        cadIdx = PullcoreMatches(mi).CadPartIndex

        If cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).hasAsmCenter Then

                If firstVal Then
                    yMin = parts(cadIdx).AsmCenterY
                    yMax = parts(cadIdx).AsmCenterY
                    zMin = parts(cadIdx).AsmCenterZ
                    zMax = parts(cadIdx).AsmCenterZ
                    firstVal = False
                Else
                    If parts(cadIdx).AsmCenterY < yMin Then yMin = parts(cadIdx).AsmCenterY
                    If parts(cadIdx).AsmCenterY > yMax Then yMax = parts(cadIdx).AsmCenterY
                    If parts(cadIdx).AsmCenterZ < zMin Then zMin = parts(cadIdx).AsmCenterZ
                    If parts(cadIdx).AsmCenterZ > zMax Then zMax = parts(cadIdx).AsmCenterZ
                End If

            End If
        End If

    Next i

    If firstVal Then
        DeterminePullcoreIdOdHeightAxis = "Y"
        Exit Function
    End If

    If Abs(zMax - zMin) > Abs(yMax - yMin) + 0.001 Then
        DeterminePullcoreIdOdHeightAxis = "Z"
    Else
        DeterminePullcoreIdOdHeightAxis = "Y"
    End If

    LogLine "Pullcore CAM center spread: Y=" & FormatNumberForCsv(Abs(yMax - yMin)) & _
            " Z=" & FormatNumberForCsv(Abs(zMax - zMin))

    Exit Function

ErrHandler:
    DeterminePullcoreIdOdHeightAxis = "Y"
End Function

Private Sub FindHighLowPullcoreMatchesByAxis(ByVal matches As Collection, _
                                             ByVal axisName As String, _
                                             ByRef highMatch As Long, _
                                             ByRef lowMatch As Long)
On Error GoTo ErrHandler

    highMatch = 0
    lowMatch = 0

    If matches Is Nothing Then Exit Sub
    If matches.count = 0 Then Exit Sub

    Dim bestHigh As Double
    Dim bestLow As Double

    bestHigh = -1E+99
    bestLow = 1E+99

    Dim i As Long
    Dim mi As Long
    Dim h As Double

    For i = 1 To matches.count

        mi = CLng(matches(i))
        h = GetPullcoreMatchHeightValue(mi, axisName)

        If h > bestHigh Then
            bestHigh = h
            highMatch = mi
        End If

        If h < bestLow Then
            bestLow = h
            lowMatch = mi
        End If

    Next i

    Exit Sub

ErrHandler:
    highMatch = 0
    lowMatch = 0
End Sub

Private Function GetPullcoreMatchHeightValue(ByVal matchIdx As Long, ByVal axisName As String) As Double
On Error GoTo ErrHandler

    GetPullcoreMatchHeightValue = 0#

    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Function

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If parts(cadIdx).hasAsmCenter = False Then Exit Function

    If UCase(axisName) = "Z" Then
        GetPullcoreMatchHeightValue = parts(cadIdx).AsmCenterZ
    Else
        GetPullcoreMatchHeightValue = parts(cadIdx).AsmCenterY
    End If

    Exit Function

ErrHandler:
    GetPullcoreMatchHeightValue = 0#
End Function

Private Sub SetPullcoreMatchLabel(ByVal matchIdx As Long, _
                                  ByVal newName As String, _
                                  ByVal axisName As String)
On Error Resume Next

    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    ' Keep TE/LE style BOM names; generic "Pullcore Cam" rows get ID/OD labels.
    If PULLCORE_PRESERVE_DESCRIPTIVE_BOM_NAMES Then
        If PullcoreNameEncodesLocation(PullcoreMatches(matchIdx).Description) Then
            LogLine "Pullcore keeping descriptive BOM name '" & _
                    PullcoreMatches(matchIdx).quoteName & "' (would have been '" & newName & "')"
            Exit Sub
        End If
    End If

    PullcoreMatches(matchIdx).quoteName = newName
    PullcoreMatches(matchIdx).Description = newName

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    If cadIdx > 0 And cadIdx <= PartCount Then

        LogLine "Pullcore labeled '" & newName & "' -> CAD '" & _
                parts(cadIdx).componentName & "'  center X/Y/Z=" & _
                FormatNumberForCsv(parts(cadIdx).AsmCenterX) & "/" & _
                FormatNumberForCsv(parts(cadIdx).AsmCenterY) & "/" & _
                FormatNumberForCsv(parts(cadIdx).AsmCenterZ) & _
                "  ID/OD axis=" & axisName & _
                "  height=" & FormatNumberForCsv(GetPullcoreMatchHeightValue(matchIdx, axisName))

    Else

        LogLine "Pullcore labeled '" & newName & "' with no CAD component."

    End If
End Sub

Private Sub LabelPullcoreKeysToNearestLabeledCams(ByVal keyMatches As Collection, _
                                                  ByVal idCamMatch As Long, _
                                                  ByVal odCamMatch As Long, _
                                                  ByVal axisName As String)
On Error GoTo ErrHandler

    If keyMatches Is Nothing Then Exit Sub
    If keyMatches.count = 0 Then Exit Sub
    If idCamMatch <= 0 Or odCamMatch <= 0 Then Exit Sub

    If keyMatches.count = 1 Then

        Dim onlyKey As Long
        onlyKey = CLng(keyMatches(1))

        If PullcoreMatchDistanceSq(onlyKey, idCamMatch) <= PullcoreMatchDistanceSq(onlyKey, odCamMatch) Then
            SetPullcoreMatchLabel onlyKey, "ID PULLCORE KEY", axisName
        Else
            SetPullcoreMatchLabel onlyKey, "OD PULLCORE KEY", axisName
        End If

        Exit Sub

    End If

    If keyMatches.count = 2 Then

        Dim k1 As Long
        Dim k2 As Long

        k1 = CLng(keyMatches(1))
        k2 = CLng(keyMatches(2))

        Dim optionA As Double
        Dim optionB As Double

        optionA = PullcoreMatchDistanceSq(k1, idCamMatch) + PullcoreMatchDistanceSq(k2, odCamMatch)
        optionB = PullcoreMatchDistanceSq(k2, idCamMatch) + PullcoreMatchDistanceSq(k1, odCamMatch)

        If optionA <= optionB Then
            SetPullcoreMatchLabel k1, "ID PULLCORE KEY", axisName
            SetPullcoreMatchLabel k2, "OD PULLCORE KEY", axisName
        Else
            SetPullcoreMatchLabel k2, "ID PULLCORE KEY", axisName
            SetPullcoreMatchLabel k1, "OD PULLCORE KEY", axisName
        End If

        Exit Sub

    End If

    Dim i As Long
    Dim mi As Long

    For i = 1 To keyMatches.count

        mi = CLng(keyMatches(i))

        If PullcoreMatchDistanceSq(mi, idCamMatch) <= PullcoreMatchDistanceSq(mi, odCamMatch) Then
            SetPullcoreMatchLabel mi, "ID PULLCORE KEY", axisName
        Else
            SetPullcoreMatchLabel mi, "OD PULLCORE KEY", axisName
        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "LabelPullcoreKeysToNearestLabeledCams error: " & Err.Description
End Sub

Private Function PullcoreMatchDistanceSq(ByVal matchA As Long, ByVal matchB As Long) As Double
On Error GoTo ErrHandler

    PullcoreMatchDistanceSq = 1E+99

    If matchA <= 0 Or matchA > PullcoreMatchCount Then Exit Function
    If matchB <= 0 Or matchB > PullcoreMatchCount Then Exit Function

    Dim cadA As Long
    Dim cadB As Long

    cadA = PullcoreMatches(matchA).CadPartIndex
    cadB = PullcoreMatches(matchB).CadPartIndex

    If cadA <= 0 Or cadA > PartCount Then Exit Function
    If cadB <= 0 Or cadB > PartCount Then Exit Function

    If parts(cadA).hasAsmCenter = False Then Exit Function
    If parts(cadB).hasAsmCenter = False Then Exit Function

    PullcoreMatchDistanceSq = _
        (parts(cadA).AsmCenterX - parts(cadB).AsmCenterX) ^ 2 + _
        (parts(cadA).AsmCenterY - parts(cadB).AsmCenterY) ^ 2 + _
        (parts(cadA).AsmCenterZ - parts(cadB).AsmCenterZ) ^ 2

    Exit Function

ErrHandler:
    PullcoreMatchDistanceSq = 1E+99
End Function

Private Sub LabelPullcoreKeysByProximity()
On Error GoTo ErrHandler

    Dim haveId As Boolean
    Dim haveOd As Boolean

    Dim idx As Double
    Dim idy As Double
    Dim idz As Double
    Dim odx As Double
    Dim ody As Double
    Dim odz As Double

    Dim i As Long
    Dim ci As Long
    Dim nm As String

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).isCam And PullcoreMatches(i).CadPartIndex > 0 Then

            ci = PullcoreMatches(i).CadPartIndex

            If parts(ci).hasAsmCenter Then

                nm = NormalizeText(PullcoreMatches(i).Description)

                If InStr(nm, "OD") > 0 Then
                    odx = parts(ci).AsmCenterX
                    ody = parts(ci).AsmCenterY
                    odz = parts(ci).AsmCenterZ
                    haveOd = True
                ElseIf InStr(nm, "ID") > 0 Then
                    idx = parts(ci).AsmCenterX
                    idy = parts(ci).AsmCenterY
                    idz = parts(ci).AsmCenterZ
                    haveId = True
                End If

            End If

        End If

    Next i

    If Not (haveId And haveOd) Then
        LogLine "Pullcore key ID/OD labeling skipped (need both ID and OD cam centers)."
        Exit Sub
    End If

    Dim keyIdx() As Long
    Dim keyN As Long

    keyN = 0
    ReDim keyIdx(1 To PullcoreMatchCount)

    For i = 1 To PullcoreMatchCount
        If (Not PullcoreMatches(i).isCam) And PullcoreMatches(i).CadPartIndex > 0 Then
            If parts(PullcoreMatches(i).CadPartIndex).hasAsmCenter Then
                keyN = keyN + 1
                keyIdx(keyN) = i
            End If
        End If
    Next i

    If keyN = 0 Then Exit Sub

    Dim bias() As Double
    ReDim bias(1 To keyN)

    Dim k As Long

    For k = 1 To keyN

        ci = PullcoreMatches(keyIdx(k)).CadPartIndex

        Dim dId As Double
        Dim dOd As Double

        dId = (parts(ci).AsmCenterX - idx) ^ 2 + (parts(ci).AsmCenterY - idy) ^ 2 + (parts(ci).AsmCenterZ - idz) ^ 2
        dOd = (parts(ci).AsmCenterX - odx) ^ 2 + (parts(ci).AsmCenterY - ody) ^ 2 + (parts(ci).AsmCenterZ - odz) ^ 2

        bias(k) = dId - dOd

    Next k

    If keyN = 2 Then

        Dim idK As Long
        Dim odK As Long

        If bias(1) <= bias(2) Then
            idK = 1
            odK = 2
        Else
            idK = 2
            odK = 1
        End If

        SetPullcoreKeyLabel keyIdx(idK), "ID PULLCORE KEY"
        SetPullcoreKeyLabel keyIdx(odK), "OD PULLCORE KEY"

    Else

        For k = 1 To keyN
            If bias(k) <= 0 Then
                SetPullcoreKeyLabel keyIdx(k), "ID PULLCORE KEY"
            Else
                SetPullcoreKeyLabel keyIdx(k), "OD PULLCORE KEY"
            End If
        Next k

    End If

    Exit Sub

ErrHandler:
    LogLine "LabelPullcoreKeysByProximity error: " & Err.Description
End Sub

Private Sub SetPullcoreKeyLabel(ByVal matchIdx As Long, ByVal newName As String)
On Error Resume Next

    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    If PULLCORE_PRESERVE_DESCRIPTIVE_BOM_NAMES Then
        If GetPullcoreLocationCode(PullcoreMatches(matchIdx).Description) <> "" Then
            LogLine "Pullcore key keeping descriptive BOM name '" & _
                    PullcoreMatches(matchIdx).quoteName & "' (would have been '" & newName & "')"
            Exit Sub
        End If
    End If

    PullcoreMatches(matchIdx).quoteName = newName

    LogLine "Pullcore key labeled '" & newName & "' -> CAD '" & _
            parts(PullcoreMatches(matchIdx).CadPartIndex).componentName & "'"
End Sub

' Ensure every pullcore match has a unique quote name. Descriptive BOM names
' are kept as-is; genuine duplicates (e.g. a qty-2 row that maps to two
' physical parts) get a trailing " (2)", " (3)", ... so exported files and
' DXFs do not silently overwrite one another.
' True only when the name already carries an ID or OD marker (the part of the
' location the BOM uses to distinguish the two physical instances). A name
' with only TE/LE is NOT considered to have ID/OD - it still needs splitting.
Private Function PullcoreNameHasIdOd(ByVal desc As String) As Boolean
    Dim loc As String
    loc = GetPullcoreLocationCode(desc)
    PullcoreNameHasIdOd = (loc = "ID" Or loc = "OD" Or _
                           loc = "IDTE" Or loc = "ODTE" Or _
                           loc = "IDLE" Or loc = "ODLE")
End Function

' Prepends "ID " or "OD " to a pullcore name, after stripping any leading
' ID/OD token so we never double it. "TE CAM" -> "ID TE CAM";
' "TE PULLCORE KEY" -> "OD TE PULLCORE KEY". Other words/casing are kept.
Private Function PullcorePrependIdOd(ByVal name As String, ByVal isId As Boolean) As String
    Dim work As String
    work = Trim(name)

    Dim up As String
    up = UCase(work)

    If Left(up, 3) = "ID " Then
        work = Trim(mid(work, 4))
    ElseIf Left(up, 3) = "OD " Then
        work = Trim(mid(work, 4))
    End If

    If isId Then
        PullcorePrependIdOd = "ID " & work
    Else
        PullcorePrependIdOd = "OD " & work
    End If
End Function

' Pullcore BOM rows that come in pairs (qty 2, or two rows sharing one name)
' but whose name has no ID/OD get split into an ID and an OD copy by assembly
' height, preserving the rest of the descriptive BOM name (TE/LE included).
' This replaces the old " (2)" suffix for these groups. Names that already
' carry ID/OD are authoritative and untouched; bare generic names are handled
' upstream by FinalizePullcoreCamIdOdNames.
Private Sub DisambiguatePullcoreDescriptiveGroupsByIdOd()
On Error GoTo ErrHandler

    If PullcoreMatchCount < 2 Then Exit Sub

    ' Group match indices by their (normalized) base name.
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = 1

    Dim i As Long
    Dim key As String

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).CadPartIndex > 0 Then
            If PullcoreNameHasIdOd(PullcoreMatches(i).Description) = False Then

                key = NormalizeText(PullcoreMatches(i).Description)

                If key <> "" Then
                    If groups.Exists(key) Then
                        groups(key) = groups(key) & "|" & CStr(i)
                    Else
                        groups(key) = CStr(i)
                    End If
                End If

            End If
        End If

    Next i

    Dim kk As Variant
    Dim parts2() As String
    Dim m As Long

    For Each kk In groups.keys

        parts2 = Split(CStr(groups(kk)), "|")

        If UBound(parts2) - LBound(parts2) + 1 >= 2 Then

            ' Order the group's members by height (higher = ID when
            ' PULLCORE_ID_IS_HIGHER), then label the top ID and the bottom OD.
            Dim idxArr() As Long
            Dim n As Long
            n = UBound(parts2) - LBound(parts2) + 1
            ReDim idxArr(1 To n)

            Dim T As Long
            For T = 1 To n
                idxArr(T) = CLng(parts2(LBound(parts2) + T - 1))
            Next T

            SortPullcoreMatchArrayByY idxArr, n

            ' SortPullcoreMatchArrayByY sorts DESCENDING by Y (highest first),
            ' and higher Y = ID, so element 1 is ID and element n is OD.
            Dim hiIdx As Long
            Dim loIdx As Long
            hiIdx = idxArr(1)
            loIdx = idxArr(n)

            Dim idMatch As Long
            Dim odMatch As Long

            If PULLCORE_ID_IS_HIGHER Then
                idMatch = hiIdx
                odMatch = loIdx
            Else
                idMatch = loIdx
                odMatch = hiIdx
            End If

            ApplyPullcoreIdOdName idMatch, True
            ApplyPullcoreIdOdName odMatch, False

            ' Any middle members (qty 3+, rare) keep their name and fall to the
            ' numeric disambiguator so nothing is silently lost.
            For T = 2 To n - 1
                LogLine "PULLCORE ID/OD split: extra member kept generic for '" & _
                        PullcoreMatches(idxArr(T)).Description & "' (qty>2)."
            Next T

        End If

    Next kk

    Exit Sub

ErrHandler:
    LogLine "DisambiguatePullcoreDescriptiveGroupsByIdOd error: " & Err.Description
End Sub

' Applies an ID/OD-prepended name to one match (both quoteName and
' Description) and logs the change.
Private Sub ApplyPullcoreIdOdName(ByVal matchIdx As Long, ByVal isId As Boolean)
On Error Resume Next

    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    Dim newName As String
    newName = PullcorePrependIdOd(PullcoreMatches(matchIdx).quoteName, isId)

    LogLine "PULLCORE ID/OD split: '" & PullcoreMatches(matchIdx).quoteName & _
            "' -> '" & newName & "'"

    PullcoreMatches(matchIdx).quoteName = newName
    PullcoreMatches(matchIdx).Description = newName
End Sub

Private Sub DisambiguatePullcoreQuoteNames()
On Error Resume Next

    If PullcoreMatchCount < 1 Then Exit Sub

    Dim i As Long
    Dim j As Long

    ' Capture the base (descriptive) name of every match FIRST, filling any
    ' blank from the BOM description. Done in its own pass so the de-dup
    ' comparison below is made against the original names, not ones we have
    ' already suffixed (otherwise the 3rd of three identical names would only
    ' see the 1st as a duplicate and collide with the 2nd at " (2)").
    Dim baseName() As String
    ReDim baseName(1 To PullcoreMatchCount)

    For i = 1 To PullcoreMatchCount
        baseName(i) = Trim(PullcoreMatches(i).quoteName)
        If baseName(i) = "" Then
            baseName(i) = CleanPullcoreDisplayName(PullcoreMatches(i).Description)
        End If
        PullcoreMatches(i).quoteName = baseName(i)
    Next i

    ' Append " (2)", " (3)", ... to the 2nd, 3rd, ... match sharing a name.
    For i = 1 To PullcoreMatchCount

        Dim dupSeen As Long
        dupSeen = 0

        For j = 1 To i - 1
            If NormalizeKey(baseName(j)) = NormalizeKey(baseName(i)) Then
                dupSeen = dupSeen + 1
            End If
        Next j

        If dupSeen > 0 Then
            PullcoreMatches(i).quoteName = baseName(i) & " (" & (dupSeen + 1) & ")"
            LogLine "Pullcore duplicate name resolved -> '" & PullcoreMatches(i).quoteName & "'"
        End If

    Next i
End Sub

' True when the BOM already names TE/LE/IDTE/ODTE etc. ? do not overwrite those.
' True when the BOM name already tells us WHERE the part goes - any ID/OD
' and/or TE/LE marker counts. Such names come straight from the BOM, are the
' ground truth, and must NEVER be overwritten by geometry-derived labels.
Private Function PullcoreNameEncodesLocation(ByVal desc As String) As Boolean
On Error Resume Next

    PullcoreNameEncodesLocation = False

    If GetPullcoreLocationCode(NormalizeText(desc)) <> "" Then
        PullcoreNameEncodesLocation = True
        Exit Function
    End If

    PullcoreNameEncodesLocation = PullcoreLocationNameIsDescriptive(desc)
End Function

Private Function PullcoreLocationNameIsDescriptive(ByVal desc As String) As Boolean
    Dim d As String
    Dim loc As String
    d = NormalizeText(desc)
    loc = GetPullcoreLocationCode(d)
    If loc = "IDTE" Or loc = "ODTE" Or loc = "IDLE" Or loc = "ODLE" Then
        PullcoreLocationNameIsDescriptive = True
        Exit Function
    End If
    If InStr(d, " TE ") > 0 Or InStr(d, " LE ") > 0 Then
        PullcoreLocationNameIsDescriptive = True
        Exit Function
    End If
    If Right(d, 3) = " TE" Or Right(d, 3) = " LE" Then
        PullcoreLocationNameIsDescriptive = True
    End If
End Function

Private Sub BlockDuplicatePullcoreCadMatches()
On Error GoTo ErrHandler

    If PullcoreMatchCount <= 1 Then Exit Sub

    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim cadIdx As Long
    Dim k As String
    Dim firstMatch As Long

    For i = 1 To PullcoreMatchCount

        cadIdx = PullcoreMatches(i).CadPartIndex

        If cadIdx > 0 Then

            k = CStr(cadIdx)

            If used.Exists(k) Then

                firstMatch = CLng(used(k))

                LogLine "PULLCORE duplicate CAD blocked. '" & _
                        PullcoreMatches(i).quoteName & "' tried to reuse CAD '" & _
                        parts(cadIdx).componentName & "', already used by '" & _
                        PullcoreMatches(firstMatch).quoteName & "'."

                PullcoreMatches(i).CadPartIndex = 0
                PullcoreMatches(i).Status = "NO CAD MATCH - duplicate CAD already used by " & _
                                            PullcoreMatches(firstMatch).quoteName

            Else

                used(k) = i

            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "BlockDuplicatePullcoreCadMatches error: " & Err.Description
End Sub

Private Sub SortPullcoreMatchArrayByY(ByRef matchArr() As Long, ByVal matchN As Long)
On Error GoTo ErrHandler

    If matchN <= 1 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    Dim yi As Double
    Dim yj As Double

    For i = 1 To matchN - 1

        For j = i + 1 To matchN

            yi = GetPullcoreMatchHeightValue(matchArr(i), "Y")
            yj = GetPullcoreMatchHeightValue(matchArr(j), "Y")

            ' Higher Y should come first because higher Y = ID.
            If yi < yj Then
                tmp = matchArr(i)
                matchArr(i) = matchArr(j)
                matchArr(j) = tmp
            End If

        Next j

    Next i

    Exit Sub

ErrHandler:
    LogLine "SortPullcoreMatchArrayByY error: " & Err.Description
End Sub

Private Sub FinalizePullcoreCamIdOdNames()
On Error GoTo ErrHandler

    If PullcoreMatchCount <= 0 Then Exit Sub

    PullcoreIdOdHeightAxisUsed = "Y"

    Dim camMatches() As Long
    Dim keyMatches() As Long
    Dim camN As Long
    Dim keyN As Long

    camN = 0
    keyN = 0

    ReDim camMatches(1 To PullcoreMatchCount)
    ReDim keyMatches(1 To PullcoreMatchCount)

    Dim usedCamCad As Object
    Dim usedKeyCad As Object

    Set usedCamCad = CreateObject("Scripting.Dictionary")
    Set usedKeyCad = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim cadIdx As Long
    Dim cadKey As String

    For i = 1 To PullcoreMatchCount

        cadIdx = PullcoreMatches(i).CadPartIndex

        If cadIdx > 0 And cadIdx <= PartCount Then

            ' Rows whose BOM name already says where they go (ID/OD, TE/LE)
            ' keep that name and stay out of the geometric ID/OD ranking;
            ' only genuinely generic rows are labeled by height.
            If PULLCORE_PRESERVE_DESCRIPTIVE_BOM_NAMES And _
               PullcoreNameEncodesLocation(PullcoreMatches(i).Description) Then
                LogLine "Pullcore ID/OD-by-height skipped for BOM-named row '" & _
                        PullcoreMatches(i).quoteName & "'"
            ElseIf parts(cadIdx).hasAsmCenter Then

                cadKey = CStr(cadIdx)

                If PullcoreMatches(i).isCam Then

                    If usedCamCad.Exists(cadKey) = False Then
                        camN = camN + 1
                        camMatches(camN) = i
                        usedCamCad(cadKey) = i
                    Else
                        LogLine "WARNING: Duplicate pullcore CAM CAD ignored for ID/OD naming: " & _
                                parts(cadIdx).componentName
                    End If

                Else

                    If usedKeyCad.Exists(cadKey) = False Then
                        keyN = keyN + 1
                        keyMatches(keyN) = i
                        usedKeyCad(cadKey) = i
                    Else
                        LogLine "WARNING: Duplicate pullcore KEY CAD ignored for ID/OD naming: " & _
                                parts(cadIdx).componentName
                    End If

                End If

            End If

        End If

    Next i

    If camN >= 2 Then

        SortPullcoreMatchArrayByY camMatches, camN

        Dim idCamMatch As Long
        Dim odCamMatch As Long

        idCamMatch = camMatches(1)
        odCamMatch = camMatches(camN)

        ForcePullcoreMatchLabel idCamMatch, "ID PULLCORE CAM"
        ForcePullcoreMatchLabel odCamMatch, "OD PULLCORE CAM"

        LogLine "Pullcore CAM ID/OD split by Y:"
        LogLine "  ID CAM = " & parts(PullcoreMatches(idCamMatch).CadPartIndex).componentName & _
                " Y=" & FormatNumberForCsv(parts(PullcoreMatches(idCamMatch).CadPartIndex).AsmCenterY)
        LogLine "  OD CAM = " & parts(PullcoreMatches(odCamMatch).CadPartIndex).componentName & _
                " Y=" & FormatNumberForCsv(parts(PullcoreMatches(odCamMatch).CadPartIndex).AsmCenterY)

        If keyN = 1 Then

            If PullcoreMatchDistanceSq(keyMatches(1), idCamMatch) <= _
               PullcoreMatchDistanceSq(keyMatches(1), odCamMatch) Then

                ForcePullcoreMatchLabel keyMatches(1), "ID PULLCORE KEY"

            Else

                ForcePullcoreMatchLabel keyMatches(1), "OD PULLCORE KEY"

            End If

        ElseIf keyN = 2 Then

            Dim k1 As Long
            Dim k2 As Long

            k1 = keyMatches(1)
            k2 = keyMatches(2)

            Dim optionA As Double
            Dim optionB As Double

            optionA = PullcoreMatchDistanceSq(k1, idCamMatch) + PullcoreMatchDistanceSq(k2, odCamMatch)
            optionB = PullcoreMatchDistanceSq(k2, idCamMatch) + PullcoreMatchDistanceSq(k1, odCamMatch)

            If optionA <= optionB Then
                ForcePullcoreMatchLabel k1, "ID PULLCORE KEY"
                ForcePullcoreMatchLabel k2, "OD PULLCORE KEY"
            Else
                ForcePullcoreMatchLabel k2, "ID PULLCORE KEY"
                ForcePullcoreMatchLabel k1, "OD PULLCORE KEY"
            End If

        ElseIf keyN > 2 Then

            Dim n As Long

            For n = 1 To keyN

                If PullcoreMatchDistanceSq(keyMatches(n), idCamMatch) <= _
                   PullcoreMatchDistanceSq(keyMatches(n), odCamMatch) Then

                    ForcePullcoreMatchLabel keyMatches(n), "ID PULLCORE KEY"

                Else

                    ForcePullcoreMatchLabel keyMatches(n), "OD PULLCORE KEY"

                End If

            Next n

        End If

    Else

        LogLine "Pullcore CAM ID/OD split skipped: need at least two CAM matches with Y centers."

        ' If cams were not available but two keys are, still split keys by Y.
        If keyN >= 2 Then

            SortPullcoreMatchArrayByY keyMatches, keyN

            ForcePullcoreMatchLabel keyMatches(1), "ID PULLCORE KEY"
            ForcePullcoreMatchLabel keyMatches(keyN), "OD PULLCORE KEY"

            LogLine "Pullcore KEY ID/OD split by Y without cams."

        End If

    End If

    Exit Sub

ErrHandler:
    LogLine "FinalizePullcoreCamIdOdNames error: " & Err.Description
End Sub

Private Sub ForcePullcoreMatchLabel(ByVal matchIdx As Long, ByVal newName As String)
On Error Resume Next
    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    ' THE BOM IS THE GROUND TRUTH. A name that already encodes a location
    ' (ID/OD and/or TE/LE) came from the BOM together with its dimensions
    ' and is never replaced by a geometry-derived generic label.
    If PULLCORE_PRESERVE_DESCRIPTIVE_BOM_NAMES Then
        If PullcoreNameEncodesLocation(PullcoreMatches(matchIdx).Description) Then
            LogLine "Pullcore keeping BOM name '" & _
                    PullcoreMatches(matchIdx).quoteName & "' (geometry label '" & _
                    newName & "' suppressed - BOM is authoritative)"
            Exit Sub
        End If
    End If

    PullcoreMatches(matchIdx).quoteName = newName
    PullcoreMatches(matchIdx).Description = newName
    LogLine "Pullcore export name finalized: '" & newName & "'"
End Sub

' Replace the trailing "Cam" in a pullcore cam name with "Key", preserving the
' location prefix and capitalisation of the rest. "OD TE Pullcore Cam" -> "OD
' TE Pullcore Key". If no "cam" token is present the name is returned as-is.
Private Function SwapCamForKey(ByVal name As String) As String
    Dim p As Long
    p = InStrRev(UCase(name), "CAM")

    If p > 0 Then
        SwapCamForKey = Left(name, p - 1) & "KEY" & mid(name, p + 3)
    Else
        SwapCamForKey = name
    End If
End Function

' Pullcore keys sit INSIDE the cams. A key is frequently the same size as a key
' in another location (e.g. an OD LE key and an OD TE key can share identical
' bounding-box dimensions), so the size matcher cannot tell them apart and
' mislabels them. The cams, however, are distinct and already correctly named
' from the BOM. This pass gives each matched key the location of the cam whose
' assembly center is closest to it (their center lines coincide because the key
' is nested in the cam), then names the key by swapping the cam's "Cam" for
' "Key". The cam's Description is used as the location source because it is the
' original BOM text and is never overwritten by the labelers.
Private Sub RelabelPullcoreKeysByNestedCam()
On Error GoTo ErrHandler

    If Not PULLCORE_NAME_KEYS_BY_NESTED_CAM Then Exit Sub
    If PullcoreMatchCount < 1 Then Exit Sub

    ' Collect cam matches that carry a usable assembly center.
    Dim camMatch() As Long
    Dim camN As Long
    camN = 0
    ReDim camMatch(1 To PullcoreMatchCount)

    Dim m As Long
    Dim cpi As Long
    For m = 1 To PullcoreMatchCount
        If PullcoreMatches(m).isCam Then
            cpi = PullcoreMatches(m).CadPartIndex
            If cpi > 0 And cpi <= PartCount Then
                If parts(cpi).hasAsmCenter Then
                    camN = camN + 1
                    camMatch(camN) = m
                End If
            End If
        End If
    Next m

    If camN = 0 Then
        LogLine "Pullcore nested-cam key naming skipped: no cams with centers."
        Exit Sub
    End If

    Dim k As Long
    Dim kpi As Long
    For k = 1 To PullcoreMatchCount

        If Not PullcoreMatches(k).isCam Then

            kpi = PullcoreMatches(k).CadPartIndex

            ' A key whose BOM name already encodes its location keeps that
            ' name - it is bound to the BOM dimensions the part was matched
            ' on. Nested-cam naming is only for generic key names.
            If PULLCORE_PRESERVE_DESCRIPTIVE_BOM_NAMES And _
               PullcoreNameEncodesLocation(PullcoreMatches(k).Description) Then

                LogLine "Pullcore key keeping BOM name '" & _
                        PullcoreMatches(k).quoteName & "' (nested-cam rename skipped)"

            ElseIf kpi > 0 And kpi <= PartCount Then
                If parts(kpi).hasAsmCenter Then

                    Dim kx As Double, ky As Double, kz As Double
                    kx = parts(kpi).AsmCenterX
                    ky = parts(kpi).AsmCenterY
                    kz = parts(kpi).AsmCenterZ

                    Dim bestM As Long
                    Dim bestDist As Double
                    Dim secondDist As Double
                    bestM = 0
                    bestDist = 1E+99
                    secondDist = 1E+99

                    Dim c As Long
                    Dim mi As Long
                    Dim ci As Long
                    Dim dX As Double, dY As Double, dZ As Double
                    Dim dist As Double

                    For c = 1 To camN
                        mi = camMatch(c)
                        ci = PullcoreMatches(mi).CadPartIndex

                        dX = parts(ci).AsmCenterX - kx
                        dY = parts(ci).AsmCenterY - ky
                        dZ = parts(ci).AsmCenterZ - kz
                        dist = Sqr(dX * dX + dY * dY + dZ * dZ)

                        If dist < bestDist Then
                            secondDist = bestDist
                            bestDist = dist
                            bestM = mi
                        ElseIf dist < secondDist Then
                            secondDist = dist
                        End If
                    Next c

                    If bestM > 0 Then

                        Dim camName As String
                        camName = PullcoreMatches(bestM).Description
                        If GetPullcoreLocationCode(camName) = "" Then
                            camName = PullcoreMatches(bestM).quoteName
                        End If

                        Dim newName As String
                        newName = SwapCamForKey(camName)

                        Dim ambNote As String
                        If bestDist <= 0.5 * secondDist Then
                            ambNote = ""
                        Else
                            ambNote = " (closest cam not clearly isolated - verify)"
                        End If

                        LogLine "Pullcore key '" & PullcoreMatches(k).quoteName & _
                                "' -> '" & newName & "'  (nested in cam '" & _
                                PullcoreMatches(bestM).Description & "', center dist=" & _
                                FormatNumberForCsv(bestDist) & ")" & ambNote

                        PullcoreMatches(k).quoteName = newName
                        PullcoreMatches(k).Description = newName

                    End If

                Else
                    LogLine "Pullcore key '" & PullcoreMatches(k).quoteName & _
                            "' has no assembly center; left as size-matched name."
                End If
            End If

        End If

    Next k

    Exit Sub

ErrHandler:
    LogLine "RelabelPullcoreKeysByNestedCam error: " & Err.Description
End Sub

Private Function PullcoreDimDistanceScore(ByVal cadL As Double, ByVal cadW As Double, ByVal cadT As Double, _
                                        ByRef b As BomInfo) As Double
    PullcoreDimDistanceScore = Abs(cadL - b.BomLength) * 3# _
                             + Abs(cadW - b.BomWidth) _
                             + Abs(cadT - b.BomThickness)
End Function

' Composite pair score: dims are the PRIMARY key (bucketed so measurement
' noise within one bucket counts as a tie), location is strictly the
' tie-breaker inside a bucket. Location can never overcome a whole bucket,
' so a fitted box that matches the BOM size always beats one that does not.
Private Function PullcorePairCompositeScore(ByVal dimScore As Double, _
                                            ByVal locScore As Double) As Double
    ' locScore range is -3..+5; shift to 0..8 and keep far below the 100-per-
    ' bucket step. dimScore is re-added at face value as the final decider.
    PullcorePairCompositeScore = Fix(dimScore / PULLCORE_DIM_SCORE_BUCKET) * 100# _
                               + (locScore + 3#) _
                               + dimScore
End Function


Private Sub MarkPullcoreCandidateSlotsUsed(ByRef candIdx() As Long, ByVal candN As Long, _
                                           ByRef candUsed() As Boolean, ByVal usedCadIdx As Long)
    Dim j As Long

    For j = 1 To candN
        If candIdx(j) = usedCadIdx Then candUsed(j) = True
    Next j
End Sub

Private Function GetPullcoreCandidateAxisValue(ByVal cadIdx As Long, ByVal axisName As String) As Double
    GetPullcoreCandidateAxisValue = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If parts(cadIdx).hasAsmCenter = False Then Exit Function

    If UCase(axisName) = "Z" Then
        GetPullcoreCandidateAxisValue = parts(cadIdx).AsmCenterZ
    Else
        GetPullcoreCandidateAxisValue = parts(cadIdx).AsmCenterY
    End If
End Function

' When BOM qty expands to 2+ identical rows, pick distinct CAD parts and rank
' them by assembly Y (higher Y = ID pullcore when PULLCORE_ID_IS_HIGHER).
Private Function TryMatchPullcoreRowsByYRank(ByRef rIdx() As Long, ByVal rowN As Long, _
                                             ByRef candIdx() As Long, _
                                             ByRef candL() As Double, _
                                             ByRef candW() As Double, _
                                             ByRef candT() As Double, _
                                             ByVal candN As Long, _
                                             ByRef candLoc() As String, _
                                             ByRef rowUsed() As Boolean, _
                                             ByRef candUsed() As Boolean) As Boolean
On Error GoTo ErrHandler

    TryMatchPullcoreRowsByYRank = False

    If rowN < 2 Then Exit Function
    If candN < rowN Then Exit Function

    Dim refBom As BomInfo
    refBom = BomRows(rIdx(1))

    Dim uniqCad() As Long
    Dim uniqScore() As Double
    Dim uniqLoc() As String
    Dim uniqN As Long

    uniqN = 0
    ReDim uniqCad(1 To candN)
    ReDim uniqScore(1 To candN)
    ReDim uniqLoc(1 To candN)

    Dim j As Long
    Dim k As Long
    Dim already As Boolean

    For j = 1 To candN

        already = False

        For k = 1 To uniqN
            If uniqCad(k) = candIdx(j) Then
                already = True
                Exit For
            End If
        Next k

        If already = False Then
            uniqN = uniqN + 1
            uniqCad(uniqN) = candIdx(j)
            uniqScore(uniqN) = PullcoreDimDistanceScore(candL(j), candW(j), candT(j), refBom)
            uniqLoc(uniqN) = candLoc(j)
        End If

    Next j

    If uniqN < rowN Then Exit Function

    ' Pick the rowN best dimension fits among unique CAD parts.
    Dim pickCad() As Long
    Dim pickScore() As Double
    Dim pickLoc() As String
    Dim pickN As Long

    pickN = 0
    ReDim pickCad(1 To rowN)
    ReDim pickScore(1 To rowN)
    ReDim pickLoc(1 To rowN)

    Dim usedUniq() As Boolean
    ReDim usedUniq(1 To uniqN)

    Dim pickStep As Long
    Dim bestU As Long
    Dim bestScore As Double

    For pickStep = 1 To rowN

        bestU = 0
        bestScore = 1E+99

        For k = 1 To uniqN
            If usedUniq(k) = False Then
                If uniqScore(k) < bestScore Then
                    bestScore = uniqScore(k)
                    bestU = k
                End If
            End If
        Next k

        If bestU = 0 Then Exit Function

        pickN = pickN + 1
        pickCad(pickN) = uniqCad(bestU)
        pickScore(pickN) = uniqScore(bestU)
        pickLoc(pickN) = uniqLoc(bestU)
        usedUniq(bestU) = True

    Next pickStep

    Dim axisName As String
    axisName = UCase(Trim(PULLCORE_ID_OD_HEIGHT_AXIS))
    If axisName <> "Y" And axisName <> "Z" Then axisName = "Y"

    ' Sort selected parts by Y/Z so ID (higher) and OD (lower) are distinct.
    Dim swapped As Boolean
    Dim needSwap As Boolean
    Dim tmpCad As Long
    Dim tmpScore As Double
    Dim tmpLoc As String
    Dim va As Double
    Dim vB As Double

    Do
        swapped = False

        For k = 1 To pickN - 1
            va = GetPullcoreCandidateAxisValue(pickCad(k), axisName)
            vB = GetPullcoreCandidateAxisValue(pickCad(k + 1), axisName)

            needSwap = False
            If PULLCORE_ID_IS_HIGHER Then
                If va < vB Then needSwap = True
            Else
                If va > vB Then needSwap = True
            End If

            If needSwap Then
                tmpCad = pickCad(k): pickCad(k) = pickCad(k + 1): pickCad(k + 1) = tmpCad
                tmpScore = pickScore(k): pickScore(k) = pickScore(k + 1): pickScore(k + 1) = tmpScore
                tmpLoc = pickLoc(k): pickLoc(k) = pickLoc(k + 1): pickLoc(k + 1) = tmpLoc
                swapped = True
            End If
        Next k
    Loop While swapped

    LogLine "PULLCORE qty-" & rowN & " Y-ranked match on axis " & axisName & ":"

    For k = 1 To pickN
        If pickCad(k) > 0 And pickCad(k) <= PartCount Then
            LogLine "  rank " & k & ": " & parts(pickCad(k)).componentName & _
                    " Y=" & FormatNumberForCsv(parts(pickCad(k)).AsmCenterY) & _
                    " score=" & FormatNumberForCsv(pickScore(k))
        End If
    Next k

    ' Fetch the FITTED (best-fit) box of every picked part up front.
    Dim pkL() As Double
    Dim pkW() As Double
    Dim pkT() As Double

    ReDim pkL(1 To pickN)
    ReDim pkW(1 To pickN)
    ReDim pkT(1 To pickN)

    For k = 1 To pickN

        pkL(k) = parts(pickCad(k)).Length
        pkW(k) = parts(pickCad(k)).Width
        pkT(k) = parts(pickCad(k)).Thickness

        If USE_PULLCORE_BEST_FIT_BBOX Then
            If TryGetPullcoreBestFitDims(pickCad(k), pkL(k), pkW(k), pkT(k)) = False Then
                pkL(k) = parts(pickCad(k)).Length
                pkW(k) = parts(pickCad(k)).Width
                pkT(k) = parts(pickCad(k)).Thickness
            End If
        End If

    Next k

    ' Height ranking can only tell apart rows that are the SAME stock size
    ' (classic qty-2 identical cams). When the rows differ in size, the name
    ' must come from the fitted-box match, not from rank position.
    Dim rowsSameSize As Boolean
    rowsSameSize = True

    For k = 2 To rowN
        If Abs(BomRows(rIdx(k)).BomLength - BomRows(rIdx(1)).BomLength) > SAME_SIZE_PAIR_TOL _
        Or Abs(BomRows(rIdx(k)).BomWidth - BomRows(rIdx(1)).BomWidth) > SAME_SIZE_PAIR_TOL _
        Or Abs(BomRows(rIdx(k)).BomThickness - BomRows(rIdx(1)).BomThickness) > SAME_SIZE_PAIR_TOL Then
            rowsSameSize = False
            Exit For
        End If
    Next k

    If rowsSameSize Then

        ' Identical rows: keep height-rank assignment (ID high / OD low).
        For k = 1 To pickN
            If rowUsed(k) = False Then
                AddPullcoreMatchRow BomRows(rIdx(k)), pickCad(k), pkL(k), pkW(k), pkT(k)
                parts(pickCad(k)).UsedForBomMatch = True
                rowUsed(k) = True
                MarkPullcoreCandidateSlotsUsed candIdx, candN, candUsed, pickCad(k)
            End If
        Next k

    Else

        ' Rows differ in stock size: pair each BOM row with the picked part
        ' whose FITTED box matches it best (location only nudges ties).
        LogLine "PULLCORE rows differ in size: assigning names by fitted-box match, not rank."

        Dim pickUsed() As Boolean
        ReDim pickUsed(1 To pickN)

        Dim stepN As Long
        Dim m As Long
        Dim bestRow As Long
        Dim bestPick As Long
        Dim bestD As Double
        Dim dd As Double

        For stepN = 1 To pickN

            bestRow = 0
            bestPick = 0
            bestD = 1E+99

            For k = 1 To rowN
                If rowUsed(k) = False Then
                    For m = 1 To pickN
                        If pickUsed(m) = False Then

                            dd = PullcoreDimDistanceScore(pkL(m), pkW(m), pkT(m), BomRows(rIdx(k)))

                            ' DIMS DECIDE: location is a tie-breaker only and
                            ' can never flip two different stock sizes.
                            If PULLCORE_USE_LOCATION_AWARE_MATCH Then
                                dd = PullcorePairCompositeScore(dd, _
                                     PullcoreLocationScore( _
                                         GetPullcoreLocationCode(BomRows(rIdx(k)).Description), _
                                         pickLoc(m)))
                            End If

                            If dd < bestD Then
                                bestD = dd
                                bestRow = k
                                bestPick = m
                            End If

                        End If
                    Next m
                End If
            Next k

            If bestRow = 0 Or bestPick = 0 Then Exit For

            LogLine "PULLCORE fitted-box assign: '" & BomRows(rIdx(bestRow)).Description & _
                    "' -> " & parts(pickCad(bestPick)).componentName & _
                    " fitted=" & FormatNumberForCsv(pkL(bestPick)) & "/" & _
                    FormatNumberForCsv(pkW(bestPick)) & "/" & _
                    FormatNumberForCsv(pkT(bestPick)) & _
                    " score=" & FormatNumberForCsv(bestD)

            AddPullcoreMatchRow BomRows(rIdx(bestRow)), pickCad(bestPick), _
                                pkL(bestPick), pkW(bestPick), pkT(bestPick)

            parts(pickCad(bestPick)).UsedForBomMatch = True
            rowUsed(bestRow) = True
            pickUsed(bestPick) = True
            MarkPullcoreCandidateSlotsUsed candIdx, candN, candUsed, pickCad(bestPick)

        Next stepN

    End If

    TryMatchPullcoreRowsByYRank = True

    Exit Function

ErrHandler:
    LogLine "TryMatchPullcoreRowsByYRank error: " & Err.Description
    TryMatchPullcoreRowsByYRank = False
End Function

' True when the (sorted) fitted box matches at least one of this class's BOM
' stock sizes within the pullcore tolerances scaled by
' PULLCORE_FITTED_VERIFY_TOL_SCALE. A fitted box should sit near-exactly on
' a stock size; anything else (pot blocks, risers, plates) is not a pullcore.
Private Function FittedDimsMatchAnyPullcoreRow(ByVal fL As Double, _
                                               ByVal fW As Double, _
                                               ByVal fT As Double, _
                                               ByRef bomRowIdx() As Long, _
                                               ByVal rowN As Long) As Boolean
On Error GoTo ErrHandler

    FittedDimsMatchAnyPullcoreRow = False

    Dim sL As Double, sW As Double, st As Double
    SortThreeDimensions fL, fW, fT, sL, sW, st

    Dim tolL As Double, tolW As Double, tolT As Double
    tolL = PULLCORE_L_TOL * PULLCORE_FITTED_VERIFY_TOL_SCALE
    tolW = PULLCORE_W_TOL * PULLCORE_FITTED_VERIFY_TOL_SCALE
    tolT = PULLCORE_T_TOL * PULLCORE_FITTED_VERIFY_TOL_SCALE

    Dim j As Long
    Dim bL As Double, bw As Double, bT As Double

    For j = 1 To rowN

        If BomRows(bomRowIdx(j)).hasDims Then

            SortThreeDimensions BomRows(bomRowIdx(j)).BomLength, _
                                BomRows(bomRowIdx(j)).BomWidth, _
                                BomRows(bomRowIdx(j)).BomThickness, bL, bw, bT

            If Abs(sL - bL) <= tolL And _
               Abs(sW - bw) <= tolW And _
               Abs(st - bT) <= tolT Then
                FittedDimsMatchAnyPullcoreRow = True
                Exit Function
            End If

        End If

    Next j

    Exit Function

ErrHandler:
    FittedDimsMatchAnyPullcoreRow = False
End Function

Private Sub AssignPullcoreClass(ByRef bomRowIdx() As Long, ByVal rowN As Long, ByVal wantCam As Boolean)
On Error GoTo ErrHandler

    Dim candIdx() As Long
    Dim candL() As Double
    Dim candW() As Double
    Dim candT() As Double
    Dim candN As Long

    candN = 0

    ReDim candIdx(1 To 1)
    ReDim candL(1 To 1)
    ReDim candW(1 To 1)
    ReDim candT(1 To 1)

    Dim i As Long
    Dim j As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim isCand As Boolean
            isCand = False

            For j = 1 To rowN
                If IsRoughPullcoreCandidateForBom(i, BomRows(bomRowIdx(j))) Then
                    isCand = True
                    Exit For
                End If
            Next j

            If isCand Then

                Dim bfL As Double
                Dim bfW As Double
                Dim bfT As Double
                Dim haveFitted As Boolean

                bfL = parts(i).Length
                bfW = parts(i).Width
                bfT = parts(i).Thickness
                haveFitted = False

                If USE_PULLCORE_BEST_FIT_BBOX Then
                    ' Opens the part individually (component doc or part file)
                    ' and computes the fitted box from its body geometry.
                    haveFitted = TryGetPullcoreBestFitDims(i, bfL, bfW, bfT)
                    If haveFitted = False Then
                        bfL = parts(i).Length
                        bfW = parts(i).Width
                        bfT = parts(i).Thickness
                    End If
                End If

                ' STRICT GATE: the rough regular-box screen only nominates.
                ' To stay a candidate the part must produce a trusted FITTED
                ' box, and that fitted box must match one of this class's BOM
                ' stock sizes. Never keep a part on regular-box guesses.
                If USE_PULLCORE_BEST_FIT_BBOX And PULLCORE_REQUIRE_FITTED_BBOX Then

                    If haveFitted = False Then
                        LogLine "PULLCORE candidate REJECTED (no trusted fitted bbox): " & _
                                parts(i).componentName
                        isCand = False
                    ElseIf FittedDimsMatchAnyPullcoreRow(bfL, bfW, bfT, bomRowIdx, rowN) = False Then
                        LogLine "PULLCORE candidate REJECTED (fitted bbox " & _
                                FormatNumberForCsv(bfL) & "/" & FormatNumberForCsv(bfW) & "/" & _
                                FormatNumberForCsv(bfT) & " matches no BOM stock size): " & _
                                parts(i).componentName
                        isCand = False
                    End If

                End If

                If isCand Then

                    candN = candN + 1

                    ReDim Preserve candIdx(1 To candN)
                    ReDim Preserve candL(1 To candN)
                    ReDim Preserve candW(1 To candN)
                    ReDim Preserve candT(1 To candN)

                    candIdx(candN) = i
                    candL(candN) = bfL
                    candW(candN) = bfW
                    candT(candN) = bfT

                    LogLine "PULLCORE candidate ACCEPTED: " & parts(i).componentName & _
                            " fitted=" & FormatNumberForCsv(bfL) & "/" & _
                            FormatNumberForCsv(bfW) & "/" & FormatNumberForCsv(bfT)

                End If

            End If

        End If

    Next i

    ' Too few candidates? The rough regular-box screen cannot predict every
    ' orientation. Sweep all remaining plausible parts: open each one
    ' individually, compute the fitted box, nominate on a fitted-box match.
    If PULLCORE_FITTED_SWEEP_FALLBACK And USE_PULLCORE_BEST_FIT_BBOX And candN < rowN Then

        LogLine "PULLCORE " & IIf(wantCam, "CAM", "KEY") & ": only " & candN & _
                " rough candidates for " & rowN & " rows - fitted-box sweep of remaining parts."

        Dim swfL As Double
        Dim swfW As Double
        Dim swfT As Double
        Dim alreadyCand As Boolean
        Dim kk As Long

        For i = 1 To PartCount

            If parts(i).UsedForBomMatch = False And parts(i).BBoxVolume <= 200# Then

                alreadyCand = False
                For kk = 1 To candN
                    If candIdx(kk) = i Then
                        alreadyCand = True
                        Exit For
                    End If
                Next kk

                If alreadyCand = False Then
                    If PartVolumePlausibleForPullcoreRows(i, bomRowIdx, rowN) Then
                        If TryGetPullcoreBestFitDims(i, swfL, swfW, swfT) Then
                            If FittedDimsMatchAnyPullcoreRow(swfL, swfW, swfT, bomRowIdx, rowN) Then

                                candN = candN + 1

                                ReDim Preserve candIdx(1 To candN)
                                ReDim Preserve candL(1 To candN)
                                ReDim Preserve candW(1 To candN)
                                ReDim Preserve candT(1 To candN)

                                candIdx(candN) = i
                                candL(candN) = swfL
                                candW(candN) = swfW
                                candT(candN) = swfT

                                LogLine "PULLCORE fitted-sweep candidate ACCEPTED: " & _
                                        parts(i).componentName & " fitted=" & _
                                        FormatNumberForCsv(swfL) & "/" & _
                                        FormatNumberForCsv(swfW) & "/" & FormatNumberForCsv(swfT)

                            End If
                        End If
                    End If
                End If

            End If

        Next i

    End If

    LogLine "Pullcore " & IIf(wantCam, "CAM", "KEY") & " candidates found: " & candN & " for " & rowN & " BOM rows"

    If candN = 0 Then
        For j = 1 To rowN
            LogLine "PULLCORE UNMATCHED (no candidate): " & BomRows(bomRowIdx(j)).Description
            AddPullcoreMatchRow BomRows(bomRowIdx(j)), 0, 0, 0, 0
        Next j
        Exit Sub
    End If

    Dim rIdx() As Long
    ReDim rIdx(1 To rowN)

    For i = 1 To rowN
        rIdx(i) = bomRowIdx(i)
    Next i

    Dim rowUsed() As Boolean
    Dim candUsed() As Boolean

    ReDim rowUsed(1 To rowN)
    ReDim candUsed(1 To candN)

    ' Spatial ID/OD code per candidate (derived from height in the assembly).
    ' Used only to nudge same-size matches onto the physically correct side.
    Dim candLoc() As String
    ReDim candLoc(1 To candN)
    ComputePullcoreCandidateSideCodes candIdx, candN, candLoc

    If PULLCORE_USE_LOCATION_AWARE_MATCH Then
        For j = 1 To candN
            If candLoc(j) <> "" And candIdx(j) > 0 Then
                LogLine "PULLCORE side-by-Y: " & parts(candIdx(j)).componentName & _
                        " Y=" & FormatNumberForCsv(parts(candIdx(j)).AsmCenterY) & _
                        " -> " & candLoc(j)
            End If
        Next j
    End If

    Dim matchedByYRank As Boolean
    matchedByYRank = False

    If rowN >= 2 And candN >= rowN Then
        matchedByYRank = TryMatchPullcoreRowsByYRank(rIdx, rowN, candIdx, candL, candW, candT, _
                                                     candN, candLoc, rowUsed, candUsed)
    End If

    If matchedByYRank = False Then

        Dim pairCount As Long

        pairCount = rowN
        If candN < pairCount Then pairCount = candN

        Dim pairStep As Long

        For pairStep = 1 To pairCount

            Dim bestR As Long
            Dim bestC As Long
            Dim bestDist As Double

            bestR = 0
            bestC = 0
            bestDist = 1E+99

            For i = 1 To rowN

                If rowUsed(i) = False Then

                    For j = 1 To candN

                        If candUsed(j) = False Then

                            Dim d As Double

                            d = PullcoreDimDistanceScore(candL(j), candW(j), candT(j), BomRows(rIdx(i)))

                            ' DIMS DECIDE: the fitted box must match the BOM
                            ' size. Location only breaks ties between rows of
                            ' the same stock size - it can never flip two
                            ' different sizes onto each other's rows.
                            If PULLCORE_USE_LOCATION_AWARE_MATCH Then
                                d = PullcorePairCompositeScore(d, _
                                    PullcoreLocationScore( _
                                        GetPullcoreLocationCode(BomRows(rIdx(i)).Description), _
                                        candLoc(j)))
                            End If

                            If d < bestDist Then
                                bestDist = d
                                bestR = i
                                bestC = j
                            End If

                        End If

                    Next j

                End If

            Next i

            If bestR = 0 Or bestC = 0 Then Exit For

            AddPullcoreMatchRow BomRows(rIdx(bestR)), candIdx(bestC), _
                                candL(bestC), candW(bestC), candT(bestC)

            parts(candIdx(bestC)).UsedForBomMatch = True

            rowUsed(bestR) = True
            MarkPullcoreCandidateSlotsUsed candIdx, candN, candUsed, candIdx(bestC)

        Next pairStep

    End If

    For i = 1 To rowN
        If rowUsed(i) = False Then
            LogLine "PULLCORE UNMATCHED (more BOM rows than candidates): " & BomRows(rIdx(i)).Description
            AddPullcoreMatchRow BomRows(rIdx(i)), 0, 0, 0, 0
        End If
    Next i

    Exit Sub

ErrHandler:
    LogLine "AssignPullcoreClass error: " & Err.Description
End Sub

' Assigns each candidate an "ID" / "OD" / "" side code based on its height in
' the assembly. The height axis is the configured one, or auto-detected as the
' axis (Y or Z) with the larger spread across the candidates. Candidates whose
' centre sits above the mid-line are tagged high, below it low; high/low maps
' to ID/OD per PULLCORE_ID_IS_HIGHER. Anything without a usable centre, or when
' the spread is negligible, is left blank so it contributes no bias.
Private Sub ComputePullcoreCandidateSideCodes(ByRef candIdx() As Long, _
                                              ByVal candN As Long, _
                                              ByRef candLoc() As String)
On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To candN
        candLoc(i) = ""
    Next i

    If candN < 2 Then Exit Sub
    If Not PULLCORE_USE_LOCATION_AWARE_MATCH Then Exit Sub

    Dim haveAny As Boolean
    haveAny = False

    Dim yMin As Double, yMax As Double
    Dim zMin As Double, zMax As Double

    yMin = 1E+99: yMax = -1E+99
    zMin = 1E+99: zMax = -1E+99

    Dim ci As Long
    For i = 1 To candN
        ci = candIdx(i)
        If ci > 0 And ci <= PartCount Then
            If parts(ci).hasAsmCenter Then
                haveAny = True
                If parts(ci).AsmCenterY < yMin Then yMin = parts(ci).AsmCenterY
                If parts(ci).AsmCenterY > yMax Then yMax = parts(ci).AsmCenterY
                If parts(ci).AsmCenterZ < zMin Then zMin = parts(ci).AsmCenterZ
                If parts(ci).AsmCenterZ > zMax Then zMax = parts(ci).AsmCenterZ
            End If
        End If
    Next i

    If Not haveAny Then Exit Sub

    Dim axis As String
    axis = UCase(Trim(PULLCORE_ID_OD_HEIGHT_AXIS))

    Dim ySpread As Double
    Dim zSpread As Double
    ySpread = yMax - yMin
    zSpread = zMax - zMin

    If axis <> "Y" And axis <> "Z" Then
        If zSpread > ySpread Then axis = "Z" Else axis = "Y"
    End If

    Dim spread As Double
    Dim mid As Double

    If axis = "Z" Then
        spread = zSpread
        mid = (zMin + zMax) / 2#
    Else
        spread = ySpread
        mid = (yMin + yMax) / 2#
    End If

    ' Negligible separation -> do not guess.
    If spread < 0.05 Then Exit Sub

    Dim hiCode As String
    Dim loCode As String

    If PULLCORE_ID_IS_HIGHER Then
        hiCode = "ID": loCode = "OD"
    Else
        hiCode = "OD": loCode = "ID"
    End If

    Dim v As Double
    For i = 1 To candN
        ci = candIdx(i)
        If ci > 0 And ci <= PartCount Then
            If parts(ci).hasAsmCenter Then
                If axis = "Z" Then v = parts(ci).AsmCenterZ Else v = parts(ci).AsmCenterY
                If v >= mid Then candLoc(i) = hiCode Else candLoc(i) = loCode
            End If
        End If
    Next i

    LogLine "Pullcore candidate side codes computed on axis " & axis & _
            " (spread=" & FormatNumberForCsv(spread) & ", mid=" & FormatNumberForCsv(mid) & ")"

    Exit Sub

ErrHandler:
    LogLine "ComputePullcoreCandidateSideCodes error: " & Err.Description
End Sub

Private Function IsRoughPullcoreCandidateForBom(ByVal cadIdx As Long, ByRef b As BomInfo) As Boolean
On Error GoTo ErrHandler

    IsRoughPullcoreCandidateForBom = False

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If parts(cadIdx).UsedForBomMatch Then Exit Function
    If b.hasDims = False Then Exit Function

    Dim cadL As Double
    Dim cadW As Double
    Dim cadT As Double

    cadL = parts(cadIdx).Length
    cadW = parts(cadIdx).Width
    cadT = parts(cadIdx).Thickness

    If parts(cadIdx).BBoxVolume > 200# Then Exit Function
    If cadT < 0.4 Then Exit Function

    ' --- 1) Fast direct axis-aligned check (original behavior) ---
    If Abs(cadL - b.BomLength) <= PULLCORE_L_TOL And _
       Abs(cadW - b.BomWidth) <= PULLCORE_W_TOL And _
       Abs(cadT - b.BomThickness) <= PULLCORE_T_TOL Then
        IsRoughPullcoreCandidateForBom = True
        Exit Function
    End If

    ' --- 2) Orientation-agnostic check: compare sorted dim triples ---
    Dim sL As Double, sW As Double, st As Double
    Dim bbL As Double, bbW As Double, bbT As Double

    SortThreeDimensions cadL, cadW, cadT, sL, sW, st
    SortThreeDimensions b.BomLength, b.BomWidth, b.BomThickness, bbL, bbW, bbT

    If Abs(sL - bbL) <= PULLCORE_L_TOL And _
       Abs(sW - bbW) <= PULLCORE_W_TOL And _
       Abs(st - bbT) <= PULLCORE_T_TOL Then
        IsRoughPullcoreCandidateForBom = True
        Exit Function
    End If

    ' --- 3) Angle-aware check: an angled pullcore inflates its axis-aligned
    '        bbox, so project the BOM block through known pullcore angles
    '        (detected on already-matched siblings, then the fallback) and
    '        compare to the part bbox. Final pairing still scores against
    '        the true best-fit bbox, so this only opens the door. ---
    If PULLCORE_ANGLE_AWARE_CANDIDATES Then
        If BomMatchesPartAabbAtKnownPullcoreAngle(bbL, bbW, bbT, sL, sW, st, _
                                                  parts(cadIdx).componentName) Then
            IsRoughPullcoreCandidateForBom = True
            Exit Function
        End If
    End If

    Exit Function

ErrHandler:
    IsRoughPullcoreCandidateForBom = False
End Function

' Store an angle measured on a successfully matched pullcore so the candidate
' filter can use it to recognize the sibling parts (same pull direction =>
' same angle). Keeps a small unique pipe-separated list.
Private Sub RememberDetectedPullcoreAngle(ByVal angDeg As Double)
On Error Resume Next

    angDeg = Abs(angDeg)

    If angDeg < PULLCORE_STRAIGHTEN_MIN_DEG Then Exit Sub
    If angDeg > 88# Then Exit Sub

    Dim token As String
    token = Format(angDeg, "0.0")

    If InStr("|" & DetectedPullcoreAnglesCsv & "|", "|" & token & "|") > 0 Then Exit Sub

    If DetectedPullcoreAnglesCsv = "" Then
        DetectedPullcoreAnglesCsv = token
    Else
        DetectedPullcoreAnglesCsv = DetectedPullcoreAnglesCsv & "|" & token
    End If

    LogLine "PULLCORE remembered detected angle " & token & _
            " deg for angle-aware sibling matching"
End Sub

' True when rotating the BOM stock block (sorted dims bL/bW/bT) by a known
' pullcore angle about one of its axes reproduces the part's sorted
' axis-aligned bbox (sL/sW/sT) within PULLCORE_ANGLE_AABB_TOL.
' Angle order tried: angles detected on matched pullcores this job,
' PULLCORE_FALLBACK_ANGLE_DEG, then an optional free angle search.
Private Function BomMatchesPartAabbAtKnownPullcoreAngle(ByVal bL As Double, _
                                                        ByVal bw As Double, _
                                                        ByVal bT As Double, _
                                                        ByVal sL As Double, _
                                                        ByVal sW As Double, _
                                                        ByVal st As Double, _
                                                        ByVal partNameForLog As String) As Boolean
On Error GoTo ErrHandler

    BomMatchesPartAabbAtKnownPullcoreAngle = False

    If bL <= 0 Or bw <= 0 Or bT <= 0 Then Exit Function
    If sL <= 0 Or sW <= 0 Or st <= 0 Then Exit Function

    ' Rotation never shrinks the bbox: sanity-check the volume ratio so we
    ' do not pull in obviously unrelated parts.
    Dim ratio As Double
    ratio = (sL * sW * st) / (bL * bw * bT)

    If ratio < PULLCORE_ANGLE_VOL_RATIO_MIN Then Exit Function
    If ratio > PULLCORE_ANGLE_VOL_RATIO_MAX Then Exit Function

    Dim angleListCsv As String
    angleListCsv = DetectedPullcoreAnglesCsv

    If angleListCsv = "" Then
        angleListCsv = Format(PULLCORE_FALLBACK_ANGLE_DEG, "0.0")
    Else
        angleListCsv = angleListCsv & "|" & Format(PULLCORE_FALLBACK_ANGLE_DEG, "0.0")
    End If

    Dim angles() As String
    angles = Split(angleListCsv, "|")

    Dim k As Long
    Dim angDeg As Double

    For k = LBound(angles) To UBound(angles)
        angDeg = val(angles(k))
        If angDeg > 0.5 And angDeg < 89.5 Then
            If RotatedBlockMatchesAabb(bL, bw, bT, sL, sW, st, angDeg) Then
                LogLine "PULLCORE angle-aware candidate: '" & partNameForLog & _
                        "' bbox explained by ~" & Format(angDeg, "0.0") & " deg rotation"
                BomMatchesPartAabbAtKnownPullcoreAngle = True
                Exit Function
            End If
        End If
    Next k

    ' Last resort: search for ANY angle that explains the bbox, then verify.
    If PULLCORE_ANGLE_INFER_ANY_ANGLE Then
        angDeg = Abs(EstimatePullcoreAngleFromFittedAndOriginalDeg(bL, bw, bT, sL, sW, st))
        If angDeg >= PULLCORE_STRAIGHTEN_MIN_DEG And angDeg <= 88# Then
            If RotatedBlockMatchesAabb(bL, bw, bT, sL, sW, st, angDeg) Then
                LogLine "PULLCORE angle-aware candidate (free search): '" & partNameForLog & _
                        "' bbox explained by ~" & Format(angDeg, "0.0") & " deg rotation"
                BomMatchesPartAabbAtKnownPullcoreAngle = True
                Exit Function
            End If
        End If
    End If

    Exit Function

ErrHandler:
    BomMatchesPartAabbAtKnownPullcoreAngle = False
End Function

' Rotate the block (bL/bW/bT) by angDeg about each of its three axes in turn
' and compare the resulting axis-aligned bbox (sorted) to the part's sorted
' bbox dims. One axis stays unchanged; the other two grow by the classic
' projection a*cos + b*sin / a*sin + b*cos.
Private Function RotatedBlockMatchesAabb(ByVal bL As Double, _
                                         ByVal bw As Double, _
                                         ByVal bT As Double, _
                                         ByVal sL As Double, _
                                         ByVal sW As Double, _
                                         ByVal st As Double, _
                                         ByVal angDeg As Double) As Boolean
On Error GoTo ErrHandler

    RotatedBlockMatchesAabb = False

    Dim rad As Double
    rad = angDeg * PI_VALUE / 180#

    Dim dims(1 To 3) As Double
    dims(1) = bL
    dims(2) = bw
    dims(3) = bT

    Dim keepAxis As Long
    Dim i As Long
    Dim c As Long
    Dim p As Double
    Dim q As Double
    Dim pA As Double
    Dim pB As Double
    Dim eL As Double
    Dim eW As Double
    Dim eT As Double

    For keepAxis = 1 To 3

        c = 0
        p = 0#
        q = 0#

        For i = 1 To 3
            If i <> keepAxis Then
                c = c + 1
                If c = 1 Then p = dims(i) Else q = dims(i)
            End If
        Next i

        pA = p * Cos(rad) + q * Sin(rad)
        pB = p * Sin(rad) + q * Cos(rad)

        SortThreeDimensions dims(keepAxis), pA, pB, eL, eW, eT

        If Abs(eL - sL) <= PULLCORE_ANGLE_AABB_TOL And _
           Abs(eW - sW) <= PULLCORE_ANGLE_AABB_TOL And _
           Abs(eT - st) <= PULLCORE_ANGLE_AABB_TOL Then
            RotatedBlockMatchesAabb = True
            Exit Function
        End If

    Next keepAxis

    Exit Function

ErrHandler:
    RotatedBlockMatchesAabb = False
End Function


' ============================================================
' MATCHING HELPERS / EXPORT ROWS
' ============================================================

' Renames alternate BOM holder names to the canonical ID/OD HOLDER quote so
' they categorize and export as holders (not as their literal "Mold Base"
' text, and never as pullcore). Other names pass through unchanged.
Private Function CanonicalHolderQuoteName(ByVal quoteName As String) As String
    CanonicalHolderQuoteName = quoteName

    Dim d As String
    d = NormalizeText(quoteName)

    If InStr(d, "MOLD BASE") > 0 Or InStr(d, "MOLDBASE") > 0 Then
        If InStr(d, "OD") > 0 Then
            CanonicalHolderQuoteName = "OD HOLDER"
        ElseIf InStr(d, "ID") > 0 Then
            CanonicalHolderQuoteName = "ID HOLDER"
        End If
    End If
End Function

Private Sub AddExportRow(ByRef b As BomInfo, ByVal cadIdx As Long)
On Error GoTo ErrHandler

    If ShouldSkipExportQuoteName(b.quoteName) Then
        LogLine "Export quote skipped (package-handled): " & b.quoteName
        If cadIdx > 0 Then RecordSpecialBomCadMatch b, cadIdx
        Exit Sub
    End If

    Dim isPyropel As Boolean
    isPyropel = (InStr(NormalizeText(b.material), "PYROPEL") > 0 Or InStr(NormalizeText(b.Description), "PYROPEL") > 0)

    If ONLY_INCLUDE_ALLOWED_STEEL_BOM_ITEMS Then
        If IsAllowedSteelType(b.material) = False Then
            If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
                LogLine "Item skipped, material not in " & ALLOWED_STEEL_TYPES & ": " & _
                        b.Description & " (" & b.material & ")"
                Exit Sub
            End If
        End If
    End If

    ExportCount = ExportCount + 1
    ReDim Preserve ExportRows(1 To ExportCount)

    ExportRows(ExportCount).quoteName = CanonicalHolderQuoteName(b.quoteName)
    ExportRows(ExportCount).Quantity = b.Quantity
    ExportRows(ExportCount).material = b.material

    ExportRows(ExportCount).CadPartIndex = cadIdx
    ExportRows(ExportCount).HasCad = (cadIdx > 0)

    ExportRows(ExportCount).BomThickness = b.BomThickness
    ExportRows(ExportCount).BomWidth = b.BomWidth
    ExportRows(ExportCount).BomLength = b.BomLength
    ExportRows(ExportCount).HasBomDims = b.hasDims

    If cadIdx > 0 Then

        ExportRows(ExportCount).Thickness = parts(cadIdx).Thickness
        ExportRows(ExportCount).Width = parts(cadIdx).Width
        ExportRows(ExportCount).Length = parts(cadIdx).Length
        ExportRows(ExportCount).Status = CompareBomToCadStatus(b, cadIdx)

        parts(cadIdx).UsedForBomMatch = True

    Else

        ExportRows(ExportCount).Status = "NO CAD MATCH"

    End If

    Exit Sub

ErrHandler:
    LogLine "AddExportRow error: " & Err.Description
End Sub

Private Sub RecordSpecialBomCadMatch(ByRef b As BomInfo, ByVal cadIdx As Long)
On Error Resume Next

    If cadIdx <= 0 Then Exit Sub
    If SpecialBomCadMatches Is Nothing Then Exit Sub

    Dim k As String
    k = NormalizeKey(b.quoteName)

    SpecialBomCadMatches(k) = cadIdx
    SpecialBomCadQuoteNames(k) = b.Description
End Sub

' True when the part's regular-box volume is plausible for at least one of
' the class's BOM stock sizes (a rotated box only grows, never shrinks much).
Private Function PartVolumePlausibleForPullcoreRows(ByVal cadIdx As Long, _
                                                    ByRef bomRowIdx() As Long, _
                                                    ByVal rowN As Long) As Boolean
On Error GoTo ErrHandler

    PartVolumePlausibleForPullcoreRows = False

    Dim aabbV As Double
    aabbV = parts(cadIdx).BBoxVolume

    If aabbV <= 0# Then Exit Function

    Dim j As Long
    Dim rv As Double

    For j = 1 To rowN
        If BomRows(bomRowIdx(j)).hasDims Then
            rv = BomRows(bomRowIdx(j)).BomLength * BomRows(bomRowIdx(j)).BomWidth * BomRows(bomRowIdx(j)).BomThickness
            If rv > 0# Then
                If aabbV >= rv * FITTED_FALLBACK_VOL_RATIO_MIN And _
                   aabbV <= rv * FITTED_FALLBACK_VOL_RATIO_MAX Then
                    PartVolumePlausibleForPullcoreRows = True
                    Exit Function
                End If
            End If
        End If
    Next j

    Exit Function

ErrHandler:
    PartVolumePlausibleForPullcoreRows = False
End Function

' Regular-row fallback: match a BOM row by opening plausible parts
' individually and comparing their FITTED box to the BOM stock size. Same
' per-axis gates and name bonuses as the regular-box matcher.
Private Function FindCadMatchForBomByFittedBox(ByRef b As BomInfo) As Long
On Error GoTo ErrHandler

    FindCadMatchForBomByFittedBox = 0

    If b.hasDims = False Then Exit Function

    Dim bomVol As Double
    bomVol = b.BomLength * b.BomWidth * b.BomThickness
    If bomVol <= 0# Then Exit Function

    Dim bL As Double, bw As Double, bT As Double
    SortThreeDimensions b.BomLength, b.BomWidth, b.BomThickness, bL, bw, bT

    Dim bestIdx As Long
    Dim bestDiff As Double

    bestIdx = 0
    bestDiff = 1E+99

    Dim i As Long
    Dim fL As Double, fW As Double, fT As Double
    Dim dL As Double, dW As Double, dT As Double
    Dim totalDiff As Double

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            If parts(i).BBoxVolume >= bomVol * FITTED_FALLBACK_VOL_RATIO_MIN And _
               parts(i).BBoxVolume <= bomVol * FITTED_FALLBACK_VOL_RATIO_MAX Then

                If TryGetPullcoreBestFitDims(i, fL, fW, fT) Then

                    dL = Abs(fL - bL)
                    dW = Abs(fW - bw)
                    dT = Abs(fT - bT)

                    If dL <= DIM_REVIEW_TOL * 4 And dW <= DIM_REVIEW_TOL * 4 And dT <= DIM_REVIEW_TOL * 4 Then

                        totalDiff = dL + dW + dT

                        If IsNameMatch(parts(i).cleanName, b.quoteName) Then totalDiff = totalDiff - 0.75
                        If IsNameMatch(parts(i).componentName, b.quoteName) Then totalDiff = totalDiff - 0.5

                        If totalDiff < bestDiff Then
                            bestDiff = totalDiff
                            bestIdx = i
                        End If

                    End If

                End If

            End If

        End If

    Next i

    If bestIdx > 0 Then
        LogLine "Matched '" & b.quoteName & "' by FITTED box: " & _
                parts(bestIdx).componentName
    End If

    FindCadMatchForBomByFittedBox = bestIdx

    Exit Function

ErrHandler:
    LogLine "FindCadMatchForBomByFittedBox error: " & Err.Description
    FindCadMatchForBomByFittedBox = 0
End Function

Private Function FindBestCadMatchForBom(ByRef b As BomInfo) As Long
On Error GoTo ErrHandler

    FindBestCadMatchForBom = 0

    If ShouldUseTcpBcpMassPairRule(b) Then

        Dim pref As String
        Dim massIdx As Long

        pref = GetMassPreferenceForQuoteName(b.quoteName)

        If pref <> "" Then
            massIdx = FindBestCadMatchForBomByDimsAndMassPreference(b, pref)

            If massIdx > 0 Then
                FindBestCadMatchForBom = massIdx
                Exit Function
            End If
        End If

    End If

    Dim bestIdx As Long
    Dim bestDiff As Double

    bestIdx = 0
    bestDiff = DIM_MAX_MATCH_TOTAL_DIFF + 1#

    Dim i As Long

    If b.hasDims Then

        For i = 1 To PartCount

            If parts(i).UsedForBomMatch = False Then

                Dim dL As Double
                Dim dW As Double
                Dim dT As Double

                dL = Abs(parts(i).Length - b.BomLength)
                dW = Abs(parts(i).Width - b.BomWidth)
                dT = Abs(parts(i).Thickness - b.BomThickness)

                Dim totalDiff As Double
                totalDiff = dL + dW + dT

                Dim nameBonus As Double
                nameBonus = 0#

                If IsNameMatch(parts(i).cleanName, b.quoteName) Then nameBonus = nameBonus - 0.75
                If IsNameMatch(parts(i).componentName, b.quoteName) Then nameBonus = nameBonus - 0.5

                totalDiff = totalDiff + nameBonus

                If dL <= DIM_REVIEW_TOL * 4 And dW <= DIM_REVIEW_TOL * 4 And dT <= DIM_REVIEW_TOL * 4 Then
                    If totalDiff < bestDiff Then
                        bestDiff = totalDiff
                        bestIdx = i
                    End If
                End If

            End If

        Next i

    End If

    ' Parts placed rotated in the assembly (typical for outsource layout
    ' assemblies) have inflated regular boxes and never match on raw dims.
    ' Open plausible parts individually and compare their FITTED box instead.
    If bestIdx = 0 And REGULAR_MATCH_FITTED_BBOX_FALLBACK Then
        bestIdx = FindCadMatchForBomByFittedBox(b)
    End If

    If bestIdx = 0 Then
        bestIdx = FindMassBasedCadMatchForBom(b)
    End If

    FindBestCadMatchForBom = bestIdx
    Exit Function

ErrHandler:
    LogLine "FindBestCadMatchForBom error: " & Err.Description
    FindBestCadMatchForBom = 0
End Function

Private Function FindBestCadMatchForBomByDimsAndMassPreference(ByRef b As BomInfo, _
                                                               ByVal massPreference As String) As Long
On Error GoTo ErrHandler

    FindBestCadMatchForBomByDimsAndMassPreference = 0

    If b.hasDims = False Then Exit Function
    If massPreference = "" Then Exit Function

    Dim i As Long
    Dim bestDimDiff As Double

    bestDimDiff = 1E+99

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim dL As Double
            Dim dW As Double
            Dim dT As Double
            Dim totalDiff As Double

            dL = Abs(parts(i).Length - b.BomLength)
            dW = Abs(parts(i).Width - b.BomWidth)
            dT = Abs(parts(i).Thickness - b.BomThickness)

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

            dL = Abs(parts(i).Length - b.BomLength)
            dW = Abs(parts(i).Width - b.BomWidth)
            dT = Abs(parts(i).Thickness - b.BomThickness)

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
        LogLine "TCP/BCP mass match: " & b.quoteName & _
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

Private Function FindMassBasedCadMatchForBom(ByRef b As BomInfo) As Long
On Error GoTo ErrHandler

    FindMassBasedCadMatchForBom = 0

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim i As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim score As Double
            score = 0

            If IsNameMatch(parts(i).cleanName, b.quoteName) Then score = score + 10
            If IsNameMatch(parts(i).componentName, b.quoteName) Then score = score + 6

            If score > bestScore Then
                bestScore = score
                bestIdx = i
            End If

        End If

    Next i

    If bestScore <= 0 Then
        FindMassBasedCadMatchForBom = 0
    Else
        FindMassBasedCadMatchForBom = bestIdx
    End If

    Exit Function

ErrHandler:
    FindMassBasedCadMatchForBom = 0
End Function

Private Function CompareBomToCadStatus(ByRef b As BomInfo, ByVal cadIdx As Long) As String
On Error GoTo ErrHandler

    If cadIdx <= 0 Then
        CompareBomToCadStatus = "NO CAD MATCH"
        Exit Function
    End If

    If b.hasDims = False Then
        CompareBomToCadStatus = "MATCH (no BOM dims)"
        Exit Function
    End If

    Dim dL As Double
    Dim dW As Double
    Dim dT As Double

    dL = Abs(parts(cadIdx).Length - b.BomLength)
    dW = Abs(parts(cadIdx).Width - b.BomWidth)
    dT = Abs(parts(cadIdx).Thickness - b.BomThickness)

    Dim worst As Double
    worst = dL

    If dW > worst Then worst = dW
    If dT > worst Then worst = dT

    If worst <= DIM_OK_TOL Then
        CompareBomToCadStatus = "OK"
    ElseIf worst <= DIM_REVIEW_TOL Then
        CompareBomToCadStatus = "REVIEW (" & FormatNumberForCsv(worst) & ")"
    Else
        CompareBomToCadStatus = "CHECK (" & FormatNumberForCsv(worst) & ")"
    End If

    Exit Function

ErrHandler:
    CompareBomToCadStatus = "ERROR"
End Function

Private Function IsPossibleJBlockByDims(ByVal L As Double, ByVal W As Double, ByVal T As Double) As Boolean
    If Abs(T - JBlockTgtT) <= J_BLOCK_DIM_MATCH_TOL Then
        If Abs(W - JBlockTgtW) <= J_BLOCK_DIM_MATCH_TOL Then
            If Abs(L - JBlockTgtL) <= J_BLOCK_DIM_MATCH_TOL Then
                IsPossibleJBlockByDims = True
            End If
        End If
    End If
End Function

Private Function IsPossibleCamByDims(ByVal L As Double, ByVal W As Double, ByVal T As Double) As Boolean
    If Abs(T - EjCamTgtT) <= EJECTOR_CAM_THICKNESS_TOL Then
        If Abs(W - EjCamTgtW) <= EJECTOR_CAM_WIDTH_TOL Then
            If Abs(L - EjCamTgtL) <= EJECTOR_CAM_LENGTH_TOL Then
                IsPossibleCamByDims = True
            End If
        End If
    End If
End Function

Private Function CleanPullcoreDisplayName(ByVal s As String) As String
    s = Trim(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanPullcoreDisplayName = s
End Function

Private Function GetPullcoreLocationCode(ByVal text As String) As String
    Dim s As String
    s = NormalizeText(text)

    If InStr(s, "IDTE") > 0 Or InStr(s, "ID TE") > 0 Then GetPullcoreLocationCode = "IDTE": Exit Function
    If InStr(s, "IDLE") > 0 Or InStr(s, "ID LE") > 0 Then GetPullcoreLocationCode = "IDLE": Exit Function
    If InStr(s, "ODTE") > 0 Or InStr(s, "OD TE") > 0 Then GetPullcoreLocationCode = "ODTE": Exit Function
    If InStr(s, "ODLE") > 0 Or InStr(s, "OD LE") > 0 Then GetPullcoreLocationCode = "ODLE": Exit Function

    If InStr(s, "ID") > 0 And InStr(s, "OD") = 0 Then GetPullcoreLocationCode = "ID": Exit Function
    If InStr(s, "OD") > 0 And InStr(s, "ID") = 0 Then GetPullcoreLocationCode = "OD": Exit Function

    Dim toks() As String
    toks = Split(s, " ")

    Dim i As Long
    For i = LBound(toks) To UBound(toks)
        If toks(i) = "TE" Then GetPullcoreLocationCode = "TE": Exit Function
        If toks(i) = "LE" Then GetPullcoreLocationCode = "LE": Exit Function
    Next i

    GetPullcoreLocationCode = ""
End Function

Private Function PullcoreLocationScore(ByVal bomLoc As String, ByVal cadLoc As String) As Double
    If bomLoc = "" Or cadLoc = "" Then
        PullcoreLocationScore = 0
        Exit Function
    End If

    If bomLoc = cadLoc Then
        PullcoreLocationScore = -3#
        Exit Function
    End If

    Dim bSide As String
    Dim cSide As String

    bSide = Left(bomLoc, 2)
    cSide = Left(cadLoc, 2)

    If bSide = cSide Then
        PullcoreLocationScore = -1#
    Else
        PullcoreLocationScore = 5#
    End If
End Function

' ============================================================
' TOP/BOT INS FALLBACK FROM CAD GEOMETRY
' ============================================================

Private Sub AddMissingTopBotInsFromCadGeometry()
On Error GoTo ErrHandler

    Dim haveTop As Boolean
    Dim haveBot As Boolean

    haveTop = ExportQuoteExists("TOP INS")
    haveBot = ExportQuoteExists("BOT INS")

    If haveTop And haveBot Then
        LogLine "TOP/BOT INS already present. Geometry fallback not needed."
        Exit Sub
    End If

    Dim candidates As Collection
    Set candidates = FindInsertCadCandidates()

    If candidates Is Nothing Or candidates.count = 0 Then
        LogLine "No insert geometry candidates found."
        Exit Sub
    End If

    LogQuarterInchCadCandidates candidates

    Dim idxTop As Long
    Dim idxBot As Long

    idxTop = 0
    idxBot = 0

    Dim tgtLarge As Double
    Dim tgtSmall As Double

    If ComputeInsTargetFromHolderAndPot(tgtLarge, tgtSmall) Then
        LogLine "INS target from formula: " & _
                FormatNumberForCsv(tgtLarge) & " x " & FormatNumberForCsv(tgtSmall) & " x 0.25"

        FindInsPairByTargetSize candidates, tgtLarge, tgtSmall, idxTop, idxBot
    End If

    If idxTop = 0 Or idxBot = 0 Then
        FindLargestSameSizeQuarterInchPairCandidates candidates, idxTop, idxBot
    End If

    If idxTop = 0 Or idxBot = 0 Then
        FindLightestAndHeaviestCandidate candidates, idxTop, idxBot
    End If

    If idxTop > 0 And haveTop = False Then
        AddExportRowFromCadPart "TOP INS", idxTop
        parts(idxTop).UsedForBomMatch = True
        LogLine "Added TOP INS from CAD geometry: " & parts(idxTop).componentName
    End If

    If idxBot > 0 And haveBot = False Then
        AddExportRowFromCadPart "BOT INS", idxBot
        parts(idxBot).UsedForBomMatch = True
        LogLine "Added BOT INS from CAD geometry: " & parts(idxBot).componentName
    End If

    Exit Sub

ErrHandler:
    LogLine "AddMissingTopBotInsFromCadGeometry error: " & Err.Description
End Sub

Private Function ComputeInsTargetFromHolderAndPot(ByRef tgtLarge As Double, ByRef tgtSmall As Double) As Boolean
On Error GoTo ErrHandler

    Dim hL As Double, hW As Double, hT As Double
    Dim pL As Double, pW As Double, pT As Double

    If Not GetMatchedDimsByQuote("ID HOLDER", hL, hW, hT) Then
        If Not GetMatchedDimsByQuote("OD HOLDER", hL, hW, hT) Then Exit Function
    End If

    If Not GetMatchedDimsByQuote("ID POT BLOCK", pL, pW, pT) Then
        If Not GetMatchedDimsByQuote("OD POT BLOCK", pL, pW, pT) Then
            If Not GetMatchedDimsByQuote("POT BLOCK", pL, pW, pT) Then Exit Function
        End If
    End If

    tgtLarge = hL
    tgtSmall = hW + pT

    ComputeInsTargetFromHolderAndPot = (tgtLarge > 0 And tgtSmall > 0)
    Exit Function

ErrHandler:
    ComputeInsTargetFromHolderAndPot = False
End Function

Private Function GetMatchedDimsByQuote(ByVal quoteName As String, _
                                       ByRef L As Double, ByRef W As Double, ByRef T As Double) As Boolean
On Error GoTo ErrHandler

    Dim key As String
    key = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = key Then
            If ExportRows(i).Length > 0 Then
                L = ExportRows(i).Length
                W = ExportRows(i).Width
                T = ExportRows(i).Thickness
            Else
                L = ExportRows(i).BomLength
                W = ExportRows(i).BomWidth
                T = ExportRows(i).BomThickness
            End If

            GetMatchedDimsByQuote = (L > 0)
            Exit Function
        End If
    Next i

    Exit Function

ErrHandler:
    GetMatchedDimsByQuote = False
End Function

Private Sub FindInsPairByTargetSize(ByVal candidates As Collection, _
                                    ByVal tgtLarge As Double, ByVal tgtSmall As Double, _
                                    ByRef idxTop As Long, ByRef idxBot As Long)
On Error GoTo ErrHandler

    Dim best1 As Long
    Dim best2 As Long
    Dim d1 As Double
    Dim d2 As Double

    best1 = 0
    best2 = 0
    d1 = 1E+99
    d2 = 1E+99

    Dim n As Long
    Dim idx As Long
    Dim cL As Double
    Dim cS As Double
    Dim dist As Double

    For n = 1 To candidates.count

        idx = CLng(candidates(n))

        cL = parts(idx).Length
        cS = parts(idx).Width

        dist = Abs(cL - tgtLarge) + Abs(cS - tgtSmall)

        If dist <= 2# Then
            If dist < d1 Then
                d2 = d1
                best2 = best1
                d1 = dist
                best1 = idx
            ElseIf dist < d2 Then
                d2 = dist
                best2 = idx
            End If
        End If

    Next n

    If best1 > 0 And best2 > 0 Then

        If parts(best1).massValue <= parts(best2).massValue Then
            idxTop = best1
            idxBot = best2
        Else
            idxTop = best2
            idxBot = best1
        End If

    Else

        idxTop = best1
        idxBot = best2

    End If

    Exit Sub

ErrHandler:
    LogLine "FindInsPairByTargetSize error: " & Err.Description
End Sub

Private Function FindInsertCadCandidates() As Collection
On Error GoTo ErrHandler

    Dim result As New Collection

    Dim i As Long

    For i = 1 To PartCount
        If parts(i).UsedForBomMatch = False Then
            If IsInsertCadGeometryCandidate(i) Then
                result.Add i
            End If
        End If
    Next i

    Set FindInsertCadCandidates = result
    Exit Function

ErrHandler:
    Set FindInsertCadCandidates = New Collection
End Function

Private Function IsInsertCadGeometryCandidate(ByVal cadIdx As Long) As Boolean
    If cadIdx <= 0 Then Exit Function

    If Abs(parts(cadIdx).Thickness - QUARTER_INCH_THICKNESS) <= INSERT_THICKNESS_TOL Then
        IsInsertCadGeometryCandidate = True
        Exit Function
    End If

    Dim hay As String
    hay = UCase(parts(cadIdx).cleanName & " " & parts(cadIdx).componentName)

    If InStr(hay, "INS") > 0 Or InStr(hay, "INSERT") > 0 Then
        IsInsertCadGeometryCandidate = True
    End If
End Function

Private Sub FindLargestSameSizeQuarterInchPairCandidates(ByVal candidates As Collection, _
                                                         ByRef idxA As Long, _
                                                         ByRef idxB As Long)
On Error GoTo ErrHandler

    idxA = 0
    idxB = 0

    If candidates Is Nothing Then Exit Sub
    If candidates.count < 2 Then Exit Sub

    Dim bestVol As Double
    bestVol = -1

    Dim i As Long
    Dim j As Long
    Dim a As Long
    Dim b As Long

    For i = 1 To candidates.count - 1

        a = CLng(candidates(i))

        For j = i + 1 To candidates.count

            b = CLng(candidates(j))

            If SameWidthLengthPair(a, b) Then

                Dim vol As Double
                vol = parts(a).BBoxVolume + parts(b).BBoxVolume

                If vol > bestVol Then
                    bestVol = vol
                    idxA = a
                    idxB = b
                End If

            End If

        Next j

    Next i

    Exit Sub

ErrHandler:
    LogLine "FindLargestSameSizeQuarterInchPairCandidates error: " & Err.Description
End Sub

Private Function SameWidthLengthPair(ByVal a As Long, ByVal b As Long) As Boolean
    If a <= 0 Or b <= 0 Then Exit Function

    If Abs(parts(a).Length - parts(b).Length) <= INSERT_LENGTH_MATCH_TOL Then
        If Abs(parts(a).Width - parts(b).Width) <= INSERT_WIDTH_MATCH_TOL Then
            SameWidthLengthPair = True
        End If
    End If
End Function

Private Sub FindLightestAndHeaviestCandidate(ByVal candidates As Collection, _
                                             ByRef lightIdx As Long, _
                                             ByRef heavyIdx As Long)
On Error GoTo ErrHandler

    lightIdx = 0
    heavyIdx = 0

    If candidates Is Nothing Then Exit Sub
    If candidates.count = 0 Then Exit Sub

    Dim i As Long
    Dim c As Long
    Dim lightMass As Double
    Dim heavyMass As Double

    lightMass = 1E+15
    heavyMass = -1

    For i = 1 To candidates.count

        c = CLng(candidates(i))

        If parts(c).massValue < lightMass Then
            lightMass = parts(c).massValue
            lightIdx = c
        End If

        If parts(c).massValue > heavyMass Then
            heavyMass = parts(c).massValue
            heavyIdx = c
        End If

    Next i

    If lightIdx = heavyIdx Then heavyIdx = 0

    Exit Sub

ErrHandler:
    LogLine "FindLightestAndHeaviestCandidate error: " & Err.Description
End Sub

Private Sub AddExportRowFromCadPart(ByVal quoteName As String, ByVal cadIdx As Long)
On Error GoTo ErrHandler

    If cadIdx <= 0 Then Exit Sub

    ExportCount = ExportCount + 1
    ReDim Preserve ExportRows(1 To ExportCount)

    ExportRows(ExportCount).quoteName = quoteName
    ExportRows(ExportCount).Quantity = parts(cadIdx).Quantity
    ExportRows(ExportCount).material = DEFAULT_STEEL_TYPE

    ExportRows(ExportCount).CadPartIndex = cadIdx
    ExportRows(ExportCount).HasCad = True

    ExportRows(ExportCount).Thickness = parts(cadIdx).Thickness
    ExportRows(ExportCount).Width = parts(cadIdx).Width
    ExportRows(ExportCount).Length = parts(cadIdx).Length

    ExportRows(ExportCount).BomThickness = 0
    ExportRows(ExportCount).BomWidth = 0
    ExportRows(ExportCount).BomLength = 0
    ExportRows(ExportCount).HasBomDims = False

    ExportRows(ExportCount).Status = "FROM CAD GEOMETRY"
    Exit Sub

ErrHandler:
    LogLine "AddExportRowFromCadPart error: " & Err.Description
End Sub

Private Function ExportQuoteExists(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            ExportQuoteExists = True
            Exit Function
        End If
    Next i
End Function

Private Sub LogQuarterInchCadCandidates(ByVal candidates As Collection)
On Error Resume Next

    If candidates Is Nothing Then Exit Sub

    Dim i As Long
    Dim c As Long

    LogLine "Insert geometry candidates: " & candidates.count

    For i = 1 To candidates.count
        c = CLng(candidates(i))

        LogLine "  CAND: " & parts(c).componentName & _
                " L=" & FormatNumberForCsv(parts(c).Length) & _
                " W=" & FormatNumberForCsv(parts(c).Width) & _
                " T=" & FormatNumberForCsv(parts(c).Thickness) & _
                " mass=" & FormatNumberForCsv(parts(c).massValue)
    Next i
End Sub

' ============================================================
' J BLOCK PACKAGE
' ============================================================

Private Sub ExportJBlockPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    Dim jbIdx As Long
    jbIdx = FindBestJBlockIndex()

    If jbIdx <= 0 Then
        LogLine "J BLOCK package: no J BLOCK CAD part found."
        Exit Sub
    End If

    Dim jbFolder As String
    jbFolder = outputFolder & "\" & CurrentJobNumber & " " & J_BLOCK_FOLDER_NAME

    EnsureFolderDeep jbFolder

    LogLine "J BLOCK CAD selected: " & parts(jbIdx).componentName

    ExportJBlockBaseExactPackage jbFolder, jbIdx

    Dim camIdx As Long
    camIdx = FindBestEjectorCamIndex()

    If camIdx > 0 Then
        LogLine "Ejector CAM CAD selected: " & parts(camIdx).componentName
        ExportJBlockCamDxfOnly jbFolder, camIdx
    Else
        LogLine "J BLOCK package: ejector CAM not found (skipping CAM DXF)."
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportJBlockPackage error: " & Err.Description
End Sub

Private Function FindBestJBlockIndex() As Long
On Error GoTo ErrHandler

    FindBestJBlockIndex = 0

    Dim specialIdx As Long

    specialIdx = GetSpecialCadMatchByQuoteKeys("EJECTOR J-BLOCK|EJ J-BLOCK|J-BLOCK|J BLOCK|JBLOCK|EJECTOR U-BLOCK|EJ U-BLOCK|U-BLOCK|U BLOCK|UBLOCK")

    If specialIdx > 0 And specialIdx <= PartCount Then
        FindBestJBlockIndex = specialIdx
        LogLine "J BLOCK selected from special CAD/library match: " & _
                parts(specialIdx).componentName
        Exit Function
    End If

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim fbIdx As Long
    Dim fbDist As Double

    fbIdx = 0
    fbDist = 1E+99

    Dim i As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch Then GoTo NextPart

        Dim dist As Double

        dist = Abs(parts(i).Length - JBlockTgtL) * 2# _
             + Abs(parts(i).Width - JBlockTgtW) _
             + Abs(parts(i).Thickness - JBlockTgtT)

        If dist < fbDist Then
            fbDist = dist
            fbIdx = i
        End If

        Dim score As Double
        score = 0

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath

        If ContainsAnyPipeKey(hay, J_BLOCK_NAME_KEYS) Then score = score + 5000

        If IsPossibleJBlockByDims(parts(i).Length, parts(i).Width, parts(i).Thickness) Then
            score = score + (1000# - dist * 100#)
        End If

        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If

NextPart:
    Next i

    LogJBlockCandidates

    If bestScore <= 0 Then

        If fbIdx > 0 And fbDist <= 1.6 Then
            FindBestJBlockIndex = fbIdx
            LogLine "J BLOCK fallback closest bbox: " & parts(fbIdx).componentName
        Else
            FindBestJBlockIndex = 0
        End If

    Else

        FindBestJBlockIndex = bestIdx
        LogLine "J BLOCK best match: " & parts(bestIdx).componentName

    End If

    Exit Function

ErrHandler:
    LogLine "FindBestJBlockIndex error: " & Err.Description
    FindBestJBlockIndex = 0
End Function

Private Sub LogJBlockCandidates()
On Error Resume Next

    Dim i As Long

    For i = 1 To PartCount

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName

        If ContainsAnyPipeKey(hay, J_BLOCK_NAME_KEYS) Or _
           IsPossibleJBlockByDims(parts(i).Length, parts(i).Width, parts(i).Thickness) Then

            LogLine "  JBLOCK CAND: " & parts(i).componentName & _
                    " L=" & FormatNumberForCsv(parts(i).Length) & _
                    " W=" & FormatNumberForCsv(parts(i).Width) & _
                    " T=" & FormatNumberForCsv(parts(i).Thickness)
        End If

    Next i
End Sub

Private Function FindBestEjectorCamIndex() As Long
On Error GoTo ErrHandler

    FindBestEjectorCamIndex = 0

    Dim specialIdx As Long

    specialIdx = GetSpecialCadMatchByQuoteKeys("EJECTOR CAM|EJ CAM|EJ. CAM|EJ.CAM|CAM")

    If specialIdx > 0 And specialIdx <= PartCount Then
        FindBestEjectorCamIndex = specialIdx
        LogLine "EJECTOR CAM selected from special CAD/library match: " & _
                parts(specialIdx).componentName
        Exit Function
    End If

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim i As Long

    For i = 1 To PartCount

        Dim score As Double
        score = 0

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath

        If ContainsAnyPipeKey(hay, EJECTOR_CAM_NAME_KEYS) Then score = score + 60
        If IsPossibleCamByDims(parts(i).Length, parts(i).Width, parts(i).Thickness) Then score = score + 30

        If InStr(UCase(hay), "PULLCORE") > 0 Or InStr(UCase(hay), "PULL CORE") > 0 Then score = score - 80

        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If

    Next i

    LogEjectorCamCandidates

    If bestScore <= 0 Then
        FindBestEjectorCamIndex = 0
    Else
        FindBestEjectorCamIndex = bestIdx
    End If

    Exit Function

ErrHandler:
    LogLine "FindBestEjectorCamIndex error: " & Err.Description
    FindBestEjectorCamIndex = 0
End Function

Private Sub LogEjectorCamCandidates()
On Error Resume Next

    Dim i As Long

    For i = 1 To PartCount

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName

        If ContainsAnyPipeKey(hay, EJECTOR_CAM_NAME_KEYS) Then
            LogLine "  EJ CAM CAND: " & parts(i).componentName
        End If

    Next i
End Sub

Private Sub ExportJBlockBaseExactPackage(ByVal jbFolder As String, ByVal jbIdx As Long)
On Error GoTo ErrHandler

    If jbIdx <= 0 Then Exit Sub

    Dim p As PartInfo
    p = parts(jbIdx)

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim jbToken As String
    jbToken = CleanQuoteTokenForFile(J_BLOCK_BASE_FILE_TOKEN)

    If jbToken = "" Then jbToken = "J BLOCK"

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = jbFolder & "\" & CurrentJobNumber & "_" & jbToken & "_" & custToken & "_" & dateToken & ".x_t"
    dxfPath = jbFolder & "\" & CurrentJobNumber & "_" & jbToken & "_" & custToken & "_" & dateToken & ".dxf"

    xtPath = GetUniqueFilePath(xtPath)
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Exporting J BLOCK base XT + DXF."

    If swModel.GetType = swDocASSEMBLY Then
        ExportAssemblyComponentBySuppressRestWithOptionalDxf swModel, p, xtPath, dxfPath, "J BLOCK", False
    ElseIf swModel.GetType = swDocPART Then
        If p.isBodyOnly Then
            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath
            ShowAllPartBodies swModel
        Else
            SaveModelAs swModel, xtPath
        End If
    End If

    If swModel.GetType = swDocASSEMBLY Then
        If CreateJBlockDxfFromAssemblyTopView(jbFolder, p, dxfPath) Then
            LogLine "J BLOCK DXF built from native assembly top-view path."
            Exit Sub
        Else
            LogLine "WARNING: J BLOCK native assembly DXF failed; falling back to part-based DXF."
        End If
    End If

    CreateProjectedDxfFromXtPath xtPath, dxfPath, "J BLOCK", _
                                  J_BLOCK_PARENT_VIEW_PRIMARY, J_BLOCK_PARENT_VIEW_FALLBACK, _
                                  True, False, False

    Exit Sub

ErrHandler:
    LogLine "ExportJBlockBaseExactPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Function CreateJBlockDxfFromAssemblyTopView(ByVal jbFolder As String, _
                                                    ByRef p As PartInfo, _
                                                    ByVal dxfPath As String) As Boolean
On Error GoTo ErrHandler

    CreateJBlockDxfFromAssemblyTopView = False

    If swModel Is Nothing Then Exit Function
    If swModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim keepNames As Collection
    Set keepNames = New Collection
    AddUniqueComponentName keepNames, p.componentName

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_JBLOCK_DXF_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    ' IMPORTANT:
    ' Use native SLDASM, not X_T, so CMS_TOP / standard views survive.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_JBLOCK_TOPVIEW_TEMP.sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "J BLOCK assembly-top DXF: could not isolate J BLOCK component."
        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    LogLine "J BLOCK DXF: saving isolated native SLDASM for right-face center view:"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(swModel, tempNativePath) = False Then GoTo CleanExit

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "J BLOCK assembly-top DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    ' J BLOCK center/base view must be the right face. Projected views unfold from there.
    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, "J BLOCK", _
                                     "*Right", "*Right", _
                                     True, False, False, False, False

    CreateJBlockDxfFromAssemblyTopView = fso.FileExists(dxfPath)

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If Not fso2 Is Nothing Then
        If tempFolder <> "" Then
            If fso2.FolderExists(tempFolder) Then fso2.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateJBlockDxfFromAssemblyTopView error: " & Err.Description
    CreateJBlockDxfFromAssemblyTopView = False
    Resume CleanExit
End Function

Private Sub ExportJBlockCamDxfOnly(ByVal jbFolder As String, ByVal camIdx As Long)
On Error GoTo ErrHandler

    If camIdx <= 0 Then Exit Sub

    Dim p As PartInfo
    p = parts(camIdx)

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim camToken As String
    camToken = CleanQuoteTokenForFile(J_BLOCK_CAM_FILE_TOKEN)

    If camToken = "" Then camToken = "CAM"

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = jbFolder & "\" & CurrentJobNumber & "_" & camToken & "_" & custToken & "_" & dateToken & ".x_t"
    dxfPath = jbFolder & "\" & CurrentJobNumber & "_" & camToken & "_" & custToken & "_" & dateToken & ".dxf"

    xtPath = GetUniqueFilePath(xtPath)
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Exporting J BLOCK ejector CAM XT + DXF."

    If swModel.GetType = swDocASSEMBLY Then
        ExportAssemblyComponentBySuppressRestWithOptionalDxf swModel, p, xtPath, dxfPath, "CAM", False
    ElseIf swModel.GetType = swDocPART Then
        If p.isBodyOnly Then
            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath
            ShowAllPartBodies swModel
        Else
            SaveModelAs swModel, xtPath
        End If
    End If

    ' FIX: the J BLOCK base DXF is projected off the RIGHT face - both
    ' J_BLOCK_PARENT_VIEW_PRIMARY and the hard-coded "*Right" in
    ' CreateJBlockDxfFromAssemblyTopView. The cam was being projected off the TOP
    ' face, so it came out rotated relative to the block it belongs to. Use the
    ' same center view as the block.
    CreateProjectedDxfFromXtPath xtPath, dxfPath, "CAM", _
                                  J_BLOCK_PARENT_VIEW_PRIMARY, J_BLOCK_PARENT_VIEW_FALLBACK, _
                                  True, False, False

    Exit Sub

ErrHandler:
    LogLine "ExportJBlockCamDxfOnly error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

' ============================================================
' END OF PART 6 OF 8
' Paste Part 7 next.
' ============================================================
' ============================================================
' PULLCORE CAM AND KEY PACKAGE
' ============================================================

Private Sub ExportPullcoreCamKeyPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If PullcoreMatchCount = 0 Then
        LogLine "PULLCORE package: no pullcore matches."
        Exit Sub
    End If

    Dim pcFolder As String
    pcFolder = outputFolder & "\" & CurrentJobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME

    EnsureFolderDeep pcFolder

    Dim usedNames As Object
    Set usedNames = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).CadPartIndex > 0 Then
            ExportPullcoreCandidateSet pcFolder, i, usedNames
        Else
            LogLine "PULLCORE skipped (no CAD): " & PullcoreMatches(i).quoteName
        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "ExportPullcoreCamKeyPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Sub ExportPullcoreCandidateSet(ByVal pcFolder As String, _
                                       ByVal matchIdx As Long, _
                                       ByVal usedNames As Object)
On Error GoTo ErrHandler

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    If cadIdx <= 0 Then Exit Sub

    Dim exportName As String
    exportName = PullcoreMatches(matchIdx).quoteName

    If exportName = "" Then exportName = CleanPullcoreDisplayName(PullcoreMatches(matchIdx).Description)

    If exportName = "" Then
        If PullcoreMatches(matchIdx).isCam Then
            exportName = "PULLCORE CAM"
        Else
            exportName = "PULLCORE KEY"
        End If
    End If

    Dim baseKey As String
    baseKey = UCase(exportName)

    If Not usedNames Is Nothing Then

        If usedNames.Exists(baseKey) Then

            Dim n As Long
            n = CLng(usedNames(baseKey)) + 1

            usedNames(baseKey) = n
            exportName = exportName & " " & CStr(n)

        Else

            usedNames(baseKey) = 1

        End If

    End If

    LogLine "PULLCORE export name: " & exportName

    ExportOneSpecialNamedPartToFolder pcFolder, cadIdx, exportName, "PULLCORE", True
    Exit Sub

ErrHandler:
    LogLine "ExportPullcoreCandidateSet error: " & Err.Description
End Sub

' ============================================================
' PULLCORE STOP / FLIPPER CAM COVER PACKAGE
' ============================================================

Private Sub ExportPullcoreStopPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    Dim stopIdx As Long
    Dim coverIdx As Long

    stopIdx = FindBestPartByNameAndDims(PULLCORE_STOP_NAME_KEYS, 0, 0, 0, 0)

    coverIdx = FindBestPartByNameAndDims(FLIPPER_CAM_COVER_NAME_KEYS, _
                                         FLIPPER_CAM_COVER_TARGET_THICKNESS, _
                                         FLIPPER_CAM_COVER_TARGET_WIDTH, _
                                         FLIPPER_CAM_COVER_TARGET_LENGTH, _
                                         FLIPPER_CAM_COVER_DIM_TOL)

    If stopIdx <= 0 And coverIdx <= 0 Then
        LogLine "PULLCORE STOP package: no stop or cover found."
        Exit Sub
    End If

    Dim stopFolder As String
    stopFolder = outputFolder & "\" & CurrentJobNumber & " " & PULLCORE_STOP_FOLDER_NAME

    EnsureFolderDeep stopFolder

    If stopIdx > 0 Then
        LogLine "Pullcore STOP CAD: " & parts(stopIdx).componentName
        ExportOneSpecialNamedPartToFolder stopFolder, stopIdx, "PULLCORE STOP", "PULLCORE STOP", True
    End If

    If coverIdx > 0 Then
        LogLine "Flipper CAM COVER CAD: " & parts(coverIdx).componentName
        ExportOneSpecialNamedPartToFolder stopFolder, coverIdx, "FLIPPER CAM COVER PLATE", "FLIPPER CAM COVER PLATE", True
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportPullcoreStopPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Function FindBestPartByNameAndDims(ByVal pipeKeys As String, _
                                           ByVal targetT As Double, _
                                           ByVal targetW As Double, _
                                           ByVal targetL As Double, _
                                           ByVal tol As Double) As Long
On Error GoTo ErrHandler

    FindBestPartByNameAndDims = 0

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim i As Long

    For i = 1 To PartCount

        Dim score As Double
        score = 0

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath

        If ContainsAnyPipeKey(hay, pipeKeys) Then score = score + 100

        If targetT > 0 And targetW > 0 And targetL > 0 And tol > 0 Then
            If Abs(parts(i).Thickness - targetT) <= tol Then
                If Abs(parts(i).Width - targetW) <= tol Then
                    If Abs(parts(i).Length - targetL) <= tol Then
                        score = score + 40
                    End If
                End If
            End If
        End If

        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If

    Next i

    If bestScore <= 0 Then
        FindBestPartByNameAndDims = 0
    Else
        FindBestPartByNameAndDims = bestIdx
    End If

    Exit Function

ErrHandler:
    LogLine "FindBestPartByNameAndDims error: " & Err.Description
    FindBestPartByNameAndDims = 0
End Function

Private Sub ExportOneSpecialNamedPartToFolder(ByVal folderPath As String, _
                                              ByVal cadIdx As Long, _
                                              ByVal exportName As String, _
                                              ByVal dxfQuoteName As String, _
                                              ByVal makeWireframe As Boolean)
On Error GoTo ErrHandler

    If cadIdx <= 0 Then Exit Sub

    EnsureFolderDeep folderPath

    Dim p As PartInfo
    p = parts(cadIdx)

    Dim quoteToken As String
    quoteToken = CleanQuoteTokenForFile(exportName)

    If quoteToken = "" Then quoteToken = "ITEM"

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = folderPath & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".x_t"
    dxfPath = folderPath & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".dxf"

    xtPath = GetUniqueFilePath(xtPath)
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Special export: " & exportName & " -> " & xtPath

    If swModel.GetType = swDocASSEMBLY Then

        ExportAssemblyComponentBySuppressRestWithOptionalDxf swModel, p, xtPath, dxfPath, dxfQuoteName, False

    ElseIf swModel.GetType = swDocPART Then

        If p.isBodyOnly Then
            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath
            ShowAllPartBodies swModel
        Else
            SaveModelAs swModel, xtPath
        End If

    End If

    Dim oldStraightenAngle As Double
    Dim oldPullcoreCadIdx As Long
    Dim oldKeepImported As Boolean

    oldStraightenAngle = CurrentDxfStraightenAngleRad
    oldPullcoreCadIdx = CurrentPullcoreStraightenCadIndex
    oldKeepImported = CurrentDxfKeepImportedOrientation

    CurrentDxfStraightenAngleRad = 0#
    CurrentPullcoreStraightenCadIndex = 0

    Dim parentPrimary As String
    Dim parentFallback As String

    parentPrimary = "*Front"
    parentFallback = "*Front"

    ' FIX: this test used to be an EXACT match on "PULLCORE" while the DXF writer
    ' decides with IsPullcoreDxfQuote(). The two could disagree, which meant the
    ' DXF writer would try to straighten a part it had been given no CAD index
    ' for. Both ends now ask the same question.
    If STRAIGHTEN_PULLCORE_DXF And IsPullcoreDxfQuote(dxfQuoteName) Then
        CurrentPullcoreStraightenCadIndex = cadIdx

        parentPrimary = PULLCORE_PARENT_VIEW_PRIMARY
        parentFallback = PULLCORE_PARENT_VIEW_FALLBACK

        ' FIX: keep the exported (assembly) orientation so the center view is
        ' projected in the plane the straighten angle lives in. Left False, the
        ' source was re-oriented to an arbitrary CMS_TOP baseline first and the
        ' rotation was applied about the wrong axis.
        CurrentDxfKeepImportedOrientation = PULLCORE_DXF_KEEP_IMPORTED_ORIENTATION

        LogLine "PULLCORE: using matched CAD component for straightening: " & _
                parts(cadIdx).componentName
        LogLine "PULLCORE: center view=" & parentPrimary & _
                "  keepImportedOrientation=" & CStr(CurrentDxfKeepImportedOrientation)
    End If

    CreateProjectedDxfFromXtPath xtPath, dxfPath, dxfQuoteName, _
                                  parentPrimary, parentFallback, makeWireframe, False, False, False, False

    CurrentDxfStraightenAngleRad = oldStraightenAngle
    CurrentPullcoreStraightenCadIndex = oldPullcoreCadIdx
    CurrentDxfKeepImportedOrientation = oldKeepImported

    Exit Sub

ErrHandler:
    LogLine "ExportOneSpecialNamedPartToFolder error (" & exportName & "): " & Err.Description
    On Error Resume Next
    CurrentDxfStraightenAngleRad = 0#
    CurrentPullcoreStraightenCadIndex = 0
    CurrentDxfKeepImportedOrientation = False
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

' ============================================================
' PULLCORE BEST-FIT BBOX / STRAIGHTENING
' ============================================================

' Returns the rotation (degrees, 0..45) that the fitted-box geometry search
' found minimizes this part's bounding box - i.e. how far it is tipped from
' square in the assembly. Valid only after TryGetPullcoreBestFitDims has run
' for the part (the matcher always calls it first). Returns -1 if unknown.
Private Function GetPullcoreFittedTipAngleDeg(ByVal cadIdx As Long) As Double
On Error Resume Next
    GetPullcoreFittedTipAngleDeg = -1#
    If PullcoreBestFitAngleCache Is Nothing Then Exit Function
    Dim ck As String
    ck = CStr(cadIdx)
    If PullcoreBestFitAngleCache.Exists(ck) Then
        GetPullcoreFittedTipAngleDeg = CDbl(PullcoreBestFitAngleCache(ck))
    End If
End Function

Private Function TryGetPullcoreBestFitDims(ByVal cadIdx As Long, _
                                           ByRef outL As Double, _
                                           ByRef outW As Double, _
                                           ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPullcoreBestFitDims = False

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    If PullcoreBestFitDimCache Is Nothing Then
        Set PullcoreBestFitDimCache = CreateObject("Scripting.Dictionary")
    End If
    If PullcoreBestFitAngleCache Is Nothing Then
        Set PullcoreBestFitAngleCache = CreateObject("Scripting.Dictionary")
    End If

    Dim cacheKey As String
    cacheKey = CStr(cadIdx)

    If PullcoreBestFitDimCache.Exists(cacheKey) Then
        Dim cached() As String
        cached = Split(CStr(PullcoreBestFitDimCache(cacheKey)), "|")

        If UBound(cached) >= 2 Then
            outL = CDbl(cached(0))
            outW = CDbl(cached(1))
            outT = CDbl(cached(2))

            TryGetPullcoreBestFitDims = (outL > 0 And outW > 0 And outT > 0)
            Exit Function
        End If
    End If

    Dim partModel As Object
    Set partModel = GetModelDocForCadIndex(cadIdx)

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim fitAngleDeg As Double
    fitAngleDeg = 0#

    If TryCreateAndReadBestFitBoundingBox(partModel, outL, outW, outT, fitAngleDeg) Then

        If LastBestFitFromPlainBBox Then
            ' Only the plain (possibly rotated) part box was obtainable. That is
            ' NOT a fitted box - refuse it so the caller drops this part instead
            ' of naming from a guess. Cache the failure so we do not retry.
            LogLine "PULLCORE: no trusted fitted bbox for " & _
                    parts(cadIdx).componentName & " (only plain bbox available) - refused."
            PullcoreBestFitDimCache(cacheKey) = "0|0|0"
            outL = 0#
            outW = 0#
            outT = 0#
            TryGetPullcoreBestFitDims = False
            Exit Function
        End If

        SortThreeDimensions outL, outW, outT, outL, outW, outT

        outL = Round(outL, DIM_DECIMALS)
        outW = Round(outW, DIM_DECIMALS)
        outT = Round(outT, DIM_DECIMALS)

        PullcoreBestFitDimCache(cacheKey) = CStr(outL) & "|" & CStr(outW) & "|" & CStr(outT)
        PullcoreBestFitAngleCache(cacheKey) = CStr(fitAngleDeg)

        LogLine "PULLCORE best-fit bbox for " & parts(cadIdx).componentName & " = " & _
                FormatNumberForCsv(outL) & "/" & FormatNumberForCsv(outW) & "/" & FormatNumberForCsv(outT) & _
                "  (tip angle " & Format(fitAngleDeg, "0.0") & " deg)"

        TryGetPullcoreBestFitDims = True

    End If

    Exit Function

ErrHandler:
    LogLine "TryGetPullcoreBestFitDims error (cadIdx " & cadIdx & "): " & Err.Description
    TryGetPullcoreBestFitDims = False
End Function

Private Function TryComputePullcoreStraightenAngleFromMatchedCad(ByVal cadIdx As Long, _
                                                                 ByRef detectedDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreStraightenAngleFromMatchedCad = False
    detectedDegOut = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    Dim matchIdx As Long
    matchIdx = FindPullcoreMatchIndexByCadIndex(cadIdx)

    If matchIdx > 0 Then

        If Abs(PullcoreMatches(matchIdx).DetectedAngleDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            ' Return the SIGNED angle (magnitude * tipped-direction sign). The
            ' caller multiplies by PULLCORE_STRAIGHTEN_SIGN to undo the tilt, so
            ' carrying the sign here straightens ID and OD the correct ways.
            Dim storedSign As Double
            storedSign = PullcoreMatches(matchIdx).DetectedAngleSign
            If storedSign = 0# Then storedSign = 1#
            detectedDegOut = storedSign * Abs(PullcoreMatches(matchIdx).DetectedAngleDeg)
            LogLine "PULLCORE angle from match record: " & _
                    Format(detectedDegOut, "0.00") & " deg (signed)"
            TryComputePullcoreStraightenAngleFromMatchedCad = True
            Exit Function
        End If

        If PullcoreMatches(matchIdx).FittedLength > 0 And _
           PullcoreMatches(matchIdx).OriginalLength > 0 Then

            detectedDegOut = EstimatePullcoreAngleFromFittedAndOriginalDeg( _
                PullcoreMatches(matchIdx).FittedLength, _
                PullcoreMatches(matchIdx).FittedWidth, _
                PullcoreMatches(matchIdx).FittedThickness, _
                PullcoreMatches(matchIdx).OriginalLength, _
                PullcoreMatches(matchIdx).OriginalWidth, _
                PullcoreMatches(matchIdx).OriginalThickness)

            ' FIX: the box-math estimator returns a MAGNITUDE only. Returned
            ' unsigned it rotated ID and OD (mirrored parts) the same way, so one
            ' of the pair ended up further from square than it started.
            detectedDegOut = GetPullcoreTiltSign(cadIdx) * Abs(detectedDegOut)

            If Abs(detectedDegOut) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
                TryComputePullcoreStraightenAngleFromMatchedCad = True
                Exit Function
            End If

        End If

    End If

    Dim partModel As Object
    Set partModel = GetModelDocForCadIndex(cadIdx)

    Dim modelEdgeDeg As Double

    If Not partModel Is Nothing Then
        If partModel.GetType = swDocPART Then
            If TryComputePullcoreLongestModelEdgeAngleFromModel(partModel, modelEdgeDeg) Then
                If Abs(modelEdgeDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
                    detectedDegOut = modelEdgeDeg
                    TryComputePullcoreStraightenAngleFromMatchedCad = True
                    Exit Function
                End If
            End If
        End If
    End If

    Dim bestL As Double
    Dim bestW As Double
    Dim bestT As Double

    If TryGetPullcoreMatchedReportDims(cadIdx, bestL, bestW, bestT) Then
        detectedDegOut = EstimatePullcoreAngleFromFittedAndOriginalDeg(bestL, bestW, bestT, _
            parts(cadIdx).OriginalAsmLength, parts(cadIdx).OriginalAsmWidth, parts(cadIdx).OriginalAsmThickness)

        ' FIX: magnitude only - carry the part's lean direction. See above.
        detectedDegOut = GetPullcoreTiltSign(cadIdx) * Abs(detectedDegOut)

        If Abs(detectedDegOut) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            TryComputePullcoreStraightenAngleFromMatchedCad = True
            Exit Function
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreStraightenAngleFromMatchedCad error: " & Err.Description
    TryComputePullcoreStraightenAngleFromMatchedCad = False
End Function

' Which way this pullcore leans, as +1 or -1. Preference order:
'   1) the sign already resolved during BOM matching (edge, else transform),
'   2) the component's own assembly placement transform,
'   3) +1.
' Some angle sources in this module return a signed value and some return a bare
' magnitude. This is what the magnitudes get multiplied by, so every source ends
' up in one convention before PULLCORE_STRAIGHTEN_SIGN is applied.
Private Function GetPullcoreTiltSign(ByVal cadIdx As Long) As Double
On Error Resume Next

    GetPullcoreTiltSign = 1#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    Dim matchIdx As Long
    matchIdx = FindPullcoreMatchIndexByCadIndex(cadIdx)

    If matchIdx > 0 Then
        If PullcoreMatches(matchIdx).DetectedAngleSign <> 0# Then
            GetPullcoreTiltSign = PullcoreMatches(matchIdx).DetectedAngleSign
            Exit Function
        End If
    End If

    If parts(cadIdx).hasAsmPullAngle Then
        If parts(cadIdx).asmPullAngleSign <> 0# Then
            GetPullcoreTiltSign = parts(cadIdx).asmPullAngleSign
        End If
    End If
End Function

Private Function FindPullcoreMatchIndexByCadIndex(ByVal cadIdx As Long) As Long
On Error GoTo ErrHandler

    FindPullcoreMatchIndexByCadIndex = 0

    If cadIdx <= 0 Then Exit Function

    Dim i As Long

    For i = 1 To PullcoreMatchCount
        If PullcoreMatches(i).CadPartIndex = cadIdx Then
            FindPullcoreMatchIndexByCadIndex = i
            Exit Function
        End If
    Next i

    Exit Function

ErrHandler:
    FindPullcoreMatchIndexByCadIndex = 0
End Function

Private Function TryGetPullcoreMatchedReportDims(ByVal cadIdx As Long, _
                                                 ByRef outL As Double, _
                                                 ByRef outW As Double, _
                                                 ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPullcoreMatchedReportDims = False

    outL = 0#
    outW = 0#
    outT = 0#

    If cadIdx <= 0 Then Exit Function

    Dim i As Long

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).CadPartIndex = cadIdx Then

            If PullcoreMatches(i).FittedLength > 0 And _
               PullcoreMatches(i).FittedWidth > 0 And _
               PullcoreMatches(i).FittedThickness > 0 Then

                outL = PullcoreMatches(i).FittedLength
                outW = PullcoreMatches(i).FittedWidth
                outT = PullcoreMatches(i).FittedThickness

                TryGetPullcoreMatchedReportDims = True
                Exit Function

            End If

        End If

    Next i

    Exit Function

ErrHandler:
    TryGetPullcoreMatchedReportDims = False
End Function

Private Function TryComputePullcoreBestFitStraightenAngleFromNativePath(ByVal nativePath As String, _
                                                                        ByVal parentViewName As String, _
                                                                        ByRef detectedDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreBestFitStraightenAngleFromNativePath = False
    detectedDegOut = 0#

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(nativePath) = False Then Exit Function

    Dim ext As String
    ext = LCase(fso.GetExtensionName(nativePath))

    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object

    If ext = "sldasm" Then
        Set mdl = swApp.OpenDoc6(nativePath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    Else
        Set mdl = swApp.OpenDoc6(nativePath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then Exit Function

    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs
    EnsureSwHidden

    ' FIX: TryComputePullcoreLongestModelEdgeAngleFromModel projects edges
    ' through mdl.ActiveView, so the ACTIVE VIEW has to be the same view the
    ' drawing's center view uses. Otherwise the angle is measured in one plane
    ' and then applied in another, which is what was happening: the source had
    ' just been forced to CMS_TOP while the drawing view was *Front.
    If parentViewName <> "" Then
        ShowNamedOrStandardView mdl, parentViewName
        StabilizeActiveView mdl, 60
        LogLine "PULLCORE: measuring tilt in the " & parentViewName & " plane."
    End If

    Dim modelEdgeDeg As Double
    Dim mathDeg As Double

    ' Primary: longest edge in the active-view screen plane (true front-view
    ' rotation, downward tilt ignored). Box math is the fallback.
    If TryComputePullcoreLongestModelEdgeAngleFromModel(mdl, modelEdgeDeg) Then
        If Abs(modelEdgeDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            detectedDegOut = modelEdgeDeg
            TryComputePullcoreBestFitStraightenAngleFromNativePath = True
            GoTo CleanExit
        End If
    End If

    If TryComputeBBoxDimensionMathAngleFromModel(mdl, mathDeg) Then
        If Abs(mathDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            detectedDegOut = mathDeg
            TryComputePullcoreBestFitStraightenAngleFromNativePath = True
            GoTo CleanExit
        End If
    End If

CleanExit:
    On Error Resume Next
    If Not mdl Is Nothing Then swApp.CloseDoc mdl.GetTitle
    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreBestFitStraightenAngleFromNativePath error: " & Err.Description
    On Error Resume Next
    If Not mdl Is Nothing Then swApp.CloseDoc mdl.GetTitle
    TryComputePullcoreBestFitStraightenAngleFromNativePath = False
End Function

' ==========================================================================
' PULLCORE TILT FROM THE ACTUAL DRAWING GEOMETRY  (minimum-area box)
' --------------------------------------------------------------------------
' Why this exists.
'
' The Pull Core Dimensions report was coming out with Original box == Fitted
' box, all three deltas 0.000, and "Detected Pullcore Angle Deg = 0" on every
' cam and key - while the DXF was visibly sitting at ~35 deg. Both of those box
' numbers are read from the component's own part FILE, and imported pullcore
' part files store the body SQUARE. The tilt lives only in the assembly
' placement transform. So the two boxes always matched, the estimator hit its
' "identical, therefore square" early-out, and the DXF had a 0 deg correction
' applied to it. The rotation code was fine; it was being handed nothing.
'
' The X_T is exported from the assembly, so the tilt IS baked into the temp
' native part the drawing is actually made from. Measuring there sees it.
'
' Method: project every edge endpoint into the center view's plane, then sweep
' for the rotation that gives the smallest bounding rectangle. That angle is the
' tilt, and the minimised rectangle is the true squared-up L x W - so the
' dimensions picked up after the rotation are the right ones.
' ==========================================================================

Private Function TryComputePullcoreStraightenAngleFromNativeGeometry(ByVal nativePath As String, _
                                                                     ByVal viewName As String, _
                                                                     ByRef detectedDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreStraightenAngleFromNativeGeometry = False
    detectedDegOut = 0#

    Dim fso As Object
    Dim ext As String
    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object
    Dim tiltDeg As Double
    Dim tightW As Double
    Dim tightH As Double
    Dim rawW As Double
    Dim rawH As Double

    Set mdl = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(nativePath) = False Then Exit Function

    ext = LCase(fso.GetExtensionName(nativePath))

    If ext = "sldasm" Then
        Set mdl = swApp.OpenDoc6(nativePath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    Else
        Set mdl = swApp.OpenDoc6(nativePath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then Exit Function

    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs
    EnsureSwHidden

    If TryComputePullcoreMinAreaAngleInViewPlane(mdl, viewName, tiltDeg, tightW, tightH, rawW, rawH) Then

        LogLine "PULLCORE box measured in the " & viewName & " plane:"
        LogLine "   as the model sits (no correction) = " & _
                FormatNumberForCsv(rawW) & " x " & FormatNumberForCsv(rawH)
        LogLine "   squared up (smallest box)         = " & _
                FormatNumberForCsv(tightW) & " x " & FormatNumberForCsv(tightH)
        LogLine "   tilt to remove                    = " & _
                Format(tiltDeg, "0.00") & " deg"

        detectedDegOut = tiltDeg
        TryComputePullcoreStraightenAngleFromNativeGeometry = True
    Else
        LogLine "PULLCORE: smallest-box measurement failed on " & nativePath
    End If

CleanExit:
    On Error Resume Next
    If Not mdl Is Nothing Then swApp.CloseDoc mdl.GetTitle
    Set mdl = Nothing
    Set fso = Nothing
    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreStraightenAngleFromNativeGeometry error: " & Err.Description
    Resume CleanExit
End Function

' Sweep for the in-plane rotation that minimises the bounding rectangle.
' Returns the tilt in degrees (what has to be REMOVED), plus the squared-up box
' and the uncorrected box, both in inches, for the log.
Private Function TryComputePullcoreMinAreaAngleInViewPlane(ByVal mdl As Object, _
                                                           ByVal viewName As String, _
                                                           ByRef angleDegOut As Double, _
                                                           ByRef tightWOut As Double, _
                                                           ByRef tightHOut As Double, _
                                                           ByRef rawWOut As Double, _
                                                           ByRef rawHOut As Double) As Boolean
On Error GoTo ErrHandler

    Dim swView As Object
    Dim mV As Variant
    Dim haveView As Boolean
    Dim bodyList As Collection
    Dim px() As Double
    Dim py() As Double
    Dim cap As Long
    Dim n As Long
    Dim bi As Long
    Dim ei As Long
    Dim k As Long
    Dim i As Long
    Dim body As Object
    Dim edge As Object
    Dim vEdges As Variant
    Dim vCP As Variant
    Dim ax As Double, ay As Double, az As Double
    Dim vx As Double, vy As Double
    Dim pass As Long
    Dim loDeg As Double, hiDeg As Double, stepDeg As Double
    Dim phi As Double
    Dim c As Double, s As Double
    Dim rx As Double, ry As Double
    Dim mnx As Double, mxx As Double, mny As Double, mxy As Double
    Dim bw As Double, bh As Double, ba As Double
    Dim bestPhi As Double, bestArea As Double, bestW As Double, bestH As Double
    Dim rawArea As Double

    TryComputePullcoreMinAreaAngleInViewPlane = False
    angleDegOut = 0#
    tightWOut = 0#
    tightHOut = 0#
    rawWOut = 0#
    rawHOut = 0#

    If mdl Is Nothing Then Exit Function

    ' Measure in the plane the drawing will actually be projected in.
    If viewName <> "" Then
        ShowNamedOrStandardView mdl, viewName
        StabilizeActiveView mdl, 60
    End If

    haveView = False
    Set swView = mdl.ActiveView

    If Not swView Is Nothing Then
        mV = swView.Orientation3.ArrayData
        If IsArray(mV) Then
            If UBound(mV) >= 8 Then haveView = True
        End If
    End If

    Set bodyList = New Collection
    CollectSolidBodiesFromModel mdl, bodyList
    If bodyList.count = 0 Then Exit Function

    cap = 8192
    n = 0
    ReDim px(1 To cap)
    ReDim py(1 To cap)

    For bi = 1 To bodyList.count
        Set body = bodyList(bi)
        If Not body Is Nothing Then
            vEdges = body.GetEdges
            If IsArray(vEdges) Then
                For ei = 0 To UBound(vEdges)
                    Set edge = vEdges(ei)
                    If Not edge Is Nothing Then
                        vCP = Empty
                        On Error Resume Next
                        vCP = edge.GetCurveParams2
                        On Error GoTo ErrHandler
                        If IsArray(vCP) Then
                            If UBound(vCP) >= 5 Then
                                ' vCP(0..2) = start point, vCP(3..5) = end point
                                For k = 0 To 3 Step 3
                                    ax = CDbl(vCP(k))
                                    ay = CDbl(vCP(k + 1))
                                    az = CDbl(vCP(k + 2))

                                    If haveView Then
                                        vx = ax * CDbl(mV(0)) + ay * CDbl(mV(3)) + az * CDbl(mV(6))
                                        vy = ax * CDbl(mV(1)) + ay * CDbl(mV(4)) + az * CDbl(mV(7))
                                    Else
                                        vx = ax
                                        vy = ay
                                    End If

                                    If n >= cap Then
                                        cap = cap * 2
                                        ReDim Preserve px(1 To cap)
                                        ReDim Preserve py(1 To cap)
                                    End If

                                    n = n + 1
                                    px(n) = vx
                                    py(n) = vy
                                Next k
                            End If
                        End If
                    End If
                Next ei
            End If
        End If
    Next bi

    If n < 8 Then Exit Function

    ' Uncorrected box, for the log and for the "is it already square" test.
    mnx = 1E+99: mxx = -1E+99
    mny = 1E+99: mxy = -1E+99

    For i = 1 To n
        If px(i) < mnx Then mnx = px(i)
        If px(i) > mxx Then mxx = px(i)
        If py(i) < mny Then mny = py(i)
        If py(i) > mxy Then mxy = py(i)
    Next i

    rawWOut = (mxx - mnx) * INCHES_PER_METER
    rawHOut = (mxy - mny) * INCHES_PER_METER
    rawArea = (mxx - mnx) * (mxy - mny)

    If rawArea <= 0.0000000001 Then Exit Function

    ' Coarse-to-fine sweep. A rectangle repeats every 90 deg, so -45..45 covers
    ' every distinct orientation.
    bestArea = 1E+99
    bestPhi = 0#
    bestW = mxx - mnx
    bestH = mxy - mny

    loDeg = -45#
    hiDeg = 45#
    stepDeg = 0.25

    For pass = 1 To 3

        phi = loDeg

        Do While phi <= hiDeg + 0.0000001

            c = Cos(phi * PI_VALUE / 180#)
            s = Sin(phi * PI_VALUE / 180#)

            mnx = 1E+99: mxx = -1E+99
            mny = 1E+99: mxy = -1E+99

            For i = 1 To n
                rx = px(i) * c - py(i) * s
                ry = px(i) * s + py(i) * c

                If rx < mnx Then mnx = rx
                If rx > mxx Then mxx = rx
                If ry < mny Then mny = ry
                If ry > mxy Then mxy = ry
            Next i

            bw = mxx - mnx
            bh = mxy - mny
            ba = bw * bh

            If ba < bestArea Then
                bestArea = ba
                bestPhi = phi
                bestW = bw
                bestH = bh
            End If

            phi = phi + stepDeg
        Loop

        loDeg = bestPhi - stepDeg
        hiDeg = bestPhi + stepDeg
        stepDeg = stepDeg / 10#

    Next pass

    tightWOut = bestW * INCHES_PER_METER
    tightHOut = bestH * INCHES_PER_METER

    ' Rotating the geometry by bestPhi squares it, so the tilt sitting on the
    ' part is the opposite of that. The caller multiplies by
    ' PULLCORE_STRAIGHTEN_SIGN to turn a tilt into a correction. If a job ever
    ' comes out mirrored (tilted twice as far instead of square), that constant
    ' is the one to flip.
    angleDegOut = -bestPhi

    ' Genuinely square parts: refuse to invent a rotation from rounding noise.
    If bestArea > rawArea * 0.9995 Then
        angleDegOut = 0#
    End If

    TryComputePullcoreMinAreaAngleInViewPlane = True
    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreMinAreaAngleInViewPlane error: " & Err.Description
    TryComputePullcoreMinAreaAngleInViewPlane = False
End Function


Private Function TryComputePullcoreLongestModelEdgeAngleFromModel(ByVal mdl As Object, _
                                                                  ByRef angleDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreLongestModelEdgeAngleFromModel = False
    angleDegOut = 0#

    If mdl Is Nothing Then Exit Function

    ' Measure the pull tilt in the ACTIVE-VIEW screen plane (the same plane the
    ' DXF straighten rotation acts in), not the body's local frame.
    '
    ' A pullcore tips about its LENGTH axis, so the length edges do NOT carry
    ' the pull - only the CROSS edges (width/thickness) rotate with it. We:
    '   1) find the part's longest edge -> its 3D direction is the length axis,
    '   2) among edges NOT parallel to that length axis, take the longest one
    '      that lies in the view plane,
    '   3) report that edge's deviation from the nearest screen axis (0 or 90),
    '      which is exactly how far the view must rotate to square the part.
    ' This gives 33 where the cross edge sits at 33, and a correct small value
    ' (not 90) when the body happens to be oriented with an edge near vertical.
    Dim swView As Object
    Set swView = mdl.ActiveView

    Dim mV As Variant
    Dim haveView As Boolean
    haveView = False
    If Not swView Is Nothing Then
        mV = swView.Orientation3.ArrayData
        If IsArray(mV) Then
            If UBound(mV) >= 8 Then haveView = True
        End If
    End If

    Dim bodyList As Collection
    Set bodyList = New Collection
    CollectSolidBodiesFromModel mdl, bodyList
    If bodyList.count = 0 Then Exit Function

    ' ---- pass 1: longest edge overall defines the length (pull) axis ----
    Dim lenDx As Double, lenDy As Double, lenDz As Double
    Dim lenMax As Double
    lenDx = 0#: lenDy = 0#: lenDz = 0#: lenMax = -1#

    Dim bi As Long, ei As Long
    Dim body As Object, edge As Object, curve As Object
    Dim vEdges As Variant, vCP As Variant
    Dim sx As Double, sy As Double, sz As Double
    Dim ex As Double, ey As Double, ez As Double
    Dim dX As Double, dY As Double, dZ As Double
    Dim fullLen As Double

    For bi = 1 To bodyList.count
        Set body = bodyList(bi)
        If Not body Is Nothing Then
            vEdges = body.GetEdges
            If IsArray(vEdges) Then
                For ei = 0 To UBound(vEdges)
                    Set edge = vEdges(ei)
                    If Not edge Is Nothing Then
                        Set curve = Nothing
                        On Error Resume Next
                        Set curve = edge.GetCurve
                        On Error GoTo ErrHandler
                        If Not curve Is Nothing Then
                            If curve.IsLine Then
                                vCP = edge.GetCurveParams2
                                If IsArray(vCP) Then
                                    If UBound(vCP) >= 5 Then
                                        dX = CDbl(vCP(3)) - CDbl(vCP(0))
                                        dY = CDbl(vCP(4)) - CDbl(vCP(1))
                                        dZ = CDbl(vCP(5)) - CDbl(vCP(2))
                                        fullLen = Sqr(dX * dX + dY * dY + dZ * dZ)
                                        If fullLen > lenMax Then
                                            lenMax = fullLen
                                            lenDx = dX: lenDy = dY: lenDz = dZ
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ei
            End If
        End If
    Next bi

    If lenMax <= 0.000001 Then Exit Function

    ' Unit length axis (for the parallel test).
    Dim uLx As Double, uLy As Double, uLz As Double
    uLx = lenDx / lenMax: uLy = lenDy / lenMax: uLz = lenDz / lenMax

    ' ---- pass 2: longest IN-VIEW edge that is NOT parallel to the length ----
    Dim bestScreenLen As Double
    Dim bestVX As Double, bestVY As Double
    Dim vdx As Double, vdy As Double, screenLen As Double
    Dim dotL As Double

    bestScreenLen = -1#: bestVX = 0#: bestVY = 0#

    For bi = 1 To bodyList.count
        Set body = bodyList(bi)
        If Not body Is Nothing Then
            vEdges = body.GetEdges
            If IsArray(vEdges) Then
                For ei = 0 To UBound(vEdges)
                    Set edge = vEdges(ei)
                    If Not edge Is Nothing Then
                        Set curve = Nothing
                        On Error Resume Next
                        Set curve = edge.GetCurve
                        On Error GoTo ErrHandler
                        If Not curve Is Nothing Then
                            If curve.IsLine Then
                                vCP = edge.GetCurveParams2
                                If IsArray(vCP) Then
                                    If UBound(vCP) >= 5 Then
                                        dX = CDbl(vCP(3)) - CDbl(vCP(0))
                                        dY = CDbl(vCP(4)) - CDbl(vCP(1))
                                        dZ = CDbl(vCP(5)) - CDbl(vCP(2))
                                        fullLen = Sqr(dX * dX + dY * dY + dZ * dZ)

                                        If fullLen > 0.000001 Then

                                            ' Skip edges parallel to the length
                                            ' axis (they do not carry the pull).
                                            dotL = Abs((dX * uLx + dY * uLy + dZ * uLz) / fullLen)

                                            If dotL < 0.9 Then

                                                If haveView Then
                                                    vdx = dX * CDbl(mV(0)) + dY * CDbl(mV(3)) + dZ * CDbl(mV(6))
                                                    vdy = dX * CDbl(mV(1)) + dY * CDbl(mV(4)) + dZ * CDbl(mV(7))
                                                Else
                                                    vdx = dX: vdy = dY
                                                End If

                                                screenLen = Sqr(vdx * vdx + vdy * vdy)

                                                If screenLen / fullLen >= 0.5 Then
                                                    If screenLen > bestScreenLen Then
                                                        bestScreenLen = screenLen
                                                        bestVX = vdx
                                                        bestVY = vdy
                                                    End If
                                                End If

                                            End If

                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ei
            End If
        End If
    Next bi

    If bestScreenLen <= 0.000001 Then Exit Function

    ' Raw angle of that cross edge in the screen plane.
    Dim rawDeg As Double
    rawDeg = RadToDeg(Atan2Safe(bestVY, bestVX))

    ' Fold to -90..90, then report the SIGNED deviation from the nearest screen
    ' axis (0 or 90). A cross edge ideally sits at +-90 to the length when the
    ' part is square; its deviation from that is the pull tilt. Equivalently we
    ' snap rawDeg to the nearest of {-90,0,90} and return the residual.
    Do While rawDeg > 90#
        rawDeg = rawDeg - 180#
    Loop
    Do While rawDeg < -90#
        rawDeg = rawDeg + 180#
    Loop

    Dim devH As Double
    Dim devV As Double
    devH = rawDeg
    If rawDeg >= 0# Then
        devV = rawDeg - 90#
    Else
        devV = rawDeg + 90#
    End If

    If Abs(devH) <= Abs(devV) Then
        angleDegOut = devH
    Else
        angleDegOut = devV
    End If

    TryComputePullcoreLongestModelEdgeAngleFromModel = True
    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreLongestModelEdgeAngleFromModel error: " & Err.Description
    TryComputePullcoreLongestModelEdgeAngleFromModel = False
End Function

Private Sub CollectSolidBodiesFromModel(ByVal mdl As Object, ByVal outBodies As Collection)
On Error Resume Next

    If mdl Is Nothing Then Exit Sub
    If outBodies Is Nothing Then Exit Sub

    If mdl.GetType = swDocPART Then

        Dim vB As Variant
        vB = mdl.GetBodies2(swSolidBody, False)

        If IsArray(vB) Then
            Dim i As Long
            For i = 0 To UBound(vB)
                If Not vB(i) Is Nothing Then outBodies.Add vB(i)
            Next i
        End If

    ElseIf mdl.GetType = swDocASSEMBLY Then

        Dim vComps As Variant
        vComps = mdl.GetComponents(False)

        If IsArray(vComps) Then

            Dim c As Long
            Dim comp As Object
            Dim cm As Object
            Dim vCB As Variant
            Dim j As Long

            For c = 0 To UBound(vComps)

                Set comp = vComps(c)

                If Not comp Is Nothing Then
                    If comp.IsSuppressed = False Then

                        Set cm = comp.GetModelDoc2

                        If Not cm Is Nothing Then
                            If cm.GetType = swDocPART Then

                                vCB = cm.GetBodies2(swSolidBody, False)

                                If IsArray(vCB) Then
                                    For j = 0 To UBound(vCB)
                                        If Not vCB(j) Is Nothing Then outBodies.Add vCB(j)
                                    Next j
                                End If

                            End If
                        End If

                    End If
                End If

            Next c

        End If

    End If
End Sub

Private Function TryComputeBBoxDimensionMathAngleFromModel(ByVal mdl As Object, _
                                                           ByRef angleDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputeBBoxDimensionMathAngleFromModel = False
    angleDegOut = 0#

    If mdl Is Nothing Then Exit Function

    If mdl.GetType = swDocPART Then
        TryComputeBBoxDimensionMathAngleFromModel = _
            TryComputeBBoxDimensionMathAngleForPart(mdl, angleDegOut)
        Exit Function
    End If

    If mdl.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = mdl.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function

    Dim bestVol As Double
    Dim bestAng As Double

    bestVol = -1#
    bestAng = 0#

    Dim i As Long
    Dim comp As Object
    Dim cm As Object
    Dim oneAng As Double
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    For i = 0 To UBound(vComps)

        Set comp = vComps(i)

        If Not comp Is Nothing Then
            If comp.IsSuppressed = False Then

                Set cm = comp.GetModelDoc2

                If Not cm Is Nothing Then
                    If cm.GetType = swDocPART Then

                        If TryComputeBBoxDimensionMathAngleForPart(cm, oneAng) Then

                            If GetPartBoundingBoxInches(cm, dX, dY, dZ) Then

                                Dim vol As Double
                                vol = dX * dY * dZ

                                If vol > bestVol Then
                                    bestVol = vol
                                    bestAng = oneAng
                                End If

                            End If

                        End If

                    End If
                End If

            End If
        End If

    Next i

    If bestVol > 0# Then
        angleDegOut = bestAng
        TryComputeBBoxDimensionMathAngleFromModel = True
    End If

    Exit Function

ErrHandler:
    TryComputeBBoxDimensionMathAngleFromModel = False
End Function

Private Function TryComputeBBoxDimensionMathAngleForPart(ByVal partModel As Object, _
                                                         ByRef angleDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputeBBoxDimensionMathAngleForPart = False
    angleDegOut = 0#

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    If GetPartBoundingBoxInches(partModel, dX, dY, dZ) = False Then Exit Function

    Dim normL As Double
    Dim normW As Double
    Dim normT As Double

    SortThreeDimensions dX, dY, dZ, normL, normW, normT

    Dim bestL As Double
    Dim bestW As Double
    Dim bestT As Double

    If TryCreateAndReadBestFitBoundingBox(partModel, bestL, bestW, bestT) = False Then Exit Function

    SortThreeDimensions bestL, bestW, bestT, bestL, bestW, bestT

    angleDegOut = EstimatePullcoreAngleFromFittedAndOriginalDeg(bestL, bestW, bestT, normL, normW, normT)

    TryComputeBBoxDimensionMathAngleForPart = (Abs(angleDegOut) > 0#)
    Exit Function

ErrHandler:
    TryComputeBBoxDimensionMathAngleForPart = False
End Function

Private Function SearchBoxRotationDeg(ByVal bestW As Double, ByVal bestT As Double, _
                                      ByVal normW As Double, ByVal normT As Double) As Double
On Error GoTo ErrHandler

    ' PRIMARY: solve the angle from the box PROPORTIONS, not absolute sizes.
    ' For a clean rotation:
    '     normW = bestW*cos + bestT*sin
    '     normT = bestW*sin + bestT*cos
    ' Adding and subtracting gives:
    '     cos+sin = (normW+normT) / (bestW+bestT)
    '     cos-sin = (normW-normT) / (bestW-bestT)
    ' so angle = atan2(sin, cos). Imported assembly boxes carry a small uniform
    ' skin on every face (their diagonal is larger than the fitted box's, which
    ' a true rotation can never do). That inflation wrecks an absolute-size
    ' search but largely cancels in these ratios, so this stays accurate.
    If (bestW + bestT) > 0.000001 And Abs(bestW - bestT) > 0.01 Then

        Dim cps As Double
        Dim cms As Double
        Dim cosV As Double
        Dim sinV As Double

        cps = (normW + normT) / (bestW + bestT)
        cms = (normW - normT) / (bestW - bestT)

        cosV = (cps + cms) / 2#
        sinV = (cps - cms) / 2#

        If Abs(cosV) > 0.000001 Or Abs(sinV) > 0.000001 Then

            Dim solvedDeg As Double
            solvedDeg = Atan2Safe(sinV, cosV) * 180# / PI_VALUE

            ' Fold into 0..45 (box has 90-degree symmetry; the pull is the
            ' acute tip).
            Do While solvedDeg < 0#
                solvedDeg = solvedDeg + 90#
            Loop
            Do While solvedDeg > 45#
                solvedDeg = 90# - solvedDeg
                If solvedDeg < 0# Then solvedDeg = 0#
            Loop

            SearchBoxRotationDeg = solvedDeg
            Exit Function

        End If

    End If

    ' FALLBACK: equal in-plane dims (bestW == bestT) make the proportion solve
    ' degenerate, so fall back to the absolute-size search.
    Dim bestAng As Double
    Dim bestErr As Double

    bestAng = 0#
    bestErr = 1E+99

    Dim a As Double
    Dim rad As Double
    Dim pW As Double
    Dim pT As Double
    Dim e As Double

    a = 0#

    Do While a <= 90#

        rad = a * PI_VALUE / 180#

        pW = bestW * Cos(rad) + bestT * Sin(rad)
        pT = bestW * Sin(rad) + bestT * Cos(rad)

        e = Abs(pW - normW) + Abs(pT - normT)

        If e < bestErr Then
            bestErr = e
            bestAng = a
        End If

        a = a + 1#

    Loop

    Dim lo As Double
    Dim hi As Double

    lo = bestAng - 1#
    hi = bestAng + 1#

    If lo < 0# Then lo = 0#
    If hi > 90# Then hi = 90#

    a = lo

    Do While a <= hi

        rad = a * PI_VALUE / 180#

        pW = bestW * Cos(rad) + bestT * Sin(rad)
        pT = bestW * Sin(rad) + bestT * Cos(rad)

        e = Abs(pW - normW) + Abs(pT - normT)

        If e < bestErr Then
            bestErr = e
            bestAng = a
        End If

        a = a + 0.1

    Loop

    SearchBoxRotationDeg = bestAng
    Exit Function

ErrHandler:
    SearchBoxRotationDeg = 0#
End Function

Private Function GetModelDocForCadIndex(ByVal cadIdx As Long) As Object
On Error GoTo ErrHandler

    Set GetModelDocForCadIndex = Nothing

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    If Not swModel Is Nothing Then

        If swModel.GetType = swDocASSEMBLY Then

            Dim swComp As Object
            Set swComp = FindAssemblyComponentByName(swModel, parts(cadIdx).componentName)

            If Not swComp Is Nothing Then
                Set GetModelDocForCadIndex = swComp.GetModelDoc2
                If Not GetModelDocForCadIndex Is Nothing Then Exit Function
            End If

        ElseIf swModel.GetType = swDocPART Then

            Set GetModelDocForCadIndex = swModel
            Exit Function

        End If

    End If

    If parts(cadIdx).filePath <> "" Then

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(parts(cadIdx).filePath) Then

            Dim errs As Long
            Dim warns As Long

            Set GetModelDocForCadIndex = swApp.OpenDoc6(parts(cadIdx).filePath, swDocPART, _
                                          swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, _
                                          parts(cadIdx).configName, errs, warns)

        End If

    End If

    Exit Function

ErrHandler:
    LogLine "GetModelDocForCadIndex error: " & Err.Description
    Set GetModelDocForCadIndex = Nothing
End Function

Private Function TryCreateAndReadBestFitBoundingBox(ByVal partModel As Object, _
                                                    ByRef outL As Double, _
                                                    ByRef outW As Double, _
                                                    ByRef outT As Double, _
                                                    Optional ByRef outAngleDeg As Double) As Boolean
On Error GoTo ErrHandler

    TryCreateAndReadBestFitBoundingBox = False
    LastBestFitFromPlainBBox = False
    outAngleDeg = 0#

    outL = 0#
    outW = 0#
    outT = 0#

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim bboxFeat As Object
    Set bboxFeat = Nothing

    partModel.ClearSelection2 True

    Dim featMgr As Object
    Set featMgr = partModel.FeatureManager

    If Not featMgr Is Nothing Then
        On Error Resume Next
        Set bboxFeat = featMgr.InsertGlobalBoundingBox(0, False, False, Nothing)
        If bboxFeat Is Nothing Then Set bboxFeat = featMgr.InsertGlobalBoundingBox(0, False, False, Nothing, False)
        If bboxFeat Is Nothing Then Set bboxFeat = featMgr.InsertGlobalBoundingBox(0, False, False)
        If bboxFeat Is Nothing Then Set bboxFeat = featMgr.InsertGlobalBoundingBox2(0, False, False, Nothing)
        On Error GoTo ErrHandler
    End If

    Dim gotDims As Boolean
    gotDims = False

    If Not bboxFeat Is Nothing Then

        partModel.ForceRebuild3 False

        gotDims = TryReadBoundingBoxCustomProps(partModel, outL, outW, outT)

        If gotDims = False Then
            gotDims = TryReadBoundingBoxFeatureSketchDims(bboxFeat, outL, outW, outT)
        End If

        On Error Resume Next
        partModel.ClearSelection2 True
        bboxFeat.Select2 False, 0
        partModel.EditDelete
        partModel.ClearSelection2 True
        On Error GoTo ErrHandler

    End If

    ' Geometry-based minimal box. Imported parts (e.g. STEP "_ASM_objects_")
    ' keep their world/assembly orientation inside the part file, so both the
    ' bbox feature and the plain part bbox can return the ROTATED (inflated)
    ' box. Searching the bodies' extreme points over single-axis rotations
    ' finds the true fitted box. Prefer it whenever it is smaller.
    Dim mbL As Double
    Dim mbW As Double
    Dim mbT As Double
    Dim mbAng As Double

    If TryComputeMinimalSingleAxisBBoxFromBodies(partModel, mbL, mbW, mbT, mbAng) Then

        If gotDims = False Then
            outL = mbL: outW = mbW: outT = mbT
            outAngleDeg = mbAng
            gotDims = True
            LogLine "Fitted bbox from body geometry: " & _
                    FormatNumberForCsv(outL) & "/" & FormatNumberForCsv(outW) & _
                    "/" & FormatNumberForCsv(outT) & "  (tip angle " & _
                    Format(mbAng, "0.0") & " deg)"
        ElseIf (mbL * mbW * mbT) < (outL * outW * outT) * 0.995 Then
            LogLine "Fitted bbox: body-geometry box " & _
                    FormatNumberForCsv(mbL) & "/" & FormatNumberForCsv(mbW) & "/" & _
                    FormatNumberForCsv(mbT) & " is smaller than feature box " & _
                    FormatNumberForCsv(outL) & "/" & FormatNumberForCsv(outW) & "/" & _
                    FormatNumberForCsv(outT) & " - using body-geometry box."
            outL = mbL: outW = mbW: outT = mbT
            outAngleDeg = mbAng
        End If

    End If

    If gotDims = False Then

        Dim dX As Double
        Dim dY As Double
        Dim dZ As Double

        If GetPartBoundingBoxInches(partModel, dX, dY, dZ) Then
            SortThreeDimensions dX, dY, dZ, outL, outW, outT
            gotDims = True
            LastBestFitFromPlainBBox = True
            LogLine "WARNING: fitted bbox unavailable - using plain part bbox " & _
                    "(may still be rotated/inflated)."
        End If

    End If

    TryCreateAndReadBestFitBoundingBox = (gotDims And outL > 0 And outW > 0 And outT > 0)
    Exit Function

ErrHandler:
    LogLine "TryCreateAndReadBestFitBoundingBox error: " & Err.Description

    On Error Resume Next
    If Not bboxFeat Is Nothing Then
        partModel.ClearSelection2 True
        bboxFeat.Select2 False, 0
        partModel.EditDelete
        partModel.ClearSelection2 True
    End If

    TryCreateAndReadBestFitBoundingBox = False
End Function

' True minimal bounding box computed straight from the solid bodies' extreme
' points (no bbox feature, no custom props). Candidate box axes are global
' X/Y/Z plus the normals of the largest planar faces, so block parts rotated
' about one OR several axes are still measured along their own faces. For
' each axis the in-plane directions are swept in coarse 3-degree steps and
' the winner refined in 0.5-degree steps; the minimum-volume box wins.
' Angle 0 about Z is included, so a straight part returns its normal box.
' Outputs inches, sorted L >= W >= T.
Private Function TryComputeMinimalSingleAxisBBoxFromBodies(ByVal partModel As Object, _
                                                           ByRef outL As Double, _
                                                           ByRef outW As Double, _
                                                           ByRef outT As Double, _
                                                           Optional ByRef outAngleDeg As Double) As Boolean
On Error GoTo ErrHandler

    TryComputeMinimalSingleAxisBBoxFromBodies = False

    outAngleDeg = 0#
    outL = 0#
    outW = 0#
    outT = 0#

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, True)

    If IsEmpty(vBodies) Then Exit Function
    If IsNull(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function

    ' ---- candidate axes: global X/Y/Z + largest planar face normals ----
    Dim axAr(1 To 12, 1 To 3) As Double
    Dim axN As Long
    axN = 0

    AddCandidateBoxAxis axAr, axN, 1#, 0#, 0#
    AddCandidateBoxAxis axAr, axN, 0#, 1#, 0#
    AddCandidateBoxAxis axAr, axN, 0#, 0#, 1#
    CollectPlanarFaceAxes vBodies, axAr, axN

    Dim bestVol As Double
    Dim bestAxis As Long
    Dim bestAng As Double
    Dim bestE0 As Double
    Dim bestE1 As Double
    Dim bestE2 As Double

    bestVol = 1E+99
    bestAxis = 0
    bestAng = 0#

    Dim k As Long
    Dim a As Double
    Dim e0 As Double
    Dim e1 As Double
    Dim e2 As Double
    Dim vol As Double

    For k = 1 To axN

        e0 = BodiesExtentAlongDir(vBodies, axAr(k, 1), axAr(k, 2), axAr(k, 3))

        If e0 > 0# Then

            a = 0#
            Do While a <= 89.5
                vol = ExtentsVolumeAboutAxis(vBodies, axAr(k, 1), axAr(k, 2), axAr(k, 3), e0, a, e1, e2)
                If vol < bestVol Then
                    bestVol = vol
                    bestAxis = k
                    bestAng = a
                    bestE0 = e0
                    bestE1 = e1
                    bestE2 = e2
                End If
                a = a + 3#
            Loop

        End If

    Next k

    If bestAxis = 0 Then Exit Function

    ' Refine around the coarse winner.
    Dim lo As Double
    Dim hi As Double

    lo = bestAng - 3#
    hi = bestAng + 3#
    If lo < 0# Then lo = 0#
    If hi > 89.9 Then hi = 89.9

    e0 = bestE0

    a = lo
    Do While a <= hi
        vol = ExtentsVolumeAboutAxis(vBodies, axAr(bestAxis, 1), axAr(bestAxis, 2), axAr(bestAxis, 3), e0, a, e1, e2)
        If vol < bestVol Then
            bestVol = vol
            bestAng = a
            bestE1 = e1
            bestE2 = e2
        End If
        a = a + 0.5
    Loop

    If bestVol >= 1E+98 Then Exit Function
    If bestE1 <= 0# Or bestE2 <= 0# Then Exit Function

    Dim dA As Double
    Dim dB As Double
    Dim dC As Double

    dA = bestE0 * INCHES_PER_METER
    dB = bestE1 * INCHES_PER_METER
    dC = bestE2 * INCHES_PER_METER

    SortThreeDimensions dA, dB, dC, outL, outW, outT

    ' PULL ANGLE = how far the part's own faces are tilted from the assembly
    ' axes (X/Y/Z), NOT the box-minimizing angle. Reconstruct the three axes
    ' of the winning fitted box: the kept axis normal plus the two in-plane
    ' directions rotated by bestAng. The tilt is the largest angle any of
    ' those box axes makes with its nearest global axis, folded to 0..45.
    Dim wnX As Double, wnY As Double, wnZ As Double
    wnX = axAr(bestAxis, 1): wnY = axAr(bestAxis, 2): wnZ = axAr(bestAxis, 3)

    Dim wuX As Double, wuY As Double, wuZ As Double
    Dim wvX As Double, wvY As Double, wvZ As Double
    BuildOrthoBasisForAxis wnX, wnY, wnZ, wuX, wuY, wuZ, wvX, wvY, wvZ

    Dim wrad As Double, wc As Double, ws As Double
    wrad = bestAng * PI_VALUE / 180#
    wc = Cos(wrad): ws = Sin(wrad)

    Dim b1x As Double, b1y As Double, b1z As Double
    Dim b2x As Double, b2y As Double, b2z As Double
    b1x = wc * wuX + ws * wvX: b1y = wc * wuY + ws * wvY: b1z = wc * wuZ + ws * wvZ
    b2x = -ws * wuX + wc * wvX: b2y = -ws * wuY + wc * wvY: b2z = -ws * wuZ + wc * wvZ

    Dim tilt As Double
    tilt = AxisTiltFromGlobalDeg(wnX, wnY, wnZ)
    If AxisTiltFromGlobalDeg(b1x, b1y, b1z) > tilt Then tilt = AxisTiltFromGlobalDeg(b1x, b1y, b1z)
    If AxisTiltFromGlobalDeg(b2x, b2y, b2z) > tilt Then tilt = AxisTiltFromGlobalDeg(b2x, b2y, b2z)

    Do While tilt > 45#
        tilt = 90# - tilt
        If tilt < 0# Then tilt = 0#
    Loop
    outAngleDeg = tilt

    TryComputeMinimalSingleAxisBBoxFromBodies = (outL > 0# And outW > 0# And outT > 0#)

    Exit Function

ErrHandler:
    LogLine "TryComputeMinimalSingleAxisBBoxFromBodies error: " & Err.Description
    TryComputeMinimalSingleAxisBBoxFromBodies = False
End Function

' Box volume (m^3) about the given unit axis with the in-plane basis rotated
' angDeg. Returns 1E+99 when extents cannot be measured.
Private Function ExtentsVolumeAboutAxis(ByRef vBodies As Variant, _
                                        ByVal nx As Double, _
                                        ByVal ny As Double, _
                                        ByVal nz As Double, _
                                        ByVal axisExtent As Double, _
                                        ByVal angDeg As Double, _
                                        ByRef e1 As Double, _
                                        ByRef e2 As Double) As Double
On Error GoTo ErrHandler

    ExtentsVolumeAboutAxis = 1E+99

    e1 = 0#
    e2 = 0#

    Dim uX As Double, uY As Double, uZ As Double
    Dim vx As Double, vy As Double, vz As Double

    BuildOrthoBasisForAxis nx, ny, nz, uX, uY, uZ, vx, vy, vz

    Dim rad As Double
    Dim c As Double
    Dim s As Double

    rad = angDeg * PI_VALUE / 180#
    c = Cos(rad)
    s = Sin(rad)

    Dim d1x As Double, d1y As Double, d1z As Double
    Dim d2x As Double, d2y As Double, d2z As Double

    d1x = c * uX + s * vx: d1y = c * uY + s * vy: d1z = c * uZ + s * vz
    d2x = -s * uX + c * vx: d2y = -s * uY + c * vy: d2z = -s * uZ + c * vz

    e1 = BodiesExtentAlongDir(vBodies, d1x, d1y, d1z)
    e2 = BodiesExtentAlongDir(vBodies, d2x, d2y, d2z)

    If e1 <= 0# Or e2 <= 0# Then Exit Function

    ExtentsVolumeAboutAxis = axisExtent * e1 * e2

    Exit Function

ErrHandler:
    ExtentsVolumeAboutAxis = 1E+99
End Function

' Builds two unit vectors orthogonal to the given (unit) axis.
Private Sub BuildOrthoBasisForAxis(ByVal nx As Double, ByVal ny As Double, ByVal nz As Double, _
                                   ByRef uX As Double, ByRef uY As Double, ByRef uZ As Double, _
                                   ByRef vx As Double, ByRef vy As Double, ByRef vz As Double)
On Error Resume Next

    Dim tx As Double, ty As Double, tz As Double

    If Abs(nx) < 0.9 Then
        tx = 1#: ty = 0#: tz = 0#
    Else
        tx = 0#: ty = 1#: tz = 0#
    End If

    ' u = normalize(cross(n, t))
    uX = ny * tz - nz * ty
    uY = nz * tx - nx * tz
    uZ = nx * ty - ny * tx

    Dim mag As Double
    mag = Sqr(uX * uX + uY * uY + uZ * uZ)
    If mag <= 0# Then Exit Sub
    uX = uX / mag: uY = uY / mag: uZ = uZ / mag

    ' v = cross(n, u)
    vx = ny * uZ - nz * uY
    vy = nz * uX - nx * uZ
    vz = nx * uY - ny * uX
End Sub

' Smallest angle (degrees, 0..90) between a unit direction and the nearest of
' the global X/Y/Z axes. Used to measure how far a fitted-box axis is tilted
' away from the assembly axes = the pull-core rotation.
Private Function AxisTiltFromGlobalDeg(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double
On Error Resume Next
    AxisTiltFromGlobalDeg = 0#
    Dim mag As Double
    mag = Sqr(x * x + y * y + z * z)
    If mag <= 0# Then Exit Function
    Dim mx As Double
    mx = Abs(x) / mag
    If Abs(y) / mag > mx Then mx = Abs(y) / mag
    If Abs(z) / mag > mx Then mx = Abs(z) / mag
    If mx > 1# Then mx = 1#
    ' Guard the Arccos identity against the divide-by-zero at mx = 1 (axis
    ' essentially aligned to a global axis => zero tilt).
    If mx >= 0.9999999 Then
        AxisTiltFromGlobalDeg = 0#
        Exit Function
    End If
    AxisTiltFromGlobalDeg = (Atn(-mx / Sqr(-mx * mx + 1#)) + 2# * Atn(1#)) * 180# / PI_VALUE
End Function

' Adds a (normalized) candidate axis, skipping near-duplicates.
Private Sub AddCandidateBoxAxis(ByRef axAr() As Double, ByRef axN As Long, _
                                ByVal x As Double, ByVal y As Double, ByVal z As Double)
On Error Resume Next

    Dim mag As Double
    mag = Sqr(x * x + y * y + z * z)

    If mag <= 0.000001 Then Exit Sub
    If axN >= 12 Then Exit Sub

    x = x / mag: y = y / mag: z = z / mag

    Dim k As Long
    Dim dT As Double

    For k = 1 To axN
        dT = Abs(x * axAr(k, 1) + y * axAr(k, 2) + z * axAr(k, 3))
        If dT > 0.999 Then Exit Sub
    Next k

    axN = axN + 1
    axAr(axN, 1) = x
    axAr(axN, 2) = y
    axAr(axN, 3) = z
End Sub

' Collects normals of the largest planar faces (up to 6) as candidate box
' axes, so a block tipped about several axes is measured along its own faces.
Private Sub CollectPlanarFaceAxes(ByRef vBodies As Variant, _
                                  ByRef axAr() As Double, ByRef axN As Long)
On Error GoTo done

    Dim topNx(1 To 6) As Double
    Dim topNy(1 To 6) As Double
    Dim topNz(1 To 6) As Double
    Dim topA(1 To 6) As Double

    Dim i As Long
    Dim f As Long
    Dim body As Object
    Dim vFaces As Variant
    Dim face As Object
    Dim surf As Object
    Dim vNorm As Variant
    Dim ar As Double
    Dim m As Long
    Dim minSlot As Long

    For i = LBound(vBodies) To UBound(vBodies)

        Set body = vBodies(i)
        If body Is Nothing Then GoTo nextBody

        vFaces = body.GetFaces
        If IsArray(vFaces) = False Then GoTo nextBody

        For f = LBound(vFaces) To UBound(vFaces)

            On Error Resume Next
            Set face = Nothing
            Set surf = Nothing
            Set face = vFaces(f)
            If Not face Is Nothing Then Set surf = face.GetSurface

            If Not surf Is Nothing Then
                If surf.IsPlane Then

                    ar = 0#
                    ar = face.GetArea
                    vNorm = face.Normal

                    If IsArray(vNorm) And ar > 0# Then

                        minSlot = 1
                        For m = 2 To 6
                            If topA(m) < topA(minSlot) Then minSlot = m
                        Next m

                        If ar > topA(minSlot) Then
                            topA(minSlot) = ar
                            topNx(minSlot) = CDbl(vNorm(0))
                            topNy(minSlot) = CDbl(vNorm(1))
                            topNz(minSlot) = CDbl(vNorm(2))
                        End If

                    End If

                End If
            End If
            On Error GoTo done

        Next f

nextBody:
    Next i

    For m = 1 To 6
        If topA(m) > 0# Then
            AddCandidateBoxAxis axAr, axN, topNx(m), topNy(m), topNz(m)
        End If
    Next m

done:
End Sub

' Total extent (meters) of all solid bodies along a unit direction, using
' Body2.GetExtremePoint in both directions.
Private Function BodiesExtentAlongDir(ByRef vBodies As Variant, _
                                      ByVal dX As Double, _
                                      ByVal dY As Double, _
                                      ByVal dZ As Double) As Double
On Error GoTo ErrHandler

    BodiesExtentAlongDir = 0#

    Dim pMax As Double
    Dim pMin As Double
    Dim gotAny As Boolean

    pMax = -1E+99
    pMin = 1E+99
    gotAny = False

    Dim i As Long
    Dim body As Object
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim proj As Double

    For i = LBound(vBodies) To UBound(vBodies)

        Set body = vBodies(i)

        If Not body Is Nothing Then

            If body.GetExtremePoint(dX, dY, dZ, x, y, z) Then
                proj = x * dX + y * dY + z * dZ
                If proj > pMax Then pMax = proj
                gotAny = True
            End If

            If body.GetExtremePoint(-dX, -dY, -dZ, x, y, z) Then
                proj = x * dX + y * dY + z * dZ
                If proj < pMin Then pMin = proj
                gotAny = True
            End If

        End If

    Next i

    If gotAny And pMax > pMin Then BodiesExtentAlongDir = pMax - pMin

    Exit Function

ErrHandler:
    BodiesExtentAlongDir = 0#
End Function

Private Function TryReadBoundingBoxCustomProps(ByVal partModel As Object, _
                                               ByRef outL As Double, _
                                               ByRef outW As Double, _
                                               ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryReadBoundingBoxCustomProps = False

    If partModel Is Nothing Then Exit Function

    Dim ext As Object
    Set ext = partModel.Extension

    If ext Is Nothing Then Exit Function

    Dim cpm As Object
    Set cpm = ext.CustomPropertyManager("")

    If cpm Is Nothing Then Exit Function

    Dim L As Double
    Dim W As Double
    Dim T As Double

    L = ReadBoundingBoxPropInches(cpm, Array("SW-BoundingBoxLength", "Total Bounding Box Length", "BoundingBoxLength"))
    W = ReadBoundingBoxPropInches(cpm, Array("SW-BoundingBoxWidth", "Total Bounding Box Width", "BoundingBoxWidth"))
    T = ReadBoundingBoxPropInches(cpm, Array("SW-BoundingBoxThickness", "Total Bounding Box Thickness", "BoundingBoxThickness", "SW-BoundingBoxHeight"))

    If L > 0 And W > 0 And T > 0 Then
        outL = L
        outW = W
        outT = T

        TryReadBoundingBoxCustomProps = True
    End If

    Exit Function

ErrHandler:
    TryReadBoundingBoxCustomProps = False
End Function

Private Function ReadBoundingBoxPropInches(ByVal cpm As Object, ByVal names As Variant) As Double
On Error Resume Next

    ReadBoundingBoxPropInches = 0#

    Dim nm As Variant
    Dim valOut As String
    Dim resolved As String

    For Each nm In names

        valOut = ""
        resolved = ""

        cpm.Get4 CStr(nm), False, valOut, resolved

        Dim useStr As String

        useStr = resolved
        If Trim(useStr) = "" Then useStr = valOut

        If Trim(useStr) <> "" Then

            Dim v As Double
            v = ParseLeadingNumber(useStr)

            If v > 0 Then
                If v < 3# Then
                    ReadBoundingBoxPropInches = v * INCHES_PER_METER
                Else
                    ReadBoundingBoxPropInches = v
                End If

                Exit Function
            End If

        End If

    Next nm
End Function

Private Function ParseLeadingNumber(ByVal s As String) As Double
On Error Resume Next

    Dim i As Long
    Dim ch As String
    Dim token As String

    token = ""

    For i = 1 To Len(s)

        ch = mid(s, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then
            token = token & ch
        ElseIf token <> "" Then
            Exit For
        End If

    Next i

    If token <> "" And IsNumeric(token) Then ParseLeadingNumber = CDbl(token)
End Function

Private Function TryReadBoundingBoxFeatureSketchDims(ByVal bboxFeat As Object, _
                                                     ByRef outL As Double, _
                                                     ByRef outW As Double, _
                                                     ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryReadBoundingBoxFeatureSketchDims = False

    outL = 0#
    outW = 0#
    outT = 0#

    If bboxFeat Is Nothing Then Exit Function

    Dim lengths As Collection
    Set lengths = New Collection

    CollectSketchSegmentLengthsFromFeature bboxFeat, lengths

    If lengths.count < 3 Then Exit Function

    Dim a As Double
    Dim b As Double
    Dim c As Double

    If GetThreeUniqueBoxEdgeLengths(lengths, a, b, c) = False Then Exit Function

    SortThreeDimensions a, b, c, outL, outW, outT

    TryReadBoundingBoxFeatureSketchDims = (outL > 0 And outW > 0 And outT > 0)
    Exit Function

ErrHandler:
    TryReadBoundingBoxFeatureSketchDims = False
End Function

Private Sub CollectSketchSegmentLengthsFromFeature(ByVal feat As Object, ByVal lengths As Collection)
On Error Resume Next

    If feat Is Nothing Then Exit Sub

    Dim spec As Object
    Set spec = feat.GetSpecificFeature2

    If Not spec Is Nothing Then

        Dim segs As Variant
        segs = spec.GetSketchSegments

        If IsEmpty(segs) = False Then

            Dim i As Long
            Dim seg As Object

            For i = 0 To UBound(segs)

                Set seg = segs(i)

                If Not seg Is Nothing Then

                    Dim p1 As Object
                    Dim p2 As Object

                    Set p1 = seg.GetStartPoint2
                    Set p2 = seg.GetEndPoint2

                    If Not p1 Is Nothing And Not p2 Is Nothing Then

                        Dim dX As Double
                        Dim dY As Double
                        Dim dZ As Double
                        Dim distIn As Double

                        dX = CDbl(p2.x) - CDbl(p1.x)
                        dY = CDbl(p2.y) - CDbl(p1.y)
                        dZ = CDbl(p2.z) - CDbl(p1.z)

                        distIn = Sqr(dX * dX + dY * dY + dZ * dZ) * INCHES_PER_METER

                        If distIn > 0.01 Then lengths.Add distIn

                    End If

                End If

            Next i

        End If

    End If

    Dim subFeat As Object
    Set subFeat = feat.GetFirstSubFeature

    Do While Not subFeat Is Nothing
        CollectSketchSegmentLengthsFromFeature subFeat, lengths
        Set subFeat = subFeat.GetNextSubFeature
    Loop
End Sub

Private Function GetThreeUniqueBoxEdgeLengths(ByVal lengths As Collection, _
                                              ByRef a As Double, ByRef b As Double, ByRef c As Double) As Boolean
On Error GoTo ErrHandler

    GetThreeUniqueBoxEdgeLengths = False

    a = 0#
    b = 0#
    c = 0#

    If lengths Is Nothing Then Exit Function
    If lengths.count < 3 Then Exit Function

    Dim uniques As Collection
    Set uniques = New Collection

    Dim i As Long

    For i = 1 To lengths.count
        AddUniqueLength uniques, CDbl(lengths(i)), 0.02
    Next i

    If uniques.count < 3 Then Exit Function

    Dim arr() As Double
    ReDim arr(1 To uniques.count)

    For i = 1 To uniques.count
        arr(i) = CDbl(uniques(i))
    Next i

    Dim j As Long
    Dim tmp As Double

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    a = arr(1)
    b = arr(2)
    c = arr(3)

    GetThreeUniqueBoxEdgeLengths = True
    Exit Function

ErrHandler:
    GetThreeUniqueBoxEdgeLengths = False
End Function

Private Sub AddUniqueLength(ByVal uniques As Collection, ByVal val As Double, ByVal tol As Double)
On Error Resume Next

    If val <= 0.01 Then Exit Sub

    Dim i As Long

    For i = 1 To uniques.count
        If Abs(CDbl(uniques(i)) - val) <= tol Then Exit Sub
    Next i

    uniques.Add val
End Sub

' ============================================================
' CAD NAMING LIBRARY
' Learns block names from BOM-matched jobs and uses them when no BOM exists.
' ============================================================

Private Function GetCadNamingLibraryPath() As String
On Error GoTo ErrHandler

    Dim root As String

    root = ResolveMatchingRoot()

    If root = "" Then root = LOCAL_WORKSPACE_ROOT

    EnsureFolderDeep root

    GetCadNamingLibraryPath = root & "\" & CAD_NAMING_LIBRARY_FILE
    Exit Function

ErrHandler:
    GetCadNamingLibraryPath = LOCAL_WORKSPACE_ROOT & "\" & CAD_NAMING_LIBRARY_FILE
End Function

Private Sub EnsureCadNamingLibraryHeader(ByVal libPath As String)
On Error GoTo ErrHandler

    If libPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    EnsureFolderDeep fso.GetParentFolderName(libPath)

    If fso.FileExists(libPath) Then Exit Sub

    Dim f As Integer
    f = FreeFile

    Open libPath For Output As #f

    Print #f, "Version,JobNumber,QuoteName,Length,Width,Thickness,CenterX,CenterY,CenterZ,NormX,NormY,NormZ,ComponentName,LearnedOn"

    Close #f

    LogLine "Created CAD naming library: " & libPath
    Exit Sub

ErrHandler:
    LogLine "EnsureCadNamingLibraryHeader error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub LearnCadNamingLibraryFromCurrentJob()
On Error GoTo ErrHandler

    If USE_CAD_NAMING_LIBRARY = False Then Exit Sub
    If LEARN_CAD_NAMING_LIBRARY_FROM_BOM = False Then Exit Sub
    If PartCount <= 0 Then Exit Sub

    Dim libPath As String
    libPath = GetCadNamingLibraryPath()

    EnsureCadNamingLibraryHeader libPath

    Dim existing As Object
    Set existing = LoadCadNamingLibraryKeyDict(libPath)

    Dim f As Integer
    f = FreeFile

    Open libPath For Append As #f

    Dim learnedCount As Long
    learnedCount = 0

    Dim i As Long

    ' Learn normal export rows.
    For i = 1 To ExportCount

        If ExportRows(i).HasCad Then

            If ShouldLearnCadLibraryQuote(ExportRows(i).quoteName) Then
                If AppendCadNamingLibraryRow(f, existing, ExportRows(i).quoteName, ExportRows(i).CadPartIndex) Then
                    learnedCount = learnedCount + 1
                End If
            End If

        End If

    Next i

    ' Learn pullcore matches too.
    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).CadPartIndex > 0 Then

            If ShouldLearnCadLibraryQuote(PullcoreMatches(i).quoteName) Then
                If AppendCadNamingLibraryRow(f, existing, PullcoreMatches(i).quoteName, PullcoreMatches(i).CadPartIndex) Then
                    learnedCount = learnedCount + 1
                End If
            End If

        End If

    Next i

    ' Learn package-style parts if found.
    Dim idx As Long

    idx = FindBestJBlockIndex()
    If idx > 0 Then
        If AppendCadNamingLibraryRow(f, existing, "J BLOCK", idx) Then learnedCount = learnedCount + 1
    End If

    idx = FindBestEjectorCamIndex()
    If idx > 0 Then
        If AppendCadNamingLibraryRow(f, existing, "EJECTOR CAM", idx) Then learnedCount = learnedCount + 1
    End If

    idx = FindBestPartByNameAndDims(PULLCORE_STOP_NAME_KEYS, 0, 0, 0, 0)
    If idx > 0 Then
        If AppendCadNamingLibraryRow(f, existing, "PULLCORE STOP", idx) Then learnedCount = learnedCount + 1
    End If

    idx = FindBestPartByNameAndDims(FLIPPER_CAM_COVER_NAME_KEYS, _
                                    FLIPPER_CAM_COVER_TARGET_THICKNESS, _
                                    FLIPPER_CAM_COVER_TARGET_WIDTH, _
                                    FLIPPER_CAM_COVER_TARGET_LENGTH, _
                                    FLIPPER_CAM_COVER_DIM_TOL)
    If idx > 0 Then
        If AppendCadNamingLibraryRow(f, existing, "FLIPPER CAM COVER PLATE", idx) Then learnedCount = learnedCount + 1
    End If

    Close #f

    LogLine "CAD naming library learned rows=" & learnedCount
    LogLine "CAD naming library path: " & libPath

    Exit Sub

ErrHandler:
    LogLine "LearnCadNamingLibraryFromCurrentJob error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Function ShouldLearnCadLibraryQuote(ByVal quoteName As String) As Boolean
On Error Resume Next

    Dim q As String
    q = NormalizeText(quoteName)

    If q = "" Then Exit Function
    If q = "MAIN ASSEMBLY" Then Exit Function
    If q = "HOLDERS" Then Exit Function

    ' Learn all meaningful named blocks/details.
    ShouldLearnCadLibraryQuote = True
End Function

Private Function AppendCadNamingLibraryRow(ByVal f As Integer, _
                                           ByVal existing As Object, _
                                           ByVal quoteName As String, _
                                           ByVal cadIdx As Long) As Boolean
On Error GoTo ErrHandler

    AppendCadNamingLibraryRow = False

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If Trim(quoteName) = "" Then Exit Function

    Dim key As String
    key = CurrentJobNumber & "|" & NormalizeKey(quoteName) & "|" & NormalizeKey(parts(cadIdx).componentName)

    If Not existing Is Nothing Then
        If existing.Exists(key) Then Exit Function
    End If

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim nx As Double
    Dim ny As Double
    Dim nz As Double

    cx = 0#: cy = 0#: cz = 0#
    nx = 0.5: ny = 0.5: nz = 0.5

    TryGetLibraryCadCenterPointInches cadIdx, True, cx, cy, cz
    TryGetCurrentPartNormLocation cadIdx, nx, ny, nz

    Print #f, CsvText("1") & "," & _
              CsvText(CurrentJobNumber) & "," & _
              CsvText(quoteName) & "," & _
              FormatNumberForCsv(parts(cadIdx).Length) & "," & _
              FormatNumberForCsv(parts(cadIdx).Width) & "," & _
              FormatNumberForCsv(parts(cadIdx).Thickness) & "," & _
              FormatNumberForCsv(cx) & "," & _
              FormatNumberForCsv(cy) & "," & _
              FormatNumberForCsv(cz) & "," & _
              FormatNumberForCsv(nx) & "," & _
              FormatNumberForCsv(ny) & "," & _
              FormatNumberForCsv(nz) & "," & _
              CsvText(parts(cadIdx).componentName) & "," & _
              CsvText(Format(Now, "yyyy-mm-dd hh:nn:ss"))

    If Not existing Is Nothing Then existing(key) = True

    LogLine "CAD library learned: " & quoteName & _
            " -> " & parts(cadIdx).componentName & _
            " L/W/T=" & FormatNumberForCsv(parts(cadIdx).Length) & "/" & _
            FormatNumberForCsv(parts(cadIdx).Width) & "/" & _
            FormatNumberForCsv(parts(cadIdx).Thickness) & _
            " Norm=" & FormatNumberForCsv(nx) & "/" & _
            FormatNumberForCsv(ny) & "/" & _
            FormatNumberForCsv(nz)

    AppendCadNamingLibraryRow = True
    Exit Function

ErrHandler:
    LogLine "AppendCadNamingLibraryRow error: " & Err.Description
    AppendCadNamingLibraryRow = False
End Function

Private Function LoadCadNamingLibraryKeyDict(ByVal libPath As String) As Object
On Error GoTo ErrHandler

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(libPath) = False Then
        Set LoadCadNamingLibraryKeyDict = dict
        Exit Function
    End If

    Dim ts As Object
    Set ts = fso.OpenTextFile(libPath, 1, False)

    Dim line As String
    Dim fields As Collection
    Dim key As String

    If Not ts.AtEndOfStream Then line = ts.ReadLine ' header

    Do While Not ts.AtEndOfStream

        line = ts.ReadLine

        Set fields = CsvSplitToCollection(line)

        If Not fields Is Nothing Then
            If fields.count >= 13 Then
                key = CStr(fields(2)) & "|" & NormalizeKey(CStr(fields(3))) & "|" & NormalizeKey(CStr(fields(13)))
                dict(key) = True
            End If
        End If

    Loop

    ts.Close

    Set LoadCadNamingLibraryKeyDict = dict
    Exit Function

ErrHandler:
    Set LoadCadNamingLibraryKeyDict = CreateObject("Scripting.Dictionary")
End Function

Private Function LoadCadNamingLibrary(ByRef lib() As CadNameLibEntry, _
                                      ByRef libCount As Long) As Boolean
On Error GoTo ErrHandler

    LoadCadNamingLibrary = False
    libCount = 0

    ReDim lib(1 To 1)

    Dim libPath As String
    libPath = GetCadNamingLibraryPath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(libPath) = False Then
        LogLine "CAD naming library not found: " & libPath
        Exit Function
    End If

    Dim ts As Object
    Set ts = fso.OpenTextFile(libPath, 1, False)

    Dim line As String
    Dim fields As Collection

    If Not ts.AtEndOfStream Then line = ts.ReadLine ' header

    Do While Not ts.AtEndOfStream

        line = ts.ReadLine

        If Trim(line) <> "" Then

            Set fields = CsvSplitToCollection(line)

            If Not fields Is Nothing Then

                If fields.count >= 13 Then

                    If val(CStr(fields(4))) > 0 And val(CStr(fields(5))) > 0 And val(CStr(fields(6))) > 0 Then

                        libCount = libCount + 1
                        ReDim Preserve lib(1 To libCount)

                        lib(libCount).sourceJob = CStr(fields(2))
                        lib(libCount).quoteName = CStr(fields(3))

                        lib(libCount).Length = CDbl(val(CStr(fields(4))))
                        lib(libCount).Width = CDbl(val(CStr(fields(5))))
                        lib(libCount).Thickness = CDbl(val(CStr(fields(6))))

                        lib(libCount).centerX = CDbl(val(CStr(fields(7))))
                        lib(libCount).centerY = CDbl(val(CStr(fields(8))))
                        lib(libCount).centerZ = CDbl(val(CStr(fields(9))))

                        lib(libCount).NormX = CDbl(val(CStr(fields(10))))
                        lib(libCount).NormY = CDbl(val(CStr(fields(11))))
                        lib(libCount).NormZ = CDbl(val(CStr(fields(12))))

                        lib(libCount).sourceComponent = CStr(fields(13))

                    End If

                End If

            End If

        End If

    Loop

    ts.Close

    LoadCadNamingLibrary = (libCount > 0)

    LogLine "Loaded CAD naming library rows=" & libCount & " from " & libPath
    Exit Function

ErrHandler:
    LogLine "LoadCadNamingLibrary error: " & Err.Description
    LoadCadNamingLibrary = False
    libCount = 0
End Function

Private Function IsNoBomLibraryPackageHandledQuote(ByVal quoteName As String) As Boolean
On Error GoTo ErrHandler

    If IsJBlockDxfQuote(quoteName) Then
        IsNoBomLibraryPackageHandledQuote = True
        Exit Function
    End If

    If IsEjectorCamPackageQuote(quoteName) Then
        IsNoBomLibraryPackageHandledQuote = True
        Exit Function
    End If

    ' PULLCORE cams and keys are package-handled too. Without this they came out
    ' of the CAD naming library as ordinary export rows, which sent them to
    ' MISC DETAILS and marked them used - so the pullcore package then reported
    ' "no pullcore matches" and produced nothing. That is what J8446 did: the
    ' names were right ("ID PULLCORE CAM"), the routing was not.
    If IsPullcoreDxfQuote(quoteName) Then
        IsNoBomLibraryPackageHandledQuote = True
        Exit Function
    End If

    IsNoBomLibraryPackageHandledQuote = False
    Exit Function

ErrHandler:
    IsNoBomLibraryPackageHandledQuote = False
End Function

Private Function IsEjectorCamPackageQuote(ByVal quoteName As String) As Boolean
On Error GoTo ErrHandler

    Dim s As String
    s = NormalizeText(quoteName)

    If s = "" Then Exit Function

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then Exit Function
    If InStr(s, "FLIPPER") > 0 Then Exit Function
    If InStr(s, "COVER") > 0 Then Exit Function
    If InStr(s, "J-BLOCK") > 0 Or InStr(s, "J BLOCK") > 0 Or InStr(s, "JBLOCK") > 0 Then Exit Function

    If InStr(s, "EJECTOR CAM") > 0 Then
        IsEjectorCamPackageQuote = True
        Exit Function
    End If

    If InStr(s, "EJ CAM") > 0 Or InStr(s, "EJ. CAM") > 0 Or InStr(s, "EJ.CAM") > 0 Then
        IsEjectorCamPackageQuote = True
        Exit Function
    End If

    ' Library may learn/export this simply as "CAM".
    If s = "CAM" Or s = "EJECTOR CAM" Then
        IsEjectorCamPackageQuote = True
        Exit Function
    End If

    IsEjectorCamPackageQuote = False
    Exit Function

ErrHandler:
    IsEjectorCamPackageQuote = False
End Function

Private Sub RecordSpecialCadMatchDirect(ByVal quoteName As String, ByVal cadIdx As Long)
On Error GoTo ErrHandler

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Sub

    If SpecialBomCadMatches Is Nothing Then
        Set SpecialBomCadMatches = CreateObject("Scripting.Dictionary")
    End If

    If SpecialBomCadQuoteNames Is Nothing Then
        Set SpecialBomCadQuoteNames = CreateObject("Scripting.Dictionary")
    End If

    PutSpecialCadAlias quoteName, cadIdx

    If IsJBlockDxfQuote(quoteName) Then
        PutSpecialCadAlias "J BLOCK", cadIdx
        PutSpecialCadAlias "J-BLOCK", cadIdx
        PutSpecialCadAlias "JBLOCK", cadIdx
        PutSpecialCadAlias "EJECTOR J-BLOCK", cadIdx
        PutSpecialCadAlias "EJ J-BLOCK", cadIdx
    End If

    If IsEjectorCamPackageQuote(quoteName) Then
        PutSpecialCadAlias "EJECTOR CAM", cadIdx
        PutSpecialCadAlias "EJ CAM", cadIdx
        PutSpecialCadAlias "EJ. CAM", cadIdx
        PutSpecialCadAlias "EJ.CAM", cadIdx
        PutSpecialCadAlias "CAM", cadIdx
    End If

    LogLine "Recorded special CAD package match: '" & quoteName & _
            "' -> " & parts(cadIdx).componentName

    Exit Sub

ErrHandler:
    LogLine "RecordSpecialCadMatchDirect error: " & Err.Description
End Sub

Private Sub PutSpecialCadAlias(ByVal quoteName As String, ByVal cadIdx As Long)
On Error Resume Next

    If quoteName = "" Then Exit Sub
    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Sub

    If SpecialBomCadMatches Is Nothing Then
        Set SpecialBomCadMatches = CreateObject("Scripting.Dictionary")
    End If

    If SpecialBomCadQuoteNames Is Nothing Then
        Set SpecialBomCadQuoteNames = CreateObject("Scripting.Dictionary")
    End If

    SpecialBomCadMatches(NormalizeKey(quoteName)) = cadIdx
    SpecialBomCadQuoteNames(NormalizeKey(quoteName)) = quoteName
End Sub

Private Function GetSpecialCadMatchByQuoteKeys(ByVal pipeKeys As String) As Long
On Error GoTo ErrHandler

    GetSpecialCadMatchByQuoteKeys = 0

    If SpecialBomCadMatches Is Nothing Then Exit Function
    If pipeKeys = "" Then Exit Function

    Dim keys() As String
    keys = Split(pipeKeys, "|")

    Dim i As Long
    Dim k As String
    Dim cadIdx As Long

    For i = LBound(keys) To UBound(keys)

        k = NormalizeKey(CStr(keys(i)))

        If k <> "" Then
            If SpecialBomCadMatches.Exists(k) Then

                cadIdx = CLng(SpecialBomCadMatches(k))

                If cadIdx > 0 And cadIdx <= PartCount Then
                    GetSpecialCadMatchByQuoteKeys = cadIdx
                    Exit Function
                End If

            End If
        End If

    Next i

    Exit Function

ErrHandler:
    GetSpecialCadMatchByQuoteKeys = 0
End Function

Private Function HasSpecialPackageCadMatches() As Boolean
On Error GoTo ErrHandler

    HasSpecialPackageCadMatches = False

    If GetSpecialCadMatchByQuoteKeys("EJECTOR J-BLOCK|EJ J-BLOCK|J-BLOCK|J BLOCK|JBLOCK") > 0 Then
        HasSpecialPackageCadMatches = True
        Exit Function
    End If

    If GetSpecialCadMatchByQuoteKeys("EJECTOR CAM|EJ CAM|EJ. CAM|EJ.CAM|CAM") > 0 Then
        HasSpecialPackageCadMatches = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    HasSpecialPackageCadMatches = False
End Function

Private Function TryBuildExportRowsFromCadNamingLibrary() As Boolean
On Error GoTo ErrHandler

    TryBuildExportRowsFromCadNamingLibrary = False

    If USE_CAD_NAMING_LIBRARY = False Then Exit Function
    If PartCount <= 0 Then Exit Function

    Dim lib() As CadNameLibEntry
    Dim libCount As Long

    If LoadCadNamingLibrary(lib, libCount) = False Then
        LogLine "No-BOM library fallback failed: library empty or missing."
        Exit Function
    End If

    ' Reserve package-handled parts (J BLOCK, EJECTOR CAM) using the same
    ' finders as the BOM flow, BEFORE the size-only matcher can consume them.
    If CAD_LIB_RESERVE_PACKAGE_PARTS_FIRST Then
        ReservePackagePartsBeforeLibraryFallback
    End If

    ' If this model is a near-exact re-run of one library source job (common
    ' when the PDFs are missing), prefer that job's names for every part.
    CadLibDominantSourceJob = ""
    If CAD_LIB_PREFER_DOMINANT_SOURCE_JOB Then
        CadLibDominantSourceJob = FindDominantLibrarySourceJob(lib, libCount)
    End If

    Dim usedQuoteCounts As Object
    Set usedQuoteCounts = CreateObject("Scripting.Dictionary")

    Dim matchedCount As Long
    matchedCount = 0

    Dim i As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim bestQuote As String
            Dim bestScore As Double
            Dim bestDimDiff As Double
            Dim bestLocDist As Double
            Dim secondScore As Double

            If FindBestLibraryQuoteForCadPart(i, lib, libCount, usedQuoteCounts, _
                                               bestQuote, bestScore, secondScore, _
                                               bestDimDiff, bestLocDist) Then

                If bestScore + CAD_LIB_AMBIGUOUS_SCORE_GAP > secondScore Then
                    LogLine "WARNING: CAD library ambiguous match for " & parts(i).componentName & _
                            ". Best='" & bestQuote & "' score=" & FormatNumberForCsv(bestScore) & _
                            " second=" & FormatNumberForCsv(secondScore)
                End If

                ' IMPORTANT:
                ' J BLOCK and EJECTOR CAM are package-handled items.
                ' Do NOT add them as normal export rows, or they go to MISC DETAILS
                ' and the package DXF flow may never run.
                If IsNoBomLibraryPackageHandledQuote(bestQuote) Then

                    RecordSpecialCadMatchDirect bestQuote, i

                    ' The pullcore package reads PullcoreMatches, not the special
                    ' CAD match table, so a pullcore also has to be registered
                    ' there or the package still finds nothing to export.
                    If IsPullcoreDxfQuote(bestQuote) Then
                        AddPullcoreMatchFromLibrary bestQuote, i
                    End If

                    ' Mark used so this library fallback does not treat the same part as a
                    ' normal export item, but package finders can still retrieve it through
                    ' SpecialBomCadMatches.
                    parts(i).UsedForBomMatch = True

                    IncrementDictionaryLong usedQuoteCounts, NormalizeKey(bestQuote)

                    matchedCount = matchedCount + 1

                    LogLine "CAD library matched PACKAGE item: " & parts(i).componentName & _
                            " -> '" & bestQuote & "'" & _
                            " score=" & FormatNumberForCsv(bestScore) & _
                            " dimDiff=" & FormatNumberForCsv(bestDimDiff) & _
                            " locDist=" & FormatNumberForCsv(bestLocDist) & _
                            " L/W/T=" & FormatNumberForCsv(parts(i).Length) & "/" & _
                            FormatNumberForCsv(parts(i).Width) & "/" & _
                            FormatNumberForCsv(parts(i).Thickness)

                Else

                    AddExportRowFromCadPart bestQuote, i
                    parts(i).UsedForBomMatch = True

                    IncrementDictionaryLong usedQuoteCounts, NormalizeKey(bestQuote)

                    matchedCount = matchedCount + 1

                    LogLine "CAD library matched: " & parts(i).componentName & _
                            " -> '" & bestQuote & "'" & _
                            " score=" & FormatNumberForCsv(bestScore) & _
                            " dimDiff=" & FormatNumberForCsv(bestDimDiff) & _
                            " locDist=" & FormatNumberForCsv(bestLocDist) & _
                            " L/W/T=" & FormatNumberForCsv(parts(i).Length) & "/" & _
                            FormatNumberForCsv(parts(i).Width) & "/" & _
                            FormatNumberForCsv(parts(i).Thickness)

                End If

            End If

        End If

    Next i

    TryBuildExportRowsFromCadNamingLibrary = (matchedCount > 0)

    LogLine "CAD naming library fallback matched count=" & matchedCount
    Exit Function

ErrHandler:
    LogLine "TryBuildExportRowsFromCadNamingLibrary error: " & Err.Description
    TryBuildExportRowsFromCadNamingLibrary = False
End Function

' In no-BOM jobs the library matcher scores by size/location only, so the
' J BLOCK or EJECTOR CAM could win some other job's detail name, get exported
' as a normal row (-> MISC DETAILS) and be marked used, which then hid it from
' the package finders (no J BLOCK / CAM DXF). Reserve them first with the SAME
' finders the BOM flow uses so categorization stays identical either way.
Private Sub ReservePackagePartsBeforeLibraryFallback()
On Error GoTo ErrHandler

    If CREATE_J_BLOCK_PACKAGE = False Then Exit Sub

    Dim jbIdx As Long
    jbIdx = FindBestJBlockIndex()

    If jbIdx > 0 And jbIdx <= PartCount Then
        RecordSpecialCadMatchDirect "J BLOCK", jbIdx
        parts(jbIdx).UsedForBomMatch = True
        LogLine "No-BOM library: reserved J BLOCK for package flow: " & _
                parts(jbIdx).componentName
    End If

    Dim camIdx As Long
    camIdx = FindBestEjectorCamIndex()

    If camIdx > 0 And camIdx <= PartCount Then
        If parts(camIdx).UsedForBomMatch = False Then
            RecordSpecialCadMatchDirect "EJECTOR CAM", camIdx
            parts(camIdx).UsedForBomMatch = True
            LogLine "No-BOM library: reserved EJECTOR CAM for package flow: " & _
                    parts(camIdx).componentName
        End If
    End If

    Exit Sub

ErrHandler:
    LogLine "ReservePackagePartsBeforeLibraryFallback error: " & Err.Description
End Sub

' Find the library source job that this model matches almost exactly: for
' each CAD part, credit every source job holding an entry within
' CAD_LIB_DOMINANT_JOB_PERFECT_DIM_DIFF total size difference (once per
' part). The job with the most credits, if it has at least
' CAD_LIB_DOMINANT_JOB_MIN_HITS, is the dominant job.
Private Function FindDominantLibrarySourceJob(ByRef lib() As CadNameLibEntry, _
                                              ByVal libCount As Long) As String
On Error GoTo ErrHandler

    FindDominantLibrarySourceJob = ""

    If libCount <= 0 Then Exit Function
    If PartCount <= 0 Then Exit Function

    Dim jobHits As Object
    Set jobHits = CreateObject("Scripting.Dictionary")
    jobHits.CompareMode = 1

    Dim creditedJobs As Object
    Dim i As Long
    Dim j As Long
    Dim dd As Double
    Dim jobKey As String

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Set creditedJobs = CreateObject("Scripting.Dictionary")
            creditedJobs.CompareMode = 1

            For j = 1 To libCount

                dd = Abs(parts(i).Length - lib(j).Length) + _
                     Abs(parts(i).Width - lib(j).Width) + _
                     Abs(parts(i).Thickness - lib(j).Thickness)

                If dd <= CAD_LIB_DOMINANT_JOB_PERFECT_DIM_DIFF Then

                    jobKey = Trim(lib(j).sourceJob)

                    If jobKey <> "" Then
                        If creditedJobs.Exists(jobKey) = False Then
                            creditedJobs(jobKey) = True

                            If jobHits.Exists(jobKey) Then
                                jobHits(jobKey) = CLng(jobHits(jobKey)) + 1
                            Else
                                jobHits(jobKey) = 1
                            End If
                        End If
                    End If

                End If

            Next j

        End If

    Next i

    Dim bestJob As String
    Dim bestHits As Long

    bestJob = ""
    bestHits = 0

    Dim k As Variant

    For Each k In jobHits.keys
        If CLng(jobHits(k)) > bestHits Then
            bestHits = CLng(jobHits(k))
            bestJob = CStr(k)
        End If
    Next k

    If bestHits >= CAD_LIB_DOMINANT_JOB_MIN_HITS Then
        FindDominantLibrarySourceJob = bestJob
        LogLine "CAD library: dominant matching source job = " & bestJob & _
                " (" & bestHits & " near-exact size matches). Preferring its names."
    Else
        LogLine "CAD library: no dominant source job (best='" & bestJob & _
                "' hits=" & bestHits & ")."
    End If

    Exit Function

ErrHandler:
    LogLine "FindDominantLibrarySourceJob error: " & Err.Description
    FindDominantLibrarySourceJob = ""
End Function

Private Function FindBestLibraryQuoteForCadPart(ByVal cadIdx As Long, _
                                                ByRef lib() As CadNameLibEntry, _
                                                ByVal libCount As Long, _
                                                ByVal usedQuoteCounts As Object, _
                                                ByRef bestQuote As String, _
                                                ByRef bestScore As Double, _
                                                ByRef secondScore As Double, _
                                                ByRef bestDimDiff As Double, _
                                                ByRef bestLocDist As Double) As Boolean
On Error GoTo ErrHandler

    FindBestLibraryQuoteForCadPart = False

    bestQuote = ""
    bestScore = 1E+99
    secondScore = 1E+99
    bestDimDiff = 0#
    bestLocDist = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If libCount <= 0 Then Exit Function

    Dim j As Long

    For j = 1 To libCount

        If ShouldUseCadLibraryQuoteForFallback(lib(j).quoteName) Then

            Dim qKey As String
            qKey = NormalizeKey(lib(j).quoteName)

            Dim alreadyUsed As Long
            alreadyUsed = 0

            If Not usedQuoteCounts Is Nothing Then
                If usedQuoteCounts.Exists(qKey) Then alreadyUsed = CLng(usedQuoteCounts(qKey))
            End If

            If alreadyUsed < CAD_LIB_MAX_MATCHES_PER_QUOTE Then

                Dim dimDiff As Double
                Dim locDist As Double
                Dim sc As Double

                sc = CadLibraryScoreForPart(cadIdx, lib(j), dimDiff, locDist)

                ' Prefer names from the dominant (near-exact matching) source
                ' job so a re-run job is categorized like the original.
                If CadLibDominantSourceJob <> "" And sc < 1E+98 Then
                    If StrComp(Trim(lib(j).sourceJob), CadLibDominantSourceJob, vbTextCompare) = 0 Then
                        sc = sc - CAD_LIB_DOMINANT_JOB_SCORE_BONUS
                    End If
                End If

                If sc < bestScore Then
                    secondScore = bestScore
                    bestScore = sc
                    bestQuote = lib(j).quoteName
                    bestDimDiff = dimDiff
                    bestLocDist = locDist
                ElseIf sc < secondScore Then
                    secondScore = sc
                End If

            End If

        End If

    Next j

    If bestQuote <> "" Then
        If bestScore <= CAD_LIB_MAX_SCORE Then
            FindBestLibraryQuoteForCadPart = True
        End If
    End If

    Exit Function

ErrHandler:
    FindBestLibraryQuoteForCadPart = False
End Function

Private Function ShouldUseCadLibraryQuoteForFallback(ByVal quoteName As String) As Boolean
On Error Resume Next

    Dim q As String
    q = NormalizeText(quoteName)

    If q = "" Then Exit Function
    If q = "MAIN ASSEMBLY" Then Exit Function
    If q = "HOLDERS" Then Exit Function

    ShouldUseCadLibraryQuoteForFallback = True
End Function

Private Function CadLibraryScoreForPart(ByVal cadIdx As Long, _
                                        ByRef e As CadNameLibEntry, _
                                        ByRef dimDiffOut As Double, _
                                        ByRef locDistOut As Double) As Double
On Error GoTo ErrHandler

    CadLibraryScoreForPart = 1E+99
    dimDiffOut = 0#
    locDistOut = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    Dim dL As Double
    Dim dW As Double
    Dim dT As Double

    dL = Abs(parts(cadIdx).Length - e.Length)
    dW = Abs(parts(cadIdx).Width - e.Width)
    dT = Abs(parts(cadIdx).Thickness - e.Thickness)

    If dL > CAD_LIB_MAX_SINGLE_DIM_DIFF_IN Then Exit Function
    If dW > CAD_LIB_MAX_SINGLE_DIM_DIFF_IN Then Exit Function
    If dT > CAD_LIB_MAX_SINGLE_DIM_DIFF_IN Then Exit Function

    dimDiffOut = dL + dW + dT

    If dimDiffOut > CAD_LIB_MAX_TOTAL_DIM_DIFF_IN Then Exit Function

    Dim nx As Double
    Dim ny As Double
    Dim nz As Double

    If TryGetCurrentPartNormLocation(cadIdx, nx, ny, nz) Then

        locDistOut = Sqr((nx - e.NormX) ^ 2 + _
                         (ny - e.NormY) ^ 2 + _
                         (nz - e.NormZ) ^ 2)

        If locDistOut > CAD_LIB_MAX_NORM_DISTANCE Then Exit Function

    Else

        ' If location is unavailable, let size decide but penalize it.
        locDistOut = 0.25

    End If

    CadLibraryScoreForPart = _
        (dimDiffOut * CAD_LIB_SCORE_DIM_WEIGHT) + _
        (locDistOut * CAD_LIB_SCORE_LOC_WEIGHT)

    Exit Function

ErrHandler:
    CadLibraryScoreForPart = 1E+99
End Function

Private Function TryGetCurrentPartNormLocation(ByVal cadIdx As Long, _
                                               ByRef nx As Double, _
                                               ByRef ny As Double, _
                                               ByRef nz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetCurrentPartNormLocation = False

    nx = 0.5
    ny = 0.5
    nz = 0.5

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    Dim minX As Double
    Dim minY As Double
    Dim minZ As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim maxZ As Double

    If TryGetCurrentCadCenterBounds(minX, minY, minZ, maxX, maxY, maxZ) = False Then Exit Function

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    If TryGetLibraryCadCenterPointInches(cadIdx, True, cx, cy, cz) = False Then Exit Function

    If Abs(maxX - minX) > 0.000001 Then nx = (cx - minX) / (maxX - minX)
    If Abs(maxY - minY) > 0.000001 Then ny = (cy - minY) / (maxY - minY)
    If Abs(maxZ - minZ) > 0.000001 Then nz = (cz - minZ) / (maxZ - minZ)

    If nx < 0# Then nx = 0#
    If nx > 1# Then nx = 1#

    If ny < 0# Then ny = 0#
    If ny > 1# Then ny = 1#

    If nz < 0# Then nz = 0#
    If nz > 1# Then nz = 1#

    TryGetCurrentPartNormLocation = True
    Exit Function

ErrHandler:
    TryGetCurrentPartNormLocation = False
End Function

Private Function TryGetCurrentCadCenterBounds(ByRef minX As Double, _
                                              ByRef minY As Double, _
                                              ByRef minZ As Double, _
                                              ByRef maxX As Double, _
                                              ByRef maxY As Double, _
                                              ByRef maxZ As Double) As Boolean
On Error GoTo ErrHandler

    TryGetCurrentCadCenterBounds = False

    Dim firstVal As Boolean
    firstVal = True

    Dim i As Long
    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    For i = 1 To PartCount

        If TryGetLibraryCadCenterPointInches(i, True, cx, cy, cz) Then

            If firstVal Then
                minX = cx: maxX = cx
                minY = cy: maxY = cy
                minZ = cz: maxZ = cz
                firstVal = False
            Else
                If cx < minX Then minX = cx
                If cx > maxX Then maxX = cx

                If cy < minY Then minY = cy
                If cy > maxY Then maxY = cy

                If cz < minZ Then minZ = cz
                If cz > maxZ Then maxZ = cz
            End If

        End If

    Next i

    TryGetCurrentCadCenterBounds = Not firstVal
    Exit Function

ErrHandler:
    TryGetCurrentCadCenterBounds = False
End Function

Private Function TryGetLibraryCadCenterPointInches(ByVal cadIdx As Long, _
                                                   ByVal preferMassCenter As Boolean, _
                                                   ByRef cx As Double, _
                                                   ByRef cy As Double, _
                                                   ByRef cz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetLibraryCadCenterPointInches = False

    cx = 0#
    cy = 0#
    cz = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    If preferMassCenter Then
        If parts(cadIdx).hasMassCenter Then
            cx = parts(cadIdx).MassCenterX
            cy = parts(cadIdx).MassCenterY
            cz = parts(cadIdx).MassCenterZ
            TryGetLibraryCadCenterPointInches = True
            Exit Function
        End If
    End If

    If parts(cadIdx).hasAsmCenter Then
        cx = parts(cadIdx).AsmCenterX
        cy = parts(cadIdx).AsmCenterY
        cz = parts(cadIdx).AsmCenterZ
        TryGetLibraryCadCenterPointInches = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    TryGetLibraryCadCenterPointInches = False
End Function

Private Sub IncrementDictionaryLong(ByVal dict As Object, ByVal key As String)
On Error Resume Next

    If dict Is Nothing Then Exit Sub
    If key = "" Then Exit Sub

    If dict.Exists(key) Then
        dict(key) = CLng(dict(key)) + 1
    Else
        dict(key) = 1
    End If
End Sub

Private Function CsvSplitToCollection(ByVal line As String) As Collection
On Error GoTo ErrHandler

    Dim result As New Collection

    Dim i As Long
    Dim ch As String
    Dim token As String
    Dim inQuotes As Boolean

    token = ""
    inQuotes = False

    i = 1

    Do While i <= Len(line)

        ch = mid(line, i, 1)

        If ch = Chr(34) Then

            If inQuotes And i < Len(line) Then
                If mid(line, i + 1, 1) = Chr(34) Then
                    token = token & Chr(34)
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = True
            End If

        ElseIf ch = "," And inQuotes = False Then

            result.Add token
            token = ""

        Else

            token = token & ch

        End If

        i = i + 1

    Loop

    result.Add token

    Set CsvSplitToCollection = result
    Exit Function

ErrHandler:
    Set CsvSplitToCollection = New Collection
End Function

' ============================================================
' CSV REPORTS
' ============================================================

Private Sub WritePartDimensionCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    csvPath = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open csvPath For Output As #f

    Print #f, "Index,Component,CleanName,Length,Width,Thickness,OriginalAsmLength,OriginalAsmWidth,OriginalAsmThickness,HasOriginalAsmBBox,BBoxVolume,Mass,Qty,BodyOnly"

    Dim i As Long

    For i = 1 To PartCount
        Print #f, CsvText(CStr(i)) & "," & _
                  CsvText(parts(i).componentName) & "," & _
                  CsvText(parts(i).cleanName) & "," & _
                  FormatNumberForCsv(parts(i).Length) & "," & _
                  FormatNumberForCsv(parts(i).Width) & "," & _
                  FormatNumberForCsv(parts(i).Thickness) & "," & _
                  FormatNumberForCsv(parts(i).OriginalAsmLength) & "," & _
                  FormatNumberForCsv(parts(i).OriginalAsmWidth) & "," & _
                  FormatNumberForCsv(parts(i).OriginalAsmThickness) & "," & _
                  CStr(parts(i).hasOriginalAsmBBox) & "," & _
                  FormatNumberForCsv(parts(i).BBoxVolume) & "," & _
                  FormatNumberForCsv(parts(i).massValue) & "," & _
                  CStr(parts(i).Quantity) & "," & _
                  CStr(parts(i).isBodyOnly)
    Next i

    Close #f

    LogLine "Wrote CAD dimension CSV: " & csvPath
    Exit Sub

ErrHandler:
    LogLine "WritePartDimensionCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteJobSignatureCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    csvPath = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open csvPath For Output As #f

    Print #f, "JobNumber,CustomerNumber,DateCode,ComponentRole,QuoteName,CadComponent,CleanName,Length,Width,Thickness,Mass,CenterX,CenterY,CenterZ,HasCenter,CenterSource,Status"

    WriteJobSignatureRow f, "TCP", "TCP", "TCP|TOP SMED|TOP CLAMPING|TOP CLAMPING PLATE|ID SMED"
    WriteJobSignatureRow f, "BCP", "BCP", "BCP|BOTTOM SMED|BOT SMED|BOTTOM CLAMPING|BOTTOM CLAMPING PLATE|OD SMED"
    WriteJobSignatureRow f, "ID HOLDER", "ID HOLDER", ID_HOLDER_KEYS
    WriteJobSignatureRow f, "OD HOLDER", "OD HOLDER", OD_HOLDER_KEYS
    WriteJobSignatureRow f, "ID POT", "ID POT BLOCK", "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT"
    WriteJobSignatureRow f, "OD POT", "OD POT BLOCK", "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT"

    Close #f

    LogLine "Wrote job signature CSV: " & csvPath
    Exit Sub

ErrHandler:
    LogLine "WriteJobSignatureCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteJobSignatureRow(ByVal f As Integer, _
                                 ByVal roleName As String, _
                                 ByVal quoteName As String, _
                                 ByVal fallbackKeys As String)
On Error Resume Next

    Dim cadIdx As Long
    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    Dim cadComponent As String
    Dim cleanName As String
    Dim statusText As String
    Dim L As Double
    Dim W As Double
    Dim T As Double
    Dim massVal As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim centerZ As Double
    Dim hasCenter As Boolean
    Dim centerSource As String

    cadComponent = ""
    cleanName = ""
    statusText = "NO CAD MATCH"
    L = 0#: W = 0#: T = 0#
    massVal = 0#
    centerX = 0#: centerY = 0#: centerZ = 0#
    hasCenter = False
    centerSource = ""

    If cadIdx > 0 And cadIdx <= PartCount Then
        cadComponent = parts(cadIdx).componentName
        cleanName = parts(cadIdx).cleanName
        L = parts(cadIdx).Length
        W = parts(cadIdx).Width
        T = parts(cadIdx).Thickness
        massVal = parts(cadIdx).massValue

        If parts(cadIdx).hasMassCenter Then
            hasCenter = True
            centerX = parts(cadIdx).MassCenterX
            centerY = parts(cadIdx).MassCenterY
            centerZ = parts(cadIdx).MassCenterZ
            centerSource = "CENTER_OF_MASS"
        ElseIf parts(cadIdx).hasAsmCenter Then
            hasCenter = True
            centerX = parts(cadIdx).AsmCenterX
            centerY = parts(cadIdx).AsmCenterY
            centerZ = parts(cadIdx).AsmCenterZ
            centerSource = "BBOX_CENTER_FALLBACK"
        Else
            hasCenter = False
            centerSource = ""
        End If

        statusText = "OK"
    End If

    Print #f, CsvText(CurrentJobNumber) & "," & _
              CsvText(CustomerNumber) & "," & _
              CsvText(DateCode) & "," & _
              CsvText(roleName) & "," & _
              CsvText(quoteName) & "," & _
              CsvText(cadComponent) & "," & _
              CsvText(cleanName) & "," & _
              FormatNumberForCsv(L) & "," & _
              FormatNumberForCsv(W) & "," & _
              FormatNumberForCsv(T) & "," & _
              FormatNumberForCsv(massVal) & "," & _
              FormatNumberForCsv(centerX) & "," & _
              FormatNumberForCsv(centerY) & "," & _
              FormatNumberForCsv(centerZ) & "," & _
              CStr(hasCenter) & "," & _
              CsvText(centerSource) & "," & _
              CsvText(statusText)
End Sub

Private Sub WriteExportCheckCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    csvPath = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open csvPath For Output As #f

    Print #f, "Section,QuoteName,Material,Qty,BomL,BomW,BomT,CadL,CadW,CadT,CadComponent,Status"

    Dim i As Long

    For i = 1 To ExportCount

        Dim cadName As String
        Dim cadL As String
        Dim cadW As String
        Dim cadT As String

        If ExportRows(i).HasCad Then
            cadName = parts(ExportRows(i).CadPartIndex).componentName
            cadL = FormatNumberForCsv(ExportRows(i).Length)
            cadW = FormatNumberForCsv(ExportRows(i).Width)
            cadT = FormatNumberForCsv(ExportRows(i).Thickness)
        Else
            cadName = ""
            cadL = ""
            cadW = ""
            cadT = ""
        End If

        Print #f, "EXPORT," & _
                  CsvText(ExportRows(i).quoteName) & "," & _
                  CsvText(ExportRows(i).material) & "," & _
                  CStr(ExportRows(i).Quantity) & "," & _
                  FormatNumberForCsv(ExportRows(i).BomLength) & "," & _
                  FormatNumberForCsv(ExportRows(i).BomWidth) & "," & _
                  FormatNumberForCsv(ExportRows(i).BomThickness) & "," & _
                  cadL & "," & cadW & "," & cadT & "," & _
                  CsvText(cadName) & "," & _
                  CsvText(ExportRows(i).Status)
    Next i

    For i = 1 To PullcoreMatchCount

        Dim pcCadName As String
        Dim pcL As String
        Dim pcW As String
        Dim pcT As String

        If PullcoreMatches(i).CadPartIndex > 0 Then
            pcCadName = parts(PullcoreMatches(i).CadPartIndex).componentName
            pcL = FormatNumberForCsv(PullcoreMatches(i).FittedLength)
            pcW = FormatNumberForCsv(PullcoreMatches(i).FittedWidth)
            pcT = FormatNumberForCsv(PullcoreMatches(i).FittedThickness)
        Else
            pcCadName = ""
            pcL = ""
            pcW = ""
            pcT = ""
        End If

        Print #f, "PULLCORE," & _
                  CsvText(PullcoreMatches(i).quoteName) & "," & _
                  CsvText(PullcoreMatches(i).material) & "," & _
                  CStr(PullcoreMatches(i).Quantity) & "," & _
                  FormatNumberForCsv(PullcoreMatches(i).BomLength) & "," & _
                  FormatNumberForCsv(PullcoreMatches(i).BomWidth) & "," & _
                  FormatNumberForCsv(PullcoreMatches(i).BomThickness) & "," & _
                  pcL & "," & pcW & "," & pcT & "," & _
                  CsvText(pcCadName) & "," & _
                  CsvText(PullcoreMatches(i).Status)
    Next i

    WritePackageReportRow f, "J BLOCK", FindBestJBlockIndex()
    WritePackageReportRow f, "EJECTOR CAM", FindBestEjectorCamIndex()
    WritePackageReportRow f, "PULLCORE STOP", FindBestPartByNameAndDims(PULLCORE_STOP_NAME_KEYS, 0, 0, 0, 0)
    WritePackageReportRow f, "FLIPPER CAM COVER PLATE", _
        FindBestPartByNameAndDims(FLIPPER_CAM_COVER_NAME_KEYS, _
                                  FLIPPER_CAM_COVER_TARGET_THICKNESS, _
                                  FLIPPER_CAM_COVER_TARGET_WIDTH, _
                                  FLIPPER_CAM_COVER_TARGET_LENGTH, _
                                  FLIPPER_CAM_COVER_DIM_TOL)

    Print #f, "RAWBOM,---,---,---,---,---,---,---,---,---,---,parsed " & CStr(BomCount) & " rows"

    For i = 1 To BomCount
        Print #f, "RAWBOM," & _
                  CsvText(BomRows(i).Description) & "," & _
                  CsvText(BomRows(i).material) & "," & _
                  CStr(BomRows(i).Quantity) & "," & _
                  FormatNumberForCsv(BomRows(i).BomLength) & "," & _
                  FormatNumberForCsv(BomRows(i).BomWidth) & "," & _
                  FormatNumberForCsv(BomRows(i).BomThickness) & ",,,," & _
                  CsvText(BomRows(i).quoteName) & "," & _
                  CsvText("isPullcore=" & CStr(IsPullcoreBomRow(BomRows(i))) & " hasDims=" & CStr(BomRows(i).hasDims))
    Next i

    Close #f

    LogLine "Wrote BOM match report CSV: " & csvPath
    Exit Sub

ErrHandler:
    LogLine "WriteExportCheckCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WritePackageReportRow(ByVal f As Integer, ByVal label As String, ByVal cadIdx As Long)
On Error Resume Next

    Dim nm As String
    Dim L As String
    Dim W As String
    Dim T As String
    Dim st As String

    If cadIdx > 0 And cadIdx <= PartCount Then
        nm = parts(cadIdx).componentName
        L = FormatNumberForCsv(parts(cadIdx).Length)
        W = FormatNumberForCsv(parts(cadIdx).Width)
        T = FormatNumberForCsv(parts(cadIdx).Thickness)
        st = "CAD MATCH FOUND"
    Else
        nm = ""
        L = ""
        W = ""
        T = ""
        st = "NO CAD MATCH"
    End If

    Print #f, "PACKAGE," & CsvText(label) & ",,,,,," & _
              L & "," & W & "," & T & "," & CsvText(nm) & "," & CsvText(st)
End Sub

' ============================================================
' PULL CORE DIMENSIONS EXCEL REPORT
' ============================================================

Private Sub WritePullCoreDimensionsExcel(ByVal reportPath As String)
On Error GoTo ErrHandler

    If CREATE_PULLCORE_DIMENSIONS_EXCEL = False Then Exit Sub

    If reportPath = "" Then
        reportPath = CurrentJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE
    End If

    LogLine "Creating Pull Core Dimensions Excel report: " & reportPath

    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False

    Set wb = xlApp.Workbooks.Add
    Set ws = wb.Worksheets(1)

    On Error Resume Next
    ws.name = "Pull Core Dimensions"
    On Error GoTo ErrHandler

    WritePullCoreDimensionsExcelHeader ws

    Dim r As Long
    r = 2

    Dim i As Long

    For i = 1 To PullcoreMatchCount
        WritePullCoreDimensionsExcelRow ws, r, i
        r = r + 1
    Next i

    With ws
        .Columns("A:AH").AutoFit
        .rows(1).Font.Bold = True
        .rows(1).Interior.Color = RGB(220, 230, 241)
        .rows(1).AutoFilter
    End With

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    EnsureFolderDeep fso.GetParentFolderName(reportPath)

    If fso.FileExists(reportPath) Then
        On Error Resume Next
        fso.DeleteFile reportPath, True
        On Error GoTo ErrHandler
    End If

    wb.SaveAs reportPath, xlOpenXMLWorkbook
    wb.Close False

    xlApp.EnableEvents = True
    xlApp.ScreenUpdating = True
    xlApp.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing

    LogLine "Pull Core Dimensions Excel report saved: " & reportPath

    Exit Sub

ErrHandler:
    LogLine "WritePullCoreDimensionsExcel error: " & Err.Description

    On Error Resume Next

    If Not wb Is Nothing Then wb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
End Sub

Private Sub WritePullCoreDimensionsExcelHeader(ByVal ws As Object)
On Error Resume Next

    If ws Is Nothing Then Exit Sub

    ws.Cells(1, 1).value = "Job"
    ws.Cells(1, 2).value = "Quote Name"
    ws.Cells(1, 3).value = "Description"
    ws.Cells(1, 4).value = "Material"
    ws.Cells(1, 5).value = "Qty"
    ws.Cells(1, 6).value = "CAD Component"
    ws.Cells(1, 7).value = "Clean Name"

    ws.Cells(1, 8).value = "BOM Length"
    ws.Cells(1, 9).value = "BOM Width"
    ws.Cells(1, 10).value = "BOM Thickness"

    ws.Cells(1, 11).value = "Original Bounding Box Dimensions"
    ws.Cells(1, 12).value = "Original Length"
    ws.Cells(1, 13).value = "Original Width"
    ws.Cells(1, 14).value = "Original Thickness"

    ws.Cells(1, 15).value = "Fitted Bounding Box Dimensions"
    ws.Cells(1, 16).value = "Fitted Length"
    ws.Cells(1, 17).value = "Fitted Width"
    ws.Cells(1, 18).value = "Fitted Thickness"

    ws.Cells(1, 19).value = "Delta Length Original-Fitted"
    ws.Cells(1, 20).value = "Delta Width Original-Fitted"
    ws.Cells(1, 21).value = "Delta Thickness Original-Fitted"

    ws.Cells(1, 22).value = "Detected Pullcore Angle Deg"
    ws.Cells(1, 23).value = "DXF Rotation Deg"

    ws.Cells(1, 24).value = "Status"
    ws.Cells(1, 25).value = "Is Cam"
    ws.Cells(1, 26).value = "CAD Index"
    ws.Cells(1, 27).value = "Mass"
    ws.Cells(1, 28).value = "BBox Volume"
    ws.Cells(1, 29).value = "Notes"

    ws.Cells(1, 30).value = "Assembly Center X"
    ws.Cells(1, 31).value = "Assembly Center Y"
    ws.Cells(1, 32).value = "Assembly Center Z"
    ws.Cells(1, 33).value = "ID/OD Height Axis Used"
    ws.Cells(1, 34).value = "ID/OD Height Value"
End Sub

Private Sub WritePullCoreDimensionsExcelRow(ByVal ws As Object, ByVal r As Long, ByVal matchIdx As Long)
On Error Resume Next

    If ws Is Nothing Then Exit Sub
    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    Dim compName As String
    Dim cleanName As String
    Dim massVal As Double
    Dim volVal As Double

    compName = ""
    cleanName = ""
    massVal = 0#
    volVal = 0#

    If cadIdx > 0 And cadIdx <= PartCount Then
        compName = parts(cadIdx).componentName
        cleanName = parts(cadIdx).cleanName
        massVal = parts(cadIdx).massValue
        volVal = parts(cadIdx).BBoxVolume
    End If

    ws.Cells(r, 1).value = CurrentJobNumber
    ws.Cells(r, 2).value = PullcoreMatches(matchIdx).quoteName
    ws.Cells(r, 3).value = PullcoreMatches(matchIdx).Description
    ws.Cells(r, 4).value = PullcoreMatches(matchIdx).material
    ws.Cells(r, 5).value = PullcoreMatches(matchIdx).Quantity

    ws.Cells(r, 6).NumberFormat = "@"
    ws.Cells(r, 6).value = compName

    ws.Cells(r, 7).NumberFormat = "@"
    ws.Cells(r, 7).value = cleanName

    ws.Cells(r, 8).value = PullcoreMatches(matchIdx).BomLength
    ws.Cells(r, 9).value = PullcoreMatches(matchIdx).BomWidth
    ws.Cells(r, 10).value = PullcoreMatches(matchIdx).BomThickness

    ws.Cells(r, 11).value = "Original Bounding Box Dimensions"
    ws.Cells(r, 12).value = PullcoreMatches(matchIdx).OriginalLength
    ws.Cells(r, 13).value = PullcoreMatches(matchIdx).OriginalWidth
    ws.Cells(r, 14).value = PullcoreMatches(matchIdx).OriginalThickness

    ws.Cells(r, 15).value = "Fitted Bounding Box Dimensions"
    ws.Cells(r, 16).value = PullcoreMatches(matchIdx).FittedLength
    ws.Cells(r, 17).value = PullcoreMatches(matchIdx).FittedWidth
    ws.Cells(r, 18).value = PullcoreMatches(matchIdx).FittedThickness

    ws.Cells(r, 19).value = PullcoreMatches(matchIdx).OriginalLength - PullcoreMatches(matchIdx).FittedLength
    ws.Cells(r, 20).value = PullcoreMatches(matchIdx).OriginalWidth - PullcoreMatches(matchIdx).FittedWidth
    ws.Cells(r, 21).value = PullcoreMatches(matchIdx).OriginalThickness - PullcoreMatches(matchIdx).FittedThickness

    ws.Cells(r, 22).value = PullcoreMatches(matchIdx).DetectedAngleDeg
    ws.Cells(r, 23).value = PullcoreMatches(matchIdx).DxfRotationDeg

    ws.Cells(r, 24).value = PullcoreMatches(matchIdx).Status
    ws.Cells(r, 25).value = CStr(PullcoreMatches(matchIdx).isCam)
    ws.Cells(r, 26).value = cadIdx
    ws.Cells(r, 27).value = massVal
    ws.Cells(r, 28).value = volVal

    If cadIdx <= 0 Then
        ws.Cells(r, 29).value = "No CAD match"
    ElseIf Abs(PullcoreMatches(matchIdx).DetectedAngleDeg) < PULLCORE_STRAIGHTEN_MIN_DEG Then
        ws.Cells(r, 29).value = "Angle below threshold or part is square to the assembly"
    Else
        ws.Cells(r, 29).value = "Angle from longest edge in front-view plane (downward tilt ignored)"
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then
        If parts(cadIdx).hasAsmCenter Then
            ws.Cells(r, 30).value = parts(cadIdx).AsmCenterX
            ws.Cells(r, 31).value = parts(cadIdx).AsmCenterY
            ws.Cells(r, 32).value = parts(cadIdx).AsmCenterZ
            ws.Cells(r, 33).value = PullcoreIdOdHeightAxisUsed

            If PullcoreIdOdHeightAxisUsed <> "" Then
                ws.Cells(r, 34).value = GetPullcoreMatchHeightValue(matchIdx, PullcoreIdOdHeightAxisUsed)
            End If
        End If
    End If
End Sub

' ============================================================
' END OF PART 7 OF 8
' Paste Part 8 next.
' ============================================================
' ============================================================
' FINDERS: JOB FOLDER + CAD
' ============================================================

Private Function TryParseJobFolderDate(ByVal folderName As String, ByRef outDate As Date) As Boolean
On Error GoTo ErrHandler

    TryParseJobFolderDate = False

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})"

    Dim matches As Object
    Set matches = re.Execute(folderName)

    If matches.count = 0 Then Exit Function

    Dim m As Object
    Set m = matches(matches.count - 1)

    Dim mo As Long
    Dim dY As Long
    Dim yr As Long

    mo = CLng(val(m.SubMatches(0)))
    dY = CLng(val(m.SubMatches(1)))
    yr = CLng(val(m.SubMatches(2)))

    If yr > 0 And yr < 100 Then yr = 2000 + yr

    If mo >= 1 And mo <= 12 And dY >= 1 And dY <= 31 And yr >= 1990 And yr <= 2100 Then
        outDate = DateSerial(yr, mo, dY)
        TryParseJobFolderDate = True
    End If

    Exit Function

ErrHandler:
    TryParseJobFolderDate = False
End Function

Private Function GetJobFolderMatchScore(ByVal folderName As String, _
                                        ByVal wantUpper As String, _
                                        ByVal depth As Long) As Long
On Error GoTo ErrHandler

    GetJobFolderMatchScore = -1

    Dim nameUpper As String
    nameUpper = UCase(Trim(folderName))

    Dim score As Long
    score = -1

    If nameUpper = wantUpper Then
        score = 1000 - depth
    ElseIf InStr(nameUpper, wantUpper) > 0 Then
        score = 500 - (depth * 20) - Abs(Len(nameUpper) - Len(wantUpper))
    End If

    If score < 0 Then Exit Function

    If InStr(nameUpper, "CHANGE") > 0 Or _
       InStr(nameUpper, "CHANGES") > 0 Or _
       InStr(nameUpper, "UPDATED") > 0 Or _
       InStr(nameUpper, "UPDATE") > 0 Then

        score = score + 5000

    End If

    Dim folderDate As Date

    If TryParseJobFolderDate(folderName, folderDate) Then
        score = score + CLng(CDbl(folderDate))
    End If

    GetJobFolderMatchScore = score
    Exit Function

ErrHandler:
    GetJobFolderMatchScore = -1
End Function

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

        nameUpper = UCase(subFolder.name)

        If ShouldSkipJobSearchTopFolder(subFolder.name, wantUpper) Then GoTo NextTop
        If nameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextTop
        If InStr(nameUpper, " PRINTS") > 0 Then GoTo NextTop
        If InStr(nameUpper, "PULLCORE") > 0 Then GoTo NextTop
        If InStr(nameUpper, "PYROPEL") > 0 Then GoTo NextTop
        If InStr(nameUpper, "J BLOCK") > 0 Then GoTo NextTop

        score = GetJobFolderMatchScore(subFolder.name, wantUpper, 0)

        If score > bestScore Then
            bestScore = score
            bestPath = subFolder.path
        End If

NextTop:
    Next subFolder

    If bestPath <> "" Then
        LogLine "Selected job folder (score=" & CStr(bestScore) & "): " & bestPath
        FindJobFolderByText = bestPath
        Exit Function
    End If

    For Each subFolder In root.SubFolders
        nameUpper = UCase(subFolder.name)
        If ShouldSkipJobSearchTopFolder(subFolder.name, wantUpper) = False Then
            If nameUpper <> UCase(EXTRACT_FOLDER_NAME) Then
                SearchJobFolderRecursive subFolder, wantUpper, 1, bestPath, bestScore
            End If
        End If
    Next subFolder

    If bestPath = "" Then
        If InStr(UCase(root.name), wantUpper) > 0 Then bestPath = rootPath
    End If

    If bestPath <> "" Then
        LogLine "Selected job folder (score=" & CStr(bestScore) & "): " & bestPath
    End If

    FindJobFolderByText = bestPath
    Exit Function

ErrHandler:
    LogLine "FindJobFolderByText error: " & Err.Description
    FindJobFolderByText = ""
End Function

Private Function ShouldSkipJobSearchTopFolder(ByVal folderName As String, ByVal wantUpper As String) As Boolean
On Error GoTo ErrHandler

    ShouldSkipJobSearchTopFolder = False

    If LIMIT_JOB_SEARCH_TO_CURRENT_AND_PREVIOUS_MONTH = False Then Exit Function

    Dim n As String
    n = UCase(folderName)

    ' Always allow a folder that directly contains the job/search text.
    If InStr(n, wantUpper) > 0 Then Exit Function

    Dim monthOffset As Long
    Dim monthText As String

    ' JOB_SEARCH_MONTH_WINDOW = 4 means:
    '   current month + previous 3 months
    For monthOffset = 0 To JOB_SEARCH_MONTH_WINDOW - 1

        monthText = UCase(Format(DateAdd("m", -monthOffset, Date), "mmmm yyyy"))

        If InStr(n, monthText) > 0 Then
            ShouldSkipJobSearchTopFolder = False
            Exit Function
        End If

    Next monthOffset

    ShouldSkipJobSearchTopFolder = True
    Exit Function

ErrHandler:
    ShouldSkipJobSearchTopFolder = False
End Function

' ============================================================
' NETWORK-AWARE PATHS + ELGIN PUBLISH
' ============================================================

Private Function ResolveRootJobPath() As String
    On Error Resume Next
    Dim root As String
    If IsOnCompanyWifi() Then
        root = PUBLIC_DOWNLOADS_PATH
    Else
        root = LOCAL_DOWNLOADS_FALLBACK
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(PUBLIC_DOWNLOADS_PATH) Then root = PUBLIC_DOWNLOADS_PATH
    End If
    ResolveRootJobPath = root
End Function

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
        If InStr(1, s, "SSID", vbTextCompare) = 1 Then
            p = InStr(s, ":")
            If p > 0 Then
                GetWifiSsidVba = Trim(mid(s, p + 1))
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
        Exit Function
    End If
    If PUBLIC_DATA_ROOT <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        If fso.FolderExists(PUBLIC_DATA_ROOT) Then IsOnCompanyWifi = True
    End If
End Function

Private Function ResolveMatchingRoot() As String
    Dim root As String

    If ALWAYS_PUBLISH_TO_PUBLIC_MATCHING_SHARE Then
        root = PUBLIC_DATA_ROOT
    Else
        If IsOnCompanyWifi() Then
            root = PUBLIC_DATA_ROOT
        Else
            root = PRIVATE_DATA_ROOT
        End If
    End If

    If root = "" Then root = PUBLIC_DATA_ROOT

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

    Dim sigPath As String
    sigPath = root & "\XT_Export_Job_Signature_" & CleanFileName(CurrentJobNumber) & ".csv"
    WriteJobSignatureCsv sigPath

    Dim jobOut As String
    jobOut = root & "\" & CleanFileName(CurrentJobNumber)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(jobOut) Then
        LogLine "Publish: replacing existing job output folder: " & jobOut
        fso.DeleteFolder jobOut, True
    End If

    EnsureFolderDeep jobOut
    CopyMatchingArtifacts CurrentJobFolder, jobOut

    ' Preserve the dedicated Match Studio STL folder too.
    Dim stlFolder As String
    Dim stlPublishFolder As String

    stlFolder = CurrentJobFolder & "\" & CurrentJobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX
    stlPublishFolder = jobOut & "\" & CurrentJobNumber & MATCH_STUDIO_STL_FOLDER_SUFFIX

    If fso.FolderExists(stlFolder) Then
        If fso.FolderExists(stlPublishFolder) Then
            fso.DeleteFolder stlPublishFolder, True
        End If

        EnsureFolderDeep stlPublishFolder
        CopyFolderContents stlFolder, stlPublishFolder

        LogLine "Published Match Studio STL folder:"
        LogLine "  " & stlPublishFolder
    End If

    LogLine "Published job outputs to: " & jobOut
    Exit Sub
eh:
    LogLine "PublishJobOutputs error: " & Err.Description
End Sub

Private Sub CopyMatchingArtifacts(ByVal srcFolder As String, ByVal dstFolder As String)
    On Error Resume Next
    CopyMatchingArtifactsRecursive srcFolder, dstFolder, 0
End Sub

Private Sub CopyMatchingArtifactsRecursive(ByVal srcFolder As String, ByVal dstFolder As String, ByVal depth As Long)
    On Error Resume Next
    If depth > 3 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcFolder) Then Exit Sub
    Dim f As Object, sub_ As Object, nm As String, ext As String, take As Boolean
    For Each f In fso.GetFolder(srcFolder).Files
        nm = f.name
        ext = UCase$(GetFileExtension(nm))
        take = False
        ' Do not publish JPEG/PNG preview images.
        If ext = "CSV" Or ext = "TXT" Or ext = "PDF" Then take = True
        If ext = "XLS" Or ext = "XLSX" Or ext = "XLSM" Then take = True
        If ext = "IGS" Or ext = "IGES" Or ext = "STL" Or ext = "EASM" Or ext = "X_T" Or ext = "XT" Or ext = "DXF" Then take = True
        If take Then fso.CopyFile f.path, dstFolder & "\" & nm, True
    Next f
    For Each sub_ In fso.GetFolder(srcFolder).SubFolders
        CopyMatchingArtifactsRecursive sub_.path, dstFolder, depth + 1
    Next sub_
End Sub

Private Sub UploadJobSignatureToElgin(ByVal csvPath As String)
On Error GoTo ErrHandler

    If AUTO_UPLOAD_JOB_SIGNATURE_TO_ELGIN = False Then Exit Sub
    If csvPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(csvPath) = False Then Exit Sub

    Dim boundary As String
    boundary = "----CMSXT" & Format(Now, "yyyymmddhhnnss")

    Dim CsvText As String
    CsvText = ReadAllTextFile(csvPath)

    If Trim(CsvText) = "" Then Exit Sub

    Dim body As String
    body = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=" & Chr(34) & "file" & Chr(34) & "; filename=" & Chr(34) & JOB_SIGNATURE_REPORT_FILE & Chr(34) & vbCrLf & _
           "Content-Type: text/csv" & vbCrLf & vbCrLf & _
           CsvText & vbCrLf & _
           "--" & boundary & "--" & vbCrLf

    Dim url As String
    url = Trim(ELGIN_API_BASE_URL)

    Do While Right(url, 1) = "/"
        url = Left(url, Len(url) - 1)
    Loop

    url = url & "/api/job-signatures/" & CurrentJobNumber & "/upload"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.SetRequestHeader "X-Elgin-Api-Key", ELGIN_SIGNATURE_API_KEY
    http.Send body

    If CLng(http.Status) >= 200 And CLng(http.Status) < 300 Then
        LogLine "Uploaded job signature to ELGIN: " & url
    Else
        LogLine "WARNING: ELGIN signature upload failed. HTTP " & CStr(http.Status) & " " & CStr(http.ResponseText)
    End If

CleanExit:
    On Error Resume Next
    Set http = Nothing
    Set fso = Nothing
    Exit Sub

ErrHandler:
    LogLine "WARNING: UploadJobSignatureToElgin error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub OrganizeJobFiles()
On Error GoTo eh

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(CurrentJobFolder) Then Exit Sub

    Dim baseDir As String
    Dim pdfDir As String
    Dim legacyBase As String

    baseDir = CurrentJobFolder & "\" & CurrentJobNumber & " Base"
    legacyBase = CurrentJobFolder & "\base"
    pdfDir = CurrentJobFolder & "\" & CurrentJobNumber & " pdfs"

    EnsureFolderDeep baseDir
    EnsureFolderDeep pdfDir

    Dim movedBase As Long
    Dim movedPdf As Long

    movedBase = 0
    movedPdf = 0

    ' Move anything from old lowercase "base" folder into the proper Base folder.
    If fso.FolderExists(legacyBase) Then
        movedBase = movedBase + MoveSolidWorksNativeFilesFromFolderToBase(legacyBase, baseDir)

        On Error Resume Next
        If fso.GetFolder(legacyBase).Files.count = 0 And fso.GetFolder(legacyBase).SubFolders.count = 0 Then
            fso.DeleteFolder legacyBase, True
        End If
        On Error GoTo eh
    End If

    ' Move all root-level SolidWorks native files into Base.
    movedBase = movedBase + MoveSolidWorksNativeFilesFromFolderToBase(CurrentJobFolder, baseDir)

    ' Move reports / PDFs / BOM files into pdf folder.
    movedPdf = movedPdf + MoveReportFilesFromFolderToPdf(CurrentJobFolder, pdfDir)

    LogLine "OrganizeJobFiles: moved " & movedBase & _
            " SolidWorks native file(s) -> " & CurrentJobNumber & _
            " Base, " & movedPdf & " report/BOM file(s) -> " & CurrentJobNumber & " pdfs"

    Exit Sub

eh:
    LogLine "OrganizeJobFiles error: " & Err.Description
End Sub

Private Function MoveSolidWorksNativeFilesFromFolderToBase(ByVal sourceFolder As String, _
                                                           ByVal baseDir As String) As Long
On Error GoTo ErrHandler

    MoveSolidWorksNativeFilesFromFolderToBase = 0

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If sourceFolder = "" Then Exit Function
    If baseDir = "" Then Exit Function
    If fso.FolderExists(sourceFolder) = False Then Exit Function

    EnsureFolderDeep baseDir

    Dim fileNames As Collection
    Set fileNames = New Collection

    Dim f As Object

    For Each f In fso.GetFolder(sourceFolder).Files
        If ShouldMoveFileToBaseFolder(f.name) Then
            fileNames.Add f.path
        End If
    Next f

    Dim i As Long
    Dim srcPath As String
    Dim targetPath As String
    Dim nm As String

    For i = 1 To fileNames.count

        srcPath = CStr(fileNames(i))
        nm = fso.GetFileName(srcPath)

        If LCase(fso.GetParentFolderName(srcPath)) <> LCase(baseDir) Then

            targetPath = GetUniqueFilePath(baseDir & "\" & nm)

            If MoveFileRobust(srcPath, targetPath) Then
                MoveSolidWorksNativeFilesFromFolderToBase = MoveSolidWorksNativeFilesFromFolderToBase + 1
                LogLine "Moved SolidWorks native file to Base: " & nm
            Else
                LogLine "WARNING: Could not move SolidWorks native file to Base: " & srcPath
            End If

        End If

    Next i

    Exit Function

ErrHandler:
    LogLine "MoveSolidWorksNativeFilesFromFolderToBase error: " & Err.Description
    MoveSolidWorksNativeFilesFromFolderToBase = 0
End Function

Private Function MoveReportFilesFromFolderToPdf(ByVal sourceFolder As String, _
                                                ByVal pdfDir As String) As Long
On Error GoTo ErrHandler

    MoveReportFilesFromFolderToPdf = 0

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If sourceFolder = "" Then Exit Function
    If pdfDir = "" Then Exit Function
    If fso.FolderExists(sourceFolder) = False Then Exit Function

    EnsureFolderDeep pdfDir

    Dim fileNames As Collection
    Set fileNames = New Collection

    Dim f As Object

    For Each f In fso.GetFolder(sourceFolder).Files
        If ShouldMoveFileToPdfFolder(f.name) Then
            fileNames.Add f.path
        End If
    Next f

    Dim i As Long
    Dim srcPath As String
    Dim targetPath As String
    Dim nm As String

    For i = 1 To fileNames.count

        srcPath = CStr(fileNames(i))
        nm = fso.GetFileName(srcPath)

        If LCase(fso.GetParentFolderName(srcPath)) <> LCase(pdfDir) Then

            targetPath = GetUniqueFilePath(pdfDir & "\" & nm)

            If MoveFileRobust(srcPath, targetPath) Then
                MoveReportFilesFromFolderToPdf = MoveReportFilesFromFolderToPdf + 1
                LogLine "Moved report/BOM file to pdf folder: " & nm
            Else
                LogLine "WARNING: Could not move report/BOM file to pdf folder: " & srcPath
            End If

        End If

    Next i

    Exit Function

ErrHandler:
    LogLine "MoveReportFilesFromFolderToPdf error: " & Err.Description
    MoveReportFilesFromFolderToPdf = 0
End Function

Private Function ShouldMoveFileToBaseFolder(ByVal fileName As String) As Boolean
On Error GoTo ErrHandler

    ShouldMoveFileToBaseFolder = False

    Dim ext As String
    Dim n As String

    ext = UCase(GetFileExtension(fileName))
    n = UCase(fileName)

    If ext = "SLDPRT" Then
        ShouldMoveFileToBaseFolder = True
        Exit Function
    End If

    If ext = "SLDASM" Then
        ShouldMoveFileToBaseFolder = True
        Exit Function
    End If

    ' Imported assembly generated part names like 861000196_MB-ASM_05-20-2026.sldasm-Part-4
    If InStr(n, ".SLDASM-PART-") > 0 Then
        ShouldMoveFileToBaseFolder = True
        Exit Function
    End If

    If InStr(n, ".SLDASM-") > 0 And InStr(n, "PART") > 0 Then
        ShouldMoveFileToBaseFolder = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    ShouldMoveFileToBaseFolder = False
End Function

Private Function ShouldMoveFileToPdfFolder(ByVal fileName As String) As Boolean
On Error GoTo ErrHandler

    ShouldMoveFileToPdfFolder = False

    Dim ext As String
    Dim n As String

    ext = UCase(GetFileExtension(fileName))
    n = UCase(fileName)

    Select Case ext

        Case "PDF"
            ShouldMoveFileToPdfFolder = True
            Exit Function

        Case "XLS", "XLSX", "XLSM", "XLSB", "XLTX", "XLTM"
            ShouldMoveFileToPdfFolder = True
            Exit Function

        Case "CSV"
            If InStr(n, "EXPORT") > 0 Or _
               InStr(n, "SIGNATURE") > 0 Or _
               InStr(n, "PULL") > 0 Or _
               InStr(n, "BOM") > 0 Or _
               InStr(n, "DIMENSION") > 0 Then

                ShouldMoveFileToPdfFolder = True
                Exit Function

            End If

        Case "TXT"
            If InStr(n, "BOM") > 0 Or _
               InStr(n, "EXPORT") > 0 Or _
               InStr(n, "LOG") > 0 Then

                ShouldMoveFileToPdfFolder = True
                Exit Function

            End If

    End Select

    Exit Function

ErrHandler:
    ShouldMoveFileToPdfFolder = False
End Function

Private Function MoveFileRobust(ByVal srcPath As String, ByVal targetPath As String) As Boolean
On Error GoTo ErrHandler

    MoveFileRobust = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If srcPath = "" Or targetPath = "" Then Exit Function
    If fso.FileExists(srcPath) = False Then Exit Function

    EnsureFolderDeep fso.GetParentFolderName(targetPath)

    On Error Resume Next
    Err.Clear
    fso.MoveFile srcPath, targetPath

    If Err.Number = 0 Then
        MoveFileRobust = fso.FileExists(targetPath)
        Exit Function
    End If

    Err.Clear
    fso.CopyFile srcPath, targetPath, True

    If Err.Number = 0 And fso.FileExists(targetPath) Then

        Err.Clear
        fso.DeleteFile srcPath, True

        MoveFileRobust = True
        Exit Function

    End If

    Exit Function

ErrHandler:
    MoveFileRobust = False
End Function

Private Sub KeepBaseAndHoldersPackageFilesInMainFolder()
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If CurrentJobFolder = "" Then Exit Sub
    If fso.FolderExists(CurrentJobFolder) = False Then Exit Sub

    ' If an old HOLDERS folder exists, move its IGS/DXF files back to main.
    MoveBaseHolderPackageFilesFromFolderToMain CurrentJobFolder & "\" & CurrentJobNumber & " HOLDERS"

    ' If an old Base folder exists, move only BASE/HOLDERS package files back to main.
    MoveBaseHolderPackageFilesFromFolderToMain CurrentJobFolder & "\" & CurrentJobNumber & " Base"

    ' If old lowercase base folder exists, move only BASE/HOLDERS package files back to main.
    MoveBaseHolderPackageFilesFromFolderToMain CurrentJobFolder & "\base"

    ' Remove old wrong lowercase pdf folder if empty.
    DeleteFolderIfEmptyOnly CurrentJobFolder & "\pdf"

    ' Delete empty old holders/base folders only.
    DeleteFolderIfEmptyOnly CurrentJobFolder & "\" & CurrentJobNumber & " HOLDERS"
    DeleteFolderIfEmptyOnly CurrentJobFolder & "\" & CurrentJobNumber & " Base"
    DeleteFolderIfEmptyOnly CurrentJobFolder & "\base"

    LogLine "Verified BASE/HOLDERS package files are in main job folder."

    Exit Sub

ErrHandler:
    LogLine "KeepBaseAndHoldersPackageFilesInMainFolder error: " & Err.Description
End Sub

Private Sub MoveBaseHolderPackageFilesFromFolderToMain(ByVal sourceFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If sourceFolder = "" Then Exit Sub
    If fso.FolderExists(sourceFolder) = False Then Exit Sub
    If LCase(sourceFolder) = LCase(CurrentJobFolder) Then Exit Sub

    Dim fileList As Collection
    Set fileList = New Collection

    Dim f As Object

    For Each f In fso.GetFolder(sourceFolder).Files
        If IsBaseOrHoldersPackageFile(f.name) Then
            fileList.Add f.path
        End If
    Next f

    Dim i As Long
    Dim srcPath As String
    Dim destPath As String

    For i = 1 To fileList.count

        srcPath = CStr(fileList(i))

        If fso.FileExists(srcPath) Then

            destPath = GetUniqueFilePath(CurrentJobFolder & "\" & fso.GetFileName(srcPath))

            If MoveFileAndDeleteOriginal(srcPath, destPath) Then
                LogLine "Moved BASE/HOLDERS package file to main folder: " & fso.GetFileName(destPath)
            Else
                LogLine "WARNING: Could not move BASE/HOLDERS package file to main folder: " & srcPath
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "MoveBaseHolderPackageFilesFromFolderToMain error: " & Err.Description
End Sub

Private Function IsBaseOrHoldersPackageFile(ByVal fileName As String) As Boolean
On Error GoTo ErrHandler

    IsBaseOrHoldersPackageFile = False

    Dim n As String
    Dim ext As String

    n = UCase(fileName)
    ext = LCase(GetFileExtension(fileName))

    If InStr(n, UCase(CurrentJobNumber & "_BASE_")) > 0 Then
        Select Case ext
            Case "dxf", "easm", "x_t", "igs", "iges", "stl"
                IsBaseOrHoldersPackageFile = True
                Exit Function
        End Select
    End If

    If InStr(n, UCase(CurrentJobNumber & "_HOLDERS_")) > 0 Then
        Select Case ext
            Case "dxf", "igs", "iges", "stl"
                IsBaseOrHoldersPackageFile = True
                Exit Function
        End Select
    End If

    Exit Function

ErrHandler:
    IsBaseOrHoldersPackageFile = False
End Function

Private Sub MoveLooseSolidWorksPartsToBaseFolder()
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If CurrentJobFolder = "" Then Exit Sub
    If fso.FolderExists(CurrentJobFolder) = False Then Exit Sub

    Dim baseDir As String
    baseDir = CurrentJobFolder & "\" & CurrentJobNumber & " Base"

    EnsureFolderDeep baseDir

    Dim fileList As Collection
    Set fileList = New Collection

    Dim f As Object

    For Each f In fso.GetFolder(CurrentJobFolder).Files
        If IsLooseGeneratedSolidWorksPartFile(f.name) Then
            fileList.Add f.path
        End If
    Next f

    Dim i As Long
    Dim srcPath As String
    Dim destPath As String

    For i = 1 To fileList.count

        srcPath = CStr(fileList(i))

        If fso.FileExists(srcPath) Then

            destPath = GetUniqueFilePath(baseDir & "\" & fso.GetFileName(srcPath))

            If MoveFileAndDeleteOriginal(srcPath, destPath) Then
                LogLine "Moved loose generated SolidWorks part to Base folder: " & fso.GetFileName(destPath)
            Else
                LogLine "WARNING: Could not move loose generated SolidWorks part: " & srcPath
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "MoveLooseSolidWorksPartsToBaseFolder error: " & Err.Description
End Sub

Private Function IsLooseGeneratedSolidWorksPartFile(ByVal fileName As String) As Boolean
On Error GoTo ErrHandler

    IsLooseGeneratedSolidWorksPartFile = False

    Dim n As String
    Dim ext As String

    n = UCase(fileName)
    ext = UCase(GetFileExtension(fileName))

    If ext = "SLDPRT" Then
        IsLooseGeneratedSolidWorksPartFile = True
        Exit Function
    End If

    If InStr(n, ".SLDASM-PART-") > 0 Then
        IsLooseGeneratedSolidWorksPartFile = True
        Exit Function
    End If

    If InStr(n, ".SLDASM-") > 0 And InStr(n, "PART") > 0 Then
        IsLooseGeneratedSolidWorksPartFile = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    IsLooseGeneratedSolidWorksPartFile = False
End Function

Private Function MoveFileAndDeleteOriginal(ByVal srcPath As String, _
                                           ByVal destPath As String) As Boolean
On Error GoTo ErrHandler

    MoveFileAndDeleteOriginal = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If srcPath = "" Then Exit Function
    If destPath = "" Then Exit Function
    If fso.FileExists(srcPath) = False Then Exit Function

    EnsureFolderDeep fso.GetParentFolderName(destPath)

    On Error Resume Next
    Err.Clear
    fso.MoveFile srcPath, destPath

    If Err.Number = 0 Then
        If fso.FileExists(destPath) And fso.FileExists(srcPath) = False Then
            MoveFileAndDeleteOriginal = True
            Exit Function
        End If
    End If

    Err.Clear
    fso.CopyFile srcPath, destPath, True

    If Err.Number <> 0 Or fso.FileExists(destPath) = False Then
        LogLine "WARNING: Copy failed while moving file: " & srcPath
        Exit Function
    End If

    Err.Clear
    fso.DeleteFile srcPath, True

    If Err.Number = 0 And fso.FileExists(srcPath) = False Then
        MoveFileAndDeleteOriginal = True
    Else
        LogLine "WARNING: File copied but original could not be deleted: " & srcPath
        MoveFileAndDeleteOriginal = False
    End If

    Exit Function

ErrHandler:
    MoveFileAndDeleteOriginal = False
End Function

Private Sub DeleteFolderIfEmptyOnly(ByVal folderPath As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If folderPath = "" Then Exit Sub
    If fso.FolderExists(folderPath) = False Then Exit Sub

    If fso.GetFolder(folderPath).Files.count = 0 And _
       fso.GetFolder(folderPath).SubFolders.count = 0 Then

        fso.DeleteFolder folderPath, True
        LogLine "Deleted empty folder: " & folderPath

    End If

    Exit Sub

ErrHandler:
    LogLine "DeleteFolderIfEmptyOnly error: " & Err.Description
End Sub

Private Sub CollectPdfExcelFilesToJobPdfsFolder()
On Error GoTo ErrHandler

    If CurrentJobFolder = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(CurrentJobFolder) = False Then Exit Sub

    Dim destFolder As String
    destFolder = CurrentJobFolder & "\" & CurrentJobNumber & " pdfs"

    EnsureFolderDeep destFolder

    Dim movedCount As Long
    movedCount = 0

    MovePdfExcelFilesRecursive CurrentJobFolder, destFolder, movedCount

    DeleteFolderIfEmptyOnly CurrentJobFolder & "\pdf"

    LogLine "Collected " & CStr(movedCount) & _
            " PDF/Excel/report file(s) into: " & destFolder

    Exit Sub

ErrHandler:
    LogLine "CollectPdfExcelFilesToJobPdfsFolder error: " & Err.Description
End Sub

Private Sub MovePdfExcelFilesRecursive(ByVal sourceFolder As String, _
                                       ByVal destFolder As String, _
                                       ByRef movedCount As Long)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If sourceFolder = "" Then Exit Sub
    If destFolder = "" Then Exit Sub
    If fso.FolderExists(sourceFolder) = False Then Exit Sub

    If LCase(sourceFolder) = LCase(destFolder) Then Exit Sub

    Dim fileList As Collection
    Set fileList = New Collection

    Dim f As Object

    For Each f In fso.GetFolder(sourceFolder).Files
        If ShouldMoveFileToPdfFolder(f.name) Then
            fileList.Add f.path
        End If
    Next f

    Dim i As Long
    Dim srcPath As String
    Dim destPath As String

    For i = 1 To fileList.count

        srcPath = CStr(fileList(i))

        If fso.FileExists(srcPath) Then

            destPath = GetUniqueFilePath(destFolder & "\" & fso.GetFileName(srcPath))

            If MoveFileAndDeleteOriginal(srcPath, destPath) Then
                movedCount = movedCount + 1
                LogLine "Moved PDF/Excel/report file to " & CurrentJobNumber & " pdfs: " & _
                        fso.GetFileName(destPath)
            Else
                LogLine "WARNING: Could not move PDF/Excel/report file: " & srcPath
            End If

        End If

    Next i

    Dim subFolder As Object
    Dim subPath As String

    For Each subFolder In fso.GetFolder(sourceFolder).SubFolders

        subPath = subFolder.path

        If LCase(subPath) <> LCase(destFolder) Then
            MovePdfExcelFilesRecursive subPath, destFolder, movedCount
        End If

    Next subFolder

    Exit Sub

ErrHandler:
    LogLine "MovePdfExcelFilesRecursive error: " & Err.Description
End Sub

Private Sub DeletePreviewImageFilesFromJob()
On Error GoTo ErrHandler

    If CurrentJobFolder = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(CurrentJobFolder) = False Then Exit Sub

    DeletePreviewImageFilesRecursive CurrentJobFolder

    Exit Sub

ErrHandler:
    LogLine "DeletePreviewImageFilesFromJob error: " & Err.Description
End Sub

Private Sub DeletePreviewImageFilesRecursive(ByVal folderPath As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) = False Then Exit Sub

    Dim fileList As Collection
    Set fileList = New Collection

    Dim f As Object
    Dim ext As String

    For Each f In fso.GetFolder(folderPath).Files
        ext = UCase(GetFileExtension(f.name))

        If ext = "JPG" Or ext = "JPEG" Or ext = "PNG" Then
            fileList.Add f.path
        End If
    Next f

    Dim i As Long
    For i = 1 To fileList.count
        fso.DeleteFile CStr(fileList(i)), True
        LogLine "Deleted preview image: " & CStr(fileList(i))
    Next i

    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(folderPath).SubFolders
        DeletePreviewImageFilesRecursive subFolder.path
    Next subFolder

    Exit Sub

ErrHandler:
    LogLine "DeletePreviewImageFilesRecursive error: " & Err.Description
End Sub

Private Sub WriteExportLogBomSummary()
On Error GoTo ErrHandler

    If ENABLE_EXPORT_LOG = False Then Exit Sub
    If BomCount <= 0 Then Exit Sub

    LogLine "========== BOM MATERIAL / PURCHASE LOG =========="
    LogLine "Type          Qty  Description                              L        W        T        Material"
    LogLine "--------------------------------------------------------------------------------"

    Dim i As Long
    Dim typ As String
    Dim line As String

    For i = 1 To BomCount
        typ = Trim(BomRows(i).TypeField)
        If typ = "" Then
            If Len(BomRows(i).material) > 0 Then
                typ = "Material"
            ElseIf IsPullcoreBomRow(BomRows(i)) Then
                typ = "Pullcore"
            Else
                typ = "Purchase"
            End If
        End If

        line = Left(typ & Space(14), 14) & _
               Right(Space(4) & CStr(BomRows(i).Quantity), 4) & "  " & _
               Left(BomRows(i).Description & Space(40), 40)

        If BomRows(i).hasDims Then
            line = line & "  " & Format(BomRows(i).BomLength, "0.000") & _
                   "  " & Format(BomRows(i).BomWidth, "0.000") & _
                   "  " & Format(BomRows(i).BomThickness, "0.000")
        End If

        If Len(BomRows(i).material) > 0 Then
            line = line & "  " & BomRows(i).material
        End If

        LogLine line
    Next i

    LogLine "========== END BOM LOG (" & BomCount & " rows) =========="
    Exit Sub

ErrHandler:
    LogLine "WriteExportLogBomSummary error: " & Err.Description
End Sub

Private Sub SearchJobFolderRecursive(ByVal folder As Object, _
                                     ByVal wantUpper As String, _
                                     ByVal depth As Long, _
                                     ByRef bestPath As String, _
                                     ByRef bestScore As Long)
On Error Resume Next

    If folder Is Nothing Then Exit Sub
    If depth > 3 Then Exit Sub

    Dim subFolder As Object
    Dim nameUpper As String
    Dim score As Long

    For Each subFolder In folder.SubFolders

        nameUpper = UCase(subFolder.name)

        If nameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextSub
        If InStr(nameUpper, " PRINTS") > 0 Then GoTo NextSub
        If InStr(nameUpper, "PULLCORE") > 0 Then GoTo NextSub
        If InStr(nameUpper, "PYROPEL") > 0 Then GoTo NextSub
        If InStr(nameUpper, "J BLOCK") > 0 Then GoTo NextSub

        score = GetJobFolderMatchScore(subFolder.name, wantUpper, depth)

        If score > bestScore Then
            bestScore = score
            bestPath = subFolder.path
        End If

        SearchJobFolderRecursive subFolder, wantUpper, depth + 1, bestPath, bestScore

NextSub:
    Next subFolder
End Sub

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

    If paths.count > 0 Then

        ReDim used(1 To paths.count)

        Dim k As Long
        Dim i As Long
        Dim bestI As Long
        Dim bestS As Long

        For k = 1 To paths.count

            bestI = 0
            bestS = -2147483647

            For i = 1 To paths.count
                If used(i) = False Then
                    If CLng(scores(i)) > bestS Then
                        bestS = CLng(scores(i))
                        bestI = i
                    End If
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

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderName As String
    folderName = UCase(folder.name)

    If InStr(folderName, " PRINTS") > 0 Then Exit Sub
    If folderName = "BASE" Or InStr(folderName, " BASE") > 0 Then Exit Sub

    Dim file As Object
    Dim ext As String
    Dim score As Long

    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.path))
        score = CadFilePriority(ext, file.name)

        If score > 0 Then
            paths.Add file.path
            scores.Add score
        End If

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
    Dim j As Long
    Dim p As String
    Dim dup As Boolean

    For i = 1 To extra.count

        p = CStr(extra(i))
        dup = False

        For j = 1 To target.count
            If LCase(CStr(target(j))) = LCase(p) Then
                dup = True
                Exit For
            End If
        Next j

        If dup = False Then target.Add p

    Next i
End Sub

Private Function CadFilePriority(ByVal ext As String, ByVal fileName As String) As Long

    Dim nameUpper As String
    nameUpper = UCase(fileName)

    If Left(fileName, 2) = "~$" Then CadFilePriority = 0: Exit Function
    If Left(nameUpper, 2) = "~$" Then CadFilePriority = 0: Exit Function

    Dim bonus As Long
    bonus = 0

    If InStr(nameUpper, CurrentJobNumber) > 0 Then bonus = bonus + 5
    If InStr(nameUpper, "ASSEM") > 0 Or InStr(nameUpper, "ASSY") > 0 Or InStr(nameUpper, "BASE") > 0 Then bonus = bonus + 3

    Select Case ext
        Case "sldasm"
            CadFilePriority = 100 + bonus
        Case "easm"
            CadFilePriority = 90 + bonus
        Case "asm"
            CadFilePriority = 85 + bonus
        Case "step", "stp"
            CadFilePriority = 80 + bonus
        Case "x_t", "x_b"
            CadFilePriority = 70 + bonus
        Case "igs", "iges"
            CadFilePriority = 60 + bonus
        Case "sldprt"
            CadFilePriority = 50 + bonus
        Case "prt"
            CadFilePriority = 45 + bonus
        Case Else
            CadFilePriority = 0
    End Select
End Function

' ============================================================
' OUTPUT NAMING
' ============================================================

Private Sub DetermineOutputNamingInfo(ByVal jobFolder As String)
On Error GoTo ErrHandler

    CustomerNumber = ""
    DateCode = ""
    NamingSourceText = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(jobFolder) Then Exit Sub

    SearchNamingInfoRecursive fso.GetFolder(jobFolder)

    If DateCode = "" Then DateCode = Format(Date, "mm-dd-yyyy")
    If CustomerNumber = "" Then CustomerNumber = CurrentJobNumber

    Exit Sub

ErrHandler:
    LogLine "DetermineOutputNamingInfo error: " & Err.Description
End Sub

Private Sub SearchNamingInfoRecursive(ByVal folder As Object)
On Error Resume Next

    Dim file As Object

    For Each file In folder.Files
        TryExtractNamingFromText file.name
        If CustomerNumber <> "" And DateCode <> "" Then Exit Sub
    Next file

    TryExtractNamingFromText folder.name

    Dim subFolder As Object

    For Each subFolder In folder.SubFolders
        SearchNamingInfoRecursive subFolder
        If CustomerNumber <> "" And DateCode <> "" Then Exit Sub
    Next subFolder
End Sub

Private Sub TryExtractNamingFromText(ByVal text As String)
On Error Resume Next

    If text = "" Then Exit Sub

    If CustomerNumber = "" Then

        Dim c As String
        c = ExtractCustomerNumberFromText(text)

        If c <> "" Then
            CustomerNumber = c
            If NamingSourceText = "" Then NamingSourceText = text
        End If

    End If

    If DateCode = "" Then

        Dim d As String
        d = ExtractDateCodeFromText(text)

        If d <> "" Then
            DateCode = d
            If NamingSourceText = "" Then NamingSourceText = text
        End If

    End If
End Sub

Private Function ExtractCustomerNumberFromText(ByVal text As String) As String
On Error GoTo ErrHandler

    Dim upper As String
    upper = UCase(text)

    Dim i As Long
    Dim runStart As Long
    Dim runLen As Long
    Dim ch As String
    Dim candidate As String

    runStart = 0
    runLen = 0

    For i = 1 To Len(upper) + 1

        If i <= Len(upper) Then
            ch = mid(upper, i, 1)
        Else
            ch = "X"
        End If

        If ch >= "0" And ch <= "9" Then

            If runLen = 0 Then runStart = i
            runLen = runLen + 1

        Else

            If runLen >= 8 Then
                candidate = mid(upper, runStart, runLen)

                If IsLikelyEightDigitDate(candidate) = False And IsMostlyZeros(candidate) = False Then
                    ExtractCustomerNumberFromText = candidate
                    Exit Function
                End If
            End If

            runLen = 0

        End If

    Next i

    Exit Function

ErrHandler:
    ExtractCustomerNumberFromText = ""
End Function

Private Function ExtractDateCodeFromText(ByVal text As String) As String
On Error GoTo ErrHandler

    Dim sep As Variant
    Dim guess As String

    For Each sep In Array("-", "_", ".", "/")

        guess = ExtractDateWithSeparator(text, CStr(sep))

        If guess <> "" Then
            ExtractDateCodeFromText = guess
            Exit Function
        End If

    Next sep

    Exit Function

ErrHandler:
    ExtractDateCodeFromText = ""
End Function

Private Function ExtractDateWithSeparator(ByVal text As String, ByVal sep As String) As String
On Error GoTo ErrHandler

    Dim cleaned As String
    cleaned = Replace(text, " ", "")

    Dim i As Long
    Dim n As Long
    Dim chunk As String
    Dim mm As String
    Dim dd As String
    Dim yyyy As String

    n = Len(cleaned)

    For i = 1 To n - 7

        chunk = mid(cleaned, i, 10)

        If Len(chunk) = 10 Then
            If mid(chunk, 3, 1) = sep And mid(chunk, 6, 1) = sep Then

                mm = Left(chunk, 2)
                dd = mid(chunk, 4, 2)
                yyyy = Right(chunk, 4)

                If IsNumeric(mm) And IsNumeric(dd) And IsNumeric(yyyy) Then
                    If IsValidMonthDayYear(CLng(mm), CLng(dd), CLng(yyyy)) Then
                        ExtractDateWithSeparator = mm & "-" & dd & "-" & yyyy
                        Exit Function
                    End If
                End If

            End If
        End If

    Next i

    Exit Function

ErrHandler:
    ExtractDateWithSeparator = ""
End Function

Private Function IsLikelyEightDigitDate(ByVal s As String) As Boolean
    If Len(s) <> 8 Then Exit Function
    If IsNumeric(s) = False Then Exit Function

    Dim mm As Long
    Dim dd As Long
    Dim yyyy As Long

    mm = CLng(Left(s, 2))
    dd = CLng(mid(s, 3, 2))
    yyyy = CLng(Right(s, 4))

    If IsValidMonthDayYear(mm, dd, yyyy) Then
        IsLikelyEightDigitDate = True
        Exit Function
    End If

    yyyy = CLng(Left(s, 4))
    mm = CLng(mid(s, 5, 2))
    dd = CLng(Right(s, 2))

    If IsValidMonthDayYear(mm, dd, yyyy) Then IsLikelyEightDigitDate = True
End Function

Private Function IsValidMonthDayYear(ByVal mm As Long, ByVal dd As Long, ByVal yyyy As Long) As Boolean
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function
    If yyyy < 2000 Or yyyy > 2100 Then Exit Function

    IsValidMonthDayYear = True
End Function

Private Function IsMostlyZeros(ByVal s As String) As Boolean
    Dim i As Long
    Dim zeros As Long

    For i = 1 To Len(s)
        If mid(s, i, 1) = "0" Then zeros = zeros + 1
    Next i

    If Len(s) > 0 Then
        If zeros / Len(s) >= 0.7 Then IsMostlyZeros = True
    End If
End Function

' ============================================================
' ARRAY / HEADER HELPERS
' ============================================================

Private Function GetArrayValue(ByVal data As Variant, ByVal r As Long, ByVal c As Long) As String
On Error Resume Next

    If c <= 0 Then Exit Function

    Dim v As Variant
    v = data(r, c)

    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function

    GetArrayValue = Trim(CStr(v))
End Function

Private Function FindBomHeaderLikeInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long, ByVal options As Variant) As Long
    Dim opt As Variant
    Dim c As Long
    Dim cellUpper As String

    For c = 1 To colCount

        cellUpper = UCase(GetArrayValue(data, r, c))

        If cellUpper <> "" Then
            For Each opt In options
                If cellUpper = UCase(CStr(opt)) Or InStr(cellUpper, UCase(CStr(opt))) > 0 Then
                    FindBomHeaderLikeInArrayRow = c
                    Exit Function
                End If
            Next opt
        End If

    Next c
End Function

Private Function FindBomQtyColumnInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long) As Long
On Error GoTo ErrHandler

    FindBomQtyColumnInArrayRow = 0

    Dim c As Long
    Dim h As String
    Dim k As String

    ' First pass: normal QTY-style headers.
    For c = 1 To colCount

        h = NormalizeText(GetArrayValue(data, r, c))
        k = NormalizeKey(h)

        If k <> "" Then
            If InStr(k, "QTY") > 0 Or _
               InStr(k, "QUANTITY") > 0 Or _
               InStr(k, "QNTY") > 0 Or _
               InStr(k, "QUAN") > 0 Then

                FindBomQtyColumnInArrayRow = c
                Exit Function

            End If
        End If

    Next c

    ' Second pass: Tempcraft / CMS style "No. Req'd" quantity header.
    For c = 1 To colCount

        h = NormalizeText(GetArrayValue(data, r, c))
        k = NormalizeKey(h)

        If k <> "" Then
            If InStr(k, "REQ") > 0 Then
                If InStr(k, "NO") > 0 Or _
                   InStr(k, "NUM") > 0 Or _
                   InStr(k, "NUMBER") > 0 Or _
                   InStr(k, "REQUIRED") > 0 Or _
                   InStr(k, "REQD") > 0 Then

                    FindBomQtyColumnInArrayRow = c
                    Exit Function

                End If
            End If
        End If

    Next c

    Exit Function

ErrHandler:
    FindBomQtyColumnInArrayRow = 0
End Function

Private Function FindBomMaterialColumnInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long) As Long
    FindBomMaterialColumnInArrayRow = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("MATERIAL", "MAT", "MTL", "STEEL"))
End Function

Private Sub FindBomDimensionColumnsInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long, _
                                              ByRef lenCol As Long, ByRef widCol As Long, ByRef thkCol As Long)

    lenCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("LENGTH", "LTH", "LEN"))
    widCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("WIDTH", "WTH", "WID"))
    thkCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("THICKNESS", "THICK", "THK", "HGT", "HEIGHT"))
End Sub

Private Function IsLikelyBomWorksheet(ByVal ws As Object) As Boolean
On Error Resume Next

    Dim n As String
    n = UCase(ws.name)

    If InStr(n, "BOM") > 0 Then IsLikelyBomWorksheet = True: Exit Function
    If InStr(n, "BILL") > 0 Then IsLikelyBomWorksheet = True: Exit Function
    If InStr(n, "MATERIAL") > 0 Then IsLikelyBomWorksheet = True: Exit Function
    If InStr(n, "QUOTE") > 0 Then IsLikelyBomWorksheet = True: Exit Function

    IsLikelyBomWorksheet = True
End Function

Private Function ShouldSkipBomWorksheet(ByVal ws As Object) As Boolean
On Error Resume Next

    Dim n As String
    n = UCase(ws.name)

    If InStr(n, "INSTRUCTION") > 0 Then ShouldSkipBomWorksheet = True: Exit Function
    If InStr(n, "NOTES") > 0 Then ShouldSkipBomWorksheet = True: Exit Function
    If InStr(n, "COVER") > 0 Then ShouldSkipBomWorksheet = True: Exit Function
End Function

' ============================================================
' NAME STANDARDIZATION
' ============================================================

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

        Dim loc As String
        Dim kind As String

        loc = GetPullcoreLocationCode(s)

        If InStr(s, "CAM") > 0 Then
            kind = "PULLCORE CAM"
        Else
            kind = "PULLCORE KEY"
        End If

        If loc <> "" Then
            StandardPlateName = loc & " " & kind
        Else
            StandardPlateName = kind
        End If

        Exit Function
    End If

    ' HOLDERS - side detected anywhere in the name, in any order.
    '
    ' The old test required the side word to come BEFORE "HOLDER" - it looked for
    ' "TOP HOLDER" and "BOTTOM HOLDER". A BOM reading
    '     "Holder Block - Top     Material"
    ' matched none of them, so the part never became ID HOLDER, never got a
    ' package, and fell through to MISC DETAILS. It also broke the top-side
    ' orientation, which tries the ID HOLDER / OD HOLDER pair FIRST and only
    ' falls back to the pots when that pair is missing.
    '
    ' Pot blocks never had this problem because the pot branch below uses
    ' IsLikelyIdSideName / IsLikelyOdSideName, which find the side word anywhere
    ' in the string. Holders now use exactly the same test, so:
    '     ANY combination of TOP + HOLDER    -> ID HOLDER
    '     ANY combination of BOTTOM + HOLDER -> OD HOLDER
    If InStr(s, "HOLDER") > 0 Then

        If IsLikelyIdSideName(s) Then
            StandardPlateName = "ID HOLDER"
            Exit Function
        End If

        If IsLikelyOdSideName(s) Then
            StandardPlateName = "OD HOLDER"
            Exit Function
        End If

        ' A holder with no side word at all: fall through rather than guess.
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

    If InStr(s, "FLIPPER CAM COVER") > 0 Or InStr(s, "CAM COVER PLATE") > 0 Then
        StandardPlateName = "FLIPPER CAM COVER PLATE"
        Exit Function
    End If

    StandardPlateName = ProperCaseText(Trim(desc))
    Exit Function

ErrHandler:
    StandardPlateName = Trim(desc)
End Function

Private Function IsLikelyIdSideName(ByVal s As String) As Boolean
    If InStr(s, "ID ") > 0 Or Left(s, 2) = "ID" Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, " TOP") > 0 Or Left(s, 3) = "TOP" Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, "TCP") > 0 Then IsLikelyIdSideName = True
End Function

Private Function IsLikelyOdSideName(ByVal s As String) As Boolean
    If InStr(s, "OD ") > 0 Or Left(s, 2) = "OD" Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "BOTTOM") > 0 Or InStr(s, "BOT ") > 0 Or Left(s, 3) = "BOT" Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "BCP") > 0 Then IsLikelyOdSideName = True
End Function

Private Function IsMainPrintsQuote(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "IDHOLDER", "ODHOLDER", "TCP", "BCP", "TOPINS", "BOTINS", _
             "SMEDPLATE", "IDPOTBLOCK", "ODPOTBLOCK", "POTBLOCK"
            IsMainPrintsQuote = True
    End Select
End Function

Private Function GetOutputFolderForRegularExport(ByVal quoteName As String, ByVal outputFolder As String) As String
    If IsMainPrintsQuote(quoteName) Then
        GetOutputFolderForRegularExport = outputFolder
    Else
        GetOutputFolderForRegularExport = CurrentJobFolder & "\" & CurrentJobNumber & " MISC DETAILS"
    End If
End Function

Private Function IsInsertQuoteName(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "TOPINS" Or k = "BOTINS" Then IsInsertQuoteName = True: Exit Function
    If InStr(k, "INSERT") > 0 Then IsInsertQuoteName = True: Exit Function
    If InStr(k, "INS") > 0 And Len(k) <= 16 Then IsInsertQuoteName = True
End Function

Private Function ShouldSkipExportQuoteName(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    If InStr(s, "J-BLOCK") > 0 Or InStr(s, "J BLOCK") > 0 Or InStr(s, "JBLOCK") > 0 Then
        ShouldSkipExportQuoteName = True
        Exit Function
    End If

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then
        ShouldSkipExportQuoteName = True
        Exit Function
    End If

    If InStr(s, "FLIPPER CAM COVER") > 0 Then
        ShouldSkipExportQuoteName = True
        Exit Function
    End If
End Function

Private Function IsNameMatch(ByVal cadName As String, ByVal quoteName As String) As Boolean
    Dim a As String
    Dim b As String

    a = NormalizeKey(cadName)
    b = NormalizeKey(quoteName)

    If a = "" Or b = "" Then Exit Function

    If InStr(a, b) > 0 Or InStr(b, a) > 0 Then IsNameMatch = True
End Function

' ============================================================
' TCP / BCP MASS HELPERS
' ============================================================

Private Function IsTcpBcpQuoteName(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "TCP", "BCP", _
             "TOPCLAMPINGPLATE", "BOTTOMCLAMPINGPLATE", _
             "TOPSMEDPLATE", "BOTTOMSMEDPLATE", _
             "TOPSMED", "BOTTOMSMED"
            IsTcpBcpQuoteName = True
    End Select
End Function

Private Function ShouldUseTcpBcpMassPairRule(ByRef b As BomInfo) As Boolean
On Error GoTo ErrHandler

    ShouldUseTcpBcpMassPairRule = False

    If b.hasDims = False Then Exit Function
    If IsTcpBcpQuoteName(b.quoteName) = False Then Exit Function

    Dim counterpart As String
    counterpart = GetCounterpartQuoteName(b.quoteName)

    If counterpart = "" Then Exit Function

    Dim i As Long

    For i = 1 To BomCount

        If NormalizeKey(BomRows(i).quoteName) = NormalizeKey(counterpart) Then
            If BomRows(i).hasDims Then
                If Abs(BomRows(i).BomLength - b.BomLength) <= SAME_SIZE_PAIR_TOL Then
                    If Abs(BomRows(i).BomWidth - b.BomWidth) <= SAME_SIZE_PAIR_TOL Then
                        If Abs(BomRows(i).BomThickness - b.BomThickness) <= SAME_SIZE_PAIR_TOL Then
                            ShouldUseTcpBcpMassPairRule = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

    Next i

    Exit Function

ErrHandler:
    ShouldUseTcpBcpMassPairRule = False
End Function

Private Function GetMassPreferenceForQuoteName(ByVal quoteName As String) As String
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "TOPINS" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "BOTINS" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function

    If k = "TCP" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "BCP" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function

    If k = "TOPCLAMPINGPLATE" Or k = "TOPSMEDPLATE" Or k = "TOPSMED" Then
        GetMassPreferenceForQuoteName = "LIGHT"
        Exit Function
    End If

    If k = "BOTTOMCLAMPINGPLATE" Or k = "BOTTOMSMEDPLATE" Or k = "BOTTOMSMED" Then
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
    If k = "TOPINS" Then GetCounterpartQuoteName = "BOT INS": Exit Function
    If k = "BOTINS" Then GetCounterpartQuoteName = "TOP INS": Exit Function
    If k = "TCP" Then GetCounterpartQuoteName = "BCP": Exit Function
    If k = "BCP" Then GetCounterpartQuoteName = "TCP": Exit Function

    If k = "TOPCLAMPINGPLATE" Or k = "TOPSMEDPLATE" Or k = "TOPSMED" Then
        GetCounterpartQuoteName = "BCP"
        Exit Function
    End If

    If k = "BOTTOMCLAMPINGPLATE" Or k = "BOTTOMSMEDPLATE" Or k = "BOTTOMSMED" Then
        GetCounterpartQuoteName = "TCP"
        Exit Function
    End If
End Function

Private Function NormalizeSteelType(ByVal matText As String) As String
    Dim s As String
    s = UCase(Trim(matText))

    If s = "" Then NormalizeSteelType = DEFAULT_STEEL_TYPE: Exit Function

    If InStr(s, "PYROPEL") > 0 Then NormalizeSteelType = "Pyropel": Exit Function
    If InStr(s, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
    If InStr(s, "4130") > 0 Then NormalizeSteelType = "4130": Exit Function
    If InStr(s, "H-13") > 0 Or InStr(s, "H13") > 0 Then NormalizeSteelType = "H13": Exit Function
    If InStr(s, "A-2") > 0 Or InStr(s, "A2") > 0 Then NormalizeSteelType = "A-2": Exit Function
    If InStr(s, "D-2") > 0 Or InStr(s, "D2") > 0 Then NormalizeSteelType = "D-2": Exit Function
    If InStr(s, "DRILL ROD") > 0 Then NormalizeSteelType = "Drill Rod": Exit Function
    If InStr(s, "STAINLESS") > 0 Or InStr(s, "SS") > 0 Then NormalizeSteelType = "Stainless": Exit Function

    NormalizeSteelType = Trim(matText)
End Function

' True when the material is one of the base steel grades in ALLOWED_STEEL_TYPES.
' An empty material falls through NormalizeSteelType to DEFAULT_STEEL_TYPE, so a
' blank BOM cell is treated as the default grade rather than being thrown out.
Private Function IsAllowedSteelType(ByVal matText As String) As Boolean
On Error Resume Next

    IsAllowedSteelType = False

    Dim norm As String
    norm = UCase(Trim(NormalizeSteelType(matText)))

    If norm = "" Then Exit Function

    Dim allowed() As String
    allowed = Split(UCase(ALLOWED_STEEL_TYPES), "|")

    Dim i As Long

    For i = LBound(allowed) To UBound(allowed)
        If Trim(allowed(i)) <> "" Then
            If norm = Trim(allowed(i)) Then
                IsAllowedSteelType = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsHardwareName(ByVal d As String) As Boolean
    If InStr(d, "SCREW") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "BOLT") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "WASHER") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "DOWEL") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "NUT ") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "SHCS") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "SOCKET HEAD") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "SPRING") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "PIN") > 0 And InStr(d, "PINION") = 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "O-RING") > 0 Or InStr(d, "ORING") > 0 Then IsHardwareName = True: Exit Function
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

Private Function ContainsAnyPipeKey(ByVal haystack As String, ByVal pipeKeys As String) As Boolean
On Error GoTo ErrHandler

    ContainsAnyPipeKey = False

    If haystack = "" Then Exit Function
    If pipeKeys = "" Then Exit Function

    Dim hayUpper As String
    hayUpper = NormalizeText(haystack)

    Dim keys() As String
    keys = Split(pipeKeys, "|")

    Dim i As Long
    Dim k1 As String

    For i = LBound(keys) To UBound(keys)
        k1 = NormalizeText(keys(i))
        If k1 <> "" Then
            If InStr(hayUpper, k1) > 0 Then
                ContainsAnyPipeKey = True
                Exit Function
            End If
        End If
    Next i

    Dim hayKey As String
    hayKey = NormalizeKey(haystack)

    Dim k2 As String

    For i = LBound(keys) To UBound(keys)
        k2 = NormalizeKey(keys(i))
        If k2 <> "" Then
            If InStr(hayKey, k2) > 0 Then
                ContainsAnyPipeKey = True
                Exit Function
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    ContainsAnyPipeKey = False
End Function

' ============================================================
' QUARTER-INCH PHYSICAL HIDE
' ============================================================

Private Sub HideQuarterInchThicknessItems(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    HideQuarterInchAssemblyComponents model
    Exit Sub

ErrHandler:
    LogLine "HideQuarterInchThicknessItems error: " & Err.Description
End Sub

Private Sub HideQuarterInchAssemblyComponents(ByVal assyModel As Object)
On Error Resume Next

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long
    Dim swComp As Object
    Dim swCompModel As Object
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    Dim L As Double
    Dim W As Double
    Dim T As Double

    For i = 0 To UBound(vComps)

        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                Set swCompModel = swComp.GetModelDoc2

                If Not swCompModel Is Nothing Then
                    If swCompModel.GetType = swDocPART Then

                        If GetPartBoundingBoxInches(swCompModel, dX, dY, dZ) Then
                            SortThreeDimensions dX, dY, dZ, L, W, T

                            If IsQuarterInchThickness(T) Then
                                swComp.Visible = swComponentHidden
                            End If
                        End If

                    End If
                End If

            End If
        End If

    Next i
End Sub

' ============================================================
' GENERAL HELPERS
' ============================================================

Private Function NormalizeText(ByVal s As String) As String
    s = UCase(Trim(s))
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    NormalizeText = NormalizeSpaces(s)
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeSpaces = Trim(s)
End Function

Private Function NormalizeKey(ByVal s As String) As String
    s = UCase(Trim(s))
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, ".", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    NormalizeKey = s
End Function

Private Function ProperCaseText(ByVal s As String) As String
On Error Resume Next

    ProperCaseText = StrConv(s, vbProperCase)

    If ProperCaseText = "" Then ProperCaseText = s
End Function

Private Function IsQuarterInchThickness(ByVal T As Double) As Boolean
    IsQuarterInchThickness = (Abs(T - QUARTER_INCH_THICKNESS) <= QUARTER_INCH_TOLERANCE)
End Function

Private Function IsComponentHidden(ByVal swComp As Object) As Boolean
On Error Resume Next

    If swComp Is Nothing Then Exit Function

    IsComponentHidden = (swComp.Visible = swComponentHidden)
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

Private Function CleanDisplayName(ByVal compName As String, ByVal compPath As String) As String
On Error Resume Next

    Dim s As String
    s = compName

    Dim dashPos As Long
    dashPos = InStrRev(s, "-")

    If dashPos > 1 Then
        Dim tail As String
        tail = mid(s, dashPos + 1)

        If IsNumeric(tail) Then s = Left(s, dashPos - 1)
    End If

    If s = "" And compPath <> "" Then s = GetFileBaseName(compPath)

    CleanDisplayName = Trim(s)
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

Private Function CleanQuoteTokenForFile(ByVal s As String) As String
    s = CleanFileName(s)
    s = NormalizeSpaces(s)

    CleanQuoteTokenForFile = Trim(s)
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

    If parent <> "" And fso.FolderExists(parent) = False Then
        EnsureFolderDeep parent
    End If

    If fso.FolderExists(folderPath) = False Then fso.CreateFolder folderPath
End Sub

Private Sub CollectJobPdfsAndReports()
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim destFolder As String
    destFolder = CurrentJobFolder & "\" & CurrentJobNumber & " pdfs"

    EnsureFolderDeep destFolder

    Dim copied As Long
    copied = 0

    copied = copied + CopyFilesByExtension(CurrentJobFolder, destFolder, "pdf", destFolder, True)

    Dim names As Variant
    names = Array("XT_Export_BOM_Match_Report.csv", _
                  "XT_Export_CAD_Dimensions.csv", _
                  "XT_Export_BOM_PDF_Text.txt", _
                  JOB_SIGNATURE_REPORT_FILE, _
                  PULLCORE_DIMENSIONS_REPORT_FILE)

    Dim i As Long
    Dim src As String

    For i = LBound(names) To UBound(names)

        src = CurrentJobFolder & "\" & CStr(names(i))

        If fso.FileExists(src) Then
            fso.CopyFile src, destFolder & "\" & CStr(names(i)), True
            copied = copied + 1
        End If

    Next i

    LogLine "Collected " & copied & " PDF/report file(s) into: " & destFolder
    Exit Sub

ErrHandler:
    LogLine "CollectJobPdfsAndReports error: " & Err.Description
End Sub

Private Function CopyFilesByExtension(ByVal srcFolder As String, ByVal destFolder As String, _
                                      ByVal ext As String, ByVal skipFolder As String, _
                                      Optional ByVal deleteSource As Boolean = False) As Long
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(srcFolder) Then Exit Function

    Dim folder As Object
    Set folder = fso.GetFolder(srcFolder)

    Dim count As Long
    count = 0

    Dim f As Object
    Dim srcPath As String

    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.path)) = LCase(ext) Then
            srcPath = f.path
            fso.CopyFile srcPath, destFolder & "\" & f.name, True
            count = count + 1

            If deleteSource Then fso.DeleteFile srcPath, True
        End If
    Next f

    Dim sub_ As Object

    For Each sub_ In folder.SubFolders
        If LCase(sub_.path) <> LCase(skipFolder) Then
            count = count + CopyFilesByExtension(sub_.path, destFolder, ext, skipFolder, deleteSource)
        End If
    Next sub_

    CopyFilesByExtension = count
End Function

Private Function GetWritableCsvPath(ByVal csvPath As String) As String
On Error GoTo ErrHandler

    Dim f As Integer
    f = FreeFile

    On Error Resume Next
    Err.Clear
    Open csvPath For Output As #f

    If Err.Number = 0 Then
        Close #f
        GetWritableCsvPath = csvPath
        Exit Function
    End If

    Err.Clear

    Dim alt As String
    alt = AppendBeforeExtension(csvPath, "_" & Format(Now, "yyyymmdd_hhnnss"))

    LogLine "CSV locked, writing to alternate file instead: " & alt

    GetWritableCsvPath = alt
    Exit Function

ErrHandler:
    GetWritableCsvPath = csvPath
End Function

Private Function AppendBeforeExtension(ByVal path As String, ByVal suffix As String) As String
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As String
    Dim base As String
    Dim ext As String

    folder = fso.GetParentFolderName(path)
    base = fso.GetBaseName(path)
    ext = fso.GetExtensionName(path)

    If ext = "" Then
        AppendBeforeExtension = folder & "\" & base & suffix
    Else
        AppendBeforeExtension = folder & "\" & base & suffix & "." & ext
    End If
End Function

Private Function GetUniqueFilePath(ByVal path As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) = False Then
        GetUniqueFilePath = path
        Exit Function
    End If

    Dim folder As String
    Dim base As String
    Dim ext As String

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

Private Function CsvText(ByVal s As String) As String
    s = Replace(s, Chr(34), Chr(34) & Chr(34))
    CsvText = Chr(34) & s & Chr(34)
End Function

Private Function FormatNumberForCsv(ByVal v As Double) As String
    FormatNumberForCsv = Format(v, "0.000")
End Function

Private Function ReadAllTextFile(ByVal path As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) = False Then Exit Function

    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1, False)

    If Not ts.AtEndOfStream Then
        ReadAllTextFile = ts.ReadAll
    End If

    ts.Close
    Exit Function

ErrHandler:
    ReadAllTextFile = ""
End Function

Private Sub ExtractDecimalNumbers(ByVal line As String, ByRef nums() As Double, ByRef numCount As Long)
On Error GoTo ErrHandler

    numCount = 0
    ReDim nums(0 To 50)

    Dim i As Long
    Dim ch As String
    Dim token As String

    token = ""

    For i = 1 To Len(line) + 1

        If i <= Len(line) Then
            ch = mid(line, i, 1)
        Else
            ch = " "
        End If

        If (ch >= "0" And ch <= "9") Or ch = "." Then

            token = token & ch

        Else

            If token <> "" Then

                If IsNumeric(token) Then
                    If InStr(token, ".") > 0 Then

                        nums(numCount) = CDbl(token)
                        numCount = numCount + 1

                        If numCount > 50 Then Exit For

                    End If
                End If

                token = ""

            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    numCount = 0
End Sub

Private Function ExtractQtyAfterOutsource(ByVal upperLine As String) As Long
On Error GoTo ErrHandler

    ExtractQtyAfterOutsource = 1

    Dim kw As Variant
    Dim p As Long

    For Each kw In Array("OUTSOURCE", "PURCHASE", "MAKE", "BUY", "STOCK")

        p = InStr(upperLine, CStr(kw))

        If p > 0 Then

            Dim tail As String
            tail = mid(upperLine, p + Len(CStr(kw)))

            Dim i As Long
            Dim ch As String
            Dim token As String

            token = ""

            For i = 1 To Len(tail)

                ch = mid(tail, i, 1)

                If ch >= "0" And ch <= "9" Then
                    token = token & ch
                ElseIf token <> "" Then
                    Exit For
                End If

            Next i

            If token <> "" Then
                ExtractQtyAfterOutsource = CLng(token)
                Exit Function
            End If

        End If

    Next kw

    Exit Function

ErrHandler:
    ExtractQtyAfterOutsource = 1
End Function

Private Function RemoveLeadingItemNumber(ByVal line As String) As String
On Error Resume Next

    Dim s As String
    s = Trim(line)

    Dim i As Long
    Dim ch As String
    Dim digits As String

    digits = ""

    For i = 1 To Len(s)
        ch = mid(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            Exit For
        End If
    Next i

    If digits <> "" And Len(digits) <= 4 Then
        s = Trim(mid(s, Len(digits) + 1))
    End If

    RemoveLeadingItemNumber = s
End Function

' ============================================================
' MATCH STUDIO MATERIAL + CENTER OF MASS HELPERS
' ============================================================

Private Sub ApplyMatchStudioCarbonSteelToDocument(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub

    If model.GetType = swDocPART Then
        ApplyMatchStudioCarbonSteelToPart model, ""
        Exit Sub
    End If

    If model.GetType <> swDocASSEMBLY Then Exit Sub

    Dim vComps As Variant
    vComps = model.GetComponents(False)

    If IsEmpty(vComps) Then Exit Sub
    If IsArray(vComps) = False Then Exit Sub

    Dim i As Long
    Dim comp As Object
    Dim partDoc As Object
    Dim cfg As String

    For i = 0 To UBound(vComps)

        Set comp = vComps(i)

        If Not comp Is Nothing Then
            If comp.IsSuppressed = False Then

                Set partDoc = comp.GetModelDoc2
                cfg = ""

                On Error Resume Next
                cfg = comp.ReferencedConfiguration
                On Error GoTo ErrHandler

                If Not partDoc Is Nothing Then
                    If partDoc.GetType = swDocPART Then
                        ApplyMatchStudioCarbonSteelToPart partDoc, cfg
                    End If
                End If

            End If
        End If

    Next i

    On Error Resume Next
    model.ForceRebuild3 False
    On Error GoTo 0

    LogLine "Match Studio material applied: " & MATCH_STUDIO_CARBON_STEEL_MATERIAL
    Exit Sub

ErrHandler:
    LogLine "ApplyMatchStudioCarbonSteelToDocument error: " & Err.Description
End Sub

Private Sub ApplyMatchStudioCarbonSteelToPart(ByVal partModel As Object, ByVal configName As String)
On Error GoTo ErrHandler

    If partModel Is Nothing Then Exit Sub
    If partModel.GetType <> swDocPART Then Exit Sub

    Dim cfg As String
    cfg = Trim(configName)

    If cfg = "" Then
        On Error Resume Next
        cfg = partModel.ConfigurationManager.ActiveConfiguration.name
        On Error GoTo ErrHandler
    End If

    If cfg = "" Then cfg = "Default"

    On Error Resume Next
    Err.Clear

    partModel.SetMaterialPropertyName2 cfg, _
                                       MATCH_STUDIO_MATERIAL_DATABASE, _
                                       MATCH_STUDIO_CARBON_STEEL_MATERIAL

    If Err.Number <> 0 Then
        LogLine "WARNING: Could not apply carbon steel material to " & _
                partModel.GetTitle & " config " & cfg & ": " & Err.Description
        Err.Clear
    End If

    partModel.ForceRebuild3 False
    On Error GoTo 0

    Exit Sub

ErrHandler:
    LogLine "ApplyMatchStudioCarbonSteelToPart error: " & Err.Description
End Sub

Private Function TryGetComponentMassCenterInches(ByVal swComp As Object, _
                                                 ByRef cx As Double, _
                                                 ByRef cy As Double, _
                                                 ByRef cz As Double, _
                                                 ByRef massOut As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentMassCenterInches = False

    cx = 0#
    cy = 0#
    cz = 0#
    massOut = 0#

    If swComp Is Nothing Then Exit Function

    Dim partModel As Object
    Set partModel = swComp.GetModelDoc2

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim mp As Object
    Set mp = partModel.Extension.CreateMassProperty

    If mp Is Nothing Then Exit Function

    Dim vCom As Variant
    vCom = mp.CenterOfMass

    If IsEmpty(vCom) Then Exit Function
    If IsArray(vCom) = False Then Exit Function
    If UBound(vCom) < 2 Then Exit Function

    Dim volCuIn As Double
    volCuIn = 0#

    On Error Resume Next
    massOut = CDbl(mp.Mass)
    volCuIn = CDbl(mp.Volume) * CUIN_PER_CUBIC_METER
    On Error GoTo ErrHandler

    If MATCH_STUDIO_USE_CARBON_STEEL_DENSITY_FOR_MASS Then
        If volCuIn > 0# Then
            massOut = volCuIn * MATCH_STUDIO_CARBON_STEEL_DENSITY_LB_PER_CUIN
        End If
    End If

    Dim xform As Object
    Set xform = swComp.Transform2

    If xform Is Nothing Then Exit Function

    Dim m As Variant
    m = xform.ArrayData

    If IsEmpty(m) Then Exit Function
    If IsArray(m) = False Then Exit Function
    If UBound(m) < 12 Then Exit Function

    Dim scaleVal As Double
    scaleVal = CDbl(m(12))
    If Abs(scaleVal) < 0.0000001 Then scaleVal = 1#

    Dim px As Double
    Dim py As Double
    Dim pz As Double

    px = CDbl(vCom(0))
    py = CDbl(vCom(1))
    pz = CDbl(vCom(2))

    Dim tx As Double
    Dim ty As Double
    Dim tz As Double

    tx = scaleVal * ((px * CDbl(m(0))) + (py * CDbl(m(3))) + (pz * CDbl(m(6)))) + CDbl(m(9))
    ty = scaleVal * ((px * CDbl(m(1))) + (py * CDbl(m(4))) + (pz * CDbl(m(7)))) + CDbl(m(10))
    tz = scaleVal * ((px * CDbl(m(2))) + (py * CDbl(m(5))) + (pz * CDbl(m(8)))) + CDbl(m(11))

    cx = tx * INCHES_PER_METER
    cy = ty * INCHES_PER_METER
    cz = tz * INCHES_PER_METER

    TryGetComponentMassCenterInches = True
    Exit Function

ErrHandler:
    TryGetComponentMassCenterInches = False
End Function

' ============================================================
' LOGGING
' ============================================================

Private Sub LogLine(ByVal msg As String)
On Error Resume Next

    If ENABLE_EXPORT_LOG = False Then Exit Sub

    Dim f As Integer
    f = FreeFile

    Dim path As String
    path = RunLogPath

    If path = "" Then path = StartupLogPath
    If path = "" Then path = Environ$("USERPROFILE") & "\Desktop\CMS_XT_Export_STARTUP_Log.txt"

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

Private Sub LogProgress(ByVal msg As String)
    If VERBOSE_PROGRESS_LOG Then LogLine "    ... " & msg
End Sub

Private Sub LogErrorText(ByVal msg As String)
    LogLine "ERROR: " & msg
    LastJobFailReason = msg

    If SHOW_ERROR_MESSAGES Then
        ' Batch-safe: no blocking message box here.
    End If
End Sub

' ============================================================
' TOP-CLAMPING-PLATE "NORMAL TO" ORIENTATION
' ------------------------------------------------------------
' The by-hand procedure, in code, with nothing moving:
'
'   1. Find the top clamping plate (by matched quote, falling back to name keys).
'   2. Take its largest flat face - the one facing away from the BCP, so it is
'      the plate's outer face and not the one against the stack.
'   3. Look Normal To that face. The face normal becomes the view direction,
'      which takes out the tilt.
'   4. Align the view to that face's longest straight edge, which takes out the
'      in-plane spin. Normal To alone does not do this, which is why a plate
'      can be face-on and still sit crooked on screen.
'   5. Save the result as the named view CMS_TCP_NORMAL and redefine *Top from
'      it. Component transforms are never touched.
'
' Then the standard orientation process runs and decides top vs bottom and the
' front face as it always did.
' ============================================================

' Squares the base off TWO real faces:
'   TCP large face normal -> UP
'   pot block side face   -> FRONT
'
' Why one write and not two Normal To presses: UpdateStandardViews regenerates the
' whole orthogonal set of standard views from the one current view. So the view is
' built ONCE as the front view - looking along the pot normal, with the TCP normal
' as screen-up - and persisted as *Front. SolidWorks then derives *Top from it,
' and because up was the TCP normal, *Top comes out looking straight down the
' clamping plate face. Both views end up square from a single write, and there is
' no second write that could spoil the first.
Private Function TryOrientFromTcpTopAndPotFront(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    ' PATCHED:
    ' Keep the real Normal To operation ONLY for the TCP/top clamping plate.
    ' Do NOT press Normal To on the pot block anymore, because SolidWorks can
    ' re-spin/rerotate the view when the pot face is normalized.
    '
    ' Workflow now:
    '   1. Pick the TCP/top clamping plate face.
    '   2. Press Normal To on that TCP face.
    '   3. Persist that as SolidWorks *Top.
    '   4. Return True so the next step can define *Front using the existing
    '      holder/pot center-of-mass logic, with pot blocks forced in front.
    '
    ' No geometry moves.

    TryOrientFromTcpTopAndPotFront = False

    If model Is Nothing Then Exit Function

    If model.GetType <> swDocASSEMBLY Then
        LogLine "TCP Normal To only skipped: document is not an assembly."
        Exit Function
    End If

    Dim tcpComp As Object
    Set tcpComp = FindTopClampingPlateComponent(model)

    If tcpComp Is Nothing Then
        LogLine "TCP Normal To only skipped: top clamping plate component not found."
        Exit Function
    End If

    Dim tcpFace As Object
    Dim tcpArea As Double

    Set tcpFace = FindBestTcpOuterPlanarFaceOnComponent(model, tcpComp, tcpArea)

    If tcpFace Is Nothing Then
        LogLine "TCP Normal To only skipped: no usable planar face found on the TCP."
        Exit Function
    End If

    Dim topX As Double
    Dim topY As Double
    Dim topZ As Double

    If TryGetFaceNormalInAssembly(tcpFace, tcpComp, topX, topY, topZ) = False Then
        LogLine "TCP Normal To only failed: could not read TCP face normal."
        Exit Function
    End If

    OrientOutwardAwayFromBcp model, tcpComp, topX, topY, topZ

    LogLine "TCP Normal To only: TCP component " & BaseStraightenComponentName(tcpComp)
    LogLine "TCP Normal To only: face area " & Format(tcpArea, "0.000") & " sq in."
    LogLine "TCP Normal To only: UP normal = " & _
            Format(topX, "0.0000") & ", " & _
            Format(topY, "0.0000") & ", " & _
            Format(topZ, "0.0000")

    Dim graphicsWereOff As Boolean
    graphicsWereOff = MainViewportGraphicsDisabled

    If graphicsWereOff Then
        RestoreMainViewportGraphics
        LogLine "TCP Normal To only: viewport graphics re-enabled so Normal To can take effect."
    End If

    Dim actErrs As Long
    On Error Resume Next
    swApp.Visible = True
    swApp.UserControl = True
    swApp.ActivateDoc3 model.GetTitle, False, 0, actErrs
    Err.Clear
    On Error GoTo ErrHandler

    If ClickFaceAndPressNormalTo(model, tcpFace, tcpComp, topX, topY, topZ, "top clamping plate") = False Then
        LogLine "TCP Normal To only FAILED: Normal To did not square the TCP face."
        GoTo CleanExit
    End If

    LogLine "TCP Normal To only: Normal To on TCP succeeded."

    On Error Resume Next
    model.DeleteNamedView TCP_NORMAL_VIEW_NAME
    Err.Clear
    model.NameView TCP_NORMAL_VIEW_NAME
    Err.Clear
    On Error GoTo ErrHandler

    LogLine "TCP Normal To only: named view " & TCP_NORMAL_VIEW_NAME & " saved."

    If PersistCurrentViewAsStandardTop(model) = False Then
        LogLine "TCP Normal To only FAILED: could not persist TCP Normal To as SolidWorks *Top."
        GoTo CleanExit
    End If

    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 200

    Dim topOffDeg As Double

    If TryGetAngleBetweenFaceNormalAndViewDeg(model, topX, topY, topZ, topOffDeg) Then

        LogLine "TCP Normal To only: persisted *Top is " & _
                Format(topOffDeg, "0.000") & _
                " deg off the TCP face."

        If topOffDeg > TCP_NORMAL_MAX_OFF_DEG Then
            LogLine "TCP Normal To only WARNING: *Top verification is outside tolerance."
            GoTo CleanExit
        End If

    Else
        LogLine "TCP Normal To only: could not verify persisted *Top, continuing because Normal To reported success."
    End If

    LogLine "TCP Normal To only DONE. *Front will be defined by holder/pot COM logic."
    TryOrientFromTcpTopAndPotFront = True

CleanExit:

    If graphicsWereOff Then
        DisableMainViewportGraphics
    End If

    Exit Function

ErrHandler:
    LogLine "TryOrientFromTcpTopAndPotFront TCP-only patch error: " & Err.Description
    TryOrientFromTcpTopAndPotFront = False

    On Error Resume Next
    If graphicsWereOff Then DisableMainViewportGraphics
End Function

' ============================================================
' CLICK A FACE, PRESS NORMAL TO
' ------------------------------------------------------------
' The literal UI action. The face is picked with a ray in assembly space - the
' API equivalent of clicking it on screen - and then the real View > Normal To
' command is run, so SolidWorks does the orientation maths itself.
'
' Why not write the view matrix: on J8494 that produced a view 90 deg out.
' Orientation3's setter does not consume the same ArrayData layout its getter
' reports, so a matrix that reads back correctly does not write back correctly.
'
' The Normal To command ID moves between releases, so the candidates in
' NORMAL_TO_COMMAND_IDS are tried in turn. After each one the face is measured
' against the view; whichever ID actually squares the face is the right one and
' gets logged by number so the list can be pinned down later.
' ============================================================

Private Function ClickFaceAndPressNormalTo(ByVal model As Object, _
                                           ByVal faceObj As Object, _
                                           ByVal comp As Object, _
                                           ByVal nx As Double, _
                                           ByVal ny As Double, _
                                           ByVal nz As Double, _
                                           ByVal label As String) As Boolean
On Error GoTo ErrHandler

    ClickFaceAndPressNormalTo = False

    If model Is Nothing Then Exit Function
    If faceObj Is Nothing Then Exit Function

    If SelectFaceByRayInAssembly(model, faceObj, comp, nx, ny, nz) = False Then
        LogLine "Normal To (" & label & "): could not select the face."
        Exit Function
    End If

    LogLine "Normal To (" & label & "): face selected."

    ' PATCH:
    ' In this SolidWorks install the guessed RunCommand IDs do not execute Normal To.
    ' The log showed every ID leaving the face at the same 36.251/54.733 deg,
    ' which means the command never fired. Use the real keyboard shortcut first.
    If USE_CTRL8_SENDKEYS_FOR_NORMAL_TO Then

        ' Re-select immediately before Ctrl+8 so SolidWorks has an active face selection.
        If SelectFaceByRayInAssembly(model, faceObj, comp, nx, ny, nz) Then

            If PressNormalToWithCtrl8(model, nx, ny, nz, label) Then
                LogLine "Normal To (" & label & "): Ctrl+8 worked."
                model.ClearSelection2 True
                ClickFaceAndPressNormalTo = True
                Exit Function
            Else
                LogLine "Normal To (" & label & "): Ctrl+8 did not square the face; trying command IDs."
            End If

        Else
            LogLine "Normal To (" & label & "): could not re-select face before Ctrl+8."
        End If

    End If

    Dim ids() As String
    ids = Split(NORMAL_TO_COMMAND_IDS, "|")

    Dim i As Long
    Dim cmdId As Long
    Dim offDeg As Double

    For i = LBound(ids) To UBound(ids)

        cmdId = 0
        On Error Resume Next
        cmdId = CLng(Trim(ids(i)))
        Err.Clear
        On Error GoTo ErrHandler

        If cmdId > 0 Then

            ' Re-select: running a command can clear the selection.
            SelectFaceByRayInAssembly model, faceObj, comp, nx, ny, nz

            On Error Resume Next
            Err.Clear
            swApp.RunCommand cmdId, ""
            Err.Clear
            On Error GoTo ErrHandler

            StabilizeActiveView model, 250

            If TryGetAngleBetweenFaceNormalAndViewDeg(model, nx, ny, nz, offDeg) Then

                If offDeg <= TCP_NORMAL_MAX_OFF_DEG Then
                    LogLine "Normal To (" & label & "): command id " & cmdId & " worked. Face is " & _
                            Format(offDeg, "0.000") & " deg off. Pin NORMAL_TO_COMMAND_IDS to " & cmdId & "."
                    model.ClearSelection2 True
                    ClickFaceAndPressNormalTo = True
                    Exit Function
                End If

                LogLine "Normal To (" & label & "): command id " & cmdId & " left the face " & _
                        Format(offDeg, "0.000") & " deg off."

            Else
                LogLine "Normal To (" & label & "): command id " & cmdId & " - could not measure."
            End If

        End If

    Next i

    model.ClearSelection2 True

    LogLine "Normal To (" & label & "): none of the candidate command ids squared the face."

    Exit Function

ErrHandler:
    LogLine "ClickFaceAndPressNormalTo error: " & Err.Description

    On Error Resume Next
    model.ClearSelection2 True

    ClickFaceAndPressNormalTo = False
End Function

' Picks the face with a ray fired at it from outside, along the inward normal.
' This is the API equivalent of clicking the face in the graphics area, and it
' works in assembly context where selecting a part-space face object does not.
Private Function SelectFaceByRayInAssembly(ByVal model As Object, _
                                           ByVal faceObj As Object, _
                                           ByVal comp As Object, _
                                           ByVal nx As Double, _
                                           ByVal ny As Double, _
                                           ByVal nz As Double) As Boolean
On Error GoTo ErrHandler

    SelectFaceByRayInAssembly = False

    If model Is Nothing Then Exit Function
    If faceObj Is Nothing Then Exit Function

    Dim px As Double
    Dim py As Double
    Dim pz As Double

    If TryGetPointOnFaceInAssembly(faceObj, comp, px, py, pz) = False Then Exit Function

    ' Start outside the part and fire back along the normal into the face.
    Dim ox As Double
    Dim oy As Double
    Dim oz As Double

    ox = px + nx * FACE_PICK_RAY_STANDOFF_M
    oy = py + ny * FACE_PICK_RAY_STANDOFF_M
    oz = pz + nz * FACE_PICK_RAY_STANDOFF_M

    model.ClearSelection2 True

    Dim ok As Boolean
    ok = False

    On Error Resume Next
    Err.Clear
    ok = CBool(model.Extension.SelectByRay(ox, oy, oz, -nx, -ny, -nz, _
                                           FACE_PICK_RAY_RADIUS_M, swSelFACES, _
                                           False, 0, 0))
    If Err.Number <> 0 Then
        Err.Clear
        ok = False
    End If
    On Error GoTo ErrHandler

    If ok = False Then
        ' Second try from the other side, in case the normal pointed inward.
        ox = px - nx * FACE_PICK_RAY_STANDOFF_M
        oy = py - ny * FACE_PICK_RAY_STANDOFF_M
        oz = pz - nz * FACE_PICK_RAY_STANDOFF_M

        On Error Resume Next
        Err.Clear
        ok = CBool(model.Extension.SelectByRay(ox, oy, oz, nx, ny, nz, _
                                               FACE_PICK_RAY_RADIUS_M, swSelFACES, _
                                               False, 0, 0))
        If Err.Number <> 0 Then
            Err.Clear
            ok = False
        End If
        On Error GoTo ErrHandler
    End If

    SelectFaceByRayInAssembly = ok

    Exit Function

ErrHandler:
    SelectFaceByRayInAssembly = False
End Function

' A point that is genuinely ON the face, expressed in assembly space. Averages
' the face's vertices then pulls that average onto the face with
' GetClosestPointOn, so it lands inside the face and not on a corner.
Private Function TryGetPointOnFaceInAssembly(ByVal faceObj As Object, _
                                             ByVal comp As Object, _
                                             ByRef outX As Double, _
                                             ByRef outY As Double, _
                                             ByRef outZ As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPointOnFaceInAssembly = False

    outX = 0#
    outY = 0#
    outZ = 0#

    If faceObj Is Nothing Then Exit Function

    Dim vVerts As Variant

    On Error Resume Next
    vVerts = faceObj.GetVertices
    Err.Clear
    On Error GoTo ErrHandler

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim n As Long

    cx = 0#: cy = 0#: cz = 0#
    n = 0

    If IsArray(vVerts) Then

        Dim i As Long
        Dim vtx As Object
        Dim vp As Variant

        For i = LBound(vVerts) To UBound(vVerts)

            Set vtx = Nothing
            On Error Resume Next
            Set vtx = vVerts(i)
            Err.Clear
            On Error GoTo ErrHandler

            If Not vtx Is Nothing Then

                On Error Resume Next
                vp = vtx.GetPoint
                Err.Clear
                On Error GoTo ErrHandler

                If IsArray(vp) Then
                    If UBound(vp) >= 2 Then
                        cx = cx + CDbl(vp(0))
                        cy = cy + CDbl(vp(1))
                        cz = cz + CDbl(vp(2))
                        n = n + 1
                    End If
                End If

            End If

        Next i

    End If

    If n > 0 Then
        cx = cx / CDbl(n)
        cy = cy / CDbl(n)
        cz = cz / CDbl(n)
    End If

    ' Pull it onto the face proper.
    Dim vOn As Variant

    On Error Resume Next
    vOn = faceObj.GetClosestPointOn(cx, cy, cz)
    Err.Clear
    On Error GoTo ErrHandler

    If IsArray(vOn) Then
        If UBound(vOn) >= 2 Then
            cx = CDbl(vOn(0))
            cy = CDbl(vOn(1))
            cz = CDbl(vOn(2))
        End If
    End If

    If n = 0 Then
        If IsArray(vOn) = False Then Exit Function
    End If

    ' Part space -> assembly space, full transform this time: rotation, scale and
    ' translation, because this is a point and not a direction.
    outX = cx
    outY = cy
    outZ = cz

    If comp Is Nothing Then
        TryGetPointOnFaceInAssembly = True
        Exit Function
    End If

    Dim xform As Object
    Set xform = comp.Transform2

    If xform Is Nothing Then
        TryGetPointOnFaceInAssembly = True
        Exit Function
    End If

    Dim m As Variant
    m = xform.ArrayData

    If IsArray(m) = False Then Exit Function
    If UBound(m) < 12 Then Exit Function

    Dim s As Double
    s = CDbl(m(12))
    If Abs(s) < 0.0000001 Then s = 1#

    outX = s * ((cx * CDbl(m(0))) + (cy * CDbl(m(3))) + (cz * CDbl(m(6)))) + CDbl(m(9))
    outY = s * ((cx * CDbl(m(1))) + (cy * CDbl(m(4))) + (cz * CDbl(m(7)))) + CDbl(m(10))
    outZ = s * ((cx * CDbl(m(2))) + (cy * CDbl(m(5))) + (cz * CDbl(m(8)))) + CDbl(m(11))

    TryGetPointOnFaceInAssembly = True

    Exit Function

ErrHandler:
    TryGetPointOnFaceInAssembly = False
End Function

' Finds a pot block: matched quote first, then name keys.
Private Function FindPotBlockComponent(ByVal model As Object) As Object
On Error GoTo ErrHandler

    Set FindPotBlockComponent = Nothing

    If model Is Nothing Then Exit Function

    Dim cadIdx As Long

    cadIdx = FindCadIndexForOrientationQuoteOrKeys("ID POT BLOCK", POT_FRONT_ORIENTATION_KEYS)

    If cadIdx <= 0 Then
        cadIdx = FindCadIndexForOrientationQuoteOrKeys("OD POT BLOCK", POT_FRONT_ORIENTATION_KEYS)
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then

        Dim comp As Object
        Set comp = FindAssemblyComponentByName(model, parts(cadIdx).componentName)

        If Not comp Is Nothing Then
            Set FindPotBlockComponent = comp
            Exit Function
        End If

    End If

    Set FindPotBlockComponent = FindComponentByKeys(model, POT_FRONT_ORIENTATION_KEYS)

    Exit Function

ErrHandler:
    LogLine "FindPotBlockComponent error: " & Err.Description
    Set FindPotBlockComponent = Nothing
End Function

' Largest planar face on the pot block whose normal is perpendicular to the TCP
' normal - i.e. a side face, not the pot's own top or bottom. Where the
' holder-to-pot direction is known, faces pointing that way are preferred and the
' normal is flipped to point from the holders toward the pots, which is the
' direction the front view has to look from for pots to read as being in front.
Private Function TryGetPotSideFaceNormal(ByVal model As Object, _
                                         ByVal potComp As Object, _
                                         ByVal topX As Double, _
                                         ByVal topY As Double, _
                                         ByVal topZ As Double, _
                                         ByRef frontX As Double, _
                                         ByRef frontY As Double, _
                                         ByRef frontZ As Double, _
                                         ByRef faceOut As Object) As Boolean
On Error GoTo ErrHandler

    TryGetPotSideFaceNormal = False

    frontX = 0#
    frontY = 0#
    frontZ = 0#
    Set faceOut = Nothing

    If potComp Is Nothing Then Exit Function

    Dim partModel As Object
    Set partModel = potComp.GetModelDoc2

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function

    ' Preferred front direction: holders -> pots, flattened into the plane
    ' perpendicular to up.
    Dim wantX As Double
    Dim wantY As Double
    Dim wantZ As Double
    Dim haveWant As Boolean

    haveWant = TryGetHolderToPotDirection(wantX, wantY, wantZ)

    If haveWant Then

        Dim dotW As Double
        dotW = (wantX * topX) + (wantY * topY) + (wantZ * topZ)

        wantX = wantX - dotW * topX
        wantY = wantY - dotW * topY
        wantZ = wantZ - dotW * topZ

        haveWant = NormalizeStraightenVector(wantX, wantY, wantZ)

    End If

    If haveWant Then
        LogLine "TCP/pot squaring: holder-to-pot direction = " & Format(wantX, "0.0000") & ", " & _
                Format(wantY, "0.0000") & ", " & Format(wantZ, "0.0000")
    Else
        LogLine "TCP/pot squaring: holder-to-pot direction unavailable; using the largest side face."
    End If

    Dim bestScore As Double
    Dim bestNX As Double
    Dim bestNY As Double
    Dim bestNZ As Double
    Dim bestFace As Object
    Dim gotAny As Boolean

    bestScore = -1#
    gotAny = False
    Set bestFace = Nothing

    Dim minArea As Double
    minArea = TCP_NORMAL_MIN_FACE_AREA_SQIN / (INCHES_PER_METER * INCHES_PER_METER)

    Dim b As Long
    Dim f As Long
    Dim body As Object
    Dim vFaces As Variant
    Dim faceObj As Object
    Dim surf As Object
    Dim areaVal As Double
    Dim nx As Double
    Dim ny As Double
    Dim nz As Double
    Dim topDot As Double
    Dim score As Double

    For b = LBound(vBodies) To UBound(vBodies)

        Set body = vBodies(b)

        If Not body Is Nothing Then

            vFaces = body.GetFaces

            If IsArray(vFaces) Then

                For f = LBound(vFaces) To UBound(vFaces)

                    Set faceObj = vFaces(f)

                    If Not faceObj Is Nothing Then

                        Set surf = Nothing
                        On Error Resume Next
                        Set surf = faceObj.GetSurface
                        Err.Clear
                        On Error GoTo ErrHandler

                        If Not surf Is Nothing Then
                            If IsPlanarSurfaceSafe(surf) Then

                                If TryGetFaceNormalInAssembly(faceObj, potComp, nx, ny, nz) Then

                                    topDot = Abs((nx * topX) + (ny * topY) + (nz * topZ))

                                    ' Side faces only - skip anything facing up or down.
                                    If topDot < POT_FRONT_MAX_TOP_DOT Then

                                        areaVal = 0#
                                        On Error Resume Next
                                        areaVal = CDbl(faceObj.GetArea)
                                        Err.Clear
                                        On Error GoTo ErrHandler

                                        If areaVal > minArea Then

                                            score = areaVal

                                            If haveWant Then
                                                score = score * (1# + 5# * Abs((nx * wantX) + (ny * wantY) + (nz * wantZ)))
                                            End If

                                            If score > bestScore Then
                                                bestScore = score
                                                bestNX = nx
                                                bestNY = ny
                                                bestNZ = nz
                                                Set bestFace = faceObj
                                                gotAny = True
                                            End If

                                        End If

                                    End If

                                End If

                            End If
                        End If

                    End If

                Next f

            End If

        End If

    Next b

    If gotAny = False Then Exit Function

    Set faceOut = bestFace

    frontX = bestNX
    frontY = bestNY
    frontZ = bestNZ

    ' Point the normal from the holders toward the pots.
    If haveWant Then
        If ((frontX * wantX) + (frontY * wantY) + (frontZ * wantZ)) < 0# Then
            frontX = -frontX
            frontY = -frontY
            frontZ = -frontZ
            LogLine "TCP/pot squaring: pot face normal flipped to face the holder-to-pot direction."
        End If
    End If

    TryGetPotSideFaceNormal = NormalizeStraightenVector(frontX, frontY, frontZ)

    Exit Function

ErrHandler:
    LogLine "TryGetPotSideFaceNormal error: " & Err.Description
    TryGetPotSideFaceNormal = False
End Function

' Average holder centre -> average pot centre, using the same index collections
' the existing front logic builds.
Private Function TryGetHolderToPotDirection(ByRef dxOut As Double, _
                                            ByRef dyOut As Double, _
                                            ByRef dzOut As Double) As Boolean
On Error GoTo ErrHandler

    TryGetHolderToPotDirection = False

    dxOut = 0#
    dyOut = 0#
    dzOut = 0#

    Dim holderIndexes As Collection
    Dim potIndexes As Collection

    BuildFrontOrientIndexCollections holderIndexes, potIndexes

    If holderIndexes Is Nothing Then Exit Function
    If potIndexes Is Nothing Then Exit Function
    If holderIndexes.count = 0 Then Exit Function
    If potIndexes.count = 0 Then Exit Function

    Dim hx As Double, hy As Double, hz As Double
    Dim px As Double, py As Double, pz As Double

    If TryAverageCadCenterPoint(holderIndexes, True, hx, hy, hz) = False Then Exit Function
    If TryAverageCadCenterPoint(potIndexes, True, px, py, pz) = False Then Exit Function

    dxOut = px - hx
    dyOut = py - hy
    dzOut = pz - hz

    TryGetHolderToPotDirection = NormalizeStraightenVector(dxOut, dyOut, dzOut)

    Exit Function

ErrHandler:
    TryGetHolderToPotDirection = False
End Function

Private Function TryAverageCadCenterPoint(ByVal cadIndexes As Collection, _
                                          ByVal preferMassCenter As Boolean, _
                                          ByRef avgX As Double, _
                                          ByRef avgY As Double, _
                                          ByRef avgZ As Double) As Boolean
On Error GoTo ErrHandler

    TryAverageCadCenterPoint = False

    avgX = 0#
    avgY = 0#
    avgZ = 0#

    If cadIndexes Is Nothing Then Exit Function
    If cadIndexes.count = 0 Then Exit Function

    Dim totalX As Double
    Dim totalY As Double
    Dim totalZ As Double
    Dim countVal As Long

    totalX = 0#
    totalY = 0#
    totalZ = 0#
    countVal = 0

    Dim i As Long
    Dim cadIdx As Long
    Dim px As Double
    Dim py As Double
    Dim pz As Double

    For i = 1 To cadIndexes.count

        cadIdx = CLng(cadIndexes(i))

        If TryGetCadCenterPointInches(cadIdx, preferMassCenter, px, py, pz) Then
            totalX = totalX + px
            totalY = totalY + py
            totalZ = totalZ + pz
            countVal = countVal + 1
        End If

    Next i

    If countVal <= 0 Then Exit Function

    avgX = totalX / CDbl(countVal)
    avgY = totalY / CDbl(countVal)
    avgZ = totalZ / CDbl(countVal)

    TryAverageCadCenterPoint = True

    Exit Function

ErrHandler:
    TryAverageCadCenterPoint = False
End Function

' After the frame is set, make sure it is the TCP side facing up and not the BCP
' side. Flips to *Bottom and saves that as *Top if it came out upside down.
Private Function EnsureNormalTopLooksAtTcpSide(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    EnsureNormalTopLooksAtTcpSide = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim tcpIdx As Long
    Dim bcpIdx As Long

    tcpIdx = FindCadIndexForOrientationQuoteOrKeys("TCP", TCP_TOP_ORIENTATION_KEYS)
    bcpIdx = FindCadIndexForOrientationQuoteOrKeys("BCP", BCP_BOTTOM_ORIENTATION_KEYS)

    If tcpIdx <= 0 Or bcpIdx <= 0 Then
        LogLine "TCP-side check skipped: TCP/BCP CAD indexes unavailable."
        Exit Function
    End If

    Dim tx As Double, ty As Double, tz As Double
    Dim bx As Double, by As Double, bz As Double

    If TryGetCadCenterPointInches(tcpIdx, True, tx, ty, tz) = False Then Exit Function
    If TryGetCadCenterPointInches(bcpIdx, True, bx, by, bz) = False Then Exit Function

    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 150

    Dim tcpDepth As Double
    Dim bcpDepth As Double

    If TryProjectPointToActiveViewDepth(model, tx, ty, tz, tcpDepth) = False Then Exit Function
    If TryProjectPointToActiveViewDepth(model, bx, by, bz, bcpDepth) = False Then Exit Function

    LogLine "TCP-side check: TCP depth=" & Format(tcpDepth, "0.000") & _
            "  BCP depth=" & Format(bcpDepth, "0.000")

    ' Larger active-view depth is closer to the viewer.
    If tcpDepth < bcpDepth Then

        LogLine "TCP-side check: BCP is on top. Flipping to *Bottom and saving that as *Top."

        model.ShowNamedView2 "*Bottom", 6
        StabilizeActiveView model, 150

        If PersistCurrentViewAsStandardTop(model) Then
            LogLine "TCP-side check: flipped side persisted as *Top."
        Else
            LogLine "WARNING: TCP-side flip could not be persisted as *Top."
        End If

        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 150

    Else
        LogLine "TCP-side check OK: the TCP side is facing up."
    End If

    EnsureNormalTopLooksAtTcpSide = True

    Exit Function

ErrHandler:
    LogLine "EnsureNormalTopLooksAtTcpSide error: " & Err.Description
    EnsureNormalTopLooksAtTcpSide = False
End Function

Private Function TryOrientTopNormalToTopClampingPlate(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    TryOrientTopNormalToTopClampingPlate = False

    If model Is Nothing Then Exit Function

    If model.GetType <> swDocASSEMBLY Then
        LogLine "TCP Normal To skipped: document is not an assembly."
        Exit Function
    End If

    Dim tcpComp As Object
    Set tcpComp = FindTopClampingPlateComponent(model)

    If tcpComp Is Nothing Then
        LogLine "TCP Normal To skipped: top clamping plate component not found. " & _
                "Standard orientation will run on the view as found."
        Exit Function
    End If

    LogLine "TCP Normal To: using component " & BaseStraightenComponentName(tcpComp)

    Dim faceObj As Object
    Dim faceArea As Double

    Set faceObj = FindLargestOuterPlanarFaceOnComponent(model, tcpComp, faceArea)

    If faceObj Is Nothing Then
        LogLine "TCP Normal To skipped: no planar face at least " & _
                TCP_NORMAL_MIN_FACE_AREA_SQIN & " sq in found on the plate."
        Exit Function
    End If

    LogLine "TCP Normal To: selected planar face, area " & Format(faceArea, "0.000") & " sq in."

    ' ---- view depth axis = face normal, in assembly space ----
    Dim dxv As Double
    Dim dyv As Double
    Dim dzv As Double

    If TryGetFaceNormalInAssembly(faceObj, tcpComp, dxv, dyv, dzv) = False Then
        LogLine "TCP Normal To failed: could not read the face normal."
        Exit Function
    End If

    If OrientOutwardAwayFromBcp(model, tcpComp, dxv, dyv, dzv) Then
        LogLine "TCP Normal To: normal flipped to point away from the BCP."
    End If

    LogLine "TCP Normal To: face normal in assembly space = " & _
            Format(dxv, "0.0000") & ", " & Format(dyv, "0.0000") & ", " & Format(dzv, "0.0000")

    ' ---- view X axis = plate's own longest straight edge, squares the spin ----
    Dim xxv As Double
    Dim xyv As Double
    Dim xzv As Double
    Dim gotEdge As Boolean

    gotEdge = False

    If TCP_NORMAL_ALIGN_TO_LONGEST_EDGE Then
        gotEdge = TryGetFaceLongestLineDirInAssembly(faceObj, tcpComp, xxv, xyv, xzv)
        If gotEdge Then
            LogLine "TCP Normal To: aligned to the plate's longest straight edge."
        Else
            LogLine "TCP Normal To: no straight edge found on that face; " & _
                    "tilt corrected but in-plane spin left as found."
        End If
    End If

    If gotEdge = False Then
        If PickAnyPerpendicularAxis(dxv, dyv, dzv, xxv, xyv, xzv) = False Then
            LogLine "TCP Normal To failed: could not build a view axis."
            Exit Function
        End If
    End If

    ' Force the chosen X axis into the face plane, then re-normalise.
    Dim dotVal As Double
    dotVal = (xxv * dxv) + (xyv * dyv) + (xzv * dzv)

    xxv = xxv - dotVal * dxv
    xyv = xyv - dotVal * dyv
    xzv = xzv - dotVal * dzv

    If NormalizeStraightenVector(xxv, xyv, xzv) = False Then
        If PickAnyPerpendicularAxis(dxv, dyv, dzv, xxv, xyv, xzv) = False Then
            LogLine "TCP Normal To failed: view axis collapsed onto the face normal."
            Exit Function
        End If
    End If

    ' ---- view Y = depth x X, so X, Y, depth are right-handed ----
    Dim yxv As Double
    Dim yyv As Double
    Dim yzv As Double

    yxv = dyv * xzv - dzv * xyv
    yyv = dzv * xxv - dxv * xzv
    yzv = dxv * xyv - dyv * xxv

    If NormalizeStraightenVector(yxv, yyv, yzv) = False Then
        LogLine "TCP Normal To failed: could not build the view up axis."
        Exit Function
    End If

    ' The main viewport has graphics updates switched OFF for speed the moment the
    ' CAD is opened, and a view with graphics disabled silently ignores an
    ' Orientation3 write - the write returns no error and the view simply does not
    ' move. That is exactly what happened before: the log read
    ' "face is 36.000 deg off the view axis", 36 deg being the plate's own tilt,
    ' i.e. the view had never moved at all. Turn graphics back on for the write,
    ' then put them back the way they were.
    Dim graphicsWereOff As Boolean
    graphicsWereOff = MainViewportGraphicsDisabled

    If graphicsWereOff Then
        RestoreMainViewportGraphics
        LogLine "TCP Normal To: viewport graphics re-enabled so the view write can take effect."
    End If

    Dim actErrs As Long
    On Error Resume Next
    swApp.ActivateDoc3 model.GetTitle, False, 0, actErrs
    Err.Clear
    On Error GoTo ErrHandler

    ' Orientation3's row/column convention is not worth guessing at. Write it,
    ' MEASURE the result, and if the face did not come square try the transpose.
    ' For a rotation matrix those are the only two possibilities and the
    ' measurement says which one this build of SolidWorks wants.
    Dim okSquare As Boolean
    Dim attempt As Long
    Dim offDeg As Double

    okSquare = False
    offDeg = 0#

    For attempt = 1 To 2

        If TrySetActiveViewOrientation(model, xxv, xyv, xzv, yxv, yyv, yzv, _
                                       dxv, dyv, dzv, (attempt = 2)) Then

            StabilizeActiveView model, 150

            If TryGetAngleBetweenFaceNormalAndViewDeg(model, dxv, dyv, dzv, offDeg) Then

                LogLine "TCP Normal To attempt " & attempt & _
                        IIf(attempt = 2, " (transposed)", "") & ": face is " & _
                        Format(offDeg, "0.000") & " deg off the view axis."

                If offDeg <= TCP_NORMAL_MAX_OFF_DEG Then
                    okSquare = True
                    Exit For
                End If

            Else
                LogLine "TCP Normal To attempt " & attempt & ": could not measure squareness."
            End If

        Else
            LogLine "TCP Normal To attempt " & attempt & ": the view orientation would not set."
        End If

    Next attempt

    If okSquare = False Then
        LogLine "TCP Normal To FAILED verification: face is still " & Format(offDeg, "0.000") & _
                " deg off after both conventions. *Top left as it was."
        GoTo CleanExit
    End If

    LogLine "TCP Normal To: view is now square to the plate face."

    ' Save the view we just took, then make it *Top.
    On Error Resume Next
    model.DeleteNamedView TCP_NORMAL_VIEW_NAME
    Err.Clear
    model.NameView TCP_NORMAL_VIEW_NAME
    Err.Clear
    On Error GoTo ErrHandler

    LogLine "TCP Normal To: named view " & TCP_NORMAL_VIEW_NAME & " saved from the Normal To orientation."

    If PersistCurrentViewAsStandardTop(model) Then
        LogLine "TCP Normal To: *Top is now the Normal To view of the clamping plate. No geometry moved."
        TryOrientTopNormalToTopClampingPlate = True
    Else
        LogLine "TCP Normal To: WARNING - *Top could not be redefined; " & _
                "orientation is still carried by " & TCP_NORMAL_VIEW_NAME & "."
    End If

CleanExit:

    If graphicsWereOff Then
        DisableMainViewportGraphics
    End If

    Exit Function

ErrHandler:
    LogLine "TryOrientTopNormalToTopClampingPlate error: " & Err.Description
    TryOrientTopNormalToTopClampingPlate = False

    On Error Resume Next
    If graphicsWereOff Then DisableMainViewportGraphics
End Function

' The top clamping plate: matched TCP quote first, then the TCP name keys, then
' the raw component-name search the world-axis path already uses.

' ============================================================
' PATCH HELPERS: TCP = LIGHTER CLAMPING PLATE + REAL CTRL+8 NORMAL TO
' ============================================================

Private Function FindLighterTcpBcpClampingPlateComponent(ByVal model As Object) As Object
On Error GoTo ErrHandler

    Set FindLighterTcpBcpClampingPlateComponent = Nothing

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim tcpIdx As Long
    Dim bcpIdx As Long

    tcpIdx = FindCadIndexForOrientationQuoteOrKeys("TCP", TCP_TOP_ORIENTATION_KEYS)
    bcpIdx = FindCadIndexForOrientationQuoteOrKeys("BCP", BCP_BOTTOM_ORIENTATION_KEYS)

    If tcpIdx <= 0 Or bcpIdx <= 0 Then
        LogLine "Lighter TCP pick skipped: TCP/BCP matched CAD indexes unavailable."
        Exit Function
    End If

    If tcpIdx > PartCount Or bcpIdx > PartCount Then Exit Function

    Dim lightIdx As Long

    If parts(tcpIdx).massValue <= parts(bcpIdx).massValue Then
        lightIdx = tcpIdx
    Else
        lightIdx = bcpIdx
    End If

    LogLine "Lighter TCP pick:"
    LogLine "  candidate A=" & parts(tcpIdx).componentName & " mass=" & FormatNumberForCsv(parts(tcpIdx).massValue)
    LogLine "  candidate B=" & parts(bcpIdx).componentName & " mass=" & FormatNumberForCsv(parts(bcpIdx).massValue)
    LogLine "  selected =" & parts(lightIdx).componentName

    Dim comp As Object
    Set comp = FindAssemblyComponentByName(model, parts(lightIdx).componentName)

    If Not comp Is Nothing Then
        Set FindLighterTcpBcpClampingPlateComponent = comp
    End If

    Exit Function

ErrHandler:
    LogLine "FindLighterTcpBcpClampingPlateComponent error: " & Err.Description
    Set FindLighterTcpBcpClampingPlateComponent = Nothing
End Function

Private Function FindBestTcpOuterPlanarFaceOnComponent(ByVal model As Object, _
                                                       ByVal tcpComp As Object, _
                                                       ByRef areaSqInOut As Double) As Object
On Error GoTo ErrHandler

    Set FindBestTcpOuterPlanarFaceOnComponent = Nothing
    areaSqInOut = 0#

    If model Is Nothing Then Exit Function
    If tcpComp Is Nothing Then Exit Function

    Dim bcpComp As Object
    Set bcpComp = FindComponentByKeys(model, BCP_BOTTOM_ORIENTATION_KEYS)

    If bcpComp Is Nothing Then
        LogLine "TCP outer face pick: BCP unavailable; falling back to largest planar face."
        Set FindBestTcpOuterPlanarFaceOnComponent = FindLargestOuterPlanarFaceOnComponent(model, tcpComp, areaSqInOut)
        Exit Function
    End If

    Dim tx As Double, ty As Double, tz As Double
    Dim bx As Double, by As Double, bz As Double

    If TryGetComponentCenterInches(tcpComp, tx, ty, tz) = False Then
        LogLine "TCP outer face pick: TCP center unavailable; falling back to largest planar face."
        Set FindBestTcpOuterPlanarFaceOnComponent = FindLargestOuterPlanarFaceOnComponent(model, tcpComp, areaSqInOut)
        Exit Function
    End If

    If TryGetComponentCenterInches(bcpComp, bx, by, bz) = False Then
        LogLine "TCP outer face pick: BCP center unavailable; falling back to largest planar face."
        Set FindBestTcpOuterPlanarFaceOnComponent = FindLargestOuterPlanarFaceOnComponent(model, tcpComp, areaSqInOut)
        Exit Function
    End If

    ' Direction from BCP toward TCP. The TCP outer face should point this way.
    Dim stackX As Double
    Dim stackY As Double
    Dim stackZ As Double

    stackX = tx - bx
    stackY = ty - by
    stackZ = tz - bz

    If NormalizeStraightenVector(stackX, stackY, stackZ) = False Then
        LogLine "TCP outer face pick: stack direction invalid; falling back to largest planar face."
        Set FindBestTcpOuterPlanarFaceOnComponent = FindLargestOuterPlanarFaceOnComponent(model, tcpComp, areaSqInOut)
        Exit Function
    End If

    Dim partModel As Object
    Set partModel = tcpComp.GetModelDoc2

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function

    Dim minAreaM2 As Double
    minAreaM2 = TCP_NORMAL_MIN_FACE_AREA_SQIN / (INCHES_PER_METER * INCHES_PER_METER)

    Dim bestFace As Object
    Dim bestScore As Double
    Dim bestAreaM2 As Double

    Set bestFace = Nothing
    bestScore = -1#
    bestAreaM2 = 0#

    Dim b As Long
    Dim f As Long
    Dim body As Object
    Dim vFaces As Variant
    Dim faceObj As Object
    Dim surf As Object

    Dim nx As Double
    Dim ny As Double
    Dim nz As Double
    Dim areaM2 As Double
    Dim normalDot As Double
    Dim score As Double

    For b = LBound(vBodies) To UBound(vBodies)

        Set body = vBodies(b)

        If Not body Is Nothing Then

            vFaces = body.GetFaces

            If IsArray(vFaces) Then

                For f = LBound(vFaces) To UBound(vFaces)

                    Set faceObj = vFaces(f)

                    If Not faceObj Is Nothing Then

                        Set surf = Nothing

                        On Error Resume Next
                        Set surf = faceObj.GetSurface
                        Err.Clear
                        On Error GoTo ErrHandler

                        If Not surf Is Nothing Then
                            If IsPlanarSurfaceSafe(surf) Then

                                If TryGetFaceNormalInAssembly(faceObj, tcpComp, nx, ny, nz) Then

                                    normalDot = (nx * stackX) + (ny * stackY) + (nz * stackZ)

                                    ' Require the face normal to point away from BCP toward TCP.
                                    ' This chooses the outside/top TCP face, not the buried inner face.
                                    If normalDot > 0.35 Then

                                        areaM2 = 0#

                                        On Error Resume Next
                                        areaM2 = CDbl(faceObj.GetArea)
                                        Err.Clear
                                        On Error GoTo ErrHandler

                                        If areaM2 >= minAreaM2 Then

                                            ' Mostly area, but strongly prefer the face pointing outward.
                                            score = areaM2 * (1# + 10# * normalDot)

                                            If score > bestScore Then
                                                bestScore = score
                                                bestAreaM2 = areaM2
                                                Set bestFace = faceObj
                                            End If

                                        End If

                                    End If

                                End If

                            End If
                        End If

                    End If

                Next f

            End If

        End If

    Next b

    If bestFace Is Nothing Then
        LogLine "TCP outer face pick: no outward planar face found; falling back to largest planar face."
        Set FindBestTcpOuterPlanarFaceOnComponent = FindLargestOuterPlanarFaceOnComponent(model, tcpComp, areaSqInOut)
        Exit Function
    End If

    areaSqInOut = bestAreaM2 * INCHES_PER_METER * INCHES_PER_METER

    LogLine "TCP outer face pick: selected outward face. Area=" & _
            Format(areaSqInOut, "0.000") & " sq in."

    Set FindBestTcpOuterPlanarFaceOnComponent = bestFace

    Exit Function

ErrHandler:
    LogLine "FindBestTcpOuterPlanarFaceOnComponent error: " & Err.Description
    Set FindBestTcpOuterPlanarFaceOnComponent = Nothing
End Function

Private Function PressNormalToWithCtrl8(ByVal model As Object, _
                                        ByVal nx As Double, _
                                        ByVal ny As Double, _
                                        ByVal nz As Double, _
                                        ByVal label As String) As Boolean
On Error GoTo ErrHandler

    PressNormalToWithCtrl8 = False

    If model Is Nothing Then Exit Function

    Dim errs As Long

    On Error Resume Next
    swApp.Visible = True
    swApp.UserControl = True
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    Err.Clear
    On Error GoTo ErrHandler

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    ' Bring SolidWorks / active document forward enough for SendKeys to land.
    On Error Resume Next
    sh.AppActivate model.GetTitle
    WaitMilliseconds 150
    Err.Clear
    On Error GoTo ErrHandler

    LogLine "Normal To (" & label & "): sending Ctrl+8."

    sh.SendKeys "^8"
    WaitMilliseconds 350

    model.GraphicsRedraw2
    DoEvents
    WaitMilliseconds 150

    Dim offDeg As Double

    If TryGetAngleBetweenFaceNormalAndViewDeg(model, nx, ny, nz, offDeg) Then

        LogLine "Normal To (" & label & "): Ctrl+8 left the face " & _
                Format(offDeg, "0.000") & " deg off."

        If offDeg <= TCP_NORMAL_MAX_OFF_DEG Then
            PressNormalToWithCtrl8 = True
            Exit Function
        End If

    Else
        LogLine "Normal To (" & label & "): Ctrl+8 sent, but squareness could not be measured."
    End If

    Exit Function

ErrHandler:
    LogLine "PressNormalToWithCtrl8 error: " & Err.Description
    PressNormalToWithCtrl8 = False
End Function

' ============================================================
' END PATCH HELPERS
' ============================================================

Private Function FindTopClampingPlateComponent(ByVal model As Object) As Object
On Error GoTo ErrHandler

    Set FindTopClampingPlateComponent = Nothing


    If FORCE_TCP_TO_LIGHTER_CLAMPING_PLATE Then
        Dim lightComp As Object
        Set lightComp = FindLighterTcpBcpClampingPlateComponent(model)

        If Not lightComp Is Nothing Then
            LogLine "TCP component forced to LIGHTER clamping plate: " & BaseStraightenComponentName(lightComp)
            Set FindTopClampingPlateComponent = lightComp
            Exit Function
        End If
    End If

    If model Is Nothing Then Exit Function

    Dim cadIdx As Long
    cadIdx = FindCadIndexForOrientationQuoteOrKeys("TCP", TCP_TOP_ORIENTATION_KEYS)

    If cadIdx > 0 And cadIdx <= PartCount Then

        Dim comp As Object
        Set comp = FindAssemblyComponentByName(model, parts(cadIdx).componentName)

        If Not comp Is Nothing Then
            Set FindTopClampingPlateComponent = comp
            Exit Function
        End If

        LogLine "TCP Normal To: matched CAD index " & cadIdx & " (" & parts(cadIdx).componentName & _
                ") but the component could not be resolved. Falling back to a name search."
    End If

    Set FindTopClampingPlateComponent = FindComponentByKeys(model, TCP_TOP_ORIENTATION_KEYS)

    Exit Function

ErrHandler:
    LogLine "FindTopClampingPlateComponent error: " & Err.Description
    Set FindTopClampingPlateComponent = Nothing
End Function

' Largest planar face on the component, in square inches.
Private Function FindLargestOuterPlanarFaceOnComponent(ByVal model As Object, _
                                                       ByVal comp As Object, _
                                                       ByRef areaSqInOut As Double) As Object
On Error GoTo ErrHandler

    Set FindLargestOuterPlanarFaceOnComponent = Nothing
    areaSqInOut = 0#

    If comp Is Nothing Then Exit Function

    Dim partModel As Object
    Set partModel = comp.GetModelDoc2

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function

    Dim bestFace As Object
    Dim bestArea As Double

    Set bestFace = Nothing
    bestArea = 0#

    Dim minArea As Double
    minArea = TCP_NORMAL_MIN_FACE_AREA_SQIN / (INCHES_PER_METER * INCHES_PER_METER)

    Dim b As Long
    Dim f As Long
    Dim body As Object
    Dim vFaces As Variant
    Dim faceObj As Object
    Dim surf As Object
    Dim aVal As Double

    For b = LBound(vBodies) To UBound(vBodies)

        Set body = vBodies(b)

        If Not body Is Nothing Then

            vFaces = body.GetFaces

            If IsArray(vFaces) Then

                For f = LBound(vFaces) To UBound(vFaces)

                    Set faceObj = vFaces(f)

                    If Not faceObj Is Nothing Then

                        Set surf = Nothing
                        On Error Resume Next
                        Set surf = faceObj.GetSurface
                        Err.Clear
                        On Error GoTo ErrHandler

                        If Not surf Is Nothing Then
                            If IsPlanarSurfaceSafe(surf) Then

                                aVal = 0#
                                On Error Resume Next
                                aVal = CDbl(faceObj.GetArea)
                                Err.Clear
                                On Error GoTo ErrHandler

                                If aVal > minArea And aVal > bestArea Then
                                    bestArea = aVal
                                    Set bestFace = faceObj
                                End If

                            End If
                        End If

                    End If

                Next f

            End If

        End If

    Next b

    If bestFace Is Nothing Then Exit Function

    areaSqInOut = bestArea * INCHES_PER_METER * INCHES_PER_METER
    Set FindLargestOuterPlanarFaceOnComponent = bestFace

    Exit Function

ErrHandler:
    LogLine "FindLargestOuterPlanarFaceOnComponent error: " & Err.Description
    Set FindLargestOuterPlanarFaceOnComponent = Nothing
End Function

Private Function IsPlanarSurfaceSafe(ByVal surf As Object) As Boolean
On Error Resume Next

    IsPlanarSurfaceSafe = False

    If surf Is Nothing Then Exit Function

    Err.Clear
    IsPlanarSurfaceSafe = CBool(surf.IsPlane)

    If Err.Number <> 0 Then
        Err.Clear
        IsPlanarSurfaceSafe = False
    End If
End Function

' Face normal, taken in part space then rotated into assembly space by the
' component transform. Same ArrayData convention the rest of the module uses.
Private Function TryGetFaceNormalInAssembly(ByVal faceObj As Object, _
                                            ByVal comp As Object, _
                                            ByRef nxOut As Double, _
                                            ByRef nyOut As Double, _
                                            ByRef nzOut As Double) As Boolean
On Error GoTo ErrHandler

    TryGetFaceNormalInAssembly = False

    nxOut = 0#
    nyOut = 0#
    nzOut = 0#

    If faceObj Is Nothing Then Exit Function

    Dim vN As Variant

    On Error Resume Next
    vN = faceObj.Normal
    Err.Clear
    On Error GoTo ErrHandler

    If IsArray(vN) = False Then Exit Function
    If UBound(vN) < 2 Then Exit Function

    nxOut = CDbl(vN(0))
    nyOut = CDbl(vN(1))
    nzOut = CDbl(vN(2))

    If RotateDirectionByComponent(comp, nxOut, nyOut, nzOut) = False Then Exit Function

    TryGetFaceNormalInAssembly = NormalizeStraightenVector(nxOut, nyOut, nzOut)

    Exit Function

ErrHandler:
    TryGetFaceNormalInAssembly = False
End Function

' Direction of the longest straight edge of the face, in assembly space.
Private Function TryGetFaceLongestLineDirInAssembly(ByVal faceObj As Object, _
                                                    ByVal comp As Object, _
                                                    ByRef dxOut As Double, _
                                                    ByRef dyOut As Double, _
                                                    ByRef dzOut As Double) As Boolean
On Error GoTo ErrHandler

    TryGetFaceLongestLineDirInAssembly = False

    dxOut = 0#
    dyOut = 0#
    dzOut = 0#

    If faceObj Is Nothing Then Exit Function

    Dim vEdges As Variant

    On Error Resume Next
    vEdges = faceObj.GetEdges
    Err.Clear
    On Error GoTo ErrHandler

    If IsArray(vEdges) = False Then Exit Function

    Dim bestLen As Double
    Dim bx As Double
    Dim by As Double
    Dim bz As Double

    bestLen = 0#

    Dim e As Long
    Dim edgeObj As Object
    Dim v1 As Object
    Dim v2 As Object
    Dim p1 As Variant
    Dim p2 As Variant
    Dim ex As Double
    Dim ey As Double
    Dim ez As Double
    Dim eLen As Double

    For e = LBound(vEdges) To UBound(vEdges)

        Set edgeObj = vEdges(e)

        If Not edgeObj Is Nothing Then

            If IsLinearEdgeSafe(edgeObj) Then

                Set v1 = Nothing
                Set v2 = Nothing

                On Error Resume Next
                Set v1 = edgeObj.GetStartVertex
                Set v2 = edgeObj.GetEndVertex
                Err.Clear
                On Error GoTo ErrHandler

                If (Not v1 Is Nothing) And (Not v2 Is Nothing) Then

                    On Error Resume Next
                    p1 = v1.GetPoint
                    p2 = v2.GetPoint
                    Err.Clear
                    On Error GoTo ErrHandler

                    If IsArray(p1) And IsArray(p2) Then
                        If UBound(p1) >= 2 And UBound(p2) >= 2 Then

                            ex = CDbl(p2(0)) - CDbl(p1(0))
                            ey = CDbl(p2(1)) - CDbl(p1(1))
                            ez = CDbl(p2(2)) - CDbl(p1(2))

                            eLen = Sqr(ex * ex + ey * ey + ez * ez)

                            If eLen > bestLen Then
                                bestLen = eLen
                                bx = ex
                                by = ey
                                bz = ez
                            End If

                        End If
                    End If

                End If

            End If

        End If

    Next e

    If bestLen <= 0# Then Exit Function

    dxOut = bx
    dyOut = by
    dzOut = bz

    If RotateDirectionByComponent(comp, dxOut, dyOut, dzOut) = False Then Exit Function

    TryGetFaceLongestLineDirInAssembly = NormalizeStraightenVector(dxOut, dyOut, dzOut)

    Exit Function

ErrHandler:
    TryGetFaceLongestLineDirInAssembly = False
End Function

Private Function IsLinearEdgeSafe(ByVal edgeObj As Object) As Boolean
On Error Resume Next

    IsLinearEdgeSafe = False

    If edgeObj Is Nothing Then Exit Function

    Dim crv As Object
    Set crv = Nothing

    Err.Clear
    Set crv = edgeObj.GetCurve

    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    If crv Is Nothing Then Exit Function

    Err.Clear
    IsLinearEdgeSafe = CBool(crv.IsLine)

    If Err.Number <> 0 Then
        Err.Clear
        IsLinearEdgeSafe = False
    End If
End Function

' Rotates a direction from part space into assembly space using the component
' transform. Translation and scale are irrelevant for a direction.
Private Function RotateDirectionByComponent(ByVal comp As Object, _
                                            ByRef vx As Double, _
                                            ByRef vy As Double, _
                                            ByRef vz As Double) As Boolean
On Error GoTo ErrHandler

    RotateDirectionByComponent = False

    If comp Is Nothing Then
        RotateDirectionByComponent = True
        Exit Function
    End If

    Dim xform As Object
    Set xform = comp.Transform2

    If xform Is Nothing Then
        RotateDirectionByComponent = True
        Exit Function
    End If

    Dim m As Variant
    m = xform.ArrayData

    If IsArray(m) = False Then Exit Function
    If UBound(m) < 8 Then Exit Function

    Dim ox As Double
    Dim oy As Double
    Dim oz As Double

    ox = vx
    oy = vy
    oz = vz

    vx = (ox * CDbl(m(0))) + (oy * CDbl(m(3))) + (oz * CDbl(m(6)))
    vy = (ox * CDbl(m(1))) + (oy * CDbl(m(4))) + (oz * CDbl(m(7)))
    vz = (ox * CDbl(m(2))) + (oy * CDbl(m(5))) + (oz * CDbl(m(8)))

    RotateDirectionByComponent = True

    Exit Function

ErrHandler:
    RotateDirectionByComponent = False
End Function

' Makes the normal point away from the BCP, so the face we look at is the
' plate's outer face rather than the one buried against the stack.
' Returns True when the normal had to be flipped.
Private Function OrientOutwardAwayFromBcp(ByVal model As Object, _
                                          ByVal tcpComp As Object, _
                                          ByRef nx As Double, _
                                          ByRef ny As Double, _
                                          ByRef nz As Double) As Boolean
On Error GoTo ErrHandler

    OrientOutwardAwayFromBcp = False

    Dim bcpComp As Object
    Set bcpComp = FindComponentByKeys(model, BCP_BOTTOM_ORIENTATION_KEYS)

    If bcpComp Is Nothing Then Exit Function
    If tcpComp Is Nothing Then Exit Function

    Dim tx As Double, ty As Double, tz As Double
    Dim bx As Double, by As Double, bz As Double

    If TryGetComponentCenterInches(tcpComp, tx, ty, tz) = False Then Exit Function
    If TryGetComponentCenterInches(bcpComp, bx, by, bz) = False Then Exit Function

    Dim dotVal As Double
    dotVal = ((tx - bx) * nx) + ((ty - by) * ny) + ((tz - bz) * nz)

    If dotVal < 0# Then
        nx = -nx
        ny = -ny
        nz = -nz
        OrientOutwardAwayFromBcp = True
    End If

    Exit Function

ErrHandler:
    OrientOutwardAwayFromBcp = False
End Function

' Writes the active view orientation from three unit axes, using the same
' Orientation3 convention TryProjectPointToActiveViewDepth reads:
'   view X     = p dot [m0, m3, m6]
'   view Y     = p dot [m1, m4, m7]
'   view depth = p dot [m2, m5, m8]
' useTranspose swaps rows for columns, which is the other of the only two
' possible conventions for a rotation matrix. The caller writes, measures, and
' retries with the other one rather than assuming either is correct.
Private Function TrySetActiveViewOrientation(ByVal model As Object, _
                                             ByVal xx As Double, ByVal xy As Double, ByVal xz As Double, _
                                             ByVal yx As Double, ByVal yy As Double, ByVal yz As Double, _
                                             ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, _
                                             ByVal useTranspose As Boolean) As Boolean
On Error GoTo ErrHandler

    TrySetActiveViewOrientation = False

    If model Is Nothing Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then
        LogLine "TrySetActiveViewOrientation: no active view."
        Exit Function
    End If

    Dim mathUtil As Object
    Set mathUtil = swApp.GetMathUtility

    If mathUtil Is Nothing Then
        LogLine "TrySetActiveViewOrientation: MathUtility unavailable."
        Exit Function
    End If

    Dim arr(0 To 15) As Double

    If useTranspose Then
        arr(0) = xx: arr(1) = xy: arr(2) = xz
        arr(3) = yx: arr(4) = yy: arr(5) = yz
        arr(6) = dX: arr(7) = dY: arr(8) = dZ
    Else
        arr(0) = xx: arr(1) = yx: arr(2) = dX
        arr(3) = xy: arr(4) = yy: arr(5) = dY
        arr(6) = xz: arr(7) = yz: arr(8) = dZ
    End If

    arr(9) = 0#
    arr(10) = 0#
    arr(11) = 0#
    arr(12) = 1#
    arr(13) = 0#
    arr(14) = 0#
    arr(15) = 0#

    Dim newXform As Object
    Set newXform = mathUtil.CreateTransform(arr)

    If newXform Is Nothing Then
        LogLine "TrySetActiveViewOrientation: CreateTransform returned nothing."
        Exit Function
    End If

    ' Read the orientation first so the write can be checked against it. A view
    ' with graphics updates disabled accepts the write without error and does not
    ' move, so "no error" is not evidence that anything happened.
    Dim before As Variant
    Dim gotBefore As Boolean

    gotBefore = False

    On Error Resume Next
    before = swView.Orientation3.ArrayData
    If Err.Number = 0 Then gotBefore = IsArray(before)
    Err.Clear
    On Error GoTo ErrHandler

    On Error Resume Next
    Err.Clear
    Set swView.Orientation3 = newXform
    If Err.Number <> 0 Then
        Err.Clear
        swView.Orientation3 = newXform
    End If
    Err.Clear
    model.GraphicsRedraw2
    Err.Clear
    On Error GoTo ErrHandler

    ' Confirm the matrix actually changed, unless it was already what we wanted.
    Dim after As Variant

    On Error Resume Next
    after = swView.Orientation3.ArrayData
    Err.Clear
    On Error GoTo ErrHandler

    If IsArray(after) = False Then
        LogLine "TrySetActiveViewOrientation: could not read the view orientation back."
        Exit Function
    End If

    If UBound(after) < 8 Then
        LogLine "TrySetActiveViewOrientation: view orientation array is too short to check."
        Exit Function
    End If

    Dim wantedAlready As Boolean
    Dim changed As Boolean
    Dim k As Long

    wantedAlready = True
    changed = False

    For k = 0 To 8
        If Abs(CDbl(after(k)) - arr(k)) > 0.000001 Then wantedAlready = False
        If gotBefore Then
            If UBound(before) >= 8 Then
                If Abs(CDbl(after(k)) - CDbl(before(k))) > 0.000001 Then changed = True
            End If
        End If
    Next k

    If changed = False And wantedAlready = False Then
        LogLine "TrySetActiveViewOrientation: the write was ignored - the view matrix did not change."
        Exit Function
    End If

    TrySetActiveViewOrientation = True

    Exit Function

ErrHandler:
    LogLine "TrySetActiveViewOrientation error: " & Err.Description
    TrySetActiveViewOrientation = False
End Function

' Angle between the face normal and the view depth axis, in degrees. 0 means the
' face is dead square to the screen.
'
' The depth axis is read as (m2, m5, m8). That is not a guess and it does not
' change when the write is transposed: it is the same convention
' TryProjectPointToActiveViewDepth uses, which is already proven correct by the
' pot-vs-holder depth checks that work off it. So this always reports the angle
' to the view axis SolidWorks is really using, whichever way the matrix went in -
' which is what makes it safe to write, measure, and retry.
Private Function TryGetAngleBetweenFaceNormalAndViewDeg(ByVal model As Object, _
                                                        ByVal nx As Double, _
                                                        ByVal ny As Double, _
                                                        ByVal nz As Double, _
                                                        ByRef degOut As Double) As Boolean
On Error GoTo ErrHandler

    TryGetAngleBetweenFaceNormalAndViewDeg = False
    degOut = 0#

    If model Is Nothing Then Exit Function

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    Dim mView As Variant
    mView = swView.Orientation3.ArrayData

    If IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    Dim vdx As Double
    Dim vdy As Double
    Dim vdz As Double

    vdx = CDbl(mView(2))
    vdy = CDbl(mView(5))
    vdz = CDbl(mView(8))

    If NormalizeStraightenVector(vdx, vdy, vdz) = False Then Exit Function

    Dim dotVal As Double
    dotVal = Abs((nx * vdx) + (ny * vdy) + (nz * vdz))

    If dotVal > 1# Then dotVal = 1#

    If dotVal >= 0.9999999 Then
        degOut = 0#
    Else
        degOut = (Atn(-dotVal / Sqr(-dotVal * dotVal + 1#)) + 2# * Atn(1#)) * 180# / PI_VALUE
    End If

    TryGetAngleBetweenFaceNormalAndViewDeg = True

    Exit Function

ErrHandler:
    TryGetAngleBetweenFaceNormalAndViewDeg = False
End Function

' Any unit vector perpendicular to the given one.
Private Function PickAnyPerpendicularAxis(ByVal nx As Double, _
                                          ByVal ny As Double, _
                                          ByVal nz As Double, _
                                          ByRef ox As Double, _
                                          ByRef oy As Double, _
                                          ByRef oz As Double) As Boolean
On Error Resume Next

    PickAnyPerpendicularAxis = False

    Dim tx As Double
    Dim ty As Double
    Dim tz As Double

    If Abs(nx) < 0.9 Then
        tx = 1#: ty = 0#: tz = 0#
    Else
        tx = 0#: ty = 1#: tz = 0#
    End If

    ox = ny * tz - nz * ty
    oy = nz * tx - nx * tz
    oz = nx * ty - ny * tx

    PickAnyPerpendicularAxis = NormalizeStraightenVector(ox, oy, oz)
End Function

Private Function NormalizeStraightenVector(ByRef x As Double, _
                                           ByRef y As Double, _
                                           ByRef z As Double) As Boolean
On Error Resume Next

    NormalizeStraightenVector = False

    Dim mag As Double
    mag = Sqr(x * x + y * y + z * z)

    If mag <= 0.000000001 Then Exit Function

    x = x / mag
    y = y / mag
    z = z / mag

    NormalizeStraightenVector = True
End Function

' ============================================================
' WHOLE-BASE STRAIGHTEN  (SUPERSEDED - kept as a fallback)
' ------------------------------------------------------------
' Only runs when STRAIGHTEN_WHOLE_BASE_BEFORE_ORIENT is True, which it is not.
' Superseded by the Normal To method above, which needs no geometry to move.
'
' Same idea as the pullcore straighten angle, applied to the entire base
' instead of one block:
'
'   1. Build the smallest bounding box that contains the whole base and take
'      the angle that produced it.
'   2. Suppress the mate group, so EditRebuild3 cannot run the solver and
'      drag mated components back toward geometry that just moved.
'   3. FLOAT every component, every level. A fixed component silently refuses
'      Transform2 - some parts move, some do not, and the mold shears.
'   4. Apply ONE rotation, about the centre of the base, to EVERY component in
'      the assembly. No part type is singled out, nothing is skipped by kind.
'      The matrix and the centre are computed once before anything moves and
'      nothing rebuilds mid-loop, so every part turns together by the same
'      amount. The only thing tracked is whether a part has already been
'      carried by a parent that turned, because turning the same body twice is
'      what shears the mold. Moved, carried and refused are all logged, so
'      every part in the assembly is accounted for.
'   5. FIX everything again, then rebuild once.
'
' After that the base is square to the assembly axes and the standard
' orientation process takes over.
' ============================================================

Private Sub StraightenWholeBaseToAssemblyAxes(ByVal model As Object)
On Error GoTo ErrHandler

    BaseStraightenApplied = False
    BaseStraightenTotalDeg = 0#
    BaseStraightenAxesUsed = ""

    If model Is Nothing Then Exit Sub

    If model.GetType <> swDocASSEMBLY Then
        LogLine "Base straighten skipped: document is a part, not an assembly."
        Exit Sub
    End If

    Dim passNo As Long
    Dim axisCode As Long
    Dim angleDeg As Double
    Dim gain As Double

    For passNo = 1 To STRAIGHTEN_BASE_MAX_PASSES

        If BuildWholeBaseAssemblyPointCloud(model) = False Then
            LogLine "Base straighten pass " & passNo & ": could not measure the base (no usable component geometry)."
            Exit For
        End If

        If SolveWholeBaseSquareRotation(axisCode, angleDeg, gain) = False Then
            LogLine "Base straighten pass " & passNo & ": bounding-box search found no orientation better than the current one."
            Exit For
        End If

        LogLine "Base straighten pass " & passNo & ": best axis=" & BaseStraightenAxisName(axisCode) & _
                "  angle=" & Format(angleDeg, "0.000") & " deg" & _
                "  box volume gain=" & Format(gain * 100#, "0.000") & "%"
        LogLine "Base straighten pass " & passNo & ": box as-is    " & BaseCloudBoxTextAtRotation(3, 0#)
        LogLine "Base straighten pass " & passNo & ": box if turned " & BaseCloudBoxTextAtRotation(axisCode, angleDeg)

        If Abs(angleDeg) < STRAIGHTEN_BASE_MIN_DEG Then
            LogLine "Base straighten pass " & passNo & ": angle under " & STRAIGHTEN_BASE_MIN_DEG & _
                    " deg. Base is already square on this axis."
            Exit For
        End If

        If Abs(angleDeg) > STRAIGHTEN_BASE_MAX_DEG Then
            LogLine "Base straighten pass " & passNo & ": REJECTED " & Format(angleDeg, "0.000") & _
                    " deg (over the " & STRAIGHTEN_BASE_MAX_DEG & " deg limit). The box search locked onto the wrong axis."
            Exit For
        End If

        If gain < STRAIGHTEN_BASE_MIN_VOLUME_GAIN Then
            LogLine "Base straighten pass " & passNo & ": volume gain under " & _
                    Format(STRAIGHTEN_BASE_MIN_VOLUME_GAIN * 100#, "0.000") & _
                    "%. Not a real skew; leaving the base as imported."
            Exit For
        End If

        If RotateWholeBaseAsOneRigidSet(model, axisCode, angleDeg) = False Then
            LogLine "Base straighten pass " & passNo & ": rotation did not take. Base left as imported."
            Exit For
        End If

        BaseStraightenApplied = True
        BaseStraightenTotalDeg = BaseStraightenTotalDeg + Abs(angleDeg)

        If BaseStraightenAxesUsed <> "" Then BaseStraightenAxesUsed = BaseStraightenAxesUsed & ", "
        BaseStraightenAxesUsed = BaseStraightenAxesUsed & BaseStraightenAxisName(axisCode) & _
                                 " " & Format(angleDeg, "0.000") & "deg"

        LogLine "Base straighten pass " & passNo & ": APPLIED " & Format(angleDeg, "0.000") & _
                " deg about " & BaseStraightenAxisName(axisCode) & "."

    Next passNo

    If BaseStraightenApplied Then
        LogLine "Base straighten DONE. Total correction " & Format(BaseStraightenTotalDeg, "0.000") & _
                " deg (" & BaseStraightenAxesUsed & ")."
        LogLine "Base is now square to the assembly axes; standard orientation can run normally."
    Else
        LogLine "Base straighten made no change; base was already square to the assembly axes."
    End If

    Exit Sub

ErrHandler:
    LogLine "StraightenWholeBaseToAssemblyAxes error: " & Err.Description
End Sub

Private Function BaseStraightenAxisName(ByVal axisCode As Long) As String
    Select Case axisCode
        Case 1
            BaseStraightenAxisName = "X"
        Case 2
            BaseStraightenAxisName = "Y"
        Case 3
            BaseStraightenAxisName = "Z"
        Case Else
            BaseStraightenAxisName = "?"
    End Select
End Function

' Every component bounding-box corner, expressed in TOP-LEVEL ASSEMBLY space.
'
' LIMITATION, worth knowing before trusting a result: each component contributes
' the 8 corners of its own part-space bounding box. For a mold base that is
' tight, because plates are modelled square in their own part files and the yaw
' lives in the component transform - which is exactly the skew this routine is
' meant to take out. If instead the yaw is baked into the part geometry itself,
' the part-space box is a loose over-box, the solver finds little or no volume
' gain, and the base is correctly left alone rather than rotated wrongly. That
' failure is logged as "not a real skew".
'
' Uses the same transform layout this module already relies on in
' TryGetComponentAssemblyBoundingBoxInches:
'   assemblyPoint = scale * A * partPoint + t
'   A column 0 = m(0),m(1),m(2)   column 1 = m(3),m(4),m(5)
'   A column 2 = m(6),m(7),m(8)   t = m(9),m(10),m(11)   scale = m(12)
Private Function BuildWholeBaseAssemblyPointCloud(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    BuildWholeBaseAssemblyPointCloud = False
    BaseCloudN = 0

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = model.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function
    If IsArray(vComps) = False Then Exit Function
    If UBound(vComps) < 0 Then Exit Function

    Dim capacity As Long
    capacity = (UBound(vComps) + 1) * 8 + 8

    ReDim BaseCloudX(1 To capacity)
    ReDim BaseCloudY(1 To capacity)
    ReDim BaseCloudZ(1 To capacity)

    Dim i As Long
    Dim k As Long
    Dim comp As Object
    Dim cm As Object
    Dim xform As Object
    Dim m As Variant

    Dim xmin As Double, yMin As Double, zMin As Double
    Dim xmax As Double, yMax As Double, zMax As Double

    Dim xs(0 To 7) As Double
    Dim ys(0 To 7) As Double
    Dim zs(0 To 7) As Double

    Dim scaleVal As Double
    Dim tx As Double, ty As Double, tz As Double
    Dim usedComps As Long

    usedComps = 0

    For i = 0 To UBound(vComps)

        Set comp = vComps(i)

        If Not comp Is Nothing Then
            If comp.IsSuppressed = False Then

                Set cm = comp.GetModelDoc2

                If Not cm Is Nothing Then
                    If cm.GetType = swDocPART Then

                        If TryGetPartRawBoxMeters(cm, xmin, yMin, zMin, xmax, yMax, zMax) Then

                            Set xform = comp.Transform2

                            If Not xform Is Nothing Then

                                m = xform.ArrayData

                                If IsArray(m) Then
                                    If UBound(m) >= 12 Then

                                        scaleVal = CDbl(m(12))
                                        If Abs(scaleVal) < 0.0000001 Then scaleVal = 1#

                                        xs(0) = xmin: ys(0) = yMin: zs(0) = zMin
                                        xs(1) = xmax: ys(1) = yMin: zs(1) = zMin
                                        xs(2) = xmin: ys(2) = yMax: zs(2) = zMin
                                        xs(3) = xmax: ys(3) = yMax: zs(3) = zMin
                                        xs(4) = xmin: ys(4) = yMin: zs(4) = zMax
                                        xs(5) = xmax: ys(5) = yMin: zs(5) = zMax
                                        xs(6) = xmin: ys(6) = yMax: zs(6) = zMax
                                        xs(7) = xmax: ys(7) = yMax: zs(7) = zMax

                                        For k = 0 To 7

                                            tx = scaleVal * ((xs(k) * CDbl(m(0))) + (ys(k) * CDbl(m(3))) + (zs(k) * CDbl(m(6)))) + CDbl(m(9))
                                            ty = scaleVal * ((xs(k) * CDbl(m(1))) + (ys(k) * CDbl(m(4))) + (zs(k) * CDbl(m(7)))) + CDbl(m(10))
                                            tz = scaleVal * ((xs(k) * CDbl(m(2))) + (ys(k) * CDbl(m(5))) + (zs(k) * CDbl(m(8)))) + CDbl(m(11))

                                            If BaseCloudN < capacity Then
                                                BaseCloudN = BaseCloudN + 1
                                                BaseCloudX(BaseCloudN) = tx
                                                BaseCloudY(BaseCloudN) = ty
                                                BaseCloudZ(BaseCloudN) = tz
                                            End If

                                        Next k

                                        usedComps = usedComps + 1

                                    End If
                                End If

                            End If

                        End If

                    End If
                End If

            End If
        End If

    Next i

    LogLine "Base straighten measure: " & usedComps & " part component(s), " & BaseCloudN & " assembly-space box corner(s)."

    BuildWholeBaseAssemblyPointCloud = (BaseCloudN >= 8)

    Exit Function

ErrHandler:
    LogLine "BuildWholeBaseAssemblyPointCloud error: " & Err.Description
    BuildWholeBaseAssemblyPointCloud = False
End Function

' Finds the single global-axis rotation that produces the smallest bounding
' box around the whole base. Sweeping -45..+45 is enough because the box
' repeats every 90 degrees.
'
' The returned angle is the angle to ROTATE BY: the volume was measured on the
' cloud after rotating it by that angle, so rotating the model by the same
' angle lands it in that box. No sign guessing anywhere.
Private Function SolveWholeBaseSquareRotation(ByRef axisCodeOut As Long, _
                                             ByRef angleDegOut As Double, _
                                             ByRef gainOut As Double) As Boolean
On Error GoTo ErrHandler

    SolveWholeBaseSquareRotation = False

    axisCodeOut = 0
    angleDegOut = 0#
    gainOut = 0#

    If BaseCloudN < 8 Then Exit Function

    Dim baseVol As Double
    baseVol = BaseCloudBoxVolumeAtRotation(3, 0#)

    If baseVol <= 0# Then Exit Function

    Dim bestVol As Double
    Dim bestAxis As Long
    Dim bestAng As Double

    bestVol = baseVol
    bestAxis = 0
    bestAng = 0#

    Dim ax As Long
    Dim a As Double
    Dim vol As Double

    ' Coarse sweep on all three global axes. Rotating about X cannot change
    ' the X extent, about Y cannot change the Y extent, and so on, so
    ' comparing full box volumes across the three axes is a fair comparison.
    For ax = 1 To 3

        a = -45#
        Do While a <= 45#

            vol = BaseCloudBoxVolumeAtRotation(ax, a)

            If vol > 0# Then
                If vol < bestVol Then
                    bestVol = vol
                    bestAxis = ax
                    bestAng = a
                End If
            End If

            a = a + STRAIGHTEN_BASE_COARSE_STEP_DEG

        Loop

    Next ax

    If bestAxis = 0 Then Exit Function

    ' Refine around the coarse winner, on the winning axis only.
    Dim lo As Double
    Dim hi As Double

    lo = bestAng - STRAIGHTEN_BASE_COARSE_STEP_DEG
    hi = bestAng + STRAIGHTEN_BASE_COARSE_STEP_DEG

    a = lo
    Do While a <= hi

        vol = BaseCloudBoxVolumeAtRotation(bestAxis, a)

        If vol > 0# Then
            If vol < bestVol Then
                bestVol = vol
                bestAng = a
            End If
        End If

        a = a + STRAIGHTEN_BASE_FINE_STEP_DEG

    Loop

    axisCodeOut = bestAxis
    angleDegOut = bestAng
    gainOut = (baseVol - bestVol) / baseVol

    SolveWholeBaseSquareRotation = True

    Exit Function

ErrHandler:
    LogLine "SolveWholeBaseSquareRotation error: " & Err.Description
    SolveWholeBaseSquareRotation = False
End Function

' Axis-aligned bounding-box volume of the cloud after rotating it about one
' global axis. Returns 0 when the cloud is degenerate.
Private Function BaseCloudBoxVolumeAtRotation(ByVal axisCode As Long, _
                                              ByVal angleDeg As Double) As Double

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    BaseCloudBoxVolumeAtRotation = 0#

    If BaseCloudBoxExtentsAtRotation(axisCode, angleDeg, dX, dY, dZ) = False Then Exit Function

    BaseCloudBoxVolumeAtRotation = dX * dY * dZ
End Function

' Axis-aligned bounding-box extents (metres) of the cloud after rotating it
' about one global axis.
Private Function BaseCloudBoxExtentsAtRotation(ByVal axisCode As Long, _
                                               ByVal angleDeg As Double, _
                                               ByRef outDx As Double, _
                                               ByRef outDy As Double, _
                                               ByRef outDz As Double) As Boolean
On Error GoTo ErrHandler

    BaseCloudBoxExtentsAtRotation = False

    outDx = 0#
    outDy = 0#
    outDz = 0#

    If BaseCloudN < 8 Then Exit Function

    Dim rad As Double
    Dim c As Double
    Dim s As Double

    rad = angleDeg * PI_VALUE / 180#
    c = Cos(rad)
    s = Sin(rad)

    Dim minX As Double, minY As Double, minZ As Double
    Dim maxX As Double, maxY As Double, maxZ As Double

    Dim i As Long
    Dim rx As Double, ry As Double, rz As Double

    minX = 0#: minY = 0#: minZ = 0#
    maxX = 0#: maxY = 0#: maxZ = 0#

    For i = 1 To BaseCloudN

        RotateBaseCloudPoint axisCode, c, s, _
                             BaseCloudX(i), BaseCloudY(i), BaseCloudZ(i), _
                             rx, ry, rz

        If i = 1 Then
            minX = rx: maxX = rx
            minY = ry: maxY = ry
            minZ = rz: maxZ = rz
        Else
            If rx < minX Then minX = rx
            If rx > maxX Then maxX = rx
            If ry < minY Then minY = ry
            If ry > maxY Then maxY = ry
            If rz < minZ Then minZ = rz
            If rz > maxZ Then maxZ = rz
        End If

    Next i

    outDx = maxX - minX
    outDy = maxY - minY
    outDz = maxZ - minZ

    If outDx <= 0# Then Exit Function
    If outDy <= 0# Then Exit Function
    If outDz <= 0# Then Exit Function

    BaseCloudBoxExtentsAtRotation = True

    Exit Function

ErrHandler:
    BaseCloudBoxExtentsAtRotation = False
End Function

' "L x W x T inches" for the cloud at a given rotation, for the log.
Private Function BaseCloudBoxTextAtRotation(ByVal axisCode As Long, _
                                            ByVal angleDeg As Double) As String
On Error Resume Next

    BaseCloudBoxTextAtRotation = "n/a"

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    If BaseCloudBoxExtentsAtRotation(axisCode, angleDeg, dX, dY, dZ) = False Then Exit Function

    Dim L As Double
    Dim W As Double
    Dim T As Double

    SortThreeDimensions dX * INCHES_PER_METER, _
                        dY * INCHES_PER_METER, _
                        dZ * INCHES_PER_METER, _
                        L, W, T

    BaseCloudBoxTextAtRotation = Format(L, "0.000") & " x " & _
                                 Format(W, "0.000") & " x " & _
                                 Format(T, "0.000") & " in"
End Function

' Rotates one point about a global axis. Right-handed. This is the ONE
' definition of the rotation; BuildBaseRotationMatrix builds the matrix form
' of exactly the same thing, so the angle that was measured is the angle that
' gets applied.
Private Sub RotateBaseCloudPoint(ByVal axisCode As Long, _
                                 ByVal c As Double, _
                                 ByVal s As Double, _
                                 ByVal x As Double, _
                                 ByVal y As Double, _
                                 ByVal z As Double, _
                                 ByRef outX As Double, _
                                 ByRef outY As Double, _
                                 ByRef outZ As Double)

    Select Case axisCode

        Case 1                      ' about X: rotate in the Y/Z plane
            outX = x
            outY = c * y - s * z
            outZ = s * y + c * z

        Case 2                      ' about Y: rotate in the Z/X plane
            outX = c * x + s * z
            outY = y
            outZ = -s * x + c * z

        Case Else                   ' about Z: rotate in the X/Y plane
            outX = c * x - s * y
            outY = s * x + c * y
            outZ = z

    End Select
End Sub

' Matrix form of RotateBaseCloudPoint: r is row-major, newPoint = r * point.
Private Sub BuildBaseRotationMatrix(ByVal axisCode As Long, _
                                    ByVal angleDeg As Double, _
                                    ByRef r() As Double)

    Dim rad As Double
    Dim c As Double
    Dim s As Double

    rad = angleDeg * PI_VALUE / 180#
    c = Cos(rad)
    s = Sin(rad)

    Dim i As Long
    Dim j As Long

    For i = 0 To 2
        For j = 0 To 2
            r(i, j) = 0#
        Next j
    Next i

    Select Case axisCode

        Case 1                      ' about X
            r(0, 0) = 1#
            r(1, 1) = c
            r(1, 2) = -s
            r(2, 1) = s
            r(2, 2) = c

        Case 2                      ' about Y
            r(0, 0) = c
            r(0, 2) = s
            r(1, 1) = 1#
            r(2, 0) = -s
            r(2, 2) = c

        Case Else                   ' about Z
            r(0, 0) = c
            r(0, 1) = -s
            r(1, 0) = s
            r(1, 1) = c
            r(2, 2) = 1#

    End Select
End Sub

' Float everything, rotate the whole base as ONE rigid set, fix everything.
Private Function RotateWholeBaseAsOneRigidSet(ByVal model As Object, _
                                              ByVal axisCode As Long, _
                                              ByVal angleDeg As Double) As Boolean
On Error GoTo ErrHandler

    RotateWholeBaseAsOneRigidSet = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    ' Rotate about the centre of the base so the mold stays where it was and
    ' does not swing away from the origin.
    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    If GetBaseCloudCenter(cx, cy, cz) = False Then
        LogLine "RotateWholeBase: no measured centre available. Rotation aborted."
        Exit Function
    End If

    Dim mathUtil As Object
    Set mathUtil = swApp.GetMathUtility

    If mathUtil Is Nothing Then
        LogLine "RotateWholeBase: MathUtility unavailable. Rotation aborted."
        Exit Function
    End If

    Dim r(0 To 2, 0 To 2) As Double
    BuildBaseRotationMatrix axisCode, angleDeg, r

    ' 1. Mates first. A live mate group means EditRebuild3 runs the solver,
    '    and the solver drags mated components back toward geometry that just
    '    moved. That is one of the two causes of a scrambled mold.
    Dim matesTouched As Long
    matesTouched = SetAllAssemblyMateGroupsSuppressed(model, True)
    LogLine "RotateWholeBase: suppressed " & matesTouched & " mate group(s)."

    ' 2. Float EVERYTHING, every level. A FIXED component silently refuses
    '    Transform2, so some parts move and some do not. That is the other cause.
    Dim floated As Long
    floated = SetAllComponentsFixed(model, False)
    LogLine "RotateWholeBase: floated " & floated & " component(s)."

    ' 3. ONE rotation, applied to EVERY component in the assembly, every level,
    '    with no attempt to tell one kind of part from another. The matrix r and
    '    the centre are computed once, above, before anything moves, and nothing
    '    rebuilds inside this loop - so every part turns by the same amount in
    '    the same operation.
    '
    '    Components are walked shallowest-first. When a parent component turns,
    '    SolidWorks carries its children with it, so a child that has already
    '    been carried is not turned a second time - turning it again is what
    '    shears the mold. Carried parts are counted and logged, so every single
    '    part in the assembly is accounted for as moved, carried, or refused.
    Dim vAll As Variant
    vAll = model.GetComponents(False)

    If IsArray(vAll) = False Then
        LogLine "RotateWholeBase: no components returned. Rotation aborted."
        SetAllComponentsFixed model, True
        Exit Function
    End If

    If UBound(vAll) < 0 Then
        LogLine "RotateWholeBase: component list is empty. Rotation aborted."
        SetAllComponentsFixed model, True
        Exit Function
    End If

    Dim moved As Long
    Dim carried As Long
    Dim refused As Long
    Dim refusedNames As String
    Dim movedNames As String

    moved = 0
    carried = 0
    refused = 0
    refusedNames = ""
    movedNames = "|"

    Dim maxDepth As Long
    Dim i As Long
    Dim comp As Object
    Dim d As Long

    maxDepth = 0
    For i = 0 To UBound(vAll)
        Set comp = vAll(i)
        If Not comp Is Nothing Then
            d = BaseStraightenComponentDepth(comp)
            If d > maxDepth Then maxDepth = d
        End If
    Next i

    Dim pass As Long
    Dim nm As String

    For pass = 0 To maxDepth

        For i = 0 To UBound(vAll)

            Set comp = vAll(i)

            If Not comp Is Nothing Then
                If comp.IsSuppressed = False Then
                    If BaseStraightenComponentDepth(comp) = pass Then

                        nm = BaseStraightenComponentName(comp)

                        If BaseStraightenNameIsUnderMoved(nm, movedNames) Then

                            carried = carried + 1

                        ElseIf ApplyBaseRotationToOneComponent(comp, mathUtil, r, cx, cy, cz) Then

                            moved = moved + 1
                            movedNames = movedNames & nm & "|"

                        Else

                            refused = refused + 1
                            If Len(refusedNames) < 400 Then
                                If refusedNames <> "" Then refusedNames = refusedNames & ", "
                                refusedNames = refusedNames & nm
                            End If

                        End If

                    End If
                End If
            End If

        Next i

    Next pass

    ' 4. Fix everything again so nothing drifts in any later rebuild.
    Dim refixed As Long
    refixed = SetAllComponentsFixed(model, True)

    ' 5. Rebuild once, now that nothing can be dragged.
    On Error Resume Next
    model.EditRebuild3
    model.ForceRebuild3 False
    model.ClearSelection2 True
    Err.Clear
    On Error GoTo ErrHandler

    StabilizeActiveView model, 150

    LogLine "RotateWholeBase: " & (UBound(vAll) + 1) & " component(s) in the assembly: " & _
            moved & " turned, " & carried & " carried by a parent that turned, " & refused & " refused."
    If refused > 0 Then LogLine "RotateWholeBase: refused component(s): " & refusedNames
    LogLine "RotateWholeBase: re-fixed " & refixed & " component(s)."

    If STRAIGHTEN_BASE_KEEP_MATES_SUPPRESSED = False Then
        LogLine "RotateWholeBase: unsuppressed " & SetAllAssemblyMateGroupsSuppressed(model, False) & " mate group(s)."
    End If

    RotateWholeBaseAsOneRigidSet = (moved > 0)

    Exit Function

ErrHandler:
    LogLine "RotateWholeBaseAsOneRigidSet error: " & Err.Description

    On Error Resume Next
    SetAllComponentsFixed model, True
    model.ClearSelection2 True

    RotateWholeBaseAsOneRigidSet = False
End Function

' newOrientation = r * oldOrientation
' newPosition    = r * (oldPosition - centre) + centre
Private Function ApplyBaseRotationToOneComponent(ByVal comp As Object, _
                                                 ByVal mathUtil As Object, _
                                                 ByRef r() As Double, _
                                                 ByVal cx As Double, _
                                                 ByVal cy As Double, _
                                                 ByVal cz As Double) As Boolean
On Error GoTo ErrHandler

    ApplyBaseRotationToOneComponent = False

    If comp Is Nothing Then Exit Function
    If mathUtil Is Nothing Then Exit Function

    Dim xform As Object
    Set xform = comp.Transform2

    If xform Is Nothing Then Exit Function

    Dim m As Variant
    m = xform.ArrayData

    If IsArray(m) = False Then Exit Function
    If UBound(m) < 12 Then Exit Function

    Dim a(0 To 2, 0 To 2) As Double

    a(0, 0) = CDbl(m(0)): a(1, 0) = CDbl(m(1)): a(2, 0) = CDbl(m(2))
    a(0, 1) = CDbl(m(3)): a(1, 1) = CDbl(m(4)): a(2, 1) = CDbl(m(5))
    a(0, 2) = CDbl(m(6)): a(1, 2) = CDbl(m(7)): a(2, 2) = CDbl(m(8))

    Dim tOldX As Double
    Dim tOldY As Double
    Dim tOldZ As Double

    tOldX = CDbl(m(9))
    tOldY = CDbl(m(10))
    tOldZ = CDbl(m(11))

    Dim scaleVal As Double
    scaleVal = CDbl(m(12))
    If Abs(scaleVal) < 0.0000001 Then scaleVal = 1#

    Dim aNew(0 To 2, 0 To 2) As Double
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim acc As Double

    For i = 0 To 2
        For j = 0 To 2
            acc = 0#
            For k = 0 To 2
                acc = acc + r(i, k) * a(k, j)
            Next k
            aNew(i, j) = acc
        Next j
    Next i

    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double

    dX = tOldX - cx
    dY = tOldY - cy
    dZ = tOldZ - cz

    Dim tNewX As Double
    Dim tNewY As Double
    Dim tNewZ As Double

    tNewX = r(0, 0) * dX + r(0, 1) * dY + r(0, 2) * dZ + cx
    tNewY = r(1, 0) * dX + r(1, 1) * dY + r(1, 2) * dZ + cy
    tNewZ = r(2, 0) * dX + r(2, 1) * dY + r(2, 2) * dZ + cz

    Dim newArr(0 To 15) As Double

    newArr(0) = aNew(0, 0): newArr(1) = aNew(1, 0): newArr(2) = aNew(2, 0)
    newArr(3) = aNew(0, 1): newArr(4) = aNew(1, 1): newArr(5) = aNew(2, 1)
    newArr(6) = aNew(0, 2): newArr(7) = aNew(1, 2): newArr(8) = aNew(2, 2)
    newArr(9) = tNewX
    newArr(10) = tNewY
    newArr(11) = tNewZ
    newArr(12) = scaleVal
    newArr(13) = 0#
    newArr(14) = 0#
    newArr(15) = 0#

    Dim newXform As Object
    Set newXform = mathUtil.CreateTransform(newArr)

    If newXform Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set comp.Transform2 = newXform
    If Err.Number <> 0 Then
        Err.Clear
        comp.Transform2 = newXform
    End If
    Err.Clear
    On Error GoTo ErrHandler

    ' Verify the write actually took. A component that is still fixed, or is
    ' driven by a live mate, silently keeps its old transform - and a silent
    ' refusal is exactly what leaves half the mold behind.
    Dim chkXform As Object
    Set chkXform = comp.Transform2

    If chkXform Is Nothing Then Exit Function

    Dim chk As Variant
    chk = chkXform.ArrayData

    If IsArray(chk) = False Then Exit Function
    If UBound(chk) < 11 Then Exit Function

    Dim tol As Double
    tol = 0.000001

    If Abs(CDbl(chk(9)) - tNewX) > tol Then Exit Function
    If Abs(CDbl(chk(10)) - tNewY) > tol Then Exit Function
    If Abs(CDbl(chk(11)) - tNewZ) > tol Then Exit Function

    If Abs(CDbl(chk(0)) - aNew(0, 0)) > tol Then Exit Function
    If Abs(CDbl(chk(4)) - aNew(1, 1)) > tol Then Exit Function
    If Abs(CDbl(chk(8)) - aNew(2, 2)) > tol Then Exit Function

    ApplyBaseRotationToOneComponent = True

    Exit Function

ErrHandler:
    ApplyBaseRotationToOneComponent = False
End Function

Private Function GetBaseCloudCenter(ByRef cx As Double, _
                                    ByRef cy As Double, _
                                    ByRef cz As Double) As Boolean
On Error GoTo ErrHandler

    GetBaseCloudCenter = False

    cx = 0#
    cy = 0#
    cz = 0#

    If BaseCloudN < 1 Then Exit Function

    Dim minX As Double, minY As Double, minZ As Double
    Dim maxX As Double, maxY As Double, maxZ As Double

    minX = BaseCloudX(1): maxX = minX
    minY = BaseCloudY(1): maxY = minY
    minZ = BaseCloudZ(1): maxZ = minZ

    Dim i As Long

    For i = 2 To BaseCloudN
        If BaseCloudX(i) < minX Then minX = BaseCloudX(i)
        If BaseCloudX(i) > maxX Then maxX = BaseCloudX(i)
        If BaseCloudY(i) < minY Then minY = BaseCloudY(i)
        If BaseCloudY(i) > maxY Then maxY = BaseCloudY(i)
        If BaseCloudZ(i) < minZ Then minZ = BaseCloudZ(i)
        If BaseCloudZ(i) > maxZ Then maxZ = BaseCloudZ(i)
    Next i

    cx = (minX + maxX) / 2#
    cy = (minY + maxY) / 2#
    cz = (minZ + maxZ) / 2#

    GetBaseCloudCenter = True

    Exit Function

ErrHandler:
    GetBaseCloudCenter = False
End Function

' True  = fix every component
' False = float (unfix) every component
'
' Selects every component at every level and acts on the whole selection in one
' call. If SolidWorks rejects a selection that includes nested components, it
' retries with top-level only rather than leaving the assembly half floated -
' a half-floated assembly is what makes some parts move and others not.
' Returns how many components were in the selection that was acted on.
Private Function SetAllComponentsFixed(ByVal model As Object, _
                                       ByVal wantFixed As Boolean) As Long
On Error GoTo ErrHandler

    SetAllComponentsFixed = 0

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim n As Long

    n = SelectAndSetFixedOnce(model, wantFixed, False)
    If n > 0 Then
        SetAllComponentsFixed = n
        Exit Function
    End If

    LogLine "SetAllComponentsFixed: all-levels " & IIf(wantFixed, "fix", "float") & _
            " did not take. Retrying top-level only."

    SetAllComponentsFixed = SelectAndSetFixedOnce(model, wantFixed, True)

    Exit Function

ErrHandler:
    LogLine "SetAllComponentsFixed error: " & Err.Description

    On Error Resume Next
    model.ClearSelection2 True

    SetAllComponentsFixed = 0
End Function

Private Function SelectAndSetFixedOnce(ByVal model As Object, _
                                       ByVal wantFixed As Boolean, _
                                       ByVal topLevelOnly As Boolean) As Long
On Error GoTo ErrHandler

    SelectAndSetFixedOnce = 0

    Dim vComps As Variant
    vComps = model.GetComponents(topLevelOnly)

    If IsArray(vComps) = False Then Exit Function
    If UBound(vComps) < 0 Then Exit Function

    model.ClearSelection2 True

    Dim selectedCount As Long
    Dim i As Long
    Dim comp As Object

    selectedCount = 0

    For i = 0 To UBound(vComps)

        Set comp = vComps(i)

        If Not comp Is Nothing Then
            If comp.IsSuppressed = False Then
                If comp.Select4(True, Nothing, False) Then selectedCount = selectedCount + 1
            End If
        End If

    Next i

    If selectedCount = 0 Then
        model.ClearSelection2 True
        Exit Function
    End If

    On Error Resume Next
    Err.Clear

    If wantFixed Then
        model.FixComponent
    Else
        model.UnfixComponent
    End If

    If Err.Number <> 0 Then
        LogLine "SelectAndSetFixedOnce: " & IIf(wantFixed, "FixComponent", "UnfixComponent") & _
                " failed (" & Err.Description & ")."
        Err.Clear
        selectedCount = 0
    End If

    On Error GoTo ErrHandler

    model.ClearSelection2 True

    SelectAndSetFixedOnce = selectedCount

    Exit Function

ErrHandler:
    LogLine "SelectAndSetFixedOnce error: " & Err.Description

    On Error Resume Next
    model.ClearSelection2 True

    SelectAndSetFixedOnce = 0
End Function

' Nesting depth of a component, from its name path. A top-level component is
' "PART-1" = depth 0; a child of a sub-assembly is "SUBASSY-1/PART-2" = depth 1.
Private Function BaseStraightenComponentDepth(ByVal comp As Object) As Long
On Error Resume Next

    BaseStraightenComponentDepth = 0

    If comp Is Nothing Then Exit Function

    Dim nm As String
    nm = BaseStraightenComponentName(comp)

    If nm = "" Then Exit Function

    Dim i As Long
    Dim d As Long

    d = 0
    For i = 1 To Len(nm)
        If mid(nm, i, 1) = "/" Then d = d + 1
    Next i

    BaseStraightenComponentDepth = d
End Function

' True when this component sits inside one that has already been turned, so it
' has already been carried and must not be turned a second time.
' movedNames is a "|name|name|" list of components already turned.
Private Function BaseStraightenNameIsUnderMoved(ByVal nm As String, _
                                                ByVal movedNames As String) As Boolean
On Error Resume Next

    BaseStraightenNameIsUnderMoved = False

    If nm = "" Then Exit Function
    If InStr(nm, "/") = 0 Then Exit Function

    Dim parentPath As String
    Dim cut As Long

    parentPath = nm

    Do
        cut = InStrRev(parentPath, "/")
        If cut <= 0 Then Exit Do

        parentPath = Left(parentPath, cut - 1)
        If parentPath = "" Then Exit Do

        If InStr(1, movedNames, "|" & parentPath & "|", vbTextCompare) > 0 Then
            BaseStraightenNameIsUnderMoved = True
            Exit Function
        End If
    Loop

End Function

' Selects the assembly's mate group feature(s) by walking the feature tree -
' no selection-type string to guess wrong - and suppresses or unsuppresses
' them. Returns how many mate groups were selected.
Private Function SetAllAssemblyMateGroupsSuppressed(ByVal model As Object, _
                                                    ByVal wantSuppressed As Boolean) As Long
On Error GoTo ErrHandler

    SetAllAssemblyMateGroupsSuppressed = 0

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    model.ClearSelection2 True

    Dim feat As Object
    Dim featTypeName As String
    Dim n As Long
    Dim guard As Long

    n = 0
    guard = 0

    Set feat = model.FirstFeature

    Do While Not feat Is Nothing

        guard = guard + 1
        If guard > 20000 Then Exit Do

        featTypeName = ""

        On Error Resume Next
        Err.Clear
        featTypeName = CStr(feat.GetTypeName2)
        If Err.Number <> 0 Then
            Err.Clear
            featTypeName = CStr(feat.GetTypeName)
            Err.Clear
        End If
        On Error GoTo ErrHandler

        If UCase(featTypeName) = "MATEGROUP" Then

            On Error Resume Next
            Err.Clear
            If CBool(feat.Select2(True, 0)) Then n = n + 1
            Err.Clear
            On Error GoTo ErrHandler

        End If

        Set feat = feat.GetNextFeature

    Loop

    If n > 0 Then

        On Error Resume Next
        Err.Clear

        If wantSuppressed Then
            model.EditSuppress2
        Else
            model.EditUnsuppress2
        End If

        Err.Clear
        On Error GoTo ErrHandler

    End If

    model.ClearSelection2 True

    SetAllAssemblyMateGroupsSuppressed = n

    Exit Function

ErrHandler:
    LogLine "SetAllAssemblyMateGroupsSuppressed error: " & Err.Description

    On Error Resume Next
    model.ClearSelection2 True

    SetAllAssemblyMateGroupsSuppressed = 0
End Function

Private Function BaseStraightenComponentName(ByVal comp As Object) As String
On Error Resume Next

    BaseStraightenComponentName = "?"

    If comp Is Nothing Then Exit Function

    BaseStraightenComponentName = CStr(comp.Name2)
End Function

' ============================================================
' END OF MODULE
' ============================================================

' ============================================================
' POTS AND HOLDERS DXF PACKAGE
' ------------------------------------------------------------
' Why this exists.
'
' The BASE assembly DXF draws the pot blocks, the holder blocks and the clamping
' plates on top of each other in every projected view. Once it is a flat DXF
' there is no component information left, so pot geometry cannot be reliably
' separated back out of holder geometry - in the views where the two overlap in
' projection there is no line you can cut along.
'
' The assembly still knows. This package isolates the pot blocks, exports their
' DXF, then isolates the holder blocks and exports theirs - the same way the
' S09-S12 block PDFs are already produced. The two files never contain each
' other's geometry, so nothing has to be subtracted afterwards.
' ============================================================



' Component-name keys used to collect the blocks. Every matching component is
' kept, so a job with top + bottom (or ID + OD) blocks exports all of them
' together into the one file.

' Names that must never be pulled into either group.

' Written next to the DXFs so the block sizes the geometry scan found are
' recorded with the files they produced.


Private Sub ExportPotsAndHoldersDxfPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "POTS/HOLDERS package: main document is not an assembly; skipped."
        Exit Sub
    End If

    Dim potIdx As Collection
    Dim holdIdx As Collection

    Set potIdx = CollectBlockIndexesByKeys(POT_BLOCK_COLLECT_KEYS, "POT")
    Set holdIdx = CollectBlockIndexesByKeys(HOLDER_BLOCK_COLLECT_KEYS, "HOLDER")

    LogLine "POTS/HOLDERS package: pot blocks found = " & potIdx.count & _
            ", holder blocks found = " & holdIdx.count

    If potIdx.count = 0 And holdIdx.count = 0 Then
        LogLine "POTS/HOLDERS package: no pot or holder blocks identified; skipped."
        Exit Sub
    End If

    Dim phFolder As String
    phFolder = outputFolder & "\" & CurrentJobNumber & " " & POTS_HOLDERS_FOLDER_NAME

    EnsureFolderDeep phFolder

    WritePotsHoldersDimensionReport phFolder & "\" & POTS_HOLDERS_DIM_REPORT, potIdx, holdIdx

    If potIdx.count > 0 Then
        LogStart "POTS DXF (isolated)"
        ExportIsolatedGroupDxf phFolder, potIdx, POTS_DXF_FILE_TOKEN, "POTS"
        LogDone "POTS DXF (isolated)"
    End If

    If holdIdx.count > 0 Then
        LogStart "HOLDERS DXF (isolated)"
        ExportIsolatedGroupDxf phFolder, holdIdx, HOLDERS_DXF_FILE_TOKEN, "HOLDERS"
        LogDone "HOLDERS DXF (isolated)"
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportPotsAndHoldersDxfPackage error: " & Err.Description
    On Error Resume Next
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
End Sub


' Every scanned component whose name matches one of the keys and none of the
' exclusions. Logs the dimensions the geometry scan recorded for each one.
Private Function CollectBlockIndexesByKeys(ByVal pipeKeys As String, _
                                           ByVal label As String) As Collection
On Error GoTo ErrHandler

    Dim out As Collection
    Set out = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim idx As Long
    Dim hay As String

    ' ---- 1. by resolved BOM quote name (the reliable route) ----------------
    ' Imported assemblies name their components "<file>.sldasm-Part-27-1", so
    ' matching on component names finds nothing - which is exactly why this
    ' package logged "pot blocks found = 0, holder blocks found = 0" on J8459.
    ' The quote name from the BOM is where "ID HOLDER" / "OD POT BLOCK" lives.
    For i = 1 To ExportCount

        If ExportRows(i).HasCad Then

            idx = ExportRows(i).CadPartIndex

            If idx > 0 And idx <= PartCount Then

                hay = NormalizeText(ExportRows(i).quoteName)

                If ContainsAnyPipeKey(hay, pipeKeys) Then
                    If ContainsAnyPipeKey(hay, POTS_HOLDERS_EXCLUDE_KEYS) = False Then
                        If seen.Exists(CStr(idx)) = False Then
                            seen(CStr(idx)) = True
                            out.Add idx
                            LogBlockPick label, idx, ExportRows(i).quoteName, "BOM quote name"
                        End If
                    End If
                End If

            End If

        End If

    Next i

    ' ---- 2. fallback: component / clean / file names -----------------------
    ' Still worth trying, for assemblies whose components ARE named properly.
    For i = 1 To PartCount

        If parts(i).isBodyOnly = False Then

            If seen.Exists(CStr(i)) = False Then

                hay = NormalizeText(parts(i).cleanName & " " & _
                                    parts(i).componentName & " " & _
                                    parts(i).filePath)

                If ContainsAnyPipeKey(hay, pipeKeys) Then
                    If ContainsAnyPipeKey(hay, POTS_HOLDERS_EXCLUDE_KEYS) = False Then
                        seen(CStr(i)) = True
                        out.Add i
                        LogBlockPick label, i, parts(i).componentName, "component name"
                    End If
                End If

            End If

        End If

    Next i

    Set CollectBlockIndexesByKeys = out
    Exit Function

ErrHandler:
    LogLine "CollectBlockIndexesByKeys error (" & label & "): " & Err.Description
    If out Is Nothing Then Set out = New Collection
    Set CollectBlockIndexesByKeys = out
End Function


Private Sub LogBlockPick(ByVal label As String, _
                         ByVal idx As Long, _
                         ByVal matchedOn As String, _
                         ByVal route As String)
On Error Resume Next

    If idx <= 0 Or idx > PartCount Then Exit Sub

    LogLine "  " & label & " block (" & route & "): " & matchedOn
    LogLine "     component  = " & parts(idx).componentName
    LogLine "     scan L/W/T = " & _
            FormatNumberForCsv(parts(idx).Length) & " / " & _
            FormatNumberForCsv(parts(idx).Width) & " / " & _
            FormatNumberForCsv(parts(idx).Thickness)

    If parts(idx).hasOriginalAsmBBox Then
        LogLine "     assembly bbox L/W/T = " & _
                FormatNumberForCsv(parts(idx).OriginalAsmLength) & " / " & _
                FormatNumberForCsv(parts(idx).OriginalAsmWidth) & " / " & _
                FormatNumberForCsv(parts(idx).OriginalAsmThickness)
    End If

    If parts(idx).hasAsmCenter Then
        LogLine "     assembly centre X/Y/Z = " & _
                FormatNumberForCsv(parts(idx).AsmCenterX) & " / " & _
                FormatNumberForCsv(parts(idx).AsmCenterY) & " / " & _
                FormatNumberForCsv(parts(idx).AsmCenterZ)
    End If
End Sub


' Hide everything except the listed components, save an isolated native SLDASM
' so the CMS_TOP / standard orientation survives, and build the DXF from it.
Private Sub ExportIsolatedGroupDxf(ByVal folderPath As String, _
                                   ByVal cadIndexes As Collection, _
                                   ByVal fileToken As String, _
                                   ByVal dxfQuoteName As String)
On Error GoTo ErrHandler

    If cadIndexes Is Nothing Then Exit Sub
    If cadIndexes.count = 0 Then Exit Sub

    Dim custToken As String
    Dim dateToken As String
    Dim token As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)
    token = CleanQuoteTokenForFile(fileToken)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")
    If token = "" Then token = "GROUP"

    Dim dxfPath As String
    Dim xtPath As String

    dxfPath = folderPath & "\" & CurrentJobNumber & "_" & token & "_" & custToken & "_" & dateToken & ".dxf"
    xtPath = folderPath & "\" & CurrentJobNumber & "_" & token & "_" & custToken & "_" & dateToken & ".x_t"

    dxfPath = GetUniqueFilePath(dxfPath)
    xtPath = GetUniqueFilePath(xtPath)

    Dim keepNames As Collection
    Set keepNames = New Collection

    Dim i As Long
    Dim idx As Long

    For i = 1 To cadIndexes.count
        idx = CLng(cadIndexes(i))
        If idx > 0 And idx <= PartCount Then
            AddUniqueComponentName keepNames, parts(idx).componentName
        End If
    Next i

    If keepNames.count = 0 Then Exit Sub

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_" & CleanFileName(token) & "_DXF_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_" & CleanFileName(token) & "_TEMP.sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine dxfQuoteName & " DXF: could not isolate the components; skipped."
        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 150

    ' X_T alongside the DXF, for anyone who wants the solid.
    SaveModelAs swModel, xtPath

    LogLine dxfQuoteName & " DXF: saving isolated native SLDASM:"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(swModel, tempNativePath) = False Then GoTo CleanExit

    LogLine dxfQuoteName & " DXF -> " & dxfPath

    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, dxfQuoteName, _
                                     CMS_TOP_VIEW_NAME, "*Top", _
                                     False, False, False, True

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If tempFolder <> "" Then
            If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
        End If
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportIsolatedGroupDxf error (" & dxfQuoteName & "): " & Err.Description
    Resume CleanExit
End Sub


' Every dimension the geometry scan holds for the blocks that went into the two
' DXFs, so the files can be checked against the BOM without re-running anything.
Private Sub WritePotsHoldersDimensionReport(ByVal reportPath As String, _
                                            ByVal potIdx As Collection, _
                                            ByVal holdIdx As Collection)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.CreateTextFile(reportPath, True)

    ts.WriteLine "Job,Group,Component,Clean Name,Scan Length,Scan Width,Scan Thickness," & _
                 "Asm BBox Length,Asm BBox Width,Asm BBox Thickness," & _
                 "Asm Center X,Asm Center Y,Asm Center Z,Mass,File"

    WritePotsHoldersDimRows ts, potIdx, "POT"
    WritePotsHoldersDimRows ts, holdIdx, "HOLDER"

    ts.Close

    LogLine "Pots/Holders dimension report: " & reportPath
    Exit Sub

ErrHandler:
    LogLine "WritePotsHoldersDimensionReport error: " & Err.Description
End Sub


Private Sub WritePotsHoldersDimRows(ByVal ts As Object, _
                                    ByVal idxList As Collection, _
                                    ByVal groupName As String)
On Error Resume Next

    If idxList Is Nothing Then Exit Sub

    Dim i As Long
    Dim k As Long

    For i = 1 To idxList.count

        k = CLng(idxList(i))
        If k <= 0 Or k > PartCount Then GoTo NextRow

        ts.WriteLine CsvText(CurrentJobNumber) & "," & _
                     CsvText(groupName) & "," & _
                     CsvText(parts(k).componentName) & "," & _
                     CsvText(parts(k).cleanName) & "," & _
                     FormatNumberForCsv(parts(k).Length) & "," & _
                     FormatNumberForCsv(parts(k).Width) & "," & _
                     FormatNumberForCsv(parts(k).Thickness) & "," & _
                     FormatNumberForCsv(parts(k).OriginalAsmLength) & "," & _
                     FormatNumberForCsv(parts(k).OriginalAsmWidth) & "," & _
                     FormatNumberForCsv(parts(k).OriginalAsmThickness) & "," & _
                     FormatNumberForCsv(parts(k).AsmCenterX) & "," & _
                     FormatNumberForCsv(parts(k).AsmCenterY) & "," & _
                     FormatNumberForCsv(parts(k).AsmCenterZ) & "," & _
                     FormatNumberForCsv(parts(k).massValue) & "," & _
                     CsvText(parts(k).filePath)
NextRow:
    Next i
End Sub

' ============================================================
' RUN THE PYROPEL FACE SCRIPT AFTER THE JOB
' ------------------------------------------------------------
' Called at the end of the job, once the POTS AND HOLDERS package has written
' its isolated DXFs. Shells out to pyro_faces.py for each of them, waits for it,
' and folds the script's own output into this log so a failure is visible here
' instead of vanishing into a closed console window.
'
' Needs, on this machine:
'   * Python 3 on PATH (or set PYRO_PYTHON_EXE to the full python.exe path)
'   * pyro_faces.py, pots_holders_split.py and dxf_split.py in PYRO_SCRIPT_FOLDER
'   * pip install ezdxf
' ============================================================


' Label line printed on each Pyropel piece. Leave blank to use the job number.

' False = fire and forget. True = wait, capture the output, log the exit code.


Private Function QuoteArg(ByVal s As String) As String
    QuoteArg = Chr(34) & s & Chr(34)
End Function


Private Sub RunPyropelFaceScriptsAfterJob(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If RUN_PYRO_AFTER_JOB = False Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim scriptPath As String
    scriptPath = PYRO_SCRIPT_FOLDER & "\pyro_faces.py"

    If fso.FileExists(scriptPath) = False Then
        LogLine "PYRO: script not found, skipping: " & scriptPath
        Exit Sub
    End If

    Dim phFolder As String
    phFolder = outputFolder & "\" & CurrentJobNumber & " " & POTS_HOLDERS_FOLDER_NAME

    If fso.FolderExists(phFolder) = False Then
        LogLine "PYRO: no POTS AND HOLDERS folder to work from: " & phFolder
        Exit Sub
    End If

    Dim pyroFolder As String
    pyroFolder = outputFolder & "\" & CurrentJobNumber & PYRO_OUTPUT_SUFFIX
    EnsureFolderDeep pyroFolder

    Dim jobsLabel As String
    jobsLabel = PYRO_JOBS_LABEL
    If Trim(jobsLabel) = "" Then jobsLabel = CurrentJobNumber

    Dim ranCount As Long
    ranCount = 0

    Dim f As Object
    For Each f In fso.GetFolder(phFolder).Files

        If LCase(fso.GetExtensionName(f.name)) = "dxf" Then

            Dim upName As String
            upName = UCase(f.name)

            Dim partName As String
            partName = ""

            If InStr(upName, "POT") > 0 Then
                partName = "POTS"
            ElseIf InStr(upName, "HOLDER") > 0 Then
                partName = "HOLDERS"
            End If

            If partName <> "" Then
                RunOnePyropelFaceScript scriptPath, f.path, partName, jobsLabel, pyroFolder
                ranCount = ranCount + 1
            End If

        End If

    Next f

    If ranCount = 0 Then
        LogLine "PYRO: no POTS/HOLDERS DXFs found in " & phFolder
    Else
        LogLine "PYRO: processed " & ranCount & " file(s) -> " & pyroFolder
    End If

    Exit Sub

ErrHandler:
    LogLine "RunPyropelFaceScriptsAfterJob error: " & Err.Description
End Sub


Private Sub RunOnePyropelFaceScript(ByVal scriptPath As String, _
                                    ByVal dxfPath As String, _
                                    ByVal partName As String, _
                                    ByVal jobsLabel As String, _
                                    ByVal outFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim logPath As String
    logPath = Environ$("TEMP") & "\CMS_PYRO_" & CleanFileName(partName) & "_" & _
              Format(Now, "yyyymmdd_hhnnss") & ".log"

    ' cmd /c needs the WHOLE command wrapped in one more pair of quotes when the
    ' individual tokens are quoted, otherwise it strips the wrong ones.
    Dim inner As String
    inner = QuoteArg(PYRO_PYTHON_EXE) & " " & _
            QuoteArg(scriptPath) & " " & _
            QuoteArg(dxfPath) & _
            " --part " & partName & _
            " --jobs " & QuoteArg(jobsLabel) & _
            " -o " & QuoteArg(outFolder) & _
            " > " & QuoteArg(logPath) & " 2>&1"

    Dim cmdLine As String
    cmdLine = "cmd /c " & QuoteArg(inner)

    LogLine "PYRO " & partName & ": " & fso.GetFileName(dxfPath)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim rc As Long
    rc = -1

    If PYRO_WAIT_FOR_FINISH Then
        rc = wsh.Run(cmdLine, 0, True)
    Else
        wsh.Run cmdLine, 0, False
        LogLine "   launched (not waiting)"
        Exit Sub
    End If

    If rc = 0 Then
        LogLine "   OK"
    Else
        LogLine "   FAILED, exit code " & rc
    End If

    ' Fold the script's output into this log.
    If fso.FileExists(logPath) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(logPath, 1, False)

        Do While ts.AtEndOfStream = False
            Dim ln As String
            ln = ts.ReadLine
            If Trim(ln) <> "" Then LogLine "   | " & ln
        Loop

        ts.Close
        On Error Resume Next
        fso.DeleteFile logPath, True
        On Error GoTo ErrHandler
    Else
        LogLine "   (no output captured - is Python on PATH?)"
    End If

    Exit Sub

ErrHandler:
    LogLine "RunOnePyropelFaceScript error (" & partName & "): " & Err.Description
End Sub

' ============================================================
' PULLCORE ID/OD FROM THE HOLDER CENTRES OF MASS
' ------------------------------------------------------------
' Replaces "pick an axis and assume ID is higher".
'
' The old method used a FIXED axis (PULLCORE_ID_OD_HEIGHT_AXIS = "Y") and a
' FIXED assumption (PULLCORE_ID_IS_HIGHER = True). When an assembly comes in
' flipped, "higher" points the wrong way and the cams come out swapped - which
' is what J8495 did.
'
' The holders are the reliable reference, because their ID/OD identity is known
' from the BOM. So:
'
'   1. Take the ID holder and OD holder centres of mass.
'   2. STANDARDISE THE DIRECTION: the ID holder must sit at positive Z. If it
'      does not, use zSign = -1 for every Z from then on. That gives one
'      consistent "up" no matter which way the assembly was built.
'   3. Assign each pullcore cam / key to whichever holder centre it is CLOSEST
'      to in standardised Z.
'   4. As a result the ID pullcore always ends up higher in standardised Z than
'      the OD one, which is the sanity check logged at the end.
' ============================================================



' Signed Z of a part's assembly centre, in the standardised direction.
Private Function StandardZ(ByVal cadIdx As Long, ByVal zSign As Double) As Double
On Error Resume Next
    StandardZ = 0#
    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    StandardZ = parts(cadIdx).AsmCenterZ * zSign
End Function


' Works out the standardised Z direction and both holder heights.
' Returns False when either holder is missing or has no centre.
Private Function TryGetHolderZReference(ByRef zSign As Double, _
                                        ByRef idHolderZ As Double, _
                                        ByRef odHolderZ As Double) As Boolean
On Error GoTo ErrHandler

    TryGetHolderZReference = False
    zSign = 1#
    idHolderZ = 0#
    odHolderZ = 0#

    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    If idIdx <= 0 Or idIdx > PartCount Then
        LogLine "Pullcore/holder Z: ID HOLDER not matched; cannot standardise direction."
        Exit Function
    End If

    If odIdx <= 0 Or odIdx > PartCount Then
        LogLine "Pullcore/holder Z: OD HOLDER not matched; cannot standardise direction."
        Exit Function
    End If

    If parts(idIdx).hasAsmCenter = False Or parts(odIdx).hasAsmCenter = False Then
        LogLine "Pullcore/holder Z: a holder has no assembly centre; cannot standardise."
        Exit Function
    End If

    Dim rawId As Double
    Dim rawOd As Double

    rawId = parts(idIdx).AsmCenterZ
    rawOd = parts(odIdx).AsmCenterZ

    ' The ID holder defines positive Z. If it is not already on the positive
    ' side of the OD holder, flip the whole axis.
    If rawId >= rawOd Then
        zSign = 1#
    Else
        zSign = -1#
    End If

    idHolderZ = rawId * zSign
    odHolderZ = rawOd * zSign

    LogLine "Pullcore/holder Z reference:"
    LogLine "   ID HOLDER " & parts(idIdx).componentName
    LogLine "      raw Z = " & FormatNumberForCsv(rawId) & _
            "   standardised Z = " & FormatNumberForCsv(idHolderZ)
    LogLine "   OD HOLDER " & parts(odIdx).componentName
    LogLine "      raw Z = " & FormatNumberForCsv(rawOd) & _
            "   standardised Z = " & FormatNumberForCsv(odHolderZ)
    LogLine "   zSign = " & FormatNumberForCsv(zSign) & _
            IIf(zSign < 0#, "  (assembly is flipped; Z inverted)", "  (assembly Z used as-is)")

    If Abs(idHolderZ - odHolderZ) < 0.001 Then
        LogLine "Pullcore/holder Z: the two holders are at the same Z; cannot separate."
        Exit Function
    End If

    TryGetHolderZReference = True
    Exit Function

ErrHandler:
    LogLine "TryGetHolderZReference error: " & Err.Description
    TryGetHolderZReference = False
End Function


' Assign one group (cams, or keys) to ID/OD by which holder centre each is
' nearest to in standardised Z.
Private Sub AssignPullcoreGroupByHolderZ(ByVal grp As Collection, _
                                         ByVal wantCam As Boolean, _
                                         ByVal zSign As Double, _
                                         ByVal idHolderZ As Double, _
                                         ByVal odHolderZ As Double)
On Error GoTo ErrHandler

    If grp Is Nothing Then Exit Sub
    If grp.count = 0 Then Exit Sub

    Dim kind As String
    If wantCam Then kind = "CAM" Else kind = "KEY"

    Dim i As Long
    Dim m As Long
    Dim z As Double
    Dim dId As Double
    Dim dOd As Double

    ' Best (nearest) candidate for each side, so a 2-part group always splits.
    Dim bestIdMatch As Long
    Dim bestOdMatch As Long
    Dim bestIdDist As Double
    Dim bestOdDist As Double

    bestIdMatch = 0
    bestOdMatch = 0
    bestIdDist = 1E+99
    bestOdDist = 1E+99

    For i = 1 To grp.count

        m = CLng(grp(i))
        z = StandardZ(PullcoreMatches(m).CadPartIndex, zSign)

        dId = Abs(z - idHolderZ)
        dOd = Abs(z - odHolderZ)

        LogLine "   " & kind & " '" & PullcoreMatches(m).quoteName & "'" & _
                "  standardised Z = " & FormatNumberForCsv(z) & _
                "   dist to ID = " & FormatNumberForCsv(dId) & _
                "   dist to OD = " & FormatNumberForCsv(dOd)

        If dId <= dOd Then
            If dId < bestIdDist Then
                bestIdDist = dId
                bestIdMatch = m
            End If
        Else
            If dOd < bestOdDist Then
                bestOdDist = dOd
                bestOdMatch = m
            End If
        End If

    Next i

    ' Everything landed on one side. Fall back to splitting the group by
    ' standardised Z: the higher one is the ID side.
    If bestIdMatch = 0 Or bestOdMatch = 0 Then

        If grp.count >= 2 Then

            Dim hiM As Long
            Dim loM As Long
            Dim hiZ As Double
            Dim loZ As Double

            hiM = 0
            loM = 0
            hiZ = -1E+99
            loZ = 1E+99

            For i = 1 To grp.count
                m = CLng(grp(i))
                z = StandardZ(PullcoreMatches(m).CadPartIndex, zSign)
                If z > hiZ Then hiZ = z: hiM = m
                If z < loZ Then loZ = z: loM = m
            Next i

            If hiM > 0 And loM > 0 And hiM <> loM Then
                bestIdMatch = hiM
                bestOdMatch = loM
                LogLine "   both " & kind & "s fell on one side; split by standardised Z instead."
            End If

        End If

    End If

    If bestIdMatch > 0 Then
        SetPullcoreMatchLabel bestIdMatch, "ID PULLCORE " & kind, "Z(holder)"
    End If

    If bestOdMatch > 0 Then
        SetPullcoreMatchLabel bestOdMatch, "OD PULLCORE " & kind, "Z(holder)"
    End If

    ' Sanity check: ID must end up higher than OD in standardised Z.
    If bestIdMatch > 0 And bestOdMatch > 0 Then

        Dim zi As Double
        Dim zo As Double

        zi = StandardZ(PullcoreMatches(bestIdMatch).CadPartIndex, zSign)
        zo = StandardZ(PullcoreMatches(bestOdMatch).CadPartIndex, zSign)

        If zi >= zo Then
            LogLine "   OK: ID " & kind & " (Z=" & FormatNumberForCsv(zi) & _
                    ") is higher than OD " & kind & " (Z=" & FormatNumberForCsv(zo) & ")"
        Else
            LogLine "   WARNING: ID " & kind & " (Z=" & FormatNumberForCsv(zi) & _
                    ") came out BELOW OD " & kind & " (Z=" & FormatNumberForCsv(zo) & _
                    "). Holder centres may be wrong for this job."
        End If

    End If

    Exit Sub

ErrHandler:
    LogLine "AssignPullcoreGroupByHolderZ error (" & kind & "): " & Err.Description
End Sub


' Primary ID/OD labeller. Returns False if the holder reference is unusable, in
' which case the caller falls back to the old height-axis method.
Private Function LabelPullcoreByHolderComZ() As Boolean
On Error GoTo ErrHandler

    LabelPullcoreByHolderComZ = False

    If LABEL_PULLCORE_FROM_HOLDER_COM = False Then Exit Function
    If PullcoreMatchCount <= 0 Then Exit Function

    Dim zSign As Double
    Dim idHolderZ As Double
    Dim odHolderZ As Double

    If TryGetHolderZReference(zSign, idHolderZ, odHolderZ) = False Then Exit Function

    Dim camGrp As Collection
    Dim keyGrp As Collection

    Set camGrp = New Collection
    Set keyGrp = New Collection

    Dim i As Long
    Dim cadIdx As Long

    For i = 1 To PullcoreMatchCount

        cadIdx = PullcoreMatches(i).CadPartIndex

        If cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).hasAsmCenter Then
                If PullcoreMatches(i).isCam Then
                    camGrp.Add i
                Else
                    keyGrp.Add i
                End If
            End If
        End If

    Next i

    If camGrp.count = 0 And keyGrp.count = 0 Then
        LogLine "Pullcore/holder Z: no pullcore matches have assembly centres."
        Exit Function
    End If

    LogLine "Pullcore ID/OD from holder centres of mass:"

    If camGrp.count > 0 Then
        AssignPullcoreGroupByHolderZ camGrp, True, zSign, idHolderZ, odHolderZ
    End If

    If keyGrp.count > 0 Then
        AssignPullcoreGroupByHolderZ keyGrp, False, zSign, idHolderZ, odHolderZ
    End If

    PullcoreIdOdHeightAxisUsed = "Z(holder COM)"

    LabelPullcoreByHolderComZ = True
    Exit Function

ErrHandler:
    LogLine "LabelPullcoreByHolderComZ error: " & Err.Description
    LabelPullcoreByHolderComZ = False
End Function

' Register a pullcore that came from the CAD naming library (no BOM row) as a
' real pullcore match, so ExportPullcoreCamKeyPackage will pick it up.
Private Sub AddPullcoreMatchFromLibrary(ByVal quoteName As String, ByVal cadIdx As Long)
On Error GoTo ErrHandler

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Sub

    Dim b As BomInfo

    b.Description = quoteName
    b.quoteName = quoteName
    b.material = ""
    b.Quantity = 1
    b.BomLength = 0#
    b.BomWidth = 0#
    b.BomThickness = 0#
    b.hasDims = False

    AddPullcoreMatchRow b, cadIdx, _
                        parts(cadIdx).Length, _
                        parts(cadIdx).Width, _
                        parts(cadIdx).Thickness

    LogLine "Pullcore registered from CAD library (no BOM row): '" & quoteName & _
            "' -> " & parts(cadIdx).componentName & _
            "  L/W/T=" & FormatNumberForCsv(parts(cadIdx).Length) & "/" & _
            FormatNumberForCsv(parts(cadIdx).Width) & "/" & _
            FormatNumberForCsv(parts(cadIdx).Thickness)

    Exit Sub

ErrHandler:
    LogLine "AddPullcoreMatchFromLibrary error (" & quoteName & "): " & Err.Description
End Sub


