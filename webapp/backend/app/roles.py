"""Canonical role -> display-name mapping shared by the quote UI, the pricing
engine, and the Module6121 VBA bridge export.

Kept in sync with geometry_classifier/qwen_classify_xt_csv.py ROLES and the
CMS naming rules in geometry_classifier/vendor_knowledge_sources.md /
mold_geometry_knowledge.md (A Plate / B Plate naming, Ejector Plate vs
Bottom Ejector Plate, latch_lock, etc).
"""

# ─────────────────────────────────────────────────────────────────────────────
# THE SHOP'S OWN NAMES, taken from a real BOM.
#
# C18616 / Dynacast 2223605, `pdf/2223605-BOM.pdf`, DESCRIPTION column:
#
#     PART / DRAWING NAME          DESCRIPTION
#     2223605_A-PLATE              A PLATE
#     2223605_B-PLATE              B PLATE
#     2223605_CLAMP-PLATE          BOTTOM CLAMP PLATE
#     2223605_RAIL-BOTTOM          BOTTOM RAIL
#     2223605_EJ-BACKUP-PLATE      EJECTOR BACK-UP PLATE
#     2223605_EJ-RET-PLATE         EJECTOR RETAINER PLATE
#     2223605_RETURN-PIN           RETURN PIN
#     LDR-PIN_2OD-X-7-3-4_PCS      LEADER PIN
#     PILLAR_D2-X-4-5_T_PCS        PILLAR
#     2223605_MOLDBASE             MOLD BASE ASSY
#
# Two of those contradict what this file used to return: EJ-RET-PLATE was being
# shown as "Ejector Plate" and EJ-BACKUP-PLATE as "Bottom Ejector Plate". The
# BOM is the document the shop actually works from, so the BOM wins for DISPLAY.
#
# The role keys are unchanged, which means quote rows and prices are unchanged --
# `StdSlotForName` still routes bottom_ejector_plate to the PIN row and
# ejector_plate to the EJECTOR row exactly as before. This is a label change, the
# same separation the rename feature uses.
#
# ⚠ The BOM PDFs in these jobs are SCANNED IMAGES with no text layer, which is
# why Module6121 logs "BOM rows read: 0". Nothing can parse them automatically;
# this table is the transcription. When a job ships a BOM with real text, prefer
# its DESCRIPTION column over this table.
BOM_DESCRIPTION_LABELS = {
    "a_plate": "A Plate",
    "b_plate": "B Plate",
    "manifold_plate": "Manifold Plate",
    "bottom_clamp_plate": "Bottom Clamp Plate",
    "top_clamp_plate": "Top Clamp Plate",
    "ejector_plate": "Ejector Retainer Plate",
    "bottom_ejector_plate": "Ejector Back-Up Plate",
    "ejector_retainer_plate": "Ejector Retainer Plate",
    "ejector_backup_plate": "Ejector Back-Up Plate",
    "return_pin": "Return Pin",
    "leader_pin": "Leader Pin",
    "leader_pin_bushing": "Leader Pin Bushing",
    "guided_ejector_bushing": "Guided Ejector Bushing",
    "support_pillar": "Pillar",
    "support_plate": "Support Plate",
    "stripper_plate": "Stripper Plate",
}

# CAD token -> the BOM description for it. Used when a role is missing or too
# generic but the CAD name carries the shop's own token.
CAD_TOKEN_LABELS = [
    ("EJ-RET-PLATE", "Ejector Retainer Plate"),
    ("EJ-BACKUP-PLATE", "Ejector Back-Up Plate"),
    ("RAIL-BOTTOM", "Bottom Rail"),
    ("RAIL-TOP", "Top Rail"),
    ("CLAMP-PLATE", "Bottom Clamp Plate"),
    ("RETURN-PIN", "Return Pin"),
    ("A-PLATE", "A Plate"),
    ("B-PLATE", "B Plate"),
    ("MOLDBASE", "Mold Base Assy"),
    ("LDR-PIN", "Leader Pin"),
    ("EJ_LDR_PIN", "Secondary Leader Pin"),
    ("PILLAR", "Pillar"),
    ("LBB", "Leader Pin Bushing"),
    ("GEB", "Guided Ejector Bushing"),
    ("TIE-STRAP", "Tie Strap"),
]


def bom_label(role: str, component: str = "") -> str:
    """The name the shop's BOM would use, or "" when nothing matches.

    Role first, then the CAD token, because the role is the more reliable signal
    when it is available.
    """
    hit = BOM_DESCRIPTION_LABELS.get((role or "").strip())
    if hit:
        return hit
    upper = (component or "").upper()
    for token, label in CAD_TOKEN_LABELS:
        if token in upper:
            return label
    return ""


ROLE_LABELS = {
    "top_clamp_plate": "Top Clamp Plate",
    "manifold_plate": "Manifold Plate",
    "a_plate": "A Plate",
    "b_plate": "B Plate",
    "stripper_plate": "Stripper Plate",
    "sc_retainer_plate": "SC Retainer Plate",
    "sc_backup_plate": "SC Backup Plate",
    "support_plate": "Support Plate",
    "bottom_clamp_plate": "Bottom Clamp Plate",
    "full_footprint_plate": "Full-Footprint Plate",
    "rail": "Rail",
    "rail_1": "Rail",
    "rail_2": "Rail",
    "pin_plate": "Pin Plate",
    "ejector_plate": "Ejector Plate",
    "bottom_ejector_plate": "Bottom Ejector Plate",
    "ejector_retainer_plate": "Bottom Ejector Plate",  # deprecated alias
    "ejector_backup_plate": "Bottom Ejector Plate",
    "latch_lock": "Latch Lock / Safety Strap",
    "leader_pin": "Leader Pin",
    "leader_pin_bushing": "Leader Pin Bushing",
    "guided_ejector_bushing": "Guided Ejector Bushing",
    "return_pin": "Return Pin",
    "ejector_pin": "Ejector Pin",
    "support_pillar": "Support Pillar",
    "pullcore": "Pull Core",
    "insert_or_core_detail": "Insert / Core Detail",
    "hardware_other": "Other Hardware",
    "other_hardware": "Other Hardware",
    "purchased_component": "Purchased Component",
    "tcp": "TCP",
    "bcp": "BCP",
    "id_holder": "ID Holder",
    "od_holder": "OD Holder",
    "id_pot": "ID Pot",
    "od_pot": "OD Pot",
    "steel_plate": "Steel Plate",
    "ignore": "Ignored",
}

# Coarse grouping used to organize the Parts table / quote sheet in the UI.
ROLE_GROUPS = {
    "Steel Plates / Mold Base": {
        "top_clamp_plate", "manifold_plate", "a_plate", "b_plate", "stripper_plate",
        "sc_retainer_plate", "sc_backup_plate", "support_plate",
        "bottom_clamp_plate", "full_footprint_plate",
        "tcp", "bcp", "id_holder", "od_holder", "id_pot", "od_pot", "steel_plate",
    },
    "Mold Base Plates": {
        "top_clamp_plate", "manifold_plate", "a_plate", "b_plate", "stripper_plate",
        "sc_retainer_plate", "sc_backup_plate", "support_plate",
        "bottom_clamp_plate", "full_footprint_plate",
    },
    "Rails": {"rail", "rail_1", "rail_2"},
    "Ejector Assembly": {
        "pin_plate", "ejector_plate", "bottom_ejector_plate",
        "ejector_retainer_plate", "ejector_backup_plate",
        "return_pin", "ejector_pin",
    },
    "Latch Locks / Safety": {"latch_lock"},
    "Guide Hardware": {
        "leader_pin", "leader_pin_bushing", "guided_ejector_bushing",
        "support_pillar",
    },
    "Pull Cores & Keys": {"pullcore"},
    "Core / Cavity Details": {"insert_or_core_detail"},
    "Purchased Components": {"purchased_component"},
    "Other Hardware": {"hardware_other", "other_hardware"},
    "Ignored": {"ignore"},
}


def role_label(role):
    return ROLE_LABELS.get(role, role.replace("_", " ").title() if role else "Unknown")


def role_group(role):
    for group, roles in ROLE_GROUPS.items():
        if role in roles:
            return group
    return "Other Hardware"
