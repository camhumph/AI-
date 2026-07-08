"""Canonical role -> display-name mapping shared by the quote UI, the pricing
engine, and the Module6121 VBA bridge export.

Kept in sync with geometry_classifier/qwen_classify_xt_csv.py ROLES and the
CMS naming rules in geometry_classifier/vendor_knowledge_sources.md /
mold_geometry_knowledge.md (A Plate / B Plate naming, Ejector Plate vs
Bottom Ejector Plate, latch_lock, etc).
"""

ROLE_LABELS = {
    "top_clamp_plate": "Top Clamp Plate",
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
    "ignore": "Ignored",
}

# Coarse grouping used to organize the Parts table / quote sheet in the UI.
ROLE_GROUPS = {
    "Mold Base Plates": {
        "top_clamp_plate", "a_plate", "b_plate", "stripper_plate",
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
    "Core / Cavity Details": {"insert_or_core_detail", "pullcore"},
    "Other Hardware": {"hardware_other"},
    "Ignored": {"ignore"},
}


def role_label(role):
    return ROLE_LABELS.get(role, role.replace("_", " ").title() if role else "Unknown")


def role_group(role):
    for group, roles in ROLE_GROUPS.items():
        if role in roles:
            return group
    return "Other Hardware"
