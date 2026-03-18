from models import UnitInfo
from unit_quiz_analysis import detect_units


def build_units(state):
    """
    Build strict chapter/unit objects from detect_units().
    detect_units() already ensures:
    - only chapter/unit boundaries are selected
    - topics/subtopics are excluded
    - quiz presence is checked within the full chapter span
    """
    raw_units = detect_units(state)
    units = []

    for item in raw_units:
        # hard safety checks
        if item["end_paragraph_index"] <= item["heading_paragraph_index"]:
            continue

        units.append(
            UnitInfo(
                title=item["title"],
                heading_paragraph_index=item["heading_paragraph_index"],
                heading_level=item["heading_level"],
                start_paragraph_index=item["start_paragraph_index"],
                end_paragraph_index=item["end_paragraph_index"],
                body_text=item["body_text"],
                quiz_present=item["quiz_present"],
                generated_quiz=[],
            )
        )

    state.units = units
    return state