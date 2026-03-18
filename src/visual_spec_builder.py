import re
from collections import OrderedDict

from config import VISUAL_ENTITY_LIMIT, VISUAL_RELATION_LIMIT
from models import VisualSpec, VisualNode, VisualEdge
from visual_classifier import classify_visual_type


ENTITY_PATTERNS = [
    r"\b[A-Z][A-Za-z0-9/_-]{2,}\b",
    r"\bAPI\b",
    r"\bERP\b",
    r"\bSCM\b",
    r"\bWMS\b",
    r"\bDB\b",
    r"\bETL\b",
]


def _clean_label(text: str, max_len: int = 40) -> str:
    text = re.sub(r"\s+", " ", (text or "")).strip()
    if len(text) > max_len:
        return text[: max_len - 3] + "..."
    return text


def _extract_entities(text: str):
    found = OrderedDict()
    for pattern in ENTITY_PATTERNS:
        for match in re.findall(pattern, text or ""):
            label = _clean_label(match)
            if label.lower() not in found:
                found[label.lower()] = label

    if len(found) < 3:
        words = re.findall(r"[A-Za-z][A-Za-z0-9/_-]+", text or "")
        for word in words:
            if len(word) > 4 and word.lower() not in found:
                found[word.lower()] = word
            if len(found) >= VISUAL_ENTITY_LIMIT:
                break

    return list(found.values())[:VISUAL_ENTITY_LIMIT]


def _build_linear_edges(node_ids):
    edges = []
    for i in range(len(node_ids) - 1):
        edges.append(VisualEdge(source=node_ids[i], target=node_ids[i + 1], label=""))
    return edges[:VISUAL_RELATION_LIMIT]


def _build_sequence_edges(node_ids):
    edges = []
    for i in range(len(node_ids) - 1):
        edges.append(VisualEdge(source=node_ids[i], target=node_ids[i + 1], label="request"))
        edges.append(VisualEdge(source=node_ids[i + 1], target=node_ids[i], label="response"))
    return edges[:VISUAL_RELATION_LIMIT]


def build_visual_specs(state):
    specs = []

    recommendations = getattr(state, "diagram_recommendations", [])
    for rec in recommendations:
        para_idx = rec.get("paragraph_index")
        if para_idx is None or para_idx < 0 or para_idx >= len(state.paragraphs):
            continue

        text = state.paragraphs[para_idx].text or ""
        visual_type = classify_visual_type(text)

        entities = _extract_entities(text)
        if len(entities) < 2:
            entities = ["Input", "Processing", "Output"]

        nodes = []
        node_ids = []
        for i, ent in enumerate(entities):
            node_id = f"n{i+1}"
            node_ids.append(node_id)
            nodes.append(VisualNode(id=node_id, label=ent, category="component"))

        if visual_type == "sequence_diagram":
            edges = _build_sequence_edges(node_ids)
        else:
            edges = _build_linear_edges(node_ids)

        annotations = []
        if len(text) > 120:
            annotations.append(_clean_label(text[:120]))

        spec = VisualSpec(
            paragraph_index=para_idx,
            visual_type=visual_type,
            title=_clean_label(text[:60]) or f"Visual for paragraph {para_idx}",
            nodes=nodes,
            edges=edges,
            annotations=annotations,
            detail_level="high" if len(entities) >= 5 else "medium",
        )
        specs.append(spec)

    state.visual_specs = specs
    return state
