import re
from collections import OrderedDict

from config import VISUAL_ENTITY_LIMIT, VISUAL_RELATION_LIMIT
from models import VisualSpec, VisualNode, VisualEdge
from visual_classifier import classify_visual_type


ENTITY_PATTERNS = [
    # Domain acronyms used elsewhere in this codebase's source material.
    r"\bAPI\b",
    r"\bERP\b",
    r"\bSCM\b",
    r"\bWMS\b",
    r"\bDB\b",
    r"\bETL\b",
    # Any general acronym: 2+ consecutive uppercase letters (AR, VR, MR, XR,
    # HCI, etc.) — far more reliable than "any capitalized word", which
    # matches ordinary sentence-leading words too.
    r"\b[A-Z]{2,}\b",
    # Multi-word capitalized phrases (e.g. "Augmented Reality", "Virtual
    # Reality") — a real concept name, not a single capitalized word that
    # could just be the first word of a sentence or a connector like
    # "Although" / "For".
    r"\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3}\b",
]

# Common sentence-leading / connector words that can accidentally match the
# multi-word-phrase pattern if they happen to be capitalized at a sentence
# start next to another capitalized word. Filtered out defensively even
# though the patterns above are already far more selective than a bare
# "starts with a capital letter" check.
_ENTITY_STOPWORDS = {
    "although", "however", "therefore", "moreover", "furthermore",
    "additionally", "for", "the", "this", "that", "these", "those",
    "in", "on", "at", "by", "with", "from", "as", "is", "are", "was",
    "were", "it", "its", "if", "when", "while", "since", "because",
    "thus", "hence", "also", "such", "each", "every", "some", "many",
}


def _clean_label(text: str, max_len: int = 40) -> str:
    text = re.sub(r"\s+", " ", (text or "")).strip()
    if len(text) > max_len:
        return text[: max_len - 3] + "..."
    return text


def _is_real_entity(label: str) -> bool:
    """
    Reject a candidate entity label if its first word is a common
    sentence-leading / connector word. A genuine acronym (all-uppercase,
    2+ letters) is never rejected by this check since it can't match a
    stopword. A multi-word phrase whose first word is "Although" or "For"
    is rejected — these are sentence fragments, not concept names.
    """
    first_word = label.split()[0].lower() if label.split() else ""
    return first_word not in _ENTITY_STOPWORDS


def _extract_entities(text: str):
    found = OrderedDict()
    for pattern in ENTITY_PATTERNS:
        for match in re.findall(pattern, text or ""):
            label = _clean_label(match)
            if not _is_real_entity(label):
                continue
            if label.lower() not in found:
                found[label.lower()] = label

    # NOTE: no fallback to grabbing arbitrary words from the text when
    # fewer than 3 entities are found. A document with no real acronyms
    # or multi-word concept phrases simply doesn't have enough structured
    # content for a meaningful entity diagram — the caller (build_visual_specs)
    # already falls back to generic placeholder labels ("Input", "Processing",
    # "Output") when len(entities) < 2, which is far better than fabricating
    # a diagram out of disconnected sentence fragments like "Although" / "For".

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
            # No real entities found for this paragraph — skip it rather
            # than fabricate a generic "Input / Processing / Output"
            # diagram. A placeholder diagram conveys no information about
            # the actual paragraph content, and since many unrelated
            # paragraphs can all equally fail to yield entities, they would
            # otherwise all produce the exact same generic 3-node diagram —
            # visually appearing as the same diagram repeated throughout
            # the document. A paragraph with no extractable structured
            # content simply isn't a good candidate for a diagram.
            continue

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
