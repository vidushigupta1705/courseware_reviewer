import os
import re
from pathlib import Path
from graphviz import Digraph

from config import (
    GENERATED_DIAGRAM_DIR,
    ENABLE_DIAGRAM_GENERATION,
    MAX_GENERATED_DIAGRAMS,
)


def _clean_label(text: str, max_len: int = 36) -> str:
    text = re.sub(r"\s+", " ", (text or "")).strip()
    if len(text) > max_len:
        return text[: max_len - 3] + "..."
    return text


def _split_sentences(text: str):
    text = text or ""
    parts = re.split(r'(?<=[.!?])\s+|\n+', text)
    return [p.strip() for p in parts if p.strip()]


def _extract_step_chunks(text: str, max_nodes: int = 5):
    sentences = _split_sentences(text)

    # prefer lines with sequence/process indicators
    priority = []
    fallback = []

    markers = [
        "first", "second", "third", "next", "then", "finally",
        "step", "process", "flow", "stage", "phase",
        "moves", "transfers", "passes", "routing", "integration"
    ]

    for s in sentences:
        lower = s.lower()
        if any(m in lower for m in markers):
            priority.append(s)
        else:
            fallback.append(s)

    selected = priority[:max_nodes]
    if len(selected) < max_nodes:
        selected.extend(fallback[: max_nodes - len(selected)])

    return [_clean_label(x) for x in selected[:max_nodes]]


def _extract_comparison_chunks(text: str, max_nodes: int = 4):
    sentences = _split_sentences(text)
    selected = []

    markers = ["compare", "comparison", "versus", "vs", "difference", "better", "whereas", "while"]
    for s in sentences:
        if any(m in s.lower() for m in markers):
            selected.append(s)

    if not selected:
        selected = sentences[:max_nodes]

    return [_clean_label(x) for x in selected[:max_nodes]]


def _infer_diagram_type(text: str) -> str:
    lower = (text or "").lower()

    if any(x in lower for x in ["compare", "comparison", "versus", "vs", "difference", "whereas"]):
        return "comparison"

    if any(x in lower for x in ["lifecycle", "stage", "phase", "cycle"]):
        return "lifecycle"

    return "flow"


def _build_flow_diagram(text: str, out_base: Path):
    nodes = _extract_step_chunks(text, max_nodes=5)
    if len(nodes) < 2:
        nodes = ["Start", "Process", "End"]

    dot = Digraph(format="png")
    dot.attr(rankdir="LR", dpi="150")
    dot.attr("graph", fontname="Arial", fontsize="14")
    dot.attr("node", shape="box", style="rounded", fontname="Arial", fontsize="14", margin="0.3,0.2")
    dot.attr("edge", fontname="Arial", fontsize="12")

    for i, label in enumerate(nodes):
        dot.node(f"n{i}", label)

    for i in range(len(nodes) - 1):
        dot.edge(f"n{i}", f"n{i+1}")

    dot.render(str(out_base), cleanup=True)
    return str(out_base) + ".png"


def _build_lifecycle_diagram(text: str, out_base: Path):
    nodes = _extract_step_chunks(text, max_nodes=5)
    if len(nodes) < 3:
        nodes = ["Input", "Process", "Output"]

    dot = Digraph(format="png")
    dot.attr(rankdir="LR", dpi="150")
    dot.attr("graph", fontname="Arial", fontsize="14")
    dot.attr("node", shape="ellipse", fontname="Arial", fontsize="14", margin="0.3,0.2")
    dot.attr("edge", fontname="Arial", fontsize="12")

    for i, label in enumerate(nodes):
        dot.node(f"n{i}", label)

    for i in range(len(nodes) - 1):
        dot.edge(f"n{i}", f"n{i+1}")
    dot.edge(f"n{len(nodes)-1}", "n0")

    dot.render(str(out_base), cleanup=True)
    return str(out_base) + ".png"


def _build_comparison_diagram(text: str, out_base: Path):
    nodes = _extract_comparison_chunks(text, max_nodes=4)
    if len(nodes) < 2:
        nodes = ["Option A", "Comparison Point", "Option B"]

    dot = Digraph(format="png")
    dot.attr(rankdir="TB", dpi="150")
    dot.attr("graph", fontname="Arial", fontsize="14")
    dot.attr("node", shape="box", fontname="Arial", fontsize="14", margin="0.3,0.2")
    dot.attr("edge", fontname="Arial", fontsize="12")

    if len(nodes) >= 3:
        dot.node("left", nodes[0])
        dot.node("mid", nodes[1])
        dot.node("right", nodes[2])
        dot.edge("left", "mid")
        dot.edge("right", "mid")
        if len(nodes) >= 4:
            dot.node("bottom", nodes[3])
            dot.edge("mid", "bottom")
    else:
        for i, label in enumerate(nodes):
            dot.node(f"n{i}", label)
        for i in range(len(nodes) - 1):
            dot.edge(f"n{i}", f"n{i+1}")

    dot.render(str(out_base), cleanup=True)
    return str(out_base) + ".png"


def _check_graphviz_available() -> bool:
    """Check if Graphviz executable is available on the system PATH."""
    import shutil
    return shutil.which("dot") is not None


def generate_diagrams_for_recommendations(state):
    if not ENABLE_DIAGRAM_GENERATION:
        state.generated_diagrams = []
        return state

    if not _check_graphviz_available():
        print(
            "\n[WARNING] Graphviz is not installed or not on your PATH. "
            "Diagram generation will be skipped.\n"
            "To enable diagrams, install Graphviz:\n"
            "  Windows: winget install graphviz\n"
            "  Mac:     brew install graphviz\n"
            "  Linux:   sudo apt install graphviz\n"
            "Then restart your terminal and run again.\n"
        )
        state.generated_diagrams = []
        return state

    GENERATED_DIAGRAM_DIR.mkdir(parents=True, exist_ok=True)
    generated = []

    recs = getattr(state, "diagram_recommendations", [])[:MAX_GENERATED_DIAGRAMS]

    for idx, rec in enumerate(recs, start=1):
        para_idx = rec.get("paragraph_index")
        if para_idx is None or para_idx < 0 or para_idx >= len(state.paragraphs):
            continue

        text = state.paragraphs[para_idx].text or ""
        diag_type = _infer_diagram_type(text)
        out_base = GENERATED_DIAGRAM_DIR / f"{state.base_filename}_diagram_{idx}"

        try:
            if diag_type == "comparison":
                png_path = _build_comparison_diagram(text, out_base)
            elif diag_type == "lifecycle":
                png_path = _build_lifecycle_diagram(text, out_base)
            else:
                png_path = _build_flow_diagram(text, out_base)

            generated.append({
                "paragraph_index": para_idx,
                "diagram_type": diag_type,
                "image_path": png_path,
                "source_text_preview": text[:250],
            })
        except Exception as e:
            generated.append({
                "paragraph_index": para_idx,
                "diagram_type": diag_type,
                "image_path": None,
                "error": str(e),
                "source_text_preview": text[:250],
            })

    state.generated_diagrams = generated
    return state