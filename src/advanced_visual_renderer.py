from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from graphviz import Digraph

from config import ADVANCED_VISUAL_DIR, ENABLE_ADVANCED_VISUALS, MAX_ADVANCED_VISUALS
from models import GeneratedVisual


def _safe_font(size=18):
    try:
        return ImageFont.truetype("DejaVuSans.ttf", size)
    except Exception:
        return ImageFont.load_default()


def _render_graphviz_visual(spec, out_base: Path):
    dot = Digraph(format="png")

    if spec.visual_type in {"architecture_diagram", "deployment_diagram"}:
        dot.attr(rankdir="TB")
        dot.attr("node", shape="box", style="rounded,filled")
    elif spec.visual_type == "sequence_diagram":
        dot.attr(rankdir="LR")
        dot.attr("node", shape="box")
    elif spec.visual_type == "comparison_visual":
        dot.attr(rankdir="TB")
        dot.attr("node", shape="box")
    elif spec.visual_type == "lifecycle_diagram":
        dot.attr(rankdir="LR")
        dot.attr("node", shape="ellipse")
    else:
        dot.attr(rankdir="LR")
        dot.attr("node", shape="box", style="rounded")

    for node in spec.nodes:
        dot.node(node.id, node.label)

    for edge in spec.edges:
        dot.edge(edge.source, edge.target, label=edge.label or "")

    dot.render(str(out_base), cleanup=True)
    return str(out_base) + ".png"


def _render_concept_visual(spec, out_base: Path):
    width, height = 1600, 900
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    title_font = _safe_font(34)
    header_font = _safe_font(24)
    body_font = _safe_font(20)

    draw.rounded_rectangle((40, 30, width - 40, 120), radius=20, outline="black", width=3)
    draw.text((70, 55), spec.title or "Concept Visual", fill="black", font=title_font)

    draw.rounded_rectangle((560, 240, 1040, 400), radius=30, outline="black", width=4)
    draw.text((600, 300), "Core Concept", fill="black", font=header_font)

    coords = [
        (120, 220, 460, 360),
        (1140, 220, 1480, 360),
        (120, 500, 460, 640),
        (1140, 500, 1480, 640),
        (560, 560, 1040, 720),
    ]

    for i, node in enumerate(spec.nodes[:5]):
        x1, y1, x2, y2 = coords[i]
        draw.rounded_rectangle((x1, y1, x2, y2), radius=24, outline="black", width=3)
        draw.text((x1 + 25, y1 + 45), node.label, fill="black", font=header_font)

    connectors = [
        ((460, 290), (560, 320)),
        ((1140, 290), (1040, 320)),
        ((460, 570), (560, 370)),
        ((1140, 570), (1040, 370)),
        ((800, 560), (800, 400)),
    ]
    for start, end in connectors[: len(spec.nodes[:5])]:
        draw.line([start, end], fill="black", width=4)

    if spec.annotations:
        draw.rounded_rectangle((80, 760, width - 80, 860), radius=18, outline="black", width=2)
        draw.text((110, 790), spec.annotations[0], fill="black", font=body_font)

    out_path = str(out_base) + ".png"
    img.save(out_path)
    return out_path


def render_advanced_visuals(state):
    if not ENABLE_ADVANCED_VISUALS:
        state.generated_visuals = []
        return state

    import shutil
    if not shutil.which("dot"):
        state.generated_visuals = []
        return state

    ADVANCED_VISUAL_DIR.mkdir(parents=True, exist_ok=True)
    generated = []

    specs = getattr(state, "visual_specs", [])[:MAX_ADVANCED_VISUALS]

    for idx, spec in enumerate(specs, start=1):
        out_base = ADVANCED_VISUAL_DIR / f"{state.base_filename}_adv_visual_{idx}"

        try:
            if spec.visual_type == "concept_visual":
                image_path = _render_concept_visual(spec, out_base)
            else:
                image_path = _render_graphviz_visual(spec, out_base)

            generated.append(
                GeneratedVisual(
                    paragraph_index=spec.paragraph_index,
                    visual_type=spec.visual_type,
                    title=spec.title,
                    image_path=image_path,
                    status="success",
                    error=None,
                )
            )
        except Exception as e:
            generated.append(
                GeneratedVisual(
                    paragraph_index=spec.paragraph_index,
                    visual_type=spec.visual_type,
                    title=spec.title,
                    image_path=None,
                    status="failed",
                    error=str(e),
                )
            )

    state.generated_visuals = generated
    return state