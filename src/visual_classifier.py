import re


ARCHITECTURE_KEYWORDS = [
    "architecture", "layer", "component", "module", "system", "service",
    "api", "database", "middleware", "integration", "queue", "broker",
    "authentication", "authorization", "gateway", "microservice"
]

DEPLOYMENT_KEYWORDS = [
    "server", "deployment", "host", "container", "kubernetes", "cluster",
    "load balancer", "environment", "production", "staging", "vm"
]

SEQUENCE_KEYWORDS = [
    "request", "response", "user", "client", "backend", "frontend",
    "calls", "invokes", "returns", "sends", "receives", "validates"
]

COMPARISON_KEYWORDS = [
    "compare", "comparison", "difference", "versus", "vs", "whereas"
]

LIFECYCLE_KEYWORDS = [
    "lifecycle", "cycle", "phase", "stage"
]

FLOW_KEYWORDS = [
    "process", "workflow", "flow", "step", "sequence", "routing"
]

CONCEPTUAL_KEYWORDS = [
    "concept", "overview", "ecosystem", "framework", "model", "foundation"
]


def classify_visual_type(text: str) -> str:
    lower = (text or "").lower()

    def count_hits(words):
        return sum(1 for w in words if w in lower)

    scores = {
        "architecture_diagram": count_hits(ARCHITECTURE_KEYWORDS),
        "deployment_diagram": count_hits(DEPLOYMENT_KEYWORDS),
        "sequence_diagram": count_hits(SEQUENCE_KEYWORDS),
        "comparison_visual": count_hits(COMPARISON_KEYWORDS),
        "lifecycle_diagram": count_hits(LIFECYCLE_KEYWORDS),
        "flow_diagram": count_hits(FLOW_KEYWORDS),
        "concept_visual": count_hits(CONCEPTUAL_KEYWORDS),
    }

    ranked = sorted(scores.items(), key=lambda x: (-x[1], x[0]))

    if ranked[0][1] <= 0:
        return "flow_diagram"

    return ranked[0][0]
