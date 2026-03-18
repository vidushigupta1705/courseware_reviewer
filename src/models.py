from dataclasses import dataclass, field, asdict
from typing import List, Optional


@dataclass
class ParagraphInfo:
    index: int
    text: str
    style_name: str
    is_heading: bool
    heading_level: Optional[int]
    font_names: List[str]
    font_sizes: List[float]
    bold_flags: List[bool]
    alignment: Optional[str]
    has_extra_spacing_issue: bool = False
    is_code_like: bool = False


@dataclass
class TableCellInfo:
    row_index: int
    col_index: int
    text: str


@dataclass
class TableInfo:
    table_index: int
    row_count: int
    col_count: int
    cells: List[TableCellInfo] = field(default_factory=list)
    paragraph_index: Optional[int] = None


@dataclass
class ImageInfo:
    rel_id: str
    filename: str
    output_path: str
    paragraph_index: Optional[int] = None
    image_type: Optional[str] = None
    ocr_text: Optional[str] = None
    ocr_confidence: Optional[float] = None
    ocr_status: Optional[str] = None
    nearest_caption_paragraph_index: Optional[int] = None
    nearest_caption_text: Optional[str] = None
    has_figure_number: Optional[bool] = None
    figure_number_text: Optional[str] = None
    has_source_link: Optional[bool] = None
    source_link_text: Optional[str] = None


@dataclass
class OCRLine:
    text: str
    confidence: float


@dataclass
class OCRResult:
    image_filename: str
    image_path: str
    image_type: str
    lines: List[OCRLine] = field(default_factory=list)
    merged_text: str = ""
    avg_confidence: Optional[float] = None
    status: str = "not_run"


@dataclass
class RewriteSuggestion:
    paragraph_index: int
    reason: str
    original_text: str
    rewritten_text: str
    applied: bool = False
    skipped: bool = False
    skip_reason: Optional[str] = None


@dataclass
class QuizItem:
    qtype: str
    question: str
    options: List[str] = field(default_factory=list)
    answer: Optional[str] = None


@dataclass
class UnitInfo:
    title: str
    heading_paragraph_index: int
    heading_level: int
    start_paragraph_index: int
    end_paragraph_index: int
    body_text: str
    quiz_present: bool = False
    generated_quiz: List[QuizItem] = field(default_factory=list)


@dataclass
class VisualNode:
    id: str
    label: str
    category: str = "component"
    group: Optional[str] = None


@dataclass
class VisualEdge:
    source: str
    target: str
    label: str = ""


@dataclass
class VisualSpec:
    paragraph_index: int
    visual_type: str
    title: str
    nodes: List[VisualNode] = field(default_factory=list)
    edges: List[VisualEdge] = field(default_factory=list)
    annotations: List[str] = field(default_factory=list)
    detail_level: str = "medium"


@dataclass
class GeneratedVisual:
    paragraph_index: int
    visual_type: str
    title: str
    image_path: Optional[str] = None
    status: str = "pending"
    error: Optional[str] = None


@dataclass
class AccuracyFinding:
    unit_title: str
    paragraph_index: int
    issue_description: str
    flagged_text: str
    suggestion: str


@dataclass
class ReviewIssue:
    issue_type: str
    severity: str
    paragraph_index: Optional[int]
    location: str
    message: str
    suggested_fix: str


@dataclass
class DocumentState:
    input_file: str
    base_filename: str
    paragraphs: List[ParagraphInfo] = field(default_factory=list)
    tables: List[TableInfo] = field(default_factory=list)
    images: List[ImageInfo] = field(default_factory=list)
    ocr_results: List[OCRResult] = field(default_factory=list)
    rewrite_suggestions: List[RewriteSuggestion] = field(default_factory=list)
    units: List[UnitInfo] = field(default_factory=list)
    visual_specs: List[VisualSpec] = field(default_factory=list)
    generated_visuals: List[GeneratedVisual] = field(default_factory=list)
    accuracy_findings: List[AccuracyFinding] = field(default_factory=list)
    issues: List[ReviewIssue] = field(default_factory=list)

    def to_dict(self):
        return asdict(self)