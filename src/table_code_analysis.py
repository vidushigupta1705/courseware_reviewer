from collections import defaultdict
from rapidfuzz.fuzz import ratio

from config import TABLE_MIN_CELL_TEXT_CHARS, TABLE_WIDE_COLUMN_THRESHOLD


def analyze_tables_and_code(state):
    table_findings = []
    code_findings = []

    # Table checks
    for table in getattr(state, "tables", []):
        row_count = table.row_count
        col_count = table.col_count

        if col_count >= TABLE_WIDE_COLUMN_THRESHOLD:
            table_findings.append({
                "type": "wide_table",
                "table_index": table.table_index,
                "severity": "medium",
                "message": f"Table {table.table_index} has {col_count} columns and may need layout review.",
            })

        # repeated row detection
        row_texts = defaultdict(list)
        row_map = defaultdict(list)

        for cell in table.cells:
            row_map[cell.row_index].append(cell.text.strip().lower())

        for r_idx, parts in row_map.items():
            joined = " | ".join(parts).strip()
            if joined:
                row_texts[joined].append(r_idx)

        for row_text, row_indexes in row_texts.items():
            if len(row_indexes) > 1:
                table_findings.append({
                    "type": "repeated_table_row",
                    "table_index": table.table_index,
                    "row_indexes": row_indexes,
                    "severity": "medium",
                    "message": f"Table {table.table_index} contains repeated row content in rows {row_indexes}.",
                })

        # sparse cells
        sparse_cells = []
        for cell in table.cells:
            if len((cell.text or "").strip()) < TABLE_MIN_CELL_TEXT_CHARS:
                sparse_cells.append((cell.row_index, cell.col_index))

        if sparse_cells:
            table_findings.append({
                "type": "sparse_table_cells",
                "table_index": table.table_index,
                "severity": "low",
                "message": f"Table {table.table_index} contains sparse/empty cells at {sparse_cells[:10]}.",
            })

    # Code paragraph checks
    code_para_indexes = [p.index for p in state.paragraphs if getattr(p, "is_code_like", False)]

    for idx in code_para_indexes:
        para = state.paragraphs[idx]
        text = para.text or ""

        if len(text.split()) > 80:
            code_findings.append({
                "type": "long_code_block",
                "paragraph_index": idx,
                "severity": "low",
                "message": "This code-like paragraph is long and may need better formatting or a dedicated code block style.",
            })

        if "  " in text and "\t" in text:
            code_findings.append({
                "type": "mixed_code_indentation",
                "paragraph_index": idx,
                "severity": "medium",
                "message": "This code-like paragraph appears to mix spaces and tabs.",
            })

    state.table_findings = table_findings
    state.code_findings = code_findings
    return state