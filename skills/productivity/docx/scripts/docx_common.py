#!/usr/bin/env python3
# MIT License. Shared helpers for the docx skill scripts.
"""Shared helpers: paragraph iteration and run-preserving text replacement."""
from __future__ import annotations


def iter_all_paragraphs(doc, include_headers_footers: bool = True):
    """Yield every paragraph in body, tables (recursively), headers, footers."""
    yield from _iter_container(doc)
    if include_headers_footers:
        for section in doc.sections:
            for part in (
                section.header, section.footer,
                section.first_page_header, section.first_page_footer,
                section.even_page_header, section.even_page_footer,
            ):
                if part is not None:
                    yield from _iter_container(part)


def _iter_container(container):
    yield from container.paragraphs
    for table in container.tables:
        yield from _iter_table(table)


def _iter_table(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from _iter_table(nested)


def iter_part_roots(doc):
    """Yield the XML root of the body plus every header/footer part."""
    yield doc.element.body
    seen = set()
    for section in doc.sections:
        for part in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ):
            if part is not None and id(part._element) not in seen:
                seen.add(id(part._element))
                yield part._element


def ensure_comment_styles(doc) -> list:
    """Define the comment styles the document is missing; return their ids.

    Adding a comment writes `w:rStyle w:val="CommentReference"`, but the
    python-docx default template never defines that style, so the package
    ends up referencing an undefined style id. Word tolerates it (latent
    built-in), but stricter consumers and `docx_validate.py` flag it, so
    define the standard Word definitions before saving.
    """
    from docx.enum.style import WD_STYLE_TYPE
    from docx.shared import Pt

    wanted = [
        ("Comment Reference", WD_STYLE_TYPE.CHARACTER, 8, False),
        ("Comment Text", WD_STYLE_TYPE.PARAGRAPH, 9, True),
    ]
    defined = {s.style_id for s in doc.styles}
    added = []
    for name, style_type, half_pt_pt, space_after_zero in wanted:
        style_id = name.replace(" ", "")
        if style_id in defined:
            continue
        try:
            style = doc.styles.add_style(name, style_type, builtin=True)
        except ValueError:  # already present under a different id
            continue
        style.font.size = Pt(half_pt_pt)
        if space_after_zero:
            style.paragraph_format.space_after = Pt(0)
        added.append(style_id)
    return added


def replace_in_paragraph(para, old: str, new: str) -> int:
    """Replace `old` with `new` in a paragraph, preserving run formatting.

    Strategy: first replace occurrences fully contained in a single run
    (formatting fully preserved). If the needle spans multiple runs, the
    matched runs are collapsed: the replacement inherits the formatting of
    the run where the match starts. Returns number of replacements made.
    """
    if not old or old not in para.text:
        return 0
    count = 0
    # Pass 1: within-run replacements.
    for run in para.runs:
        if old in run.text:
            count += run.text.count(old)
            run.text = run.text.replace(old, new)
    # Pass 2: cross-run occurrences.
    while old in para.text:
        runs = para.runs
        # Map paragraph text offsets to (run_index, offset_in_run).
        full = "".join(r.text for r in runs)
        start = full.find(old)
        if start < 0:
            break
        end = start + len(old)
        pos = 0
        spans = []  # (run_idx, cut_start, cut_end) portions inside the match
        for i, r in enumerate(runs):
            r_start, r_end = pos, pos + len(r.text)
            if r_end > start and r_start < end:
                spans.append((i, max(start, r_start) - r_start,
                              min(end, r_end) - r_start))
            pos = r_end
        first = True
        for i, cs, ce in spans:
            t = runs[i].text
            if first:
                runs[i].text = t[:cs] + new + t[ce:]
                first = False
            else:
                runs[i].text = t[:cs] + t[ce:]
        count += 1
    return count
