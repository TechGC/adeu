# FILE: src/adeu/redline/mapper.py
import re
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, List, Optional, Tuple

import structlog
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from adeu.domain import build_structural_appendix
from adeu.redline.comments import CommentsManager
from adeu.utils.docx import (
    DocxEvent,
    get_paragraph_prefix,
    get_run_style_markers,
    get_run_text,
    iter_block_items,
    iter_document_parts,
    iter_paragraph_content,
)

logger = structlog.get_logger(__name__)


@dataclass
class TextSpan:
    start: int
    end: int
    text: str
    run: Optional[Run]
    paragraph: Optional[Paragraph]
    ins_id: Optional[str] = None
    del_id: Optional[str] = None
    hyperlink_id: Optional[str] = None


class DocumentMapper:
    _FIELD_MARKUP_RE = re.compile(r"\[~(?P<xref>.*?)~\]\(#[^)]*\)|\{#[^}]*\}")

    def __init__(self, doc: DocumentObject, clean_view: bool = False):
        self.doc = doc
        self.clean_view = clean_view
        self.comments_mgr = CommentsManager(doc)
        self.comments_map = self.comments_mgr.extract_comments_data()
        self.full_text = ""
        self.spans: List[TextSpan] = []
        self.appendix_start_index: int = -1
        self._build_map()

    def _build_map(self):
        current_offset = 0
        self.spans = []
        self.full_text = ""

        for part in iter_document_parts(self.doc):
            current_offset = self._map_blocks(part, current_offset)

            # Add part separator if needed, or rely on block separators
            if self.spans and self.spans[-1].text != "\n\n":
                self._add_virtual_text("\n\n", current_offset, None)
                current_offset += 2

        # Cleanup trailing newlines
        while self.spans and self.spans[-1].text == "\n\n":
            self.spans.pop()
            self.full_text = self.full_text[:-2]

        appendix_text = build_structural_appendix(self.doc, self.full_text)
        if appendix_text:
            self.appendix_start_index = len(self.full_text)
            self._add_virtual_text(appendix_text, self.appendix_start_index, None)

    def _map_blocks(self, container, offset: int) -> int:
        current = offset
        c_type = type(container).__name__

        if c_type == "NotesPart":
            header = "## Footnotes" if container.note_type == "fn" else "## Endnotes"
            sep = f"---\n{header}"
            self._add_virtual_text(sep, current, None)
            current += len(sep)
            self._add_virtual_text("\n\n", current, None)
            current += 2

        is_first_para = True

        previous_item: Any = None
        for item in iter_block_items(container):
            i_type = type(item).__name__

            if i_type == "FootnoteItem":
                current = self._map_blocks(item, current)
            elif isinstance(item, Paragraph):
                if not is_first_para:
                    # Attach the newline to the previous paragraph so merges work correctly
                    prev_para = previous_item if isinstance(previous_item, Paragraph) else None
                    self._add_virtual_text("\n\n", current, prev_para)
                    current += 2

                prefix = get_paragraph_prefix(item)
                if is_first_para and c_type == "FootnoteItem":
                    prefix = f"[^{container.note_type}-{container.id}]: " + prefix
                if prefix:
                    self._add_virtual_text(prefix, current, item)
                    current += len(prefix)

                current = self._map_paragraph_content(item, current)
                is_first_para = False
                previous_item = item
            elif isinstance(item, Table):
                if not is_first_para:
                    # Attach the newline to the previous paragraph so merges work correctly
                    prev_para = previous_item if isinstance(previous_item, Paragraph) else None
                    self._add_virtual_text("\n\n", current, prev_para)
                    current += 2

                current = self._map_table(item, current)
                is_first_para = False
                previous_item = item

        return current

    def _map_table(self, table: Table, offset: int) -> int:
        current = offset
        rows_processed = 0

        for row in table.rows:
            # Structural Row Tracking
            tr = row._element
            trPr = tr.find(qn("w:trPr"))
            ins = trPr.find(qn("w:ins")) if trPr is not None else None
            del_node = trPr.find(qn("w:del")) if trPr is not None else None

            if self.clean_view and del_node is not None:
                continue

            if rows_processed > 0:
                # Newline separator BETWEEN rows (matches "\n".join in ingest)
                self._add_virtual_text("\n", current, None)
                current += 1

            if ins is not None and not self.clean_view:
                self._add_virtual_text("{++ ", current, None)
                current += 4
            elif del_node is not None and not self.clean_view:
                self._add_virtual_text("{-- ", current, None)
                current += 4

            seen_cells = set()
            cells_processed = 0

            for cell in row.cells:
                if cell in seen_cells:
                    continue
                seen_cells.add(cell)

                if cells_processed > 0:
                    self._add_virtual_text(" | ", current, None)
                    current += 3

                current = self._map_blocks(cell, current)
                cells_processed += 1

            if ins is not None and not self.clean_view:
                suffix = f" |Chg:{ins.get(qn('w:id'))}++}}"
                self._add_virtual_text(suffix, current, None)
                current += len(suffix)
            elif del_node is not None and not self.clean_view:
                suffix = f" |Chg:{del_node.get(qn('w:id'))}--}}"
                self._add_virtual_text(suffix, current, None)
                current += len(suffix)

            rows_processed += 1

        return current

    def _strip_markdown_formatting(self, text: str) -> str:
        """
        Strips markdown formatting markers from text for matching purposes.
        Handles: **bold**, __bold__, _italic_, *italic*, # headers
        Only strips when content looks like actual formatted text (2+ word chars).
        """
        result = text

        # Strip header markers at start of lines
        result = re.sub(r"^#+\s*", "", result, flags=re.MULTILINE)

        # Strip bold markers - only when wrapping word content (not single chars)
        result = re.sub(r"\*\*(\w[\w\s]*\w|\w{2,})\*\*", r"\1", result)
        result = re.sub(r"__(\w[\w\s]*\w|\w{2,})__", r"\1", result)

        # Strip italic markers - only when wrapping word content
        result = re.sub(r"(?<!\w)_(\w[\w\s]*\w|\w{2,})_(?!\w)", r"\1", result)
        result = re.sub(r"(?<!\w)\*(\w[\w\s]*\w|\w{2,})\*(?!\w)", r"\1", result)

        return result

    def _map_paragraph_content(self, paragraph: Paragraph, start_offset: int) -> int:
        """
        Maps Runs to Spans, handling Flattened CriticMarkup generation.
        """
        current = start_offset

        # Inject zero-length paragraph anchor so empty paragraphs/cells are still map-able
        span = TextSpan(start=current, end=current, text="", run=None, paragraph=paragraph)
        self.spans.append(span)

        active_ids: set[str] = set()
        active_ins: dict[str, DocxEvent] = {}
        active_del: dict[str, DocxEvent] = {}

        deferred_meta_states: List[Tuple] = []
        current_wrappers = ("", "")
        current_style = ("", "")  # Trailing style in pending_runs (for elision)
        active_hyperlink_id = None
        pending_runs: List[Tuple[str, str, Optional[Run], Optional[str], Optional[str]]] = []

        def flush_pending_runs():
            """Emits pending_runs + wrappers as spans. Resets pending_runs."""
            nonlocal current, pending_runs
            if not pending_runs:
                return
            s_tok, e_tok = current_wrappers
            if s_tok:
                self._add_virtual_text(s_tok, current, paragraph)
                current += len(s_tok)
            for kind, txt, r_obj, i_id, d_id in pending_runs:
                if kind == "virtual":
                    self._add_virtual_text(txt, current, paragraph, hyperlink_id=active_hyperlink_id)
                else:
                    span = TextSpan(
                        start=current,
                        end=current + len(txt),
                        text=txt,
                        run=r_obj,
                        paragraph=paragraph,
                        ins_id=i_id,
                        del_id=d_id,
                        hyperlink_id=active_hyperlink_id,
                    )
                    self.spans.append(span)
                    self.full_text += txt
                current += len(txt)
            if e_tok:
                self._add_virtual_text(e_tok, current, paragraph)
                current += len(e_tok)
            pending_runs = []

        items = list(iter_paragraph_content(paragraph))

        for i, item in enumerate(items):
            if isinstance(item, Run):
                prefix, suffix = get_run_style_markers(item)
                run_parts: List[Tuple[str, str, Optional[Run]]] = []

                text = get_run_text(item)

                if "\n" in text and (prefix or suffix):
                    parts = text.split("\n")
                    for idx, part in enumerate(parts):
                        if idx > 0:
                            run_parts.append(("real", "\n", item))
                        if part:
                            if prefix:
                                run_parts.append(("virtual", prefix, None))
                            run_parts.append(("real", part, item))
                            if suffix:
                                run_parts.append(("virtual", suffix, None))
                else:
                    if prefix:
                        run_parts.append(("virtual", prefix, None))
                    if text:
                        run_parts.append(("real", text, item))
                    if suffix:
                        run_parts.append(("virtual", suffix, None))

                if self.clean_view and active_del:
                    pass

                full_seg_text = "".join(x[1] for x in run_parts)

                curr_ins_id = list(active_ins.keys())[-1] if active_ins else None
                curr_del_id = list(active_del.keys())[-1] if active_del else None

                if full_seg_text and not (self.clean_view and curr_del_id):
                    if self.clean_view:
                        new_wrappers = ("", "")
                    else:
                        start_token, end_token = self._get_wrappers(curr_ins_id, curr_del_id, active_ids)
                        new_wrappers = (start_token, end_token)
                    new_style = (prefix, suffix)

                    if pending_runs and new_wrappers == current_wrappers:
                        # MERGE: same redline wrapper group.
                        skip_leading_prefix = False
                        if (
                            new_style == current_style
                            and current_style != ("", "")
                            and pending_runs
                            and pending_runs[-1][0] == "virtual"
                            and pending_runs[-1][1] == current_style[1]
                        ):
                            pending_runs.pop()
                            skip_leading_prefix = True

                        for kind, txt, r_obj in run_parts:
                            if skip_leading_prefix and kind == "virtual" and txt == new_style[0]:
                                skip_leading_prefix = False
                                continue
                            pending_runs.append((kind, txt, r_obj, curr_ins_id, curr_del_id))

                        current_style = new_style
                    else:
                        # FLUSH and open new wrapper group.
                        flush_pending_runs()
                        current_wrappers = new_wrappers
                        current_style = new_style
                        for kind, txt, r_obj in run_parts:
                            pending_runs.append((kind, txt, r_obj, curr_ins_id, curr_del_id))

                # Metadata Handling
                if not self.clean_view:
                    state_snapshot = (
                        active_ins.copy(),
                        active_del.copy(),
                        active_ids.copy(),
                    )
                    deferred_meta_states.append(state_snapshot)

                    should_defer = False
                    is_redline = bool(curr_ins_id) or bool(curr_del_id)

                    if is_redline:
                        j = i + 1
                        next_is_redline = False
                        temp_ins_count = len(active_ins)
                        temp_del_count = len(active_del)

                        while j < len(items):
                            next_item = items[j]
                            if isinstance(next_item, Run):
                                if not get_run_text(next_item):
                                    j += 1
                                    continue
                                if temp_ins_count > 0 or temp_del_count > 0:
                                    next_is_redline = True
                                break
                            elif isinstance(next_item, DocxEvent):
                                if next_item.type == "ins_start":
                                    temp_ins_count += 1
                                elif next_item.type == "ins_end":
                                    temp_ins_count = max(0, temp_ins_count - 1)
                                elif next_item.type == "del_start":
                                    temp_del_count += 1
                                elif next_item.type == "del_end":
                                    temp_del_count = max(0, temp_del_count - 1)
                            j += 1

                        if next_is_redline:
                            should_defer = True

                    if not should_defer:
                        meta_block = self._build_merged_meta_block(deferred_meta_states)
                        if meta_block:
                            flush_pending_runs()
                            current_wrappers = ("", "")
                            current_style = ("", "")
                            full_meta = f"{{>>{meta_block}<<}}"
                            self._add_virtual_text(full_meta, current, paragraph)
                            current += len(full_meta)
                        deferred_meta_states = []

            elif isinstance(item, DocxEvent):
                flush_pending_runs()
                current_wrappers = ("", "")
                current_style = ("", "")

                if item.type == "start":
                    active_ids.add(item.id)
                elif item.type == "end":
                    if item.id in active_ids:
                        active_ids.remove(item.id)
                elif item.type == "ins_start":
                    active_ins[item.id] = item
                elif item.type == "ins_end":
                    active_ins.pop(item.id, None)
                elif item.type == "del_start":
                    active_del[item.id] = item
                elif item.type == "del_end":
                    active_del.pop(item.id, None)
                elif item.type in ("footnote", "endnote"):
                    flush_pending_runs()
                    current_wrappers = ("", "")
                    current_style = ("", "")
                    prefix_str = "fn" if item.type == "footnote" else "en"
                    txt = f"[^{prefix_str}-{item.id}]"
                    self._add_virtual_text(txt, current, paragraph)
                    current += len(txt)
                elif item.type == "hyperlink_start":
                    flush_pending_runs()
                    current_wrappers = ("", "")
                    current_style = ("", "")
                    self._add_virtual_text("[", current, paragraph, hyperlink_id=item.id)
                    current += 1
                    active_hyperlink_id = item.id
                elif item.type == "hyperlink_end":
                    flush_pending_runs()
                    current_wrappers = ("", "")
                    current_style = ("", "")
                    txt = f"]({item.date})"
                    self._add_virtual_text(txt, current, paragraph, hyperlink_id=item.id)
                    current += len(txt)
                    active_hyperlink_id = None
                elif item.type == "xref_start":
                    flush_pending_runs()
                    current_wrappers = ("", "")
                    current_style = ("", "")
                    self._add_virtual_text("[~", current, paragraph)
                    current += 2
                elif item.type == "xref_end":
                    flush_pending_runs()
                    current_wrappers = ("", "")
                    current_style = ("", "")
                    txt = f"~](#{item.id})"
                    self._add_virtual_text(txt, current, paragraph)
                    current += len(txt)
                elif item.type == "bookmark":
                    flush_pending_runs()
                    current_wrappers = ("", "")
                    current_style = ("", "")
                    txt = f"{{#{item.id}}}"
                    self._add_virtual_text(txt, current, paragraph)
                    current += len(txt)

        flush_pending_runs()

        if deferred_meta_states:
            meta_block = self._build_merged_meta_block(deferred_meta_states)
            if meta_block:
                full_meta = f"{{>>{meta_block}<<}}"
                self._add_virtual_text(full_meta, current, paragraph)
                current += len(full_meta)

        return current

    def _get_wrappers(self, ins_id, del_id, active_ids):
        if del_id:
            return "{--", "--}"
        elif ins_id:
            return "{++", "++}"
        elif active_ids:
            return "{==", "==}"
        return "", ""

    def _build_merged_meta_block(self, states_list) -> str:
        change_lines = []
        comment_lines = []
        seen_sigs = set()

        for ins_map, del_map, comments_set in states_list:
            for map_obj in (ins_map, del_map):
                for uid, meta in map_obj.items():
                    sig = f"Chg:{uid}"
                    if sig not in seen_sigs:
                        auth = meta.author or "Unknown"
                        change_lines.append(f"[{sig}] {auth}")
                        seen_sigs.add(sig)

            sorted_ids = sorted(list(comments_set))
            for c_id in sorted_ids:
                if c_id not in self.comments_map:
                    continue
                sig = f"Com:{c_id}"
                if sig not in seen_sigs:
                    data = self.comments_map[c_id]
                    header = f"[{sig}] {data['author']}"
                    if data["date"]:
                        header += f" @ {data['date']}"
                    if data["resolved"]:
                        header += "(RESOLVED)"
                    comment_lines.append(f"{header}: {data['text']}")
                    seen_sigs.add(sig)

        return "\n".join(change_lines + comment_lines)

    def _add_virtual_text(
        self,
        text: str,
        offset: int,
        context_paragraph: Optional[Paragraph],
        hyperlink_id: Optional[str] = None,
    ):
        span = TextSpan(
            start=offset,
            end=offset + len(text),
            text=text,
            run=None,  # Virtual
            paragraph=context_paragraph,
            hyperlink_id=hyperlink_id,
        )
        self.spans.append(span)
        self.full_text += text

    def _replace_smart_quotes(self, text: str) -> str:
        return text.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")

    def _build_field_visible_projection(self, text: str) -> Tuple[str, List[int]]:
        """Projects Word field markup to its visible text, keeping an index map back to `text`.

        Cross-reference fields are emitted as ``[~10.16~](#_Ref…)`` and bookmark
        anchors as ``{#_Ref…}``. An LLM-driven caller anchors on the *visible*
        text (``10.16``; bookmarks render nothing), so matching that anchor
        against the raw markup fails. Returns ``(visible, index_map)`` where
        ``visible`` is the field-collapsed projection and ``index_map[i]`` is the
        offset in ``text`` of visible char ``i`` (with a trailing sentinel
        ``index_map[len(visible)] == len(text)``), so a match found in the
        projection can be mapped back to real ``full_text`` offsets.
        """
        visible_chars: List[str] = []
        index_map: List[int] = []
        last = 0
        for mo in self._FIELD_MARKUP_RE.finditer(text):
            for i in range(last, mo.start()):
                visible_chars.append(text[i])
                index_map.append(i)
            xref = mo.group("xref")
            if xref:
                inner_start = mo.start() + 2  # past the leading "[~"
                for k, ch in enumerate(xref):
                    visible_chars.append(ch)
                    index_map.append(inner_start + k)
            last = mo.end()
        for i in range(last, len(text)):
            visible_chars.append(text[i])
            index_map.append(i)
        index_map.append(len(text))
        return "".join(visible_chars), index_map

    def _map_visible_span_to_full(self, v_start: int, v_len: int, index_map: List[int]) -> Tuple[int, int]:
        """Maps a (start, length) span in a field-visible projection back to raw ``full_text`` offsets."""
        orig_start = index_map[v_start]
        if v_len <= 0:
            return orig_start, 0
        orig_end = index_map[v_start + v_len - 1] + 1
        return orig_start, orig_end - orig_start

    def _find_index_in(self, target_text: str, haystack: str) -> Tuple[int, int]:
        """Runs the exact → smart-quote → markdown-strip → fuzzy-regex ladder against `haystack`."""
        start_idx = haystack.find(target_text)
        if start_idx != -1:
            return start_idx, len(target_text)

        norm_full = self._replace_smart_quotes(haystack)
        norm_target = self._replace_smart_quotes(target_text)
        start_idx = norm_full.find(norm_target)
        if start_idx != -1:
            return start_idx, len(target_text)

        stripped_target = self._strip_markdown_formatting(target_text)
        if stripped_target in haystack:
            return haystack.find(stripped_target), len(stripped_target)

        try:
            pattern = self._make_fuzzy_regex(target_text)
            match = re.search(pattern, haystack)
            if match:
                return match.start(), match.end() - match.start()
        except re.error:
            pass

        return -1, 0

    def _find_all_indices_in(self, target_text: str, haystack: str) -> List[Tuple[int, int]]:
        """Runs the exact → smart-quote → markdown-strip → fuzzy-regex ladder against `haystack`, all matches."""
        matches = [m.span() for m in re.finditer(re.escape(target_text), haystack)]
        if matches:
            return [(s, e - s) for s, e in matches]

        norm_full = self._replace_smart_quotes(haystack)
        norm_target = self._replace_smart_quotes(target_text)
        matches = [m.span() for m in re.finditer(re.escape(norm_target), norm_full)]
        if matches:
            return [(s, e - s) for s, e in matches]

        stripped_target = self._strip_markdown_formatting(target_text)
        matches = [m.span() for m in re.finditer(re.escape(stripped_target), haystack)]
        if matches:
            return [(s, e - s) for s, e in matches]

        try:
            pattern = self._make_fuzzy_regex(target_text)
            matches = [m.span() for m in re.finditer(pattern, haystack)]
            if matches:
                return [(s, e - s) for s, e in matches]
        except re.error:
            pass

        return []

    def _make_fuzzy_regex(self, target_text: str) -> str:
        """
        Constructs a regex from target text permitting variable whitespace,
        variable underscores in placeholders, smart quotes, intervening
        markdown markers, and punctuation boundaries.
        """
        target_text = self._strip_markdown_formatting(target_text)
        target_text = self._replace_smart_quotes(target_text)

        parts = []
        token_pattern = re.compile(r"(\[_+\])|(\s+)|(['\"])|([.,;:])")

        last_idx = 0
        for match in token_pattern.finditer(target_text):
            literal = target_text[last_idx : match.start()]
            if literal:
                escaped = re.escape(literal)
                parts.append(escaped)

            g_placeholder, g_space, g_quote, g_punct = match.groups()

            if g_placeholder:
                parts.append(r"\[_+\]")
            elif g_space:
                parts.append(r"(?>\*\*|__|\*|_)?")
                parts.append(r"\s+")
                parts.append(r"(?>\*\*|__|\*|_)?")
            elif g_quote:
                if g_quote == "'":
                    parts.append(r"[\u2018\u2019']")
                else:
                    parts.append(r"[\"\u201c\u201d]")
            elif g_punct:
                parts.append(r"(?>\*\*|__|\*|_)?")
                parts.append(re.escape(g_punct))
                parts.append(r"(?>\*\*|__|\*|_)?")

            last_idx = match.end()

        remaining = target_text[last_idx:]
        if remaining:
            parts.append(re.escape(remaining))

        return "".join(parts)

    def find_match_index(self, target_text: str) -> Tuple[int, int]:
        """
        Returns (start_index, match_length).
        Returns (-1, 0) if not found.
        """
        start_idx, length = self._find_index_in(target_text, self.full_text)
        if start_idx != -1:
            return start_idx, length

        # Field-markup fallback: callers anchor on visible text, but Word cross-ref
        # fields/bookmarks inject CriticMarkup ([~N~](#_Ref…), {#_Ref…}) into full_text,
        # so the raw ladder above misses. Match against the field-collapsed projection
        # and map the hit back to real full_text offsets (markup-inclusive span).
        visible, index_map = self._build_field_visible_projection(self.full_text)
        if len(visible) != len(self.full_text):
            v_start, v_len = self._find_index_in(target_text, visible)
            if v_start != -1:
                return self._map_visible_span_to_full(v_start, v_len, index_map)

        return -1, 0

    def find_all_match_indices(self, target_text: str) -> List[Tuple[int, int]]:
        """
        Returns a list of all non-overlapping matches as (start_index, match_length).
        Returns an empty list if not found.
        """
        if not target_text:
            return []

        matches = self._find_all_indices_in(target_text, self.full_text)
        if matches:
            return matches

        # Field-markup fallback — see find_match_index for rationale.
        visible, index_map = self._build_field_visible_projection(self.full_text)
        if len(visible) != len(self.full_text):
            v_matches = self._find_all_indices_in(target_text, visible)
            if v_matches:
                return [self._map_visible_span_to_full(s, length, index_map) for s, length in v_matches]

        return []

    def find_target_runs(self, target_text: str) -> List[Run]:
        start_idx, length = self.find_match_index(target_text)
        if start_idx == -1:
            return []
        return self._resolve_runs_at_range(start_idx, start_idx + length)

    def find_target_runs_by_index(self, start_index: int, length: int, rebuild_map: bool = True) -> List[Run]:
        end_index = start_index + length
        return self._resolve_runs_at_range(start_index, end_index, rebuild_map=rebuild_map)

    def get_virtual_spans_in_range(self, start_index: int, length: int) -> List[TextSpan]:
        """
        Returns any virtual spans (run is None) that fall completely within the
        provided range. Used primarily for detecting deleted paragraph boundaries.
        """
        end_index = start_index + length
        return [
            s
            for s in self.spans
            if s.run is None and s.text == "\n\n" and s.start >= start_index and s.end <= end_index
        ]

    def _resolve_runs_at_range(self, start_idx: int, end_idx: int, rebuild_map: bool = True) -> List[Run]:
        affected_spans = [s for s in self.spans if s.end > start_idx and s.start < end_idx]
        if not affected_spans:
            return []

        working_runs = [s.run for s in affected_spans if s.run is not None]
        if not working_runs:
            return []

        dom_modified = False

        # 1. Start Split
        first_real_span = next((s for s in affected_spans if s.run is not None), None)
        start_split_adjustment = 0

        if first_real_span:
            local_start = start_idx - first_real_span.start
            if local_start > 0:
                idx_in_working = 0
                _, right_run = self._split_run_at_index(working_runs[idx_in_working], local_start)
                working_runs[idx_in_working] = right_run
                dom_modified = True
                start_split_adjustment = local_start

        # 2. End Split
        last_real_span = next((s for s in reversed(affected_spans) if s.run is not None), None)

        if last_real_span:
            is_same_run = first_real_span is last_real_span
            run_to_split = working_runs[-1]
            overlap_end = min(last_real_span.end, end_idx)
            local_end = overlap_end - last_real_span.start

            if is_same_run and start_split_adjustment > 0:
                local_end -= start_split_adjustment

            if 0 < local_end < len(run_to_split.text):
                left_run, _ = self._split_run_at_index(run_to_split, local_end)
                working_runs[-1] = left_run
                dom_modified = True

        if dom_modified and rebuild_map:
            self._build_map()

        return working_runs

    def get_insertion_anchor(self, index: int, rebuild_map: bool = True) -> Tuple[Optional[Run], Optional[Paragraph]]:
        preceding = [s for s in self.spans if s.end == index]
        if preceding:
            for s in reversed(preceding):
                if s.run:
                    return s.run, s.paragraph
            for s in reversed(preceding):
                if s.paragraph:
                    return None, s.paragraph

        containing = [s for s in self.spans if s.start < index < s.end]
        if containing:
            span = containing[0]
            if span.run is None:
                if span.paragraph is None:
                    # We are inside a virtual string (like " | " or "\n").
                    # Push the insertion point to the end of this virtual boundary.
                    return self.get_insertion_anchor(span.end, rebuild_map=False)
                return None, span.paragraph
            else:
                offset = index - span.start
                left, _ = self._split_run_at_index(span.run, offset)
                if rebuild_map:
                    self._build_map()
                return left, span.paragraph

        if index == 0 and self.spans:
            for s in self.spans:
                if s.run:
                    return s.run, s.paragraph
            for s in self.spans:
                if s.paragraph:
                    return None, s.paragraph
            return None, None

        preceding_gap = [s for s in self.spans if s.end < index]
        if preceding_gap:
            for s in reversed(preceding_gap):
                if s.run:
                    return s.run, s.paragraph
            for s in reversed(preceding_gap):
                if s.paragraph:
                    return None, s.paragraph
        return None, None

    def _split_run_at_index(self, run: Run, split_index: int) -> Tuple[Run, Run]:
        text = run.text
        left_text = text[:split_index]
        right_text = text[split_index:]

        run.text = left_text
        new_r_element = deepcopy(run._element)
        t_list = new_r_element.findall(qn("w:t"))
        for t in t_list:
            new_r_element.remove(t)

        new_t = OxmlElement("w:t")
        new_t.text = right_text
        if right_text.strip() != right_text:
            new_t.set(qn("xml:space"), "preserve")
        new_r_element.append(new_t)
        run._element.addnext(new_r_element)
        new_run = Run(new_r_element, run._parent)
        return run, new_run

    def get_context_at_range(self, start_idx: int, end_idx: int) -> Optional[TextSpan]:
        real_spans = [s for s in self.spans if s.run and s.end > start_idx and s.start < end_idx]
        if real_spans:
            return real_spans[0]
        return None
