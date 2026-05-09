from __future__ import annotations

import json
import os
import re
import uuid
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


WORKBOOK_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
APP_DATA_DIR_NAME = "ActionsTui"
STATE_FILE_NAME = "actions_state.json"


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def excel_date_to_iso(raw_value: str) -> str:
    if not raw_value:
        return ""
    try:
        serial = float(raw_value)
    except ValueError:
        return raw_value
    if serial <= 0:
        return ""
    converted = date(1899, 12, 30) + timedelta(days=int(serial))
    return converted.isoformat()


def slugify_name(name: str) -> str:
    cleaned = "".join(ch.lower() if ch.isalnum() else "_" for ch in name.strip())
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.strip("_") or "table"


def display_table_name(raw_name: str) -> str:
    text = (raw_name or "").replace("_", " ").strip()
    return text or "Untitled Table"


def default_data_dir(base_dir: Path) -> Path:
    override = os.environ.get("ACTIONS_TUI_DATA_DIR", "").strip()
    if override:
        return Path(override).expanduser()
    repo_dir = Path(__file__).resolve().parent
    if base_dir.resolve() != repo_dir.resolve():
        return base_dir / "data"
    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app_data:
        return Path(local_app_data) / APP_DATA_DIR_NAME
    return base_dir / "data"


NOTE_BULLET_RE = re.compile(r"^(?P<prefix>\s*(?:->|>>|[-*•]|\d+[.)])\s*)(?P<text>.*)$")
NOTE_BULLET_COLOR_RE = re.compile(r"^\[color:(?P<color>(?:[a-z]+|#[0-9a-fA-F]{6}))\]\s*(?P<text>.*)$")


def normalize_note_bullet_color(color: str) -> str:
    return color.strip().lower()


def format_note_bullet(note_text: str, prefix: str = "- ", color: str = "") -> str:
    lines = [line.rstrip() for line in note_text.strip().splitlines()]
    if not lines:
        return ""
    normalized_color = normalize_note_bullet_color(color)
    continuation_prefix = " " * len(prefix)
    first_line = lines[0]
    if normalized_color:
        first_line = f"[color:{normalized_color}] {first_line}"
    formatted = [f"{prefix}{first_line}"]
    for line in lines[1:]:
        formatted.append(f"{continuation_prefix}{line}" if line else "")
    return "\n".join(formatted).rstrip()


def append_note_text(existing_notes: str, note_text: str, prefix: str = "- ", color: str = "") -> str:
    bullet = format_note_bullet(note_text, prefix=prefix, color=color)
    existing = existing_notes.rstrip()
    if not bullet:
        return existing
    if not existing:
        return bullet
    return f"{existing}\n{bullet}"


def extract_note_bullets(notes: str) -> list[NoteBullet]:
    lines = notes.splitlines()
    bullets: list[NoteBullet] = []
    current_start: int | None = None
    current_prefix = ""
    current_lines: list[str] = []
    current_color = ""

    def flush(end_line: int) -> None:
        nonlocal current_start, current_prefix, current_lines, current_color
        if current_start is None:
            return
        bullets.append(
            NoteBullet(
                index=len(bullets),
                start_line=current_start,
                end_line=end_line,
                prefix=current_prefix,
                text="\n".join(current_lines).rstrip(),
                color=current_color,
            )
        )
        current_start = None
        current_prefix = ""
        current_lines = []
        current_color = ""

    for index, line in enumerate(lines):
        match = NOTE_BULLET_RE.match(line)
        if (
            match
            and current_start is not None
            and current_prefix.strip().startswith(">>")
            and line.startswith(" ")
            and match.group("prefix").strip().startswith("-")
        ):
            current_lines.append(line)
            continue
        if match:
            flush(index)
            current_start = index
            current_prefix = match.group("prefix")
            text = match.group("text")
            color_match = NOTE_BULLET_COLOR_RE.match(text)
            if color_match:
                current_color = normalize_note_bullet_color(color_match.group("color"))
                text = color_match.group("text")
            current_lines = [text]
            continue
        if current_start is not None:
            current_lines.append(line)
            continue
        flush(index)
    flush(len(lines))
    return bullets


def replace_note_bullet(existing_notes: str, bullet_index: int, note_text: str, color: str | None = None) -> str:
    bullets = extract_note_bullets(existing_notes)
    if bullet_index < 0 or bullet_index >= len(bullets):
        raise IndexError("Unknown note bullet.")
    bullet = bullets[bullet_index]
    replacement = format_note_bullet(
        note_text,
        prefix=bullet.prefix,
        color=bullet.color if color is None else color,
    )
    lines = existing_notes.splitlines()
    replacement_lines = replacement.splitlines() if replacement else []
    updated_lines = lines[: bullet.start_line] + replacement_lines + lines[bullet.end_line :]
    return "\n".join(updated_lines).rstrip()


def insert_note_bullet_after(
    existing_notes: str,
    bullet_index: int,
    note_text: str,
    prefix: str | None = None,
    color: str = "",
) -> str:
    bullets = extract_note_bullets(existing_notes)
    if bullet_index < 0 or bullet_index >= len(bullets):
        raise IndexError("Unknown note bullet.")
    reference_bullet = bullets[bullet_index]
    new_bullet = format_note_bullet(note_text, prefix=prefix or reference_bullet.prefix, color=color)
    if not new_bullet:
        return existing_notes.rstrip()
    lines = existing_notes.splitlines()
    insert_at = reference_bullet.end_line
    updated_lines = lines[:insert_at] + new_bullet.splitlines() + lines[insert_at:]
    return "\n".join(updated_lines).rstrip()


def update_note_bullet_color(existing_notes: str, bullet_index: int, color: str) -> str:
    bullets = extract_note_bullets(existing_notes)
    if bullet_index < 0 or bullet_index >= len(bullets):
        raise IndexError("Unknown note bullet.")
    bullet = bullets[bullet_index]
    replacement = format_note_bullet(bullet.text, prefix=bullet.prefix, color=color)
    lines = existing_notes.splitlines()
    replacement_lines = replacement.splitlines() if replacement else []
    updated_lines = lines[: bullet.start_line] + replacement_lines + lines[bullet.end_line :]
    return "\n".join(updated_lines).rstrip()


def delete_note_bullet(existing_notes: str, bullet_index: int) -> str:
    bullets = extract_note_bullets(existing_notes)
    if bullet_index < 0 or bullet_index >= len(bullets):
        raise IndexError("Unknown note bullet.")
    bullet = bullets[bullet_index]
    lines = existing_notes.splitlines()
    updated_lines = lines[: bullet.start_line] + lines[bullet.end_line :]
    return "\n".join(updated_lines).rstrip()


def display_note_text(notes: str) -> str:
    display_lines: list[str] = []
    for line in notes.splitlines():
        match = NOTE_BULLET_RE.match(line)
        if not match:
            display_lines.append(line)
            continue
        text = match.group("text")
        color_match = NOTE_BULLET_COLOR_RE.match(text)
        if color_match:
            text = color_match.group("text")
        display_lines.append(f"{match.group('prefix')}{text}")
    return "\n".join(display_lines)


@dataclass
class TableRecord:
    id: str
    name: str
    created_at: str


@dataclass
class ActionRecord:
    id: str
    table_id: str
    title: str
    due_date: str = ""
    estimate: str = ""
    notes: str = ""
    status: str = ""
    created_at: str = field(default_factory=utc_now_iso)
    updated_at: str = field(default_factory=utc_now_iso)
    completed_at: str = ""
    retired_at: str = ""


@dataclass
class ActionState:
    version: int
    tables: list[TableRecord]
    active: list[ActionRecord]
    archive: list[ActionRecord]
    graveyard: list[ActionRecord]
    retired_tables: list[TableRecord] = field(default_factory=list)


@dataclass(frozen=True)
class NoteBullet:
    index: int
    start_line: int
    end_line: int
    prefix: str
    text: str
    color: str = ""


class ActionsStore:
    def __init__(self, base_dir: Path, workbook_path: Path | None = None) -> None:
        self.base_dir = base_dir
        self.repo_data_dir = base_dir / "data"
        self.repo_state_path = self.repo_data_dir / STATE_FILE_NAME
        self.data_dir = default_data_dir(base_dir)
        self.state_path = self.data_dir / STATE_FILE_NAME
        self.workbook_path = workbook_path
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.state = self._load_or_seed()
        changed = self.normalize_legacy_graveyard()
        if self.normalize_active_action_dates():
            changed = True
        if changed:
            self.save()

    def _load_or_seed(self) -> ActionState:
        if self.state_path.exists():
            return self._load_state()
        if self.repo_state_path != self.state_path and self.repo_state_path.exists():
            state = self._load_state_path(self.repo_state_path)
            self.state = state
            self.save()
            return state
        state = self.import_workbook(self.workbook_path) if self.workbook_path and self.workbook_path.exists() else self.empty_state()
        self.state = state
        self.save()
        return state

    def empty_state(self) -> ActionState:
        return ActionState(version=1, tables=[], active=[], archive=[], graveyard=[], retired_tables=[])

    def _load_state(self) -> ActionState:
        return self._load_state_path(self.state_path)

    def _load_state_path(self, state_path: Path) -> ActionState:
        raw = json.loads(state_path.read_text(encoding="utf-8-sig"))
        return ActionState(
            version=raw.get("version", 1),
            tables=[TableRecord(**item) for item in raw.get("tables", [])],
            active=[ActionRecord(**item) for item in raw.get("active", [])],
            archive=[ActionRecord(**item) for item in raw.get("archive", [])],
            graveyard=[ActionRecord(**item) for item in raw.get("graveyard", [])],
            retired_tables=[TableRecord(**item) for item in raw.get("retired_tables", [])],
        )

    def save(self) -> None:
        payload = {
            "version": self.state.version,
            "tables": [asdict(item) for item in self.state.tables],
            "active": [asdict(item) for item in self.state.active],
            "archive": [asdict(item) for item in self.state.archive],
            "graveyard": [asdict(item) for item in self.state.graveyard],
            "retired_tables": [asdict(item) for item in self.state.retired_tables],
        }
        self.state_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def normalize_active_action_dates(self) -> bool:
        changed = False
        today = date.today()
        today_iso = today.isoformat()
        now = utc_now_iso()
        for action in self.state.active:
            if not action.due_date:
                continue
            try:
                due = date.fromisoformat(action.due_date)
            except ValueError:
                continue
            if due >= today:
                continue
            action.due_date = today_iso
            action.status = "Transferred"
            action.updated_at = now
            changed = True
        return changed

    @staticmethod
    def clear_transferred_status_if_needed(action: ActionRecord) -> bool:
        if action.status != "Transferred":
            return False
        action.status = ""
        action.updated_at = utc_now_iso()
        return True

    def reseed_from_workbook(self) -> None:
        if not self.workbook_path or not self.workbook_path.exists():
            raise FileNotFoundError("Workbook path is not available.")
        self.state = self.import_workbook(self.workbook_path)
        self.normalize_active_action_dates()
        self.save()

    def import_workbook(self, workbook_path: Path) -> ActionState:
        sheets = self._read_workbook_sheets(workbook_path)
        tables: list[TableRecord] = []
        table_ids: dict[str, str] = {}
        active: list[ActionRecord] = []
        archive: list[ActionRecord] = []
        seen_titles: set[tuple[str, str]] = set()

        def ensure_table(raw_name: str) -> str:
            display_name = display_table_name(raw_name)
            key = slugify_name(display_name)
            if key not in table_ids:
                table_id = f"table_{key}"
                table_ids[key] = table_id
                tables.append(TableRecord(id=table_id, name=display_name, created_at=utc_now_iso()))
            return table_ids[key]

        for row in sheets.get("Sheet1", [])[1:]:
            if len(row) < 2:
                continue
            source_table = row[0].strip()
            title = row[1].strip()
            if not source_table or not title:
                continue
            table_id = ensure_table(source_table)
            dedupe_key = (table_id, title.lower())
            if dedupe_key in seen_titles:
                continue
            seen_titles.add(dedupe_key)
            active.append(
                ActionRecord(
                    id=new_id("act"),
                    table_id=table_id,
                    title=title,
                    due_date=excel_date_to_iso(row[2].strip() if len(row) > 2 else ""),
                    estimate=row[3].strip() if len(row) > 3 else "",
                    notes=row[4].strip() if len(row) > 4 else "",
                    status=row[5].strip() if len(row) > 5 else "",
                )
            )

        for row in sheets.get("Sheet3", [])[1:]:
            if len(row) < 2:
                continue
            source_table = row[0].strip()
            title = row[1].strip()
            if not source_table or not title:
                continue
            table_id = ensure_table(source_table)
            archive.append(
                ActionRecord(
                    id=new_id("arc"),
                    table_id=table_id,
                    title=title,
                    due_date=excel_date_to_iso(row[2].strip() if len(row) > 2 else ""),
                    estimate=row[3].strip() if len(row) > 3 else "",
                    notes=row[4].strip() if len(row) > 4 else "",
                    status=row[5].strip() if len(row) > 5 else "Complete",
                    completed_at=utc_now_iso(),
                )
            )

        return ActionState(version=1, tables=tables, active=active, archive=archive, graveyard=[])

    def _read_workbook_sheets(self, workbook_path: Path) -> dict[str, list[list[str]]]:
        with zipfile.ZipFile(workbook_path) as zf:
            shared_strings = self._read_shared_strings(zf)
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
            relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}
            sheet_rows: dict[str, list[list[str]]] = {}
            sheets = workbook.find("a:sheets", WORKBOOK_NS)
            if sheets is None:
                return {}
            for sheet in sheets:
                name = sheet.attrib.get("name", "")
                rel_id = sheet.attrib.get(REL_NS, "")
                target = rel_map.get(rel_id)
                if not target:
                    continue
                xml_root = ET.fromstring(zf.read(f"xl/{target}"))
                rows: list[list[str]] = []
                for row_node in xml_root.findall(".//a:row", WORKBOOK_NS):
                    values = [self._read_cell_value(cell, shared_strings) for cell in row_node.findall("a:c", WORKBOOK_NS)]
                    rows.append(values)
                sheet_rows[name] = rows
            return sheet_rows

    def _read_shared_strings(self, workbook_zip: zipfile.ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in workbook_zip.namelist():
            return []
        root = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
        values: list[str] = []
        for si in root.findall("a:si", WORKBOOK_NS):
            values.append("".join(node.text or "" for node in si.iterfind(".//a:t", WORKBOOK_NS)))
        return values

    def _read_cell_value(self, cell: ET.Element, shared_strings: list[str]) -> str:
        cell_type = cell.attrib.get("t", "")
        if cell_type == "inlineStr":
            node = cell.find("a:is/a:t", WORKBOOK_NS)
            return node.text or "" if node is not None else ""
        value_node = cell.find("a:v", WORKBOOK_NS)
        if value_node is None:
            return ""
        value = value_node.text or ""
        if cell_type == "s":
            try:
                return shared_strings[int(value)]
            except (IndexError, ValueError):
                return value
        return value

    def all_tables(self) -> list[TableRecord]:
        return list(self.state.tables)

    def all_known_tables(self) -> list[TableRecord]:
        known = list(self.state.tables)
        active_ids = {table.id for table in known}
        known.extend(table for table in self.state.retired_tables if table.id not in active_ids)
        return known

    def active_actions(self) -> list[ActionRecord]:
        return sorted(self.state.active, key=self._sort_key)

    def archive_actions(self) -> list[ActionRecord]:
        return sorted(self.state.archive, key=self._sort_key)

    def graveyard_actions(self) -> list[ActionRecord]:
        return []

    def today_actions(self, selected_date: date) -> list[ActionRecord]:
        target = selected_date.isoformat()
        return [item for item in self.active_actions() if item.due_date == target]

    def actions_for_table(self, table_id: str) -> list[ActionRecord]:
        return [item for item in self.active_actions() if item.table_id == table_id]

    def table_for_id(self, table_id: str) -> TableRecord | None:
        for table in self.state.tables:
            if table.id == table_id:
                return table
        for table in self.state.retired_tables:
            if table.id == table_id:
                return table
        return None

    def action_by_id(self, action_id: str) -> tuple[str, ActionRecord] | None:
        for bucket_name in ("active", "archive"):
            bucket = getattr(self.state, bucket_name)
            for item in bucket:
                if item.id == action_id:
                    return bucket_name, item
        return None

    def ensure_table(self, table_name: str) -> TableRecord:
        cleaned = display_table_name(table_name)
        for table in self.state.tables:
            if table.name.casefold() == cleaned.casefold():
                return table
        table = TableRecord(id=f"table_{slugify_name(cleaned)}_{uuid.uuid4().hex[:4]}", name=cleaned, created_at=utc_now_iso())
        self.state.tables.append(table)
        self.save()
        return table

    def rename_table(self, table_id: str, new_name: str) -> TableRecord:
        table = self.table_for_id(table_id)
        if table is None:
            raise KeyError("Unknown table.")
        table.name = display_table_name(new_name)
        self.save()
        return table

    def retire_table(self, table_id: str) -> tuple[TableRecord, int]:
        table = next((item for item in self.state.tables if item.id == table_id), None)
        if table is None:
            raise KeyError("Unknown table.")
        now = utc_now_iso()
        moved_count = 0
        remaining_active: list[ActionRecord] = []
        for action in self.state.active:
            if action.table_id != table_id:
                remaining_active.append(action)
                continue
            action.status = "Completed"
            action.completed_at = now
            action.updated_at = now
            self.state.archive.append(action)
            moved_count += 1
        self.state.active = remaining_active
        self.state.tables = [item for item in self.state.tables if item.id != table_id]
        if not any(item.id == table.id for item in self.state.retired_tables):
            self.state.retired_tables.append(table)
        self.save()
        return table, moved_count

    def add_action(
        self,
        *,
        table_name: str,
        title: str,
        due_date: str,
        estimate: str,
        notes: str,
        status: str,
    ) -> ActionRecord:
        table = self.ensure_table(table_name)
        action = ActionRecord(
            id=new_id("act"),
            table_id=table.id,
            title=title.strip(),
            due_date=due_date.strip(),
            estimate=estimate.strip(),
            notes=notes.rstrip(),
            status=status.strip(),
        )
        self.state.active.append(action)
        self.normalize_active_action_dates()
        self.save()
        return action

    def update_action(
        self,
        action_id: str,
        *,
        table_name: str,
        title: str,
        due_date: str,
        estimate: str,
        notes: str,
        status: str,
    ) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        table = self.ensure_table(table_name)
        action.table_id = table.id
        action.title = title.strip()
        action.due_date = due_date.strip()
        action.estimate = estimate.strip()
        action.notes = notes.rstrip()
        action.updated_at = utc_now_iso()
        self.normalize_active_action_dates()
        return self.set_action_status(action_id, status)

    def update_action_notes(self, action_id: str, notes: str) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        self.clear_transferred_status_if_needed(action)
        action.notes = notes.rstrip()
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def update_action_due_date(self, action_id: str, due_date: str) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        bucket_name, action = found
        if bucket_name == "active":
            self.clear_transferred_status_if_needed(action)
        action.due_date = due_date.strip()
        action.updated_at = utc_now_iso()
        if bucket_name == "active":
            self.normalize_active_action_dates()
        self.save()
        return action

    def set_action_status(self, action_id: str, status: str) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        bucket_name, action = found
        normalized = status.strip()
        now = utc_now_iso()

        if bucket_name == "active" and normalized == "Completed":
            action = self._remove_from_bucket("active", action_id)
            action.status = normalized
            action.completed_at = now
            action.updated_at = now
            self.state.archive.append(action)
            self.save()
            return action

        if bucket_name == "archive" and normalized != "Completed":
            action = self._remove_from_bucket("archive", action_id)
            self._restore_retired_table(action.table_id)
            action.status = normalized
            action.completed_at = ""
            action.updated_at = now
            self.state.active.append(action)
            self.normalize_active_action_dates()
            self.save()
            return action

        action.status = normalized
        if bucket_name == "archive" and not action.completed_at:
            action.completed_at = now
        action.updated_at = now
        self.save()
        return action

    def append_action_note(self, action_id: str, note_text: str, *, prefix: str = "- ", color: str = "") -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        self.clear_transferred_status_if_needed(action)
        action.notes = append_note_text(action.notes, note_text, prefix=prefix, color=color)
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def insert_action_note_after_bullet(
        self,
        action_id: str,
        bullet_index: int,
        note_text: str,
        *,
        prefix: str | None = None,
        color: str = "",
    ) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        self.clear_transferred_status_if_needed(action)
        action.notes = insert_note_bullet_after(action.notes, bullet_index, note_text, prefix=prefix, color=color)
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def update_action_note_bullet(
        self,
        action_id: str,
        bullet_index: int,
        note_text: str,
        *,
        color: str | None = None,
    ) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        self.clear_transferred_status_if_needed(action)
        action.notes = replace_note_bullet(action.notes, bullet_index, note_text, color=color)
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def update_action_note_bullet_color(self, action_id: str, bullet_index: int, color: str) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        self.clear_transferred_status_if_needed(action)
        action.notes = update_note_bullet_color(action.notes, bullet_index, color)
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def delete_action_note_bullet(self, action_id: str, bullet_index: int) -> ActionRecord:
        found = self.action_by_id(action_id)
        if found is None:
            raise KeyError("Unknown action.")
        _, action = found
        self.clear_transferred_status_if_needed(action)
        action.notes = delete_note_bullet(action.notes, bullet_index)
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def complete_action(self, action_id: str) -> ActionRecord:
        action = self._remove_from_bucket("active", action_id)
        action.completed_at = utc_now_iso()
        action.updated_at = action.completed_at
        action.status = "Completed"
        self.state.archive.append(action)
        self.save()
        return action

    def retire_action(self, action_id: str) -> ActionRecord:
        action = self._remove_from_bucket("archive", action_id)
        action.updated_at = utc_now_iso()
        self.save()
        return action

    def restore_action(self, action_id: str) -> ActionRecord:
        action = self._remove_from_bucket("archive", action_id)
        self._restore_retired_table(action.table_id)
        action.updated_at = utc_now_iso()
        self.state.active.append(action)
        self.normalize_active_action_dates()
        self.save()
        return action

    def delete_graveyard_action(self, action_id: str) -> None:
        self._remove_from_bucket("archive", action_id)
        self.save()

    def normalize_legacy_graveyard(self) -> bool:
        if not self.state.graveyard:
            return False
        for action in self.state.graveyard:
            if not action.completed_at:
                action.completed_at = action.retired_at or utc_now_iso()
            action.retired_at = ""
            if not action.status:
                action.status = "Completed"
            self.state.archive.append(action)
        self.state.graveyard = []
        return True

    def _remove_any(self, action_id: str, *, allowed: tuple[str, ...]) -> tuple[str, ActionRecord]:
        for bucket_name in allowed:
            bucket = getattr(self.state, bucket_name)
            for index, action in enumerate(bucket):
                if action.id == action_id:
                    return bucket_name, bucket.pop(index)
        raise KeyError("Unknown action.")

    def _remove_from_bucket(self, bucket_name: str, action_id: str) -> ActionRecord:
        bucket = getattr(self.state, bucket_name)
        for index, action in enumerate(bucket):
            if action.id == action_id:
                return bucket.pop(index)
        raise KeyError("Unknown action.")

    def _restore_retired_table(self, table_id: str) -> TableRecord:
        current = next((item for item in self.state.tables if item.id == table_id), None)
        if current is not None:
            return current
        retired = next((item for item in self.state.retired_tables if item.id == table_id), None)
        if retired is not None:
            self.state.retired_tables = [item for item in self.state.retired_tables if item.id != table_id]
            self.state.tables.append(retired)
            return retired
        fallback = TableRecord(id=table_id, name="Recovered Project", created_at=utc_now_iso())
        self.state.tables.append(fallback)
        return fallback

    def _sort_key(self, item: ActionRecord) -> tuple[Any, ...]:
        due = item.due_date or "9999-99-99"
        table = self.table_for_id(item.table_id)
        table_name = table.name if table else ""
        return due, table_name.casefold(), item.title.casefold()
