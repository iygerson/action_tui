from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable, Optional

from rich.console import Group
from rich.panel import Panel
from rich.text import Text
from textual import on
from textual import events
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.screen import ModalScreen
from textual.widget import Widget
from textual.widgets import Button, Footer, Input, Label, ListItem, ListView, Static, TextArea

from actions_store import ActionRecord, ActionsStore, NoteBullet, TableRecord, display_note_text, extract_note_bullets


DEFAULT_WORKBOOK = Path(r"C:\Users\ilany\Desktop\Actions_Tool.xlsm")
STATUS_OPTIONS = ["Not Started", "Started", "Completed"]
SUBACTION_COLOR = "#d97706"
SUBACTION_NOT_STARTED_COLOR = "#6b7280"
SUBACTION_COMPLETED_COLOR = "#00a651"
TABLE_BORDER_COLOR = "#2f6b3c"
NOTE_BULLET_COLORS = {
    "r": ("red", "Red"),
    "p": ("#8a5cff", "Purple"),
    "b": ("#1f4ba8", "Blue"),
    "c": ("cyan", "Cyan"),
    "w": ("white", "White"),
    "d": ("", "Default"),
}


def normalize_due_date_input(raw_value: str, *, default_year: int | None = None) -> str:
    value = raw_value.strip()
    if not value:
        return ""
    if default_year is None:
        default_year = date.today().year
    try:
        return date.fromisoformat(value).isoformat()
    except ValueError:
        pass
    separator = "/" if "/" in value else "-" if "-" in value else None
    if separator is None:
        return value
    parts = [part.strip() for part in value.split(separator)]
    try:
        if len(parts) == 2:
            month, day = (int(part) for part in parts)
            return date(default_year, month, day).isoformat()
        if len(parts) == 3:
            first, second, third = parts
            if len(first) == 4:
                year, month, day = (int(first), int(second), int(third))
                return date(year, month, day).isoformat()
            month, day, year = (int(first), int(second), int(third))
            if year < 100:
                year += 2000
            return date(year, month, day).isoformat()
    except ValueError:
        return value
    return value


def normalize_status_value(raw_value: str) -> str:
    value = raw_value.strip()
    lowered = value.casefold()
    if not lowered:
        return "Not Started"
    if lowered == "not started":
        return "Not Started"
    if lowered == "started":
        return "Started"
    if lowered in {"completed", "complete", "done"}:
        return "Completed"
    return "Started"


def subaction_color_for_status(status: str) -> str:
    normalized = normalize_status_value(status)
    if normalized == "Not Started":
        return SUBACTION_NOT_STARTED_COLOR
    if normalized == "Completed":
        return SUBACTION_COMPLETED_COLOR
    return SUBACTION_COLOR


def subaction_status_from_color(color: str) -> str:
    normalized = color.strip().lower()
    if normalized == SUBACTION_NOT_STARTED_COLOR.lower():
        return "Not Started"
    if normalized == SUBACTION_COMPLETED_COLOR.lower():
        return "Completed"
    return "Started"


@dataclass
class RepresentativeEntry:
    kind: str
    group: str
    label: str
    action_id: str | None = None
    table_id: str | None = None
    transferred: bool = False


class StatusSelect(Static):
    can_focus = True

    def __init__(self, value: str = "", **kwargs) -> None:
        super().__init__("", **kwargs)
        self.value = normalize_status_value(value)

    def step(self, delta: int) -> None:
        current_index = STATUS_OPTIONS.index(self.value)
        self.value = STATUS_OPTIONS[(current_index + delta) % len(STATUS_OPTIONS)]
        self.refresh()

    def render(self) -> Text:
        text = Text()
        for index, option in enumerate(STATUS_OPTIONS):
            style = "dim"
            if option == self.value:
                style = "bold white on #00a651" if self.has_focus else "bold green"
            text.append(f" {option} ", style=style)
            if index < len(STATUS_OPTIONS) - 1:
                text.append("  ")
        return text

    def on_focus(self, _: events.Focus) -> None:
        self.refresh()

    def on_blur(self, _: events.Blur) -> None:
        self.refresh()

    async def _on_key(self, event: events.Key) -> None:
        normalized = NoteWriteTextArea.normalized_event_keys(event)
        if "escape" in normalized:
            event.stop()
            event.prevent_default()
            dismiss = getattr(self.screen, "dismiss_with_focus_restore", None)
            if callable(dismiss):
                dismiss(None)
            return
        if "tab" in normalized:
            event.stop()
            event.prevent_default()
            submit = getattr(self.screen, "submit", None)
            if callable(submit):
                submit()
            return
        if "left" in normalized:
            event.stop()
            event.prevent_default()
            self.step(-1)
            return
        if "right" in normalized:
            event.stop()
            event.prevent_default()
            self.step(1)
            return
        focus_previous = getattr(self.screen, "focus_previous_subaction_field", None)
        if "up" in normalized and callable(focus_previous):
            event.stop()
            event.prevent_default()
            focus_previous()
            return
        focus_next = getattr(self.screen, "focus_next_subaction_field", None)
        if ("down" in normalized or {"enter", "return"} & normalized) and callable(focus_next):
            event.stop()
            event.prevent_default()
            focus_next()
            return
        await super()._on_key(event)


class PromptScreen(ModalScreen[Optional[str]]):
    BINDINGS = [
        Binding("escape", "cancel", "Cancel"),
        Binding("enter", "submit", "Save"),
    ]

    def __init__(self, title: str, value: str = "") -> None:
        super().__init__()
        self.title = title
        self.value = value

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label(self.title, id="prompt_title"),
            Input(value=self.value, id="prompt_input"),
            Static("Enter saves, Esc cancels", id="prompt_hint"),
            id="prompt_dialog",
        )

    def on_mount(self) -> None:
        self.query_one("#prompt_input", Input).focus()

    @on(Input.Submitted, "#prompt_input")
    def input_submitted(self) -> None:
        self.action_submit()

    def action_submit(self) -> None:
        self.dismiss(self.query_one("#prompt_input", Input).value)

    def action_cancel(self) -> None:
        self.dismiss(None)


class ConfirmScreen(ModalScreen[bool]):
    BINDINGS = [
        Binding("left", "prev_choice", show=False, priority=True),
        Binding("up", "prev_choice", show=False, priority=True),
        Binding("right", "next_choice", show=False, priority=True),
        Binding("down", "next_choice", show=False, priority=True),
        Binding("tab", "submit", show=False, priority=True),
        Binding("enter", "submit", show=False, priority=True),
        Binding("y", "accept", "Yes"),
        Binding("n", "cancel", "No"),
        Binding("escape", "cancel", "Cancel"),
    ]

    def __init__(self, message: str, *, default_accept: bool = True) -> None:
        super().__init__()
        self.message = message
        self.selected_choice = default_accept

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label("Confirm", id="confirm_title"),
            Static(self.message, id="confirm_message"),
            Static(id="confirm_choices"),
            Static("Arrows switch choice. Tab or Enter confirms. Y = yes, N or Esc = cancel.", id="confirm_hint"),
            id="confirm_dialog",
        )

    def on_mount(self) -> None:
        self.refresh_choices()

    def refresh_choices(self) -> None:
        yes_style = "bold white on #00a651" if self.selected_choice else "bold"
        no_style = "bold white on #7a1f1f" if not self.selected_choice else "bold"
        text = Text()
        text.append(" Yes ", style=yes_style)
        text.append("   ")
        text.append(" No ", style=no_style)
        self.query_one("#confirm_choices", Static).update(text)

    def action_prev_choice(self) -> None:
        self.selected_choice = not self.selected_choice
        self.refresh_choices()

    def action_next_choice(self) -> None:
        self.selected_choice = not self.selected_choice
        self.refresh_choices()

    def action_submit(self) -> None:
        if self.selected_choice:
            self.action_accept()
            return
        self.action_cancel()

    def action_accept(self) -> None:
        self.dismiss(True)

    def action_cancel(self) -> None:
        self.dismiss(False)


class HelpScreen(ModalScreen[None]):
    BINDINGS = [Binding("escape", "close", "Close"), Binding("q", "close", "Close")]

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label("Actions TUI Prototype", id="help_title"),
            Static(
                "\n".join(
                    [
                        "Up/Down: switch between Today, Projects, and Completed while on the left",
                        "Enter or Tab: move from group focus into the current group's item list",
                        "Up/Down inside a group: move between items in that group",
                        "Right: open the focused group or item in the detail pane",
                        "When Today is open on the right, Up/Down moves across action rows",
                        "When a project table is open on the right, Up/Down moves across action rows",
                        "Right on a focused project-table row opens that action's detail",
                        "In the all-projects right-pane view, Up/Down moves between project tables",
                        "Right, Enter, or Tab enters the focused project's action rows; Left returns to table focus",
                        "When an action with note bullets is open on the right, Up/Down moves across those bullets",
                        "Right, Enter, or Tab on a focused note bullet opens it for editing",
                        "Left: back out of detail, or back out of a group's item list",
                        "N: add a note on a focused action; from a project table with no focused action it adds a new action",
                        "R: change the color of the focused note bullet",
                        "S: open the status picker for the focused action",
                        "Inside the action editor, arrows or Enter move between fields, Left/Right on Status changes it, and Tab saves",
                        "E: edit the selected action or rename the selected project",
                        "A: in full action detail add a subaction; otherwise add action to the selected project",
                        "C: complete active action to Completed",
                        "X: delete the focused note bullet, or in Completed move the action to Graveyard after confirmation",
                        "V: toggle the Completed section between Completed and Graveyard",
                        "U: restore from Completed or Graveyard",
                        "D: edit the due date of the focused action",
                        "Delete: permanently delete from Graveyard",
                        "[ or ]: move the Today date backward or forward",
                        "Q: quit",
                    ]
                ),
                id="help_body",
            ),
            Static("Esc closes help", id="help_hint"),
            id="help_dialog",
        )

    def action_close(self) -> None:
        self.dismiss(None)


class ActionEditorScreen(ModalScreen[Optional[dict[str, str]]]):
    BINDINGS = [
        Binding("up", "editor_up", show=False, priority=True),
        Binding("down", "editor_down", show=False, priority=True),
        Binding("left", "editor_left", show=False, priority=True),
        Binding("right", "editor_right", show=False, priority=True),
        Binding("enter", "editor_enter", show=False, priority=True),
        Binding("tab", "submit", show=False, priority=True),
        Binding("ctrl+s", "submit", "Save"),
        Binding("escape", "cancel", "Cancel"),
    ]

    def __init__(self, initial: dict[str, str], title: str) -> None:
        super().__init__()
        self.initial = initial
        self.title = title

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label(self.title, id="editor_title"),
            Label("Table", classes="editor_label"),
            Input(value=self.initial.get("table_name", ""), id="table_input"),
            Label("Action", classes="editor_label"),
            Input(value=self.initial.get("title", ""), id="action_input"),
            Horizontal(
                Vertical(
                    Label("Due Date (YYYY-MM-DD or M/D)", classes="editor_label"),
                    Input(value=self.initial.get("due_date", ""), id="due_input"),
                ),
                Vertical(
                    Label("Time", classes="editor_label"),
                    Input(value=self.initial.get("estimate", ""), id="estimate_input"),
                ),
                Vertical(
                    Label("Status", classes="editor_label"),
                    StatusSelect(value=self.initial.get("status", ""), id="status_input"),
                ),
                classes="editor_row",
            ),
            Label("Notes", classes="editor_label"),
            TextArea(self.initial.get("notes", ""), id="notes_input", placeholder="Add notes for this action"),
            Horizontal(
                Button("Save", id="save_button", variant="success"),
                Button("Cancel", id="cancel_button"),
                id="editor_buttons",
            ),
            Static("Arrows or Enter move between fields. On Status, Left/Right changes the option. Tab or Ctrl+S saves. Esc cancels.", id="editor_hint"),
            id="editor_dialog",
        )

    def on_mount(self) -> None:
        table_input = self.query_one("#table_input", Input)
        if table_input.value.strip():
            self.query_one("#action_input", Input).focus()
        else:
            table_input.focus()

    def field_order(self) -> list[Widget]:
        return [
            self.query_one("#table_input", Input),
            self.query_one("#action_input", Input),
            self.query_one("#due_input", Input),
            self.query_one("#estimate_input", Input),
            self.query_one("#status_input", StatusSelect),
            self.query_one("#notes_input", TextArea),
        ]

    def focus_relative_field(self, delta: int) -> None:
        fields = self.field_order()
        focused = self.app.focused
        try:
            current_index = fields.index(focused)
        except ValueError:
            current_index = 0 if delta >= 0 else len(fields) - 1
        next_index = (current_index + delta) % len(fields)
        fields[next_index].focus()

    def action_editor_down(self) -> None:
        self.focus_relative_field(1)

    def action_editor_up(self) -> None:
        self.focus_relative_field(-1)

    def action_editor_left(self) -> None:
        focused = self.app.focused
        if isinstance(focused, StatusSelect):
            focused.step(-1)
            return
        self.focus_relative_field(-1)

    def action_editor_right(self) -> None:
        focused = self.app.focused
        if isinstance(focused, StatusSelect):
            focused.step(1)
            return
        self.focus_relative_field(1)

    def action_editor_enter(self) -> None:
        self.focus_relative_field(1)

    @on(Button.Pressed, "#save_button")
    def save_pressed(self) -> None:
        self.action_submit()

    @on(Button.Pressed, "#cancel_button")
    def cancel_pressed(self) -> None:
        self.action_cancel()

    def action_submit(self) -> None:
        self.dismiss(
            {
                "table_name": self.query_one("#table_input", Input).value.strip(),
                "title": self.query_one("#action_input", Input).value.strip(),
                "due_date": normalize_due_date_input(self.query_one("#due_input", Input).value),
                "estimate": self.query_one("#estimate_input", Input).value.strip(),
                "status": self.query_one("#status_input", StatusSelect).value,
                "notes": self.query_one("#notes_input", TextArea).text.rstrip(),
            }
        )

    def action_cancel(self) -> None:
        self.dismiss(None)


class StatusPickerScreen(ModalScreen[Optional[str]]):
    BINDINGS = [
        Binding("up", "prev_option", show=False, priority=True),
        Binding("left", "prev_option", show=False, priority=True),
        Binding("down", "next_option", show=False, priority=True),
        Binding("right", "next_option", show=False, priority=True),
        Binding("enter", "submit", show=False, priority=True),
        Binding("tab", "submit", show=False, priority=True),
        Binding("escape", "cancel", "Cancel"),
    ]

    def __init__(self, title: str, value: str = "") -> None:
        super().__init__()
        self.title = title
        self.value = value

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label(self.title, id="status_picker_title"),
            StatusSelect(value=self.value, id="status_picker_input"),
            Static("Left/Right or Up/Down changes status. Enter or Tab saves. Esc cancels.", id="status_picker_hint"),
            id="status_picker_dialog",
        )

    def on_mount(self) -> None:
        self.query_one("#status_picker_input", StatusSelect).focus()

    def action_prev_option(self) -> None:
        self.query_one("#status_picker_input", StatusSelect).step(-1)

    def action_next_option(self) -> None:
        self.query_one("#status_picker_input", StatusSelect).step(1)

    def action_submit(self) -> None:
        self.dismiss(self.query_one("#status_picker_input", StatusSelect).value)

    def action_cancel(self) -> None:
        self.dismiss(None)


class NoteColorPaletteScreen(ModalScreen[Optional[str]]):
    BINDINGS = [
        Binding("escape", "cancel", "Cancel"),
        Binding("r", "pick_red", show=False, priority=True),
        Binding("p", "pick_purple", show=False, priority=True),
        Binding("b", "pick_blue", show=False, priority=True),
        Binding("c", "pick_cyan", show=False, priority=True),
        Binding("w", "pick_white", show=False, priority=True),
        Binding("d", "pick_default", show=False, priority=True),
        Binding("shift+r", "pick_red", show=False, priority=True),
        Binding("shift+p", "pick_purple", show=False, priority=True),
        Binding("shift+b", "pick_blue", show=False, priority=True),
        Binding("shift+c", "pick_cyan", show=False, priority=True),
        Binding("shift+w", "pick_white", show=False, priority=True),
        Binding("shift+d", "pick_default", show=False, priority=True),
    ]

    def __init__(self, title: str) -> None:
        super().__init__()
        self.title = title

    def compose(self) -> ComposeResult:
        text = Text()
        text.append("R ", style="bold white on red")
        text.append("Red\n", style="red")
        text.append("P ", style="bold white on #8a5cff")
        text.append("Purple\n", style="#8a5cff")
        text.append("B ", style="bold white on #1f4ba8")
        text.append("Blue\n", style="#1f4ba8")
        text.append("C ", style="bold black on cyan")
        text.append("Cyan\n", style="cyan")
        text.append("W ", style="bold black on white")
        text.append("White\n", style="white")
        text.append("D ", style="bold")
        text.append("Default")
        yield Vertical(
            Label(self.title, id="note_color_title"),
            Static(text, id="note_color_body"),
            Static("Press the letter for the color. Esc cancels.", id="note_color_hint"),
            id="note_color_dialog",
        )

    def pick_color(self, key: str) -> None:
        self.dismiss(NOTE_BULLET_COLORS[key][0])

    def action_pick_red(self) -> None:
        self.pick_color("r")

    def action_pick_purple(self) -> None:
        self.pick_color("p")

    def action_pick_blue(self) -> None:
        self.pick_color("b")

    def action_pick_cyan(self) -> None:
        self.pick_color("c")

    def action_pick_white(self) -> None:
        self.pick_color("w")

    def action_pick_default(self) -> None:
        self.pick_color("d")

    def action_cancel(self) -> None:
        self.dismiss(None)


class NoteWriteTextArea(TextArea):
    NEWLINE_KEYS = {"shift+enter", "shift+return", "ctrl+j", "meta+j", "newline"}

    @staticmethod
    def normalized_event_keys(event: events.Key) -> set[str]:
        keys = {event.key.lower()}
        keys.update(alias.lower() for alias in event.aliases)
        keys.update(alias.lower().replace("_", "+") for alias in event.name_aliases)
        return keys

    async def _on_key(self, event: events.Key) -> None:
        normalized = self.normalized_event_keys(event)
        if "escape" in normalized:
            event.stop()
            event.prevent_default()
            dismiss = getattr(self.screen, "dismiss_with_focus_restore", None)
            if callable(dismiss):
                dismiss(None)
            return
        if "tab" in normalized:
            event.stop()
            event.prevent_default()
            submit = getattr(self.screen, "submit", None)
            if callable(submit):
                submit()
            return
        focus_previous = getattr(self.screen, "focus_previous_subaction_field", None)
        if "up" in normalized and callable(focus_previous):
            event.stop()
            event.prevent_default()
            focus_previous()
            return
        focus_next = getattr(self.screen, "focus_next_subaction_field", None)
        if "down" in normalized and callable(focus_next):
            event.stop()
            event.prevent_default()
            focus_next()
            return
        if self.NEWLINE_KEYS & normalized:
            event.stop()
            event.prevent_default()
            self.insert("\n")
            return
        if {"enter", "return"} & normalized:
            event.stop()
            event.prevent_default()
            self.insert("\n")
            return
        await super()._on_key(event)

    def on_focus(self, _: events.Focus) -> None:
        refresh = getattr(self.screen, "refresh_controls_display", None)
        if callable(refresh):
            refresh()

    def on_blur(self, _: events.Blur) -> None:
        refresh = getattr(self.screen, "refresh_controls_display", None)
        if callable(refresh):
            refresh()


class NoteWriteControls(Static):
    can_focus = True

    async def _on_key(self, event: events.Key) -> None:
        normalized = NoteWriteTextArea.normalized_event_keys(event)
        if "escape" in normalized:
            event.stop()
            event.prevent_default()
            dismiss = getattr(self.screen, "dismiss_with_focus_restore", None)
            if callable(dismiss):
                dismiss(None)
            return
        if "tab" in normalized:
            event.stop()
            event.prevent_default()
            submit = getattr(self.screen, "submit", None)
            if callable(submit):
                submit()
            return
        if {"enter", "return"} & normalized:
            event.stop()
            event.prevent_default()
            submit = getattr(self.screen, "submit", None)
            if callable(submit):
                submit()
            return
        await super()._on_key(event)

    def on_focus(self, _: events.Focus) -> None:
        refresh = getattr(self.screen, "refresh_controls_display", None)
        if callable(refresh):
            refresh()

    def on_blur(self, _: events.Blur) -> None:
        refresh = getattr(self.screen, "refresh_controls_display", None)
        if callable(refresh):
            refresh()


class SubactionTitleInput(Input):
    async def _on_key(self, event: events.Key) -> None:
        normalized = NoteWriteTextArea.normalized_event_keys(event)
        if "escape" in normalized:
            event.stop()
            event.prevent_default()
            dismiss = getattr(self.screen, "dismiss_with_focus_restore", None)
            if callable(dismiss):
                dismiss(None)
            return
        if "tab" in normalized:
            event.stop()
            event.prevent_default()
            submit = getattr(self.screen, "submit", None)
            if callable(submit):
                submit()
            return
        focus_previous = getattr(self.screen, "focus_previous_subaction_field", None)
        if "up" in normalized and callable(focus_previous):
            event.stop()
            event.prevent_default()
            focus_previous()
            return
        focus_next = getattr(self.screen, "focus_next_subaction_field", None)
        if "down" in normalized and callable(focus_next):
            event.stop()
            event.prevent_default()
            focus_next()
            return
        await super()._on_key(event)

    def on_focus(self, _: events.Focus) -> None:
        refresh = getattr(self.screen, "refresh_controls_display", None)
        if callable(refresh):
            refresh()

    def on_blur(self, _: events.Blur) -> None:
        refresh = getattr(self.screen, "refresh_controls_display", None)
        if callable(refresh):
            refresh()


class ActionNoteScreen(ModalScreen[Optional[str]]):
    BINDINGS: list[Binding] = []

    DEFAULT_HINT = "Enter adds a new line. Tab saves and closes the note."

    def __init__(self, title: str, initial_text: str = "") -> None:
        super().__init__()
        self.title = title
        self.initial_text = initial_text

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label(self.title, id="write_title"),
            NoteWriteControls("", id="write_controls"),
            NoteWriteTextArea(
                self.initial_text,
                soft_wrap=True,
                show_line_numbers=False,
                highlight_cursor_line=False,
                placeholder="Write a new note bullet here",
                id="write_text_input",
            ),
            Static(self.DEFAULT_HINT, id="write_hint"),
            id="write_dialog",
        )

    def on_mount(self) -> None:
        text_area = self.text_area()
        text_area.focus()
        if self.initial_text:
            lines = self.initial_text.splitlines()
            text_area.cursor_location = (len(lines) - 1, len(lines[-1]))
        self.refresh_controls_display()

    def text_area(self) -> NoteWriteTextArea:
        return self.query_one("#write_text_input", NoteWriteTextArea)

    def controls(self) -> NoteWriteControls:
        return self.query_one("#write_controls", NoteWriteControls)

    def focus_relative_field(self, delta: int) -> None:
        fields: list[Widget] = [self.controls(), self.text_area()]
        focused = self.app.focused
        try:
            current_index = fields.index(focused)
        except ValueError:
            current_index = 0 if delta >= 0 else len(fields) - 1
        next_index = (current_index + delta) % len(fields)
        fields[next_index].focus()
        self.refresh_controls_display()

    def action_focus_next_field(self) -> None:
        self.focus_relative_field(1)

    def action_focus_prev_field(self) -> None:
        self.focus_relative_field(-1)

    def refresh_controls_display(self) -> None:
        controls = self.controls()
        controls_focused = self.app.focused is controls
        hint_style = "bold green" if controls_focused else "bold yellow"
        text = Text()
        text.append("Notes editor", style="bold")
        text.append("    ")
        text.append("Enter", style="bold yellow")
        text.append(" new line    ")
        text.append("Tab", style="bold yellow")
        text.append(" save and close    ")
        text.append("Esc", style="bold yellow")
        text.append(" cancel")
        text.append("\n")
        if controls_focused:
            text.append("Controls active. Enter or Tab saves the note.", style=hint_style)
        else:
            text.append("Editor active. Enter inserts a new line. Tab saves and closes.", style=hint_style)
        controls.update(text)
        if controls_focused:
            controls.add_class("write-controls-focused")
        else:
            controls.remove_class("write-controls-focused")

    def dismiss_with_focus_restore(self, result: str | None = None) -> None:
        self.dismiss(result)
        restore_focus = getattr(self.app, "restore_navigation_focus", None)
        if callable(restore_focus):
            self.app.call_after_refresh(restore_focus)

    def key_escape(self) -> None:
        self.dismiss_with_focus_restore(None)

    def submit(self) -> None:
        self.dismiss_with_focus_restore(self.text_area().text.rstrip())


class ActionSubactionScreen(ModalScreen[Optional[dict[str, str]]]):
    BINDINGS: list[Binding] = []

    DEFAULT_HINT = "Tab saves and closes. Esc cancels."

    def __init__(
        self,
        title: str,
        *,
        initial_title: str = "",
        initial_notes: str = "",
        initial_status: str = "Not Started",
        focus_notes: bool = False,
    ) -> None:
        super().__init__()
        self.title = title
        self.initial_title = initial_title
        self.initial_notes = initial_notes
        self.initial_status = normalize_status_value(initial_status)
        self.focus_notes = focus_notes

    def compose(self) -> ComposeResult:
        yield Vertical(
            Label(self.title, id="subaction_title"),
            NoteWriteControls("", id="subaction_controls"),
            Label("Subaction", classes="editor_label"),
            SubactionTitleInput(
                value=f">> {self.initial_title}".rstrip(),
                placeholder=">> Subaction",
                id="subaction_title_input",
            ),
            Label("Status", classes="editor_label"),
            StatusSelect(value=self.initial_status, id="subaction_status_input"),
            Label("Subaction Notes", classes="editor_label"),
            NoteWriteTextArea(
                self.prepared_notes_text(),
                soft_wrap=True,
                show_line_numbers=False,
                highlight_cursor_line=False,
                placeholder="Additional notes for this subaction",
                id="subaction_notes_input",
            ),
            Static(self.DEFAULT_HINT, id="subaction_hint"),
            id="subaction_dialog",
        )

    def on_mount(self) -> None:
        title_input = self.title_input()
        if self.focus_notes:
            notes_input = self.notes_input()
            notes_input.focus()
            if notes_input.text:
                lines = notes_input.text.split("\n")
                notes_input.cursor_location = (len(lines) - 1, len(lines[-1]))
        else:
            title_input.focus()
            title_input.cursor_position = len(title_input.value)
        self.refresh_controls_display()

    @on(Input.Submitted, "#subaction_title_input")
    def subaction_title_submitted(self) -> None:
        self.status_input().focus()

    def controls(self) -> NoteWriteControls:
        return self.query_one("#subaction_controls", NoteWriteControls)

    def title_input(self) -> SubactionTitleInput:
        return self.query_one("#subaction_title_input", SubactionTitleInput)

    def notes_input(self) -> NoteWriteTextArea:
        return self.query_one("#subaction_notes_input", NoteWriteTextArea)

    def status_input(self) -> StatusSelect:
        return self.query_one("#subaction_status_input", StatusSelect)

    def focus_previous_subaction_field(self) -> None:
        focused = self.app.focused
        if focused is self.notes_input():
            self.status_input().focus()
            return
        self.title_input().focus()

    def focus_next_subaction_field(self) -> None:
        focused = self.app.focused
        if focused is self.title_input():
            self.status_input().focus()
            return
        self.notes_input().focus()

    def prepared_notes_text(self) -> str:
        if not self.focus_notes:
            return self.initial_notes
        if self.initial_notes:
            return f"{self.initial_notes.rstrip()}\n"
        return ""

    def refresh_controls_display(self) -> None:
        controls = self.controls()
        controls_focused = self.app.focused is controls
        text = Text()
        text.append("Subaction editor", style="bold")
        text.append("    ")
        text.append("Status", style="bold yellow")
        text.append(" Left/Right changes state    ")
        text.append("Tab", style="bold yellow")
        text.append(" save and close    ")
        text.append("Esc", style="bold yellow")
        text.append(" cancel")
        text.append("\n")
        if controls_focused:
            text.append("Controls active. Enter or Tab saves the subaction.", style="bold green")
        else:
            text.append("Click either field to edit. Tab saves and closes.", style="bold yellow")
        controls.update(text)
        if controls_focused:
            controls.add_class("write-controls-focused")
        else:
            controls.remove_class("write-controls-focused")

    def dismiss_with_focus_restore(self, result: dict[str, str] | None = None) -> None:
        self.dismiss(result)
        restore_focus = getattr(self.app, "restore_navigation_focus", None)
        if callable(restore_focus):
            self.app.call_after_refresh(restore_focus)

    def key_escape(self) -> None:
        self.dismiss_with_focus_restore(None)

    def submit(self) -> None:
        raw_title = self.title_input().value.strip()
        if raw_title.startswith(">>"):
            raw_title = raw_title[2:].lstrip()
        self.dismiss_with_focus_restore(
            {
                "title": raw_title,
                "status": self.status_input().value,
                "notes": self.notes_input().text.rstrip(),
            }
        )


class ActionsTuiApp(App[None]):
    CSS = """
    App {
        background: ansi_default;
    }
    Screen {
        background: ansi_default;
    }
    ModalScreen {
        align: center middle;
        background: rgba(8, 12, 18, 0.72);
    }
    #layout {
        layout: vertical;
        height: 1fr;
    }
    #content {
        height: 3fr;
    }
    #representative_container {
        width: 34;
        padding-right: 1;
    }
    #representative_groups {
        height: 1fr;
    }
    #detail_pane {
        width: 1fr;
    }
    .pane {
        border: round yellow;
        background: ansi_default;
    }
    .pane-active {
        border: round green;
    }
    .pane-title {
        padding: 0 1;
        text-style: bold;
    }
    .group-box {
        height: 1fr;
        border: round yellow;
        background: ansi_default;
        margin-top: 1;
    }
    .group-box-active {
        border: round green;
    }
    .group-title {
        padding: 0 1;
        text-style: bold;
        color: yellow;
    }
    .group-title-active {
        color: green;
    }
    .group-divider {
        height: 1;
        color: yellow;
        margin: 0 1;
    }
    .group-divider-active {
        color: green;
    }
    ListView {
        height: 1fr;
        background: ansi_default;
    }
    ListItem {
        background: transparent;
    }
    ListView > ListItem.-highlight,
    ListView:focus > ListItem.-highlight {
        background: transparent;
        color: $foreground;
        text-style: none;
    }
    ListItem.row-current,
    ListView > ListItem.row-current.-highlight,
    ListView:focus > ListItem.row-current.-highlight {
        background: #00a651;
        color: #ffffff;
        text-style: bold;
    }
    ListItem.row-current > Label,
    ListView > ListItem.row-current.-highlight > Label,
    ListView:focus > ListItem.row-current.-highlight > Label {
        background: #00a651;
        color: #ffffff;
        text-style: bold;
    }
    ListItem.transferred-row,
    ListView > ListItem.transferred-row.-highlight,
    ListView:focus > ListItem.transferred-row.-highlight {
        background: #6f5a00;
        color: #fff3a0;
    }
    ListItem.transferred-row > Label,
    ListView > ListItem.transferred-row.-highlight > Label,
    ListView:focus > ListItem.transferred-row.-highlight > Label {
        background: #6f5a00;
        color: #fff3a0;
    }
    ListItem.row-current.transferred-row,
    ListView > ListItem.row-current.transferred-row.-highlight,
    ListView:focus > ListItem.row-current.transferred-row.-highlight {
        background: #c9ab00;
        color: #1f1a00;
        text-style: bold;
    }
    ListItem.row-current.transferred-row > Label,
    ListView > ListItem.row-current.transferred-row.-highlight > Label,
    ListView:focus > ListItem.row-current.transferred-row.-highlight > Label {
        background: #c9ab00;
        color: #1f1a00;
        text-style: bold;
    }
    #detail_title {
        text-style: bold;
        color: $accent;
    }
    #detail_scroll {
        height: 1fr;
        padding: 0 1 1 1;
        scrollbar-size-vertical: 1;
    }
    #detail_body {
        width: 100%;
    }
    #status_bar {
        height: 1;
        padding: 0 1;
        background: #1b1f24;
        color: #f2f2f2;
    }
    #prompt_dialog,
    #confirm_dialog,
    #help_dialog,
    #editor_dialog,
    #status_picker_dialog,
    #note_color_dialog {
        width: 88;
        max-width: 110;
        height: auto;
        border: round $accent;
        padding: 1;
        background: rgba(24, 30, 38, 0.9);
    }
    #editor_dialog {
        width: 110;
        height: 28;
    }
    #write_dialog {
        width: 92;
        height: 22;
        min-height: 18;
        border: round $accent;
        padding: 1;
        background: rgba(24, 30, 38, 0.9);
    }
    #subaction_dialog {
        width: 92;
        height: 26;
        min-height: 22;
        border: round $accent;
        padding: 1;
        background: rgba(24, 30, 38, 0.9);
    }
    #subaction_title_input {
        color: #d97706;
    }
    .editor_label,
    #prompt_hint,
    #confirm_hint,
    #note_color_hint,
    #status_picker_hint,
    #write_controls,
    #write_hint,
    #subaction_controls,
    #subaction_hint,
    #help_hint,
    #editor_hint {
        margin-top: 1;
    }
    #write_controls {
        padding: 0 1;
        border: round transparent;
        background: transparent;
    }
    #write_controls.write-controls-focused {
        border: round green;
    }
    #notes_input {
        height: 10;
        background: $surface;
        margin-top: 1;
        border: round $panel;
        padding: 0 1;
    }
    #status_input {
        height: auto;
        min-height: 3;
        background: $surface;
        margin-top: 1;
        border: round $panel;
        padding: 0 1;
    }
    #status_picker_input {
        height: auto;
        min-height: 3;
        background: $surface;
        margin-top: 1;
        border: round $panel;
        padding: 0 1;
    }
    #write_text_input {
        height: 1fr;
        min-height: 10;
        background: $surface;
        margin-top: 1;
    }
    #write_hint {
        width: 1fr;
        height: auto;
    }
    #editor_buttons {
        margin-top: 1;
        height: auto;
    }
    .editor_row {
        height: auto;
        margin-top: 1;
    }
    #today_group {
        margin-top: 0;
    }
    Footer {
        background: #1b1f24;
        color: #f2f2f2;
    }
    Footer .footer--key {
        background: #2b3138;
        color: #ffb14a;
        text-style: bold;
    }
    Footer .footer--description {
        color: #f2f2f2;
    }
    """

    BINDINGS = [
        Binding("up", "nav_up", "Up", priority=True),
        Binding("down", "nav_down", "Down", priority=True),
        Binding("left", "nav_left", "Back", priority=True),
        Binding("right", "nav_right", "Open", priority=True),
        Binding("enter", "nav_enter", "Enter", priority=True),
        Binding("tab", "nav_tab", show=False, priority=True),
        Binding("e", "open_selected", "Edit"),
        Binding("n", "context_note", "Notes"),
        Binding("shift+n", "context_note", show=False, priority=True),
        Binding("s", "edit_status", "Status"),
        Binding("shift+s", "edit_status", show=False, priority=True),
        Binding("a", "new_action", "Add"),
        Binding("c", "complete_action", "Complete"),
        Binding("x", "retire_action", "Retire"),
        Binding("u", "restore_action", "Restore"),
        Binding("d", "edit_due_date", "Due Date"),
        Binding("shift+d", "edit_due_date", show=False, priority=True),
        Binding("delete", "delete_graveyard", "Delete"),
        Binding("v", "toggle_archive_mode", "Completed/Grave"),
        Binding("r", "refresh_data", "Color"),
        Binding("shift+r", "refresh_data", show=False, priority=True),
        Binding("open_bracket", "prev_day", "Prev Day"),
        Binding("close_bracket", "next_day", "Next Day"),
        Binding("question_mark", "show_help", "Help"),
        Binding("q", "quit", "Quit"),
    ]

    def __init__(self, base_dir: Path | None = None, workbook_path: Path | None = None) -> None:
        super().__init__()
        self.base_dir = base_dir or Path(__file__).resolve().parent
        self.store = ActionsStore(self.base_dir, workbook_path=workbook_path or DEFAULT_WORKBOOK)
        self.selected_date = self.default_selected_date()
        self.archive_mode = "archive"
        self.group_ids = ["today", "projects", "archive"]
        self.group_list_ids = {
            "today": "today_list",
            "projects": "projects_list",
            "archive": "archive_list",
        }
        self.group_box_ids = {
            "today": "today_group",
            "projects": "projects_group",
            "archive": "archive_group",
        }
        self.group_title_ids = {
            "today": "today_group_title",
            "projects": "projects_group_title",
            "archive": "archive_group_title",
        }
        self.group_divider_ids = {
            "today": "today_group_divider",
            "projects": "projects_group_divider",
            "archive": "archive_group_divider",
        }
        self.group_entries: dict[str, list[RepresentativeEntry]] = {group: [] for group in self.group_ids}
        self.group_indexes: dict[str, int] = {group: 0 for group in self.group_ids}
        self.active_group = "today"
        self.detail_origin_group = "today"
        self.navigation_mode = "groups"
        self.return_mode = "groups"
        self.selected_key: str | None = "group:today"
        self.opened_key: str | None = "group:today"
        self.detail_back_stack: list[str] = []
        self.detail_table_indexes: dict[str, int] = {}
        self.detail_note_indexes: dict[str, int] = {}
        self.detail_today_index = 0
        self.detail_projects_table_index = 0
        self.detail_projects_mode = "tables"
        self.detail_projects_action_indexes: dict[str, int] = {}
        self.detail_archive_table_index = 0
        self.detail_archive_mode = "tables"
        self.detail_archive_action_indexes: dict[str, int] = {}
        self.refreshing_lists = False
        self.status_text = "Loaded actions prototype"

    def default_selected_date(self) -> date:
        actual_today = date.today()
        due_dates = []
        for action in self.store.active_actions():
            if not action.due_date:
                continue
            try:
                due_dates.append(date.fromisoformat(action.due_date))
            except ValueError:
                continue
        if actual_today in due_dates:
            return actual_today
        return max(due_dates) if due_dates else actual_today

    def default_new_due_date(self) -> str:
        return date.today().isoformat()

    def representative_today_date(self) -> date:
        return date.today()

    def today_detail_date(self) -> date:
        if self.navigation_mode == "detail" and self.opened_key == "group:today":
            return self.selected_date
        return self.representative_today_date()

    def check_action(self, action: str, parameters: tuple[object, ...]) -> bool | None:
        if action.startswith("nav_") and isinstance(self.screen, ModalScreen):
            return False
        return True

    def restore_navigation_focus(self) -> None:
        self.focus_navigation_target()

    def compose(self) -> ComposeResult:
        yield Vertical(
            Horizontal(
                Vertical(
                    Vertical(
                        Vertical(
                            Static("", id="today_group_title", classes="group-title"),
                            Static("─" * 28, id="today_group_divider", classes="group-divider"),
                            ListView(id="today_list"),
                            id="today_group",
                            classes="group-box group-box-active",
                        ),
                        Vertical(
                            Static("", id="projects_group_title", classes="group-title"),
                            Static("─" * 28, id="projects_group_divider", classes="group-divider"),
                            ListView(id="projects_list"),
                            id="projects_group",
                            classes="group-box",
                        ),
                        Vertical(
                            Static("", id="archive_group_title", classes="group-title"),
                            Static("─" * 28, id="archive_group_divider", classes="group-divider"),
                            ListView(id="archive_list"),
                            id="archive_group",
                            classes="group-box",
                        ),
                        id="representative_groups",
                    ),
                    id="representative_container",
                ),
                Vertical(
                    Static("Details", id="detail_title", classes="pane-title"),
                    VerticalScroll(Static("", id="detail_body"), id="detail_scroll"),
                    id="detail_pane",
                    classes="pane",
                ),
                id="content",
            ),
            Static("", id="status_bar"),
            Footer(),
            id="layout",
        )

    def on_mount(self) -> None:
        self.refresh_all()

    @on(ListView.Selected)
    def list_selected(self, event: ListView.Selected) -> None:
        if self.refreshing_lists or self.navigation_mode == "groups" or not event.list_view.has_focus:
            return
        group = self.group_for_list_id(event.list_view.id or "")
        if group:
            self.active_group = group
            self.group_indexes[group] = event.list_view.index or 0
            if self.navigation_mode != "detail":
                self.navigation_mode = "items"
            entry = self.current_left_entry()
            self.selected_key = self.entry_key(entry) if entry else None
            self.apply_focus_state()

    @on(ListView.Highlighted)
    def list_highlighted(self, event: ListView.Highlighted) -> None:
        if self.refreshing_lists or self.navigation_mode == "groups" or not event.list_view.has_focus:
            return
        group = self.group_for_list_id(event.list_view.id or "")
        if group:
            self.active_group = group
            self.group_indexes[group] = event.list_view.index or 0
            if self.navigation_mode == "items":
                entry = self.current_left_entry()
                self.selected_key = self.entry_key(entry) if entry else None
            self.apply_focus_state()

    def refresh_all(self) -> None:
        self.refresh_representative_pane()
        self.apply_focus_state()
        self.update_details()
        self.update_status(self.status_text)

    def refresh_representative_pane(self) -> None:
        self.group_entries = self.build_representative_entries()
        selected_group = self.find_group_for_key(self.selected_key)
        if selected_group:
            self.active_group = selected_group
        self.query_one("#today_group_title", Static).update(f"Today ({len(self.group_entries['today'])})")
        self.query_one("#projects_group_title", Static).update(f"Projects ({len(self.group_entries['projects'])})")
        archive_mode = "Completed" if self.archive_mode == "archive" else "Graveyard"
        self.query_one("#archive_group_title", Static).update(f"Completed [{archive_mode}] ({len(self.group_entries['archive'])})")
        self.refreshing_lists = True
        for group in self.group_ids:
            self.populate_group_list(group, self.group_entries[group])
        self.refreshing_lists = False
        if selected_group:
            self.active_group = selected_group
        if self.navigation_mode == "detail":
            self.selected_key = self.selected_key or f"group:{self.active_group}"
        elif self.navigation_mode == "groups":
            self.selected_key = f"group:{self.active_group}"
            self.opened_key = f"group:{self.active_group}"
        else:
            entry = self.current_left_entry()
            self.selected_key = self.entry_key(entry) if entry else None

    def build_representative_entries(self) -> dict[str, list[RepresentativeEntry]]:
        entries: dict[str, list[RepresentativeEntry]] = {group: [] for group in self.group_ids}
        today_actions = self.store.today_actions(self.representative_today_date())
        for action in today_actions:
            entries["today"].append(
                RepresentativeEntry(
                    kind="action",
                    group="today",
                    label=action.title,
                    action_id=action.id,
                    table_id=action.table_id,
                    transferred=action.status == "Transferred",
                )
            )

        tables = self.store.all_tables()
        for table in tables:
            entries["projects"].append(
                RepresentativeEntry(
                    kind="table",
                    group="projects",
                    label=table.name,
                    table_id=table.id,
                )
            )

        for table in self.current_archive_tables():
            entries["archive"].append(
                RepresentativeEntry(
                    kind="table",
                    group="archive",
                    label=table.name,
                    table_id=table.id,
                )
            )
        return entries

    def populate_group_list(self, group: str, entries: Iterable[RepresentativeEntry]) -> None:
        list_view = self.query_one(f"#{self.group_list_ids[group]}", ListView)
        previous_key = self.selected_key if self.find_group_for_key(self.selected_key) == group else None
        list_view.clear()
        entry_list = list(entries)
        for entry in entry_list:
            item_classes = "transferred-row" if entry.transferred else ""
            list_view.append(ListItem(Label(entry.label), classes=item_classes))
        if not entry_list:
            list_view.append(ListItem(Label("(empty)")))
            list_view.index = 0
            self.group_indexes[group] = 0
            return
        target_index = 0
        if previous_key:
            for index, entry in enumerate(entry_list):
                if self.entry_key(entry) == previous_key:
                    target_index = index
                    break
        else:
            target_index = min(self.group_indexes[group], len(entry_list) - 1)
        list_view.index = target_index
        self.group_indexes[group] = target_index

    def refresh_group_classes(self) -> None:
        for group in self.group_ids:
            box = self.query_one(f"#{self.group_box_ids[group]}", Vertical)
            title = self.query_one(f"#{self.group_title_ids[group]}", Static)
            divider = self.query_one(f"#{self.group_divider_ids[group]}", Static)
            is_active = group == self.active_group
            if is_active and not box.has_class("group-box-active"):
                box.add_class("group-box-active")
            if not is_active and box.has_class("group-box-active"):
                box.remove_class("group-box-active")
            if is_active and not title.has_class("group-title-active"):
                title.add_class("group-title-active")
            if not is_active and title.has_class("group-title-active"):
                title.remove_class("group-title-active")
            if is_active and not divider.has_class("group-divider-active"):
                divider.add_class("group-divider-active")
            if not is_active and divider.has_class("group-divider-active"):
                divider.remove_class("group-divider-active")

    def refresh_detail_class(self) -> None:
        detail_pane = self.query_one("#detail_pane", Vertical)
        is_active = self.navigation_mode == "detail"
        if is_active and not detail_pane.has_class("pane-active"):
            detail_pane.add_class("pane-active")
        if not is_active and detail_pane.has_class("pane-active"):
            detail_pane.remove_class("pane-active")

    def decorate_current_rows(self) -> None:
        for group in self.group_ids:
            list_view = self.query_one(f"#{self.group_list_ids[group]}", ListView)
            active_index = list_view.index if self.navigation_mode == "items" and group == self.active_group else -1
            for index, child in enumerate(list_view.children):
                if isinstance(child, ListItem):
                    is_current = index == active_index
                    if is_current and not child.has_class("row-current"):
                        child.add_class("row-current")
                    if not is_current and child.has_class("row-current"):
                        child.remove_class("row-current")

    def focus_navigation_target(self) -> None:
        if self.navigation_mode == "detail":
            self.set_focus(self.query_one("#detail_scroll", VerticalScroll))
        elif self.navigation_mode == "groups":
            self.set_focus(None)
        else:
            self.set_focus(self.query_one(f"#{self.group_list_ids[self.active_group]}", ListView))

    def apply_focus_state(self) -> None:
        self.refresh_group_classes()
        self.refresh_detail_class()
        self.decorate_current_rows()
        self.focus_navigation_target()

    def current_left_entry(self) -> RepresentativeEntry | None:
        if self.navigation_mode == "groups":
            return RepresentativeEntry(kind="group", group=self.active_group, label=self.active_group.title())
        entries = self.group_entries.get(self.active_group, [])
        if not entries:
            return None
        list_view = self.query_one(f"#{self.group_list_ids[self.active_group]}", ListView)
        index = min(max(list_view.index or 0, 0), len(entries) - 1)
        self.group_indexes[self.active_group] = index
        return entries[index]

    def current_entry(self) -> RepresentativeEntry | None:
        if self.navigation_mode == "detail":
            return self.entry_for_key(self.opened_key)
        return self.current_left_entry()

    def group_for_list_id(self, list_id: str) -> str | None:
        for group, candidate in self.group_list_ids.items():
            if candidate == list_id:
                return group
        return None

    def find_group_for_key(self, key: str | None) -> str | None:
        if key is None:
            return None
        if key.startswith("group:"):
            return key.split(":", 1)[1]
        for group, entries in self.group_entries.items():
            for entry in entries:
                if self.entry_key(entry) == key:
                    return group
        return None

    def entry_for_key(self, key: str | None) -> RepresentativeEntry | None:
        if key is None:
            return None
        if key.startswith("group:"):
            group = key.split(":", 1)[1]
            return RepresentativeEntry(kind="group", group=group, label=group.title())
        for entries in self.group_entries.values():
            for entry in entries:
                if self.entry_key(entry) == key:
                    return entry
        if key.startswith("action:"):
            action_id = key.split(":", 1)[1]
            found = self.store.action_by_id(action_id)
            if found is None:
                return None
            bucket_name, action = found
            table = self.store.table_for_id(action.table_id)
            group = "archive" if bucket_name in {"archive", "graveyard"} else "projects"
            return RepresentativeEntry(
                kind="action",
                group=group,
                label=action.title,
                action_id=action.id,
                table_id=action.table_id,
            )
        if key.startswith("table:"):
            parts = key.split(":", 2)
            if len(parts) == 3:
                _, group, table_id = parts
            else:
                group = "projects"
                table_id = parts[1]
            table = self.store.table_for_id(table_id)
            if table is None:
                return None
            return RepresentativeEntry(kind="table", group=group, label=table.name, table_id=table.id)
        return None

    def current_table_context(self) -> TableRecord | None:
        entry = self.current_entry()
        if entry and entry.table_id:
            return self.store.table_for_id(entry.table_id)
        return None

    def current_project_table_context(self) -> TableRecord | None:
        if self.navigation_mode == "detail" and self.opened_key == "group:projects" and self.detail_projects_mode == "tables":
            return self.current_detail_projects_table()
        entry = self.current_entry()
        if entry is None or entry.kind != "table" or entry.group != "projects" or not entry.table_id:
            return None
        return self.store.table_for_id(entry.table_id)

    def current_new_action_table(self) -> TableRecord | None:
        if self.navigation_mode == "detail" and self.opened_key == "group:projects":
            return self.current_detail_projects_table()
        return self.current_table_context()

    def current_action_detail_action(self) -> ActionRecord | None:
        if self.navigation_mode != "detail" or not (self.opened_key or "").startswith("action:"):
            return None
        found = self.store.action_by_id(self.opened_key.split(":", 1)[1])
        return found[1] if found else None

    def current_action(self) -> ActionRecord | None:
        detail_action = self.current_detail_projects_action()
        if detail_action is not None:
            return detail_action
        detail_action = self.current_detail_archive_action()
        if detail_action is not None:
            return detail_action
        detail_action = self.current_detail_today_action()
        if detail_action is not None:
            return detail_action
        detail_action = self.current_detail_table_action()
        if detail_action is not None:
            return detail_action
        entry = self.current_entry()
        if entry is None or not entry.action_id:
            return None
        found = self.store.action_by_id(entry.action_id)
        return found[1] if found else None

    def focused_action_for_notes(self) -> ActionRecord | None:
        return self.current_action()

    def focused_action_for_due_date(self) -> ActionRecord | None:
        return self.current_action()

    def focused_action_for_status(self) -> ActionRecord | None:
        return self.focused_action_for_due_date()

    def status_edit_origin(self) -> str:
        if self.navigation_mode == "detail" and self.opened_key == "group:today":
            return "today"
        if self.navigation_mode != "detail" and self.active_group == "today":
            return "today"
        return "default"

    def entry_key(self, entry: RepresentativeEntry | None) -> str | None:
        if entry is None:
            return None
        if entry.kind == "group":
            return f"group:{entry.group}"
        if entry.action_id:
            return f"action:{entry.action_id}"
        if entry.table_id:
            return f"table:{entry.group}:{entry.table_id}"
        return entry.label

    def update_details(self, *, reset_scroll: bool = True) -> None:
        title, body = self.render_entry_details(self.entry_for_key(self.opened_key))
        self.query_one("#detail_title", Static).update(title)
        self.query_one("#detail_body", Static).update(body)
        detail_scroll = self.query_one("#detail_scroll", VerticalScroll)
        target_line = self.current_detail_scroll_line()
        if target_line is not None:
            self.ensure_detail_line_visible(detail_scroll, target_line, reset_scroll=reset_scroll)
            return
        if reset_scroll:
            detail_scroll.scroll_home(animate=False)

    def ensure_detail_line_visible(self, detail_scroll: VerticalScroll, target_line: int, *, reset_scroll: bool) -> None:
        viewport_height = max(int(getattr(detail_scroll.size, "height", 0)), 1)
        visible_top = 0 if reset_scroll else int(detail_scroll.scroll_y)
        visible_bottom = visible_top + viewport_height - 1
        if target_line < visible_top:
            detail_scroll.scroll_to(y=target_line, animate=False)
            return
        if target_line > visible_bottom:
            detail_scroll.scroll_to(y=max(target_line - viewport_height + 1, 0), animate=False)
            return
        if reset_scroll:
            detail_scroll.scroll_to(y=visible_top, animate=False)

    def render_entry_details(self, entry: RepresentativeEntry | None) -> tuple[str, str | Text]:
        if entry is None:
            return "Details", "No selection."
        if entry.kind == "group":
            if entry.group == "today":
                return f"Today {self.today_detail_date().isoformat()}", self.render_today_overview()
            if entry.group == "projects":
                return "Projects", self.render_projects_overview()
            return self.archive_title_word(), self.render_archive_overview()
        if entry.kind == "table":
            table = self.store.table_for_id(entry.table_id or "")
            if table is None:
                return "Project", "Unknown project."
            return table.name, self.render_table_section(table, group=entry.group)
        if entry.action_id:
            found = self.store.action_by_id(entry.action_id)
            if found is None:
                return "Action", "Action not found."
            bucket_name, action = found
            table = self.store.table_for_id(action.table_id)
            lines = [
                f"Table: {table.name if table else 'Unknown'}",
                f"State: {bucket_name}",
                f"Due: {action.due_date or 'unscheduled'}",
                f"Time: {action.estimate or '-'}",
                f"Status: {action.status or '-'}",
            ]
            if action.completed_at:
                lines.append(f"Completed At: {action.completed_at}")
            if action.retired_at:
                lines.append(f"Retired At: {action.retired_at}")
            return action.title, Group(
                self.render_action_info_panel(lines),
                Text(""),
                self.render_action_notes_detail(action),
            )
        return "Details", entry.label

    def render_today_overview(self) -> str | Text:
        display_date = self.today_detail_date()
        actions = self.store.today_actions(display_date)
        summary = f"Actions due: {len(actions)}"
        if not actions:
            block_lines = ["", *self.plain_table_block_lines(summary, ["No actions due on this date."])]
            return self.render_lines_with_selection(block_lines)
        block_lines = ["", *self.plain_table_block_lines(summary, self.format_action_table(actions))]
        if self.opened_key != "group:today" or self.navigation_mode != "detail":
            return self.render_lines_with_selection(block_lines)
        selected_index = self.detail_today_index_for(len(actions))
        return self.render_lines_with_selection(block_lines, selected_line_index=6 + selected_index)

    def render_projects_overview(self) -> str | Text:
        tables = self.store.all_tables()
        if not tables:
            return "No projects yet."
        if self.opened_key != "group:projects" or self.navigation_mode != "detail":
            lines: list[str | Text] = []
            for index, table in enumerate(tables):
                if index:
                    lines.append("")
                lines.extend(self.table_overview_lines(table))
            return self.render_lines_with_selection(lines)
        selected_table_index = self.detail_projects_table_index_for(len(tables))
        selected_line_indexes: set[int] = set()
        lines: list[str | Text] = []
        for table_index, table in enumerate(tables):
            if lines:
                lines.append("")
            block_start = len(lines)
            block_lines = self.table_overview_lines(table)
            lines.extend(block_lines)
            if table_index != selected_table_index:
                continue
            if self.detail_projects_mode == "actions":
                actions = self.store.actions_for_table(table.id)
                if actions:
                    selected_action_index = self.detail_projects_action_index_for(table.id, len(actions))
                    selected_line_indexes.add(block_start + 5 + selected_action_index)
                else:
                    selected_line_indexes.add(block_start)
            else:
                selected_line_indexes.add(block_start)
        return self.render_lines_with_selection(lines, selected_line_indexes=selected_line_indexes)

    def render_archive_overview(self) -> str | Text:
        actions = self.current_archive_actions()
        lines = [
            f"View: {self.archive_title_word()}",
            f"Actions: {len(actions)}",
            "",
        ]
        if not actions:
            lines.append(f"No actions in {self.archive_title_word().lower()}.")
            return "\n".join(lines)
        tables = self.current_archive_tables()
        if self.opened_key != "group:archive" or self.navigation_mode != "detail":
            detail_lines: list[str | Text] = list(lines)
            for table in tables:
                if detail_lines:
                    detail_lines.append("")
                detail_lines.extend(self.table_overview_lines(table, group="archive"))
            return self.render_lines_with_selection(detail_lines)
        selected_table_index = self.detail_archive_table_index_for(len(tables))
        selected_line_indexes: set[int] = set()
        detail_lines: list[str | Text] = list(lines)
        for table_index, table in enumerate(tables):
            if detail_lines:
                detail_lines.append("")
            block_start = len(detail_lines)
            block_lines = self.table_overview_lines(table, group="archive")
            detail_lines.extend(block_lines)
            if table_index != selected_table_index:
                continue
            if self.detail_archive_mode == "actions":
                table_actions = self.actions_for_table_entry(
                    RepresentativeEntry(kind="table", group="archive", label=table.name, table_id=table.id)
                )
                if table_actions:
                    selected_action_index = self.detail_archive_action_index_for(table.id, len(table_actions))
                    selected_line_indexes.add(block_start + 5 + selected_action_index)
                else:
                    selected_line_indexes.add(block_start)
            else:
                selected_line_indexes.add(block_start)
        return self.render_lines_with_selection(detail_lines, selected_line_indexes=selected_line_indexes)

    def render_table_section(self, table: TableRecord, *, group: str = "projects") -> str | Text:
        entry = RepresentativeEntry(kind="table", group=group, label=table.name, table_id=table.id)
        actions = self.actions_for_table_entry(entry)
        block_lines = self.table_overview_lines(table, group=group)
        if self.navigation_mode != "detail":
            return self.render_lines_with_selection(block_lines)
        if not actions:
            return self.render_lines_with_selection(block_lines)
        selected_index = self.detail_table_index_for(table.id, len(actions))
        return self.render_lines_with_selection(block_lines, selected_line_index=5 + selected_index)

    def table_overview_lines(self, table: TableRecord, *, group: str = "projects") -> list[str | Text]:
        entry = RepresentativeEntry(kind="table", group=group, label=table.name, table_id=table.id)
        actions = self.actions_for_table_entry(entry)
        action_label = "Open actions" if group == "projects" else f"{self.archive_title_word()} actions"
        summary = f"{action_label}: {len(actions)}"
        if not actions:
            empty_label = "active actions" if group == "projects" else self.archive_title_word().lower()
            content_lines = [f"No {empty_label}."]
            return self.table_block_lines(table.name, summary, content_lines)
        content_lines = self.format_action_table(actions)
        return self.table_block_lines(table.name, summary, content_lines)

    def render_table_overview(self, table: TableRecord, *, group: str = "projects") -> Text:
        return self.render_lines_with_selection(self.table_overview_lines(table, group=group))

    def current_archive_actions(self) -> list[ActionRecord]:
        return self.store.archive_actions() if self.archive_mode == "archive" else self.store.graveyard_actions()

    def current_archive_tables(self) -> list[TableRecord]:
        archive_table_ids = {action.table_id for action in self.current_archive_actions()}
        return [table for table in self.store.all_known_tables() if table.id in archive_table_ids]

    def archive_title_word(self) -> str:
        return "Completed" if self.archive_mode == "archive" else "Graveyard"

    def current_detail_scroll_line(self) -> int | None:
        if self.opened_key == "group:today" and self.navigation_mode == "detail":
            actions = self.store.today_actions(self.selected_date)
            if actions:
                return 6 + self.detail_today_index_for(len(actions))
            return None
        if self.opened_key == "group:projects" and self.navigation_mode == "detail":
            tables = self.store.all_tables()
            if not tables:
                return None
            lines_before = 0
            selected_table_index = self.detail_projects_table_index_for(len(tables))
            for table_index, table in enumerate(tables):
                if table_index == selected_table_index:
                    if self.detail_projects_mode == "actions":
                        actions = self.store.actions_for_table(table.id)
                        if actions:
                            return lines_before + 5 + self.detail_projects_action_index_for(table.id, len(actions))
                    return lines_before
                block_lines = self.table_overview_lines(table)
                lines_before += len(block_lines) + 1
            return None
        if self.opened_key == "group:archive" and self.navigation_mode == "detail":
            tables = self.current_archive_tables()
            if not tables:
                return None
            lines_before = 4
            selected_table_index = self.detail_archive_table_index_for(len(tables))
            for table_index, table in enumerate(tables):
                if table_index == selected_table_index:
                    if self.detail_archive_mode == "actions":
                        actions = self.actions_for_table_entry(
                            RepresentativeEntry(kind="table", group="archive", label=table.name, table_id=table.id)
                        )
                        if actions:
                            return lines_before + 5 + self.detail_archive_action_index_for(table.id, len(actions))
                    return lines_before
                block_lines = self.table_overview_lines(table, group="archive")
                lines_before += len(block_lines) + 1
            return None
        if self.navigation_mode != "detail":
            return None
        entry = self.entry_for_key(self.opened_key)
        if entry is None or entry.kind != "table" or not entry.table_id:
            return None
        actions = self.actions_for_table_entry(entry)
        if not actions:
            return None
        return 5 + self.detail_table_index_for(entry.table_id, len(actions))

    def detail_today_index_for(self, total_actions: int) -> int:
        if total_actions <= 0:
            return 0
        self.detail_today_index = min(max(self.detail_today_index, 0), total_actions - 1)
        return self.detail_today_index

    def detail_table_index_for(self, table_id: str, total_actions: int) -> int:
        if total_actions <= 0:
            return 0
        stored_index = self.detail_table_indexes.get(table_id, 0)
        clamped_index = min(max(stored_index, 0), total_actions - 1)
        self.detail_table_indexes[table_id] = clamped_index
        return clamped_index

    def detail_projects_table_index_for(self, total_tables: int) -> int:
        if total_tables <= 0:
            return 0
        self.detail_projects_table_index = min(max(self.detail_projects_table_index, 0), total_tables - 1)
        return self.detail_projects_table_index

    def detail_projects_action_index_for(self, table_id: str, total_actions: int) -> int:
        if total_actions <= 0:
            return 0
        stored_index = self.detail_projects_action_indexes.get(table_id, 0)
        clamped_index = min(max(stored_index, 0), total_actions - 1)
        self.detail_projects_action_indexes[table_id] = clamped_index
        return clamped_index

    def detail_archive_table_index_for(self, total_tables: int) -> int:
        if total_tables <= 0:
            return 0
        self.detail_archive_table_index = min(max(self.detail_archive_table_index, 0), total_tables - 1)
        return self.detail_archive_table_index

    def detail_archive_action_index_for(self, table_id: str, total_actions: int) -> int:
        if total_actions <= 0:
            return 0
        stored_index = self.detail_archive_action_indexes.get(table_id, 0)
        clamped_index = min(max(stored_index, 0), total_actions - 1)
        self.detail_archive_action_indexes[table_id] = clamped_index
        return clamped_index

    def current_detail_projects_table(self) -> TableRecord | None:
        if self.opened_key != "group:projects":
            return None
        tables = self.store.all_tables()
        if not tables:
            return None
        return tables[self.detail_projects_table_index_for(len(tables))]

    def current_detail_projects_action(self) -> ActionRecord | None:
        if self.opened_key != "group:projects" or self.detail_projects_mode != "actions":
            return None
        table = self.current_detail_projects_table()
        if table is None:
            return None
        actions = self.store.actions_for_table(table.id)
        if not actions:
            return None
        selected_index = self.detail_projects_action_index_for(table.id, len(actions))
        return actions[selected_index]

    def current_detail_archive_table(self) -> TableRecord | None:
        if self.opened_key != "group:archive":
            return None
        tables = self.current_archive_tables()
        if not tables:
            return None
        return tables[self.detail_archive_table_index_for(len(tables))]

    def current_detail_archive_action(self) -> ActionRecord | None:
        if self.opened_key != "group:archive" or self.detail_archive_mode != "actions":
            return None
        table = self.current_detail_archive_table()
        if table is None:
            return None
        actions = self.actions_for_table_entry(
            RepresentativeEntry(kind="table", group="archive", label=table.name, table_id=table.id)
        )
        if not actions:
            return None
        selected_index = self.detail_archive_action_index_for(table.id, len(actions))
        return actions[selected_index]

    def current_detail_today_action(self) -> ActionRecord | None:
        if self.opened_key != "group:today" or self.navigation_mode != "detail":
            return None
        actions = self.store.today_actions(self.selected_date)
        if not actions:
            return None
        return actions[self.detail_today_index_for(len(actions))]

    def current_detail_table_action(self) -> ActionRecord | None:
        entry = self.entry_for_key(self.opened_key)
        if entry is None or entry.kind != "table" or not entry.table_id:
            return None
        actions = self.actions_for_table_entry(entry)
        if not actions:
            return None
        selected_index = self.detail_table_index_for(entry.table_id, len(actions))
        return actions[selected_index]

    def actions_for_table_entry(self, entry: RepresentativeEntry) -> list[ActionRecord]:
        if not entry.table_id:
            return []
        if entry.group == "archive":
            return [action for action in self.current_archive_actions() if action.table_id == entry.table_id]
        return self.store.actions_for_table(entry.table_id)

    def detail_note_bullets_for(self, action: ActionRecord) -> list[NoteBullet]:
        return extract_note_bullets(action.notes)

    def detail_note_index_for(self, action_id: str, total_bullets: int) -> int:
        if total_bullets <= 0:
            return 0
        stored_index = self.detail_note_indexes.get(action_id, 0)
        clamped_index = min(max(stored_index, 0), total_bullets - 1)
        self.detail_note_indexes[action_id] = clamped_index
        return clamped_index

    def seed_last_note_focus(self, action_id: str) -> None:
        found = self.store.action_by_id(action_id)
        if found is None:
            return
        bullets = self.detail_note_bullets_for(found[1])
        if not bullets:
            self.detail_note_indexes.pop(action_id, None)
            return
        self.detail_note_indexes[action_id] = len(bullets) - 1

    def current_detail_note_bullet(self) -> tuple[ActionRecord, NoteBullet] | None:
        entry = self.entry_for_key(self.opened_key)
        if entry is None or not entry.action_id:
            return None
        found = self.store.action_by_id(entry.action_id)
        if found is None:
            return None
        action = found[1]
        bullets = self.detail_note_bullets_for(action)
        if not bullets:
            return None
        selected_index = self.detail_note_index_for(action.id, len(bullets))
        return action, bullets[selected_index]

    def move_detail_table_selection(self, offset: int) -> bool:
        entry = self.entry_for_key(self.opened_key)
        if entry is None or entry.kind != "table" or not entry.table_id:
            return False
        actions = self.actions_for_table_entry(entry)
        if not actions:
            return False
        current_index = self.detail_table_index_for(entry.table_id, len(actions))
        next_index = min(max(current_index + offset, 0), len(actions) - 1)
        self.detail_table_indexes[entry.table_id] = next_index
        return next_index != current_index

    def move_detail_today_selection(self, offset: int) -> bool:
        if self.opened_key != "group:today" or self.navigation_mode != "detail":
            return False
        actions = self.store.today_actions(self.selected_date)
        if not actions:
            return False
        current_index = self.detail_today_index_for(len(actions))
        next_index = min(max(current_index + offset, 0), len(actions) - 1)
        self.detail_today_index = next_index
        return next_index != current_index

    def move_detail_projects_selection(self, offset: int) -> bool:
        if self.opened_key != "group:projects":
            return False
        if self.detail_projects_mode == "actions":
            table = self.current_detail_projects_table()
            if table is None:
                return False
            actions = self.store.actions_for_table(table.id)
            if not actions:
                self.detail_projects_mode = "tables"
                return False
            current_index = self.detail_projects_action_index_for(table.id, len(actions))
            next_index = min(max(current_index + offset, 0), len(actions) - 1)
            self.detail_projects_action_indexes[table.id] = next_index
            return next_index != current_index
        tables = self.store.all_tables()
        if not tables:
            return False
        current_index = self.detail_projects_table_index_for(len(tables))
        next_index = min(max(current_index + offset, 0), len(tables) - 1)
        self.detail_projects_table_index = next_index
        return next_index != current_index

    def move_detail_archive_selection(self, offset: int) -> bool:
        if self.opened_key != "group:archive":
            return False
        if self.detail_archive_mode == "actions":
            table = self.current_detail_archive_table()
            if table is None:
                return False
            actions = self.actions_for_table_entry(
                RepresentativeEntry(kind="table", group="archive", label=table.name, table_id=table.id)
            )
            if not actions:
                self.detail_archive_mode = "tables"
                return False
            current_index = self.detail_archive_action_index_for(table.id, len(actions))
            next_index = min(max(current_index + offset, 0), len(actions) - 1)
            self.detail_archive_action_indexes[table.id] = next_index
            return next_index != current_index
        tables = self.current_archive_tables()
        if not tables:
            return False
        current_index = self.detail_archive_table_index_for(len(tables))
        next_index = min(max(current_index + offset, 0), len(tables) - 1)
        self.detail_archive_table_index = next_index
        return next_index != current_index

    def enter_detail_projects_actions(self) -> bool:
        if self.opened_key != "group:projects":
            return False
        table = self.current_detail_projects_table()
        if table is None:
            return False
        actions = self.store.actions_for_table(table.id)
        if not actions:
            return False
        self.detail_projects_mode = "actions"
        self.detail_projects_action_index_for(table.id, len(actions))
        return True

    def exit_detail_projects_actions(self) -> bool:
        if self.opened_key != "group:projects" or self.detail_projects_mode != "actions":
            return False
        self.detail_projects_mode = "tables"
        return True

    def enter_detail_archive_actions(self) -> bool:
        if self.opened_key != "group:archive":
            return False
        table = self.current_detail_archive_table()
        if table is None:
            return False
        actions = self.actions_for_table_entry(
            RepresentativeEntry(kind="table", group="archive", label=table.name, table_id=table.id)
        )
        if not actions:
            return False
        self.detail_archive_mode = "actions"
        self.detail_archive_action_index_for(table.id, len(actions))
        return True

    def exit_detail_archive_actions(self) -> bool:
        if self.opened_key != "group:archive" or self.detail_archive_mode != "actions":
            return False
        self.detail_archive_mode = "tables"
        return True

    def move_detail_note_selection(self, offset: int) -> bool:
        current = self.current_detail_note_bullet()
        if current is None:
            return False
        action, bullet = current
        bullets = self.detail_note_bullets_for(action)
        next_index = min(max(bullet.index + offset, 0), len(bullets) - 1)
        self.detail_note_indexes[action.id] = next_index
        return next_index != bullet.index

    def render_action_notes_detail(self, action: ActionRecord, lines: list[str]) -> str | Text:
        bullets = self.detail_note_bullets_for(action)
        note_lines = display_note_text(action.notes).splitlines() if action.notes else ["(no notes)"]
        lines = list(lines)
        lines.append("")
        lines.append("──────────────── Notes ────────────────")
        if bullets:
            lines.append("Notes: Up/Down selects bullets. Right, Enter, or Tab edits the focused bullet.")
        else:
            lines.append("Notes:")
        note_start_index = len(lines)
        lines.extend(note_lines)
        if not bullets:
            return "\n".join(lines)
        selected_bullet_index = self.detail_note_index_for(action.id, len(bullets))
        selected_bullet = bullets[selected_bullet_index]
        line_styles: dict[int, str] = {}
        for bullet in bullets:
            if not bullet.color:
                continue
            style = bullet.color
            if bullet.index == selected_bullet_index:
                style = f"bold {bullet.color} on #5a7a1f"
            for line_index in range(bullet.start_line, bullet.end_line):
                line_styles[note_start_index + line_index] = style
        selected_line_indexes = {
            note_start_index + line_index
            for line_index in range(selected_bullet.start_line, selected_bullet.end_line)
        }
        return self.render_lines_with_selection(
            lines,
            selected_line_indexes=selected_line_indexes,
            line_styles=line_styles,
        )

    def render_action_info_panel(self, lines: list[str]) -> Panel:
        return Panel(
            Text("\n".join(lines)),
            title="Info",
            title_align="left",
            border_style="yellow",
            expand=True,
            padding=(0, 1),
        )

    def render_action_notes_detail(self, action: ActionRecord) -> str | Text:
        bullets = self.detail_note_bullets_for(action)
        note_lines = display_note_text(action.notes).splitlines() if action.notes else []
        lines: list[str] = []
        note_start_index = 0
        lines.extend(note_lines)
        line_styles: dict[int, str] = {}
        selected_line_styles: dict[int, str] = {}
        if not bullets:
            return self.render_lines_with_selection(lines, line_styles=line_styles)
        selected_bullet_index = self.detail_note_index_for(action.id, len(bullets))
        selected_bullet = bullets[selected_bullet_index]
        is_detail_active = self.navigation_mode == "detail"
        for bullet in bullets:
            if not bullet.color:
                if is_detail_active and bullet.index == selected_bullet_index:
                    for line_index in range(bullet.start_line, bullet.end_line):
                        selected_line_styles[note_start_index + line_index] = "bold on #496131"
                continue
            style = bullet.color
            for line_index in range(bullet.start_line, bullet.end_line):
                line_styles[note_start_index + line_index] = style
                if is_detail_active and bullet.index == selected_bullet_index:
                    selected_line_styles[note_start_index + line_index] = f"bold {bullet.color} on #496131"
        selected_line_indexes = (
            {
                note_start_index + line_index
                for line_index in range(selected_bullet.start_line, selected_bullet.end_line)
            }
            if is_detail_active
            else set()
        )
        return self.render_lines_with_selection(
            lines,
            selected_line_indexes=selected_line_indexes,
            line_styles=line_styles,
            selected_line_styles=selected_line_styles,
        )

    def render_lines_with_selection(
        self,
        lines: list[str | Text],
        *,
        selected_line_index: int | None = None,
        selected_line_indexes: set[int] | None = None,
        line_styles: dict[int, str] | None = None,
        selected_line_styles: dict[int, str] | None = None,
    ) -> Text:
        text = Text()
        active_indexes = set(selected_line_indexes or set())
        style_map = line_styles or {}
        selected_style_map = selected_line_styles or {}
        if selected_line_index is not None:
            active_indexes.add(selected_line_index)
        for index, line in enumerate(lines):
            rendered = line.copy() if isinstance(line, Text) else Text(line)
            style = style_map.get(index)
            if style:
                rendered.stylize(style)
            if index in active_indexes:
                selected_style = selected_style_map.get(index)
                if selected_style:
                    rendered.stylize(selected_style)
                elif isinstance(line, Text) and rendered.plain.startswith("╭─┤ "):
                    title_start = len("╭─┤ ")
                    title_end = rendered.plain.find(" ├", title_start)
                    if title_end > title_start:
                        rendered.stylize("bold white on #00a651", title_start, title_end)
                    else:
                        rendered.stylize("bold white on #00a651")
                elif isinstance(line, Text) and rendered.plain.startswith("│ ") and rendered.plain.endswith(" │"):
                    rendered.stylize("bold white on #00a651", 2, len(rendered.plain) - 2)
                else:
                    rendered.stylize("bold white on #00a651")
            elif index == 0:
                rendered.stylize("bold")
            text.append_text(rendered)
            if index < len(lines) - 1:
                text.append("\n")
        return text

    def table_block_inner_width(self) -> int:
        return max(self.detail_table_width() - 4, 56)

    def style_table_row_line(self, line: str) -> Text:
        text = Text()
        border_chars = {"│", "─", "┼"}
        for char in line:
            style = TABLE_BORDER_COLOR if char in border_chars else None
            text.append(char, style=style)
        return text

    def wrap_table_content_line(self, line: str, inner_width: int) -> Text:
        text = Text()
        text.append("│ ", style=TABLE_BORDER_COLOR)
        text.append_text(self.style_table_row_line(line.ljust(inner_width)))
        text.append(" │", style=TABLE_BORDER_COLOR)
        return text

    def table_block_lines(self, title: str, summary: str, content_lines: list[str]) -> list[str | Text]:
        inner_width = self.table_block_inner_width()
        title_text = self.summary_cell(title, max(inner_width - 4, 12))
        top_fill = max(inner_width - len(title_text) - 3, 0)
        top = Text(style=f"bold {TABLE_BORDER_COLOR}")
        top.append("╭─┤ ")
        top.append(title_text, style=f"bold {TABLE_BORDER_COLOR}")
        top.append(" ├")
        top.append("─" * top_fill)
        top.append("╮")
        divider = Text("├" + "─" * (inner_width + 2) + "┤", style=TABLE_BORDER_COLOR)
        bottom = Text("╰" + "─" * (inner_width + 2) + "╯", style=TABLE_BORDER_COLOR)
        lines: list[str | Text] = [
            top,
            self.wrap_table_content_line(self.summary_cell(summary, inner_width), inner_width),
            divider,
        ]
        lines.extend(self.wrap_table_content_line(line, inner_width) for line in content_lines)
        lines.append(bottom)
        return lines

    def plain_table_block_lines(self, summary: str, content_lines: list[str]) -> list[str | Text]:
        inner_width = self.table_block_inner_width()
        top = Text("╭" + "─" * (inner_width + 2) + "╮", style=TABLE_BORDER_COLOR)
        divider = Text("├" + "─" * (inner_width + 2) + "┤", style=TABLE_BORDER_COLOR)
        bottom = Text("╰" + "─" * (inner_width + 2) + "╯", style=TABLE_BORDER_COLOR)
        lines: list[str | Text] = [
            top,
            self.wrap_table_content_line(self.summary_cell(summary, inner_width), inner_width),
            divider,
        ]
        lines.extend(self.wrap_table_content_line(line, inner_width) for line in content_lines)
        lines.append(bottom)
        return lines

    def format_action_table(
        self,
        actions: Iterable[ActionRecord],
        *,
        include_table: bool = False,
        include_completed: bool = False,
        include_retired: bool = False,
    ) -> list[str]:
        headers = ["Action"]
        if include_table:
            headers.append("Table")
        headers.extend(["Due", "Time", "Status", "Last Note"])
        if include_completed:
            headers.append("Completed")
        if include_retired:
            headers.append("Retired")

        rows: list[list[str]] = []
        for action in actions:
            table = self.store.table_for_id(action.table_id)
            row = [self.summary_cell(action.title, 40)]
            if include_table:
                row.append(self.summary_cell(table.name if table else "Unknown", 24))
            row.extend(
                [
                    self.summary_cell(action.due_date or "-", 10),
                    self.summary_cell(action.estimate or "-", 24),
                    self.summary_cell(action.status or "-", 18),
                    self.summary_cell(self.last_note_line(action.notes), 80),
                ]
            )
            if include_completed:
                row.append(self.summary_cell(action.completed_at or "-", 10))
            if include_retired:
                row.append(self.summary_cell(action.retired_at or "-", 10))
            rows.append(row)
        return self.format_columns(headers, rows, total_width=self.table_block_inner_width())

    def detail_table_width(self) -> int:
        fallback_width = 92
        try:
            detail_scroll = self.query_one("#detail_scroll", VerticalScroll)
            width = int(getattr(detail_scroll.size, "width", 0))
            if width > 0:
                return max(width - 2, 64)
        except Exception:
            pass
        return fallback_width

    def column_widths(self, headers: list[str], total_width: int) -> list[int]:
        preferred = {
            "Action": 18,
            "Table": 18,
            "Due": 10,
            "Time": 4,
            "Status": 11,
            "Last Note": 36,
            "Completed": 10,
            "Retired": 10,
        }
        minimum = {
            "Action": 12,
            "Table": 10,
            "Due": 10,
            "Time": 4,
            "Status": 11,
            "Last Note": 18,
            "Completed": 10,
            "Retired": 10,
        }
        widths = [preferred.get(header, len(header)) for header in headers]
        available = max(total_width - (3 * len(headers) + 1), len(headers))
        shrink_order = ["Last Note", "Action", "Table"]
        grow_order = ["Last Note", "Action", "Table"]
        while sum(widths) > available:
            changed = False
            for header in shrink_order:
                for index, current in enumerate(widths):
                    if headers[index] != header:
                        continue
                    if current <= minimum.get(header, len(header)):
                        continue
                    widths[index] -= 1
                    changed = True
                    if sum(widths) <= available:
                        return widths
            if not changed:
                break
        while sum(widths) < available:
            changed = False
            for header in grow_order:
                for index in range(len(widths)):
                    if headers[index] != header:
                        continue
                    widths[index] += 1
                    changed = True
                    if sum(widths) >= available:
                        return widths
            if not changed:
                break
        return widths

    def format_columns(self, headers: list[str], rows: list[list[str]], *, total_width: int) -> list[str]:
        widths = self.column_widths(headers, total_width)

        def fit(value: str, width: int) -> str:
            return self.summary_cell(value, width).ljust(width)

        def render_row(values: list[str]) -> str:
            cells = [fit(values[index], widths[index]) for index in range(len(widths))]
            return " │ ".join(cells)

        divider = "─┼─".join("─" * width for width in widths)
        lines = [render_row(headers), divider]
        lines.extend(render_row(row) for row in rows)
        return lines

    @staticmethod
    def last_note_line(notes: str) -> str:
        visible_notes = display_note_text(notes)
        for line in reversed(visible_notes.splitlines()):
            stripped = line.strip()
            if stripped:
                return stripped
        return "-"

    @staticmethod
    def summary_cell(value: str, limit: int) -> str:
        collapsed = " ".join((value or "-").split())
        if not collapsed:
            collapsed = "-"
        if len(collapsed) <= limit:
            return collapsed
        return collapsed[: max(limit - 3, 1)].rstrip() + "..."

    @staticmethod
    def truncate(value: str, limit: int) -> str:
        if len(value) <= limit:
            return value
        return value[: max(limit - 3, 1)].rstrip() + "..."

    def update_status(self, message: str) -> None:
        self.status_text = message
        self.query_one("#status_bar", Static).update(message)

    def step_group(self, offset: int) -> None:
        current_index = self.group_ids.index(self.active_group)
        self.active_group = self.group_ids[(current_index + offset) % len(self.group_ids)]
        self.selected_key = f"group:{self.active_group}"
        self.opened_key = f"group:{self.active_group}"
        self.apply_focus_state()
        self.update_details()

    def step_item(self, offset: int) -> None:
        entries = self.group_entries.get(self.active_group, [])
        if not entries:
            return
        list_view = self.query_one(f"#{self.group_list_ids[self.active_group]}", ListView)
        next_index = min(max((list_view.index or 0) + offset, 0), len(entries) - 1)
        list_view.index = next_index
        self.group_indexes[self.active_group] = next_index
        entry = self.current_left_entry()
        self.selected_key = self.entry_key(entry) if entry else None
        self.apply_focus_state()

    def action_nav_up(self) -> None:
        if self.navigation_mode == "groups":
            self.step_group(-1)
            return
        if self.navigation_mode == "items":
            self.step_item(-1)
            return
        if self.move_detail_today_selection(-1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_projects_selection(-1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_archive_selection(-1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_table_selection(-1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_note_selection(-1):
            self.update_details(reset_scroll=False)
            return
        self.query_one("#detail_scroll", VerticalScroll).scroll_up(animate=False)

    def action_nav_down(self) -> None:
        if self.navigation_mode == "groups":
            self.step_group(1)
            return
        if self.navigation_mode == "items":
            self.step_item(1)
            return
        if self.move_detail_today_selection(1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_projects_selection(1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_archive_selection(1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_table_selection(1):
            self.update_details(reset_scroll=False)
            return
        if self.move_detail_note_selection(1):
            self.update_details(reset_scroll=False)
            return
        self.query_one("#detail_scroll", VerticalScroll).scroll_down(animate=False)

    def action_nav_left(self) -> None:
        if self.navigation_mode == "detail":
            if self.exit_detail_projects_actions():
                self.update_details(reset_scroll=False)
                return
            if self.exit_detail_archive_actions():
                self.update_details(reset_scroll=False)
                return
            if self.detail_back_stack:
                self.opened_key = self.detail_back_stack.pop()
                self.update_details()
                return
            self.active_group = self.detail_origin_group
            self.navigation_mode = self.return_mode
            if self.navigation_mode == "groups":
                self.selected_key = f"group:{self.active_group}"
            else:
                entry = self.current_left_entry()
                self.selected_key = self.entry_key(entry) if entry else f"group:{self.active_group}"
            self.apply_focus_state()
            self.update_details(reset_scroll=False)
            return
        if self.navigation_mode == "items":
            self.navigation_mode = "groups"
            self.selected_key = f"group:{self.active_group}"
            self.opened_key = f"group:{self.active_group}"
            self.apply_focus_state()
            self.update_details()

    def action_nav_right(self) -> None:
        if self.navigation_mode == "detail":
            if self.opened_key == "group:today":
                detail_action = self.current_detail_today_action()
                if detail_action is not None:
                    self.detail_back_stack.append(self.opened_key or "")
                    self.seed_last_note_focus(detail_action.id)
                    self.opened_key = f"action:{detail_action.id}"
                    self.update_details()
                return
            if self.opened_key == "group:projects":
                if self.detail_projects_mode == "tables":
                    if self.enter_detail_projects_actions():
                        self.update_details(reset_scroll=False)
                    return
                detail_action = self.current_detail_projects_action()
                if detail_action is not None:
                    self.detail_back_stack.append(self.opened_key or "")
                    self.seed_last_note_focus(detail_action.id)
                    self.opened_key = f"action:{detail_action.id}"
                    self.update_details()
                return
            if self.opened_key == "group:archive":
                if self.detail_archive_mode == "tables":
                    if self.enter_detail_archive_actions():
                        self.update_details(reset_scroll=False)
                    return
                detail_action = self.current_detail_archive_action()
                if detail_action is not None:
                    self.detail_back_stack.append(self.opened_key or "")
                    self.seed_last_note_focus(detail_action.id)
                    self.opened_key = f"action:{detail_action.id}"
                    self.update_details()
                return
            if self.open_selected_note_bullet():
                return
            detail_action = self.current_detail_table_action()
            if detail_action is not None:
                self.detail_back_stack.append(self.opened_key or "")
                self.seed_last_note_focus(detail_action.id)
                self.opened_key = f"action:{detail_action.id}"
                self.update_details()
            return
        entry = self.current_left_entry()
        if entry is None:
            return
        self.detail_back_stack.clear()
        self.detail_origin_group = self.active_group
        if entry.kind == "group" and entry.group == "today":
            self.selected_date = date.today()
        if entry.action_id:
            self.seed_last_note_focus(entry.action_id)
        self.opened_key = self.entry_key(entry)
        if self.opened_key == "group:projects":
            self.detail_projects_mode = "tables"
        if self.opened_key == "group:archive":
            self.detail_archive_mode = "tables"
        self.return_mode = self.navigation_mode
        self.navigation_mode = "detail"
        self.apply_focus_state()
        self.update_details()

    def action_nav_enter(self) -> None:
        if self.navigation_mode == "detail":
            self.action_nav_right()
            return
        if self.navigation_mode != "groups":
            return
        self.navigation_mode = "items"
        list_view = self.query_one(f"#{self.group_list_ids[self.active_group]}", ListView)
        list_view.index = self.group_indexes[self.active_group]
        entry = self.current_left_entry()
        self.selected_key = self.entry_key(entry) if entry else None
        self.apply_focus_state()

    def action_nav_tab(self) -> None:
        if self.navigation_mode == "groups":
            self.action_nav_enter()
            return
        if self.navigation_mode == "detail":
            self.action_nav_right()

    def action_prev_day(self) -> None:
        self.selected_key = "group:today"
        self.selected_date -= timedelta(days=1)
        self.refresh_all()
        self.update_status(f"Today pane moved to {self.selected_date.isoformat()}")

    def action_next_day(self) -> None:
        self.selected_key = "group:today"
        self.selected_date += timedelta(days=1)
        self.refresh_all()
        self.update_status(f"Today pane moved to {self.selected_date.isoformat()}")

    def action_toggle_archive_mode(self) -> None:
        self.selected_key = "group:archive"
        self.archive_mode = "graveyard" if self.archive_mode == "archive" else "archive"
        self.refresh_all()
        self.update_status(f"Completed pane switched to {self.archive_title_word()}")

    def action_new_table(self) -> None:
        if self.active_group != "projects":
            self.update_status("New table works from the Projects section")
            return
        self.push_screen(PromptScreen("New table name"), self.finish_new_table)

    def action_context_note(self) -> None:
        if self.navigation_mode != "detail" and self.active_group == "projects":
            self.action_new_table()
            return
        current_bullet = self.current_detail_note_bullet()
        if current_bullet is not None:
            action, bullet = current_bullet
            table = self.store.table_for_id(action.table_id)
            table_name = table.name if table else "Unknown"
            self.push_screen(
                ActionNoteScreen(f"Notes: {action.title} [{table_name}]"),
                lambda value: self.finish_insert_note_below(action.id, bullet.index, value),
            )
            return
        action = self.focused_action_for_notes()
        if action is not None:
            table = self.store.table_for_id(action.table_id)
            table_name = table.name if table else "Unknown"
            self.push_screen(ActionNoteScreen(f"Notes: {action.title} [{table_name}]"), lambda value: self.finish_note_edit(action.id, value))
            return
        if self.navigation_mode == "detail":
            table = self.current_new_action_table()
            if table is not None:
                self.action_new_action()
                return
        self.update_status("Notes open for a focused action")

    def finish_note_edit(self, action_id: str, value: str | None) -> None:
        if value is None:
            self.update_status("Notes cancelled")
            return
        if not value.strip():
            self.update_status("No note saved")
            return
        action = self.store.append_action_note(action_id, value)
        self.refresh_all()
        self.update_status(f"Added note to {action.title}")

    def finish_insert_note_below(self, action_id: str, bullet_index: int, value: str | None) -> None:
        if value is None:
            self.update_status("Notes cancelled")
            return
        if not value.strip():
            self.update_status("No note saved")
            return
        action = self.store.insert_action_note_after_bullet(action_id, bullet_index, value, prefix="- ", color="")
        self.detail_note_indexes[action.id] = bullet_index + 1
        self.refresh_all()
        self.update_status(f"Inserted note under {action.title}")

    def action_note_bullet_color(self) -> None:
        current = self.current_detail_note_bullet()
        if current is None:
            self.update_status("Color works on a focused note bullet")
            return
        action, bullet = current
        if bullet.prefix.strip().startswith(">>"):
            self.update_status("Subaction color is controlled by its status")
            return
        self.push_screen(
            NoteColorPaletteScreen(f"Bullet Color: {action.title}"),
            lambda value: self.finish_note_bullet_color(action.id, bullet.index, value),
        )

    def finish_note_bullet_color(self, action_id: str, bullet_index: int, value: str | None) -> None:
        if value is None:
            self.update_status("Cancelled bullet color change")
            return
        action = self.store.update_action_note_bullet_color(action_id, bullet_index, value)
        self.detail_note_indexes[action.id] = bullet_index
        self.refresh_all()
        if value:
            self.update_status(f"Changed bullet color in {action.title} to {value}")
            return
        self.update_status(f"Cleared bullet color in {action.title}")

    def open_selected_note_bullet(self) -> bool:
        current = self.current_detail_note_bullet()
        if current is None:
            return False
        action, bullet = current
        table = self.store.table_for_id(action.table_id)
        table_name = table.name if table else "Unknown"
        if bullet.prefix.strip().startswith(">>"):
            subaction_title, subaction_notes = self.parse_subaction_note_text(bullet.text)
            self.push_screen(
                ActionSubactionScreen(
                    f"Edit subaction: {action.title} [{table_name}]",
                    initial_title=subaction_title,
                    initial_notes=subaction_notes,
                    initial_status=subaction_status_from_color(bullet.color),
                    focus_notes=True,
                ),
                lambda value: self.finish_subaction_bullet_edit(action.id, bullet.index, value),
            )
            return True
        self.push_screen(
            ActionNoteScreen(f"Edit note: {action.title} [{table_name}]", bullet.text),
            lambda value: self.finish_note_bullet_edit(action.id, bullet.index, value),
        )
        return True

    def finish_note_bullet_edit(self, action_id: str, bullet_index: int, value: str | None) -> None:
        if value is None:
            self.update_status("Note edit cancelled")
            return
        if not value.strip():
            self.update_status("No note saved")
            return
        action = self.store.update_action_note_bullet(action_id, bullet_index, value)
        self.refresh_all()
        self.update_status(f"Updated note on {action.title}")

    def finish_subaction_bullet_edit(self, action_id: str, bullet_index: int, payload: dict[str, str] | None) -> None:
        if payload is None:
            self.update_status("Subaction edit cancelled")
            return
        note_text = self.compose_subaction_note_text(payload.get("title", ""), payload.get("notes", ""))
        if not note_text.strip():
            self.update_status("No subaction saved")
            return
        action = self.store.update_action_note_bullet(
            action_id,
            bullet_index,
            note_text,
            color=subaction_color_for_status(payload.get("status", "Started")),
        )
        self.refresh_all()
        self.update_status(f"Updated subaction on {action.title}")

    def finish_new_table(self, value: str | None) -> None:
        if not value:
            self.update_status("Cancelled new table")
            return
        table = self.store.ensure_table(value)
        self.selected_key = f"table:projects:{table.id}"
        self.opened_key = f"table:projects:{table.id}"
        self.refresh_all()
        self.update_status(f"Created table {table.name}")
        self.push_screen(
            ActionEditorScreen(
                {
                    "table_name": table.name,
                    "title": "",
                    "due_date": self.default_new_due_date(),
                    "estimate": "",
                    "status": "",
                    "notes": "",
                },
                f"First action for {table.name}",
            ),
            self.finish_new_action,
        )

    @staticmethod
    def compose_subaction_note_text(title: str, notes: str) -> str:
        cleaned_title = title.strip()
        note_lines = [line.strip() for line in notes.splitlines() if line.strip()]
        cleaned_notes = "\n".join(f"- {line}" for line in note_lines)
        if cleaned_title and cleaned_notes:
            return f"{cleaned_title}\n{cleaned_notes}"
        return cleaned_title or cleaned_notes

    @staticmethod
    def parse_subaction_note_text(note_text: str) -> tuple[str, str]:
        lines = note_text.splitlines()
        if not lines:
            return "", ""
        title = lines[0].strip()
        note_lines: list[str] = []
        for line in lines[1:]:
            stripped = line.strip()
            if stripped.startswith("- "):
                note_lines.append(stripped[2:])
            elif stripped:
                note_lines.append(stripped)
        return title, "\n".join(note_lines)

    def action_new_subaction(self) -> None:
        action = self.current_action_detail_action()
        if action is None:
            self.update_status("Subactions open from full action detail")
            return
        table = self.store.table_for_id(action.table_id)
        table_name = table.name if table else "Unknown"
        current_bullet = self.current_detail_note_bullet()
        if current_bullet is not None and current_bullet[0].id == action.id:
            _, bullet = current_bullet
            self.push_screen(
                ActionSubactionScreen(f"Subaction: {action.title} [{table_name}]"),
                lambda value: self.finish_insert_subaction_below(action.id, bullet.index, value),
            )
            return
        self.push_screen(
            ActionSubactionScreen(f"Subaction: {action.title} [{table_name}]"),
            lambda value: self.finish_subaction_edit(action.id, value),
        )

    def action_new_action(self) -> None:
        if self.current_action_detail_action() is not None:
            self.action_new_subaction()
            return
        table = self.current_new_action_table()
        initial_table = table.name if table else ""
        self.push_screen(
            ActionEditorScreen(
                {
                    "table_name": initial_table,
                    "title": "",
                    "due_date": self.default_new_due_date(),
                    "estimate": "",
                    "status": "",
                    "notes": "",
                },
                "New action",
            ),
            self.finish_new_action,
        )

    def finish_new_action(self, payload: dict[str, str] | None) -> None:
        if not payload:
            self.update_status("Cancelled new action")
            return
        if not payload["table_name"] or not payload["title"]:
            self.update_status("Table and action title are required")
            return
        action = self.store.add_action(**payload)
        self.selected_key = f"action:{action.id}"
        self.seed_last_note_focus(action.id)
        self.opened_key = f"action:{action.id}"
        self.refresh_all()
        self.update_status(f"Added action {action.title}")

    def finish_subaction_edit(self, action_id: str, payload: dict[str, str] | None) -> None:
        if payload is None:
            self.update_status("Subaction cancelled")
            return
        note_text = self.compose_subaction_note_text(payload.get("title", ""), payload.get("notes", ""))
        if not note_text.strip():
            self.update_status("No subaction saved")
            return
        action = self.store.append_action_note(
            action_id,
            note_text,
            prefix=">> ",
            color=subaction_color_for_status(payload.get("status", "Not Started")),
        )
        self.seed_last_note_focus(action.id)
        self.selected_key = f"action:{action.id}"
        self.opened_key = f"action:{action.id}"
        self.refresh_all()
        self.update_status(f"Added subaction to {action.title}")

    def finish_insert_subaction_below(self, action_id: str, bullet_index: int, payload: dict[str, str] | None) -> None:
        if payload is None:
            self.update_status("Subaction cancelled")
            return
        note_text = self.compose_subaction_note_text(payload.get("title", ""), payload.get("notes", ""))
        if not note_text.strip():
            self.update_status("No subaction saved")
            return
        action = self.store.insert_action_note_after_bullet(
            action_id,
            bullet_index,
            note_text,
            prefix=">> ",
            color=subaction_color_for_status(payload.get("status", "Not Started")),
        )
        self.detail_note_indexes[action.id] = bullet_index + 1
        self.selected_key = f"action:{action.id}"
        self.opened_key = f"action:{action.id}"
        self.refresh_all()
        self.update_status(f"Inserted subaction under {action.title}")

    def action_open_selected(self) -> None:
        entry = self.current_entry()
        if entry is None:
            return
        if entry.kind == "group":
            self.update_status("Select a project or action to edit")
            return
        if entry.kind == "table":
            table = self.store.table_for_id(entry.table_id or "")
            if table is None:
                self.update_status("No project selected")
                return
            self.push_screen(PromptScreen("Rename table", table.name), lambda value: self.finish_rename_table(table.id, value))
            return
        action = self.current_action()
        if action is None:
            self.update_status("No action selected")
            return
        table = self.store.table_for_id(action.table_id)
        self.push_screen(
            ActionEditorScreen(
                {
                    "table_name": table.name if table else "",
                    "title": action.title,
                    "due_date": action.due_date,
                    "estimate": action.estimate,
                    "status": action.status,
                    "notes": action.notes,
                },
                "Edit action",
            ),
            lambda payload: self.finish_edit_action(action.id, payload),
        )

    def finish_rename_table(self, table_id: str, value: str | None) -> None:
        if not value:
            self.update_status("Cancelled table rename")
            return
        table = self.store.rename_table(table_id, value)
        self.selected_key = f"table:projects:{table.id}"
        self.opened_key = f"table:projects:{table.id}"
        self.refresh_all()
        self.update_status(f"Renamed table to {table.name}")

    def finish_edit_action(self, action_id: str, payload: dict[str, str] | None) -> None:
        if not payload:
            self.update_status("Cancelled action edit")
            return
        if not payload["table_name"] or not payload["title"]:
            self.update_status("Table and action title are required")
            return
        action = self.store.update_action(action_id, **payload)
        self.selected_key = f"action:{action.id}"
        self.seed_last_note_focus(action.id)
        self.opened_key = f"action:{action.id}"
        self.refresh_all()
        bucket_name, _ = self.store.action_by_id(action.id) or ("active", action)
        if bucket_name == "archive":
            self.update_status(f"Updated action {action.title} and moved it to Completed")
            return
        self.update_status(f"Updated action {action.title}")

    def action_edit_due_date(self) -> None:
        action = self.focused_action_for_due_date()
        if action is None:
            self.update_status("Due date works on a focused action")
            return
        self.push_screen(
            PromptScreen("Due date (YYYY-MM-DD or M/D)", action.due_date),
            lambda value: self.finish_edit_due_date(action.id, value),
        )

    def finish_edit_due_date(self, action_id: str, value: str | None) -> None:
        if value is None:
            self.update_status("Cancelled due date edit")
            return
        due_date = normalize_due_date_input(value)
        action = self.store.update_action_due_date(action_id, due_date)
        self.refresh_all()
        self.update_status(f"Updated due date for {action.title} to {action.due_date or 'cleared'}")

    def action_edit_status(self) -> None:
        action = self.focused_action_for_status()
        if action is None:
            self.update_status("Status works on a focused action")
            return
        origin = self.status_edit_origin()
        self.push_screen(
            StatusPickerScreen(f"Status: {action.title}", action.status),
            lambda value: self.finish_edit_status(action.id, value, origin),
        )

    def restore_today_focus(self, action_id: str) -> None:
        self.active_group = "today"
        today_entries = self.group_entries.get("today", [])
        target_index = next((index for index, entry in enumerate(today_entries) if entry.action_id == action_id), None)
        list_view = self.query_one("#today_list", ListView)
        if target_index is None:
            if today_entries:
                target_index = min(self.group_indexes.get("today", 0), len(today_entries) - 1)
                self.group_indexes["today"] = target_index
                list_view.index = target_index
                self.selected_key = self.entry_key(today_entries[target_index])
            else:
                self.group_indexes["today"] = 0
                list_view.index = 0
                self.selected_key = "group:today"
        else:
            self.group_indexes["today"] = target_index
            list_view.index = target_index
            self.selected_key = f"action:{action_id}"
        self.opened_key = "group:today"
        self.apply_focus_state()
        self.update_details(reset_scroll=False)

    def finish_edit_status(self, action_id: str, value: str | None, origin: str = "default") -> None:
        if value is None:
            self.update_status("Cancelled status edit")
            return
        action = self.store.set_action_status(action_id, value)
        if origin == "today":
            self.refresh_all()
            self.restore_today_focus(action.id)
        else:
            self.selected_key = f"action:{action.id}"
            self.seed_last_note_focus(action.id)
            self.opened_key = f"action:{action.id}"
            self.refresh_all()
        bucket_name, _ = self.store.action_by_id(action.id) or ("active", action)
        if bucket_name == "archive":
            self.update_status(f"Marked {action.title} completed and moved it to Completed")
            return
        self.update_status(f"Updated status for {action.title} to {action.status}")

    def action_complete_action(self) -> None:
        action = self.current_action()
        if action is None:
            self.update_status("Complete works on active actions only")
            return
        found = self.store.action_by_id(action.id)
        if found is None or found[0] != "active":
            self.update_status("Complete works on active actions only")
            return
        self.store.complete_action(action.id)
        self.selected_key = "group:archive"
        self.opened_key = "group:archive"
        self.navigation_mode = "detail"
        self.return_mode = "groups"
        self.refresh_all()
        self.update_status(f"Completed {action.title}")

    def action_retire_action(self) -> None:
        current_bullet = self.current_detail_note_bullet()
        if current_bullet is not None:
            action, bullet = current_bullet
            self.push_screen(
                ConfirmScreen("Delete this bullet?", default_accept=False),
                lambda confirmed: self.finish_delete_note_bullet(action.id, action.title, bullet.index, confirmed),
            )
            return
        project_table = self.current_project_table_context()
        if project_table is not None:
            self.push_screen(
                ConfirmScreen(
                    f"Delete project '{project_table.name}'? Any open actions in it will be completed and moved to Completed.",
                    default_accept=False,
                ),
                lambda confirmed: self.finish_retire_table(project_table.id, project_table.name, confirmed),
            )
            return
        action = self.current_action()
        if action is None:
            self.update_status("No action selected")
            return
        found = self.store.action_by_id(action.id)
        if found is None:
            self.update_status("No action selected")
            return
        if found[0] == "graveyard":
            self.update_status("Action is already in the graveyard")
            return
        if found[0] != "archive" or self.archive_mode != "archive" or self.find_group_for_key(self.opened_key) != "archive":
            self.update_status("Retire works from Completed only")
            return
        self.push_screen(
            ConfirmScreen(f"Delete '{action.title}' from Completed and move it to Graveyard?"),
            lambda confirmed: self.finish_retire_action(action.id, action.title, confirmed),
        )

    def finish_retire_table(self, table_id: str, title: str, confirmed: bool) -> None:
        if not confirmed:
            self.update_status("Cancelled project delete")
            return
        table, moved_count = self.store.retire_table(table_id)
        if moved_count:
            self.active_group = "archive"
            self.selected_key = f"table:archive:{table.id}"
            self.opened_key = f"table:archive:{table.id}"
            self.navigation_mode = "detail"
            self.return_mode = "groups"
            self.refresh_all()
            self.update_status(f"Deleted project {title} and moved {moved_count} action(s) to Completed")
            return
        self.active_group = "projects"
        self.selected_key = "group:projects"
        self.opened_key = "group:projects"
        self.navigation_mode = "detail"
        self.return_mode = "groups"
        self.refresh_all()
        self.update_status(f"Deleted project {title}")

    def finish_retire_action(self, action_id: str, title: str, confirmed: bool) -> None:
        if not confirmed:
            self.update_status("Cancelled retire")
            return
        self.store.retire_action(action_id)
        self.selected_key = "group:archive"
        self.opened_key = "group:archive"
        self.navigation_mode = "detail"
        self.return_mode = "groups"
        self.refresh_all()
        self.update_status(f"Moved {title} to graveyard")

    def finish_delete_note_bullet(self, action_id: str, title: str, bullet_index: int, confirmed: bool) -> None:
        if not confirmed:
            self.update_status("Cancelled bullet delete")
            return
        action = self.store.delete_action_note_bullet(action_id, bullet_index)
        remaining_bullets = self.detail_note_bullets_for(action)
        if remaining_bullets:
            self.detail_note_indexes[action.id] = min(bullet_index, len(remaining_bullets) - 1)
        else:
            self.detail_note_indexes.pop(action.id, None)
        self.selected_key = f"action:{action.id}"
        self.opened_key = f"action:{action.id}"
        self.refresh_all()
        self.update_status(f"Deleted note bullet from {title}")

    def action_restore_action(self) -> None:
        action = self.current_action()
        if action is None:
            self.update_status("No action selected")
            return
        found = self.store.action_by_id(action.id)
        if found is None or found[0] not in {"archive", "graveyard"}:
            self.update_status("Restore works from Completed or Graveyard")
            return
        self.store.restore_action(action.id)
        self.selected_key = f"action:{action.id}"
        self.seed_last_note_focus(action.id)
        self.opened_key = f"action:{action.id}"
        self.refresh_all()
        self.update_status(f"Restored {action.title} to active")

    def action_delete_graveyard(self) -> None:
        action = self.current_action()
        found = self.store.action_by_id(action.id) if action else None
        if self.archive_mode != "graveyard" or found is None or found[0] != "graveyard":
            self.update_status("Delete is only available in Graveyard view")
            return
        self.push_screen(
            ConfirmScreen(f"Permanently delete '{action.title}' from the graveyard?"),
            lambda confirmed: self.finish_delete_graveyard(action.id, action.title, confirmed),
        )

    def finish_delete_graveyard(self, action_id: str, title: str, confirmed: bool) -> None:
        if not confirmed:
            self.update_status("Cancelled delete")
            return
        self.store.delete_graveyard_action(action_id)
        self.selected_key = "group:archive"
        self.opened_key = "group:archive"
        self.navigation_mode = "detail"
        self.return_mode = "groups"
        self.refresh_all()
        self.update_status(f"Deleted {title} from graveyard")

    def action_refresh_data(self) -> None:
        if self.current_detail_note_bullet() is not None:
            self.action_note_bullet_color()
            return
        self.update_status("Color works on a focused note bullet")

    def action_show_help(self) -> None:
        self.push_screen(HelpScreen())


def main(argv: list[str] | None = None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    workbook_path = DEFAULT_WORKBOOK if not args else Path(args[0])
    ActionsTuiApp(workbook_path=workbook_path).run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
