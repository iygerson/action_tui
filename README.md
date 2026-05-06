# Actions TUI Prototype

This prototype combines the `notes_tool` pane-driven Textual feel with the Excel actions workflow from `Actions_Tool.xlsm`.

## What it does

- shows a two-pane TUI:
- a left representative pane with three boxed groups: `Today`, `Projects`, and `Completed`
  - a right detail pane that shows either the full selected group view or the selected action/project detail
- seeds itself from `C:\Users\ilany\Desktop\Actions_Tool.xlsm` on first run
- stores local prototype state in `data/actions_state.json`

## Controls

- `Up` / `Down`: switch between `Today`, `Projects`, and `Completed` while focused on the left
- `Enter` or `Tab`: move from group focus into the current group's item list
- `Up` / `Down` inside a group: move between items in that group
- `Right`: open the focused group or item on the right and move focus to the detail pane
- when `Today` is open on the right, `Up` / `Down` moves focus across its action rows
- when a project is open on the right, `Up` / `Down` moves focus across its action rows
- `Right` on a focused project-table row opens that action's detailed view
- in the all-projects right-pane view, `Up` / `Down` moves between project tables
- `Right`, `Enter`, or `Tab` on a project table enters that table's action rows, and `Left` returns to table focus
- when an action with note bullets is open on the right, `Up` / `Down` moves across those bullets
- `Right`, `Enter`, or `Tab` on a focused note bullet opens it for editing
- `r` / `R` on a focused note bullet opens a color palette; picking a letter changes that bullet to the selected color
- `Left`: return from the detail pane to the left, or back out of a group's item list
- `n` / `N`: add a new note bullet to the focused action; when a note bullet is focused on the right, the new note is inserted directly below it; from a project table context with no focused action, `n` opens a new action for that table; from `Projects` group focus on the left, `n` creates a new project
- inside the notes editor, `Enter` adds a new line and `Tab` saves/closes the note
- inside the action editor, arrow keys or `Enter` move between fields, `Left` / `Right` on the status field chooses `Not Started`, `Started`, or `Completed`, `Completed` moves the action to the `Completed` group, and `Tab` saves
- `e`: edit the selected action, or rename the selected project
- `s` / `S`: open the status picker for the focused action
- `a`: add an action
- `c`: mark an active action complete and send it to `Completed`
- `x`: on a focused note bullet, opens a confirmation window and deletes that bullet; otherwise, only from `Completed`, opens a confirmation window and moves the action to graveyard
- `v`: toggle the `Completed` group between `Completed` and graveyard
- `u`: restore a completed or graveyard action to active
- `d` / `D`: open the due date prompt for the focused action
- `Delete`: permanently delete from graveyard
- `[` / `]`: move the Today group backward or forward by one day
- `r`: change the color of the focused note bullet
- `?`: help
- `q`: quit

## Running

Use the included launcher:

```powershell
.\action-tui.cmd
```

You can also point it at a different workbook:

```powershell
.\action-tui.cmd C:\path\to\Actions_Tool.xlsm
```
