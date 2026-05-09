# Actions TUI

`Actions TUI` is a two-pane terminal tool for managing action tables, daily work, notes, and completed items.

## Behavior

- a fresh clone starts empty
- local state is stored outside the repo in `%LOCALAPPDATA%\ActionsTui\actions_state.json`
- pulling code updates does not overwrite your notes, because runtime data is no longer tracked with the repo
- you create your own tables, actions, notes, and completed items inside the app

## Setup

From the repo root, run:

```powershell
.\setup.ps1
```

What setup does:

- creates a local `.venv`
- installs `textual`
- prepares `%LOCALAPPDATA%\ActionsTui\actions_state.json`
- creates an `actions-tui` command shim in `%USERPROFILE%\bin`
- adds `%USERPROFILE%\bin` to your user `PATH` if needed
- creates a desktop shortcut named `Actions TUI`

If you want setup to wipe any existing local action data and reset the tool to empty:

```powershell
.\setup.ps1 -ResetData
```

After setup, open a new terminal window and run:

```powershell
actions-tui
```

You can still launch it from the repo with:

```powershell
.\actions-tui.cmd
```

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
- `r` / `R` on a focused note bullet opens a color palette
- `Left`: return from the detail pane to the left, or back out of a group's item list
- `n` / `N`: add a new note bullet to the focused action; from a project table context with no focused action, `n` opens a new action for that table; from `Projects` group focus on the left, `n` creates a new project
- inside the notes editor, `Enter` adds a new line and `Tab` saves/closes the note
- inside the action editor, arrow keys or `Enter` move between fields, `Left` / `Right` on the status field chooses `Not Started`, `Started`, or `Completed`, and `Tab` saves
- `e`: edit the selected action, or rename the selected project
- `s` / `S`: open the status picker for the focused action
- `a`: in full action detail add a subaction; otherwise add an action
- `c`: mark an active action complete and send it to `Completed`
- `x`: on a focused note bullet, opens a confirmation window and deletes that bullet; from `Completed`, permanently deletes the focused completed action after confirmation
- `u`: restore a completed action to active
- `d` / `D`: open the due date prompt for the focused action
- `[` / `]`: move the Today group backward or forward by one day
- `?`: help
- `q`: quit
