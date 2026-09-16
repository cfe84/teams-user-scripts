# Teams user scripts

This repository is a catalogue and distribution source for UserScripts that
work in the Microsoft Teams web app. The root [`index.json`](./index.json)
lists each available script, its description, and its download path.

To load these scripts in the Teams desktop client, install
[Teamsmonkey](https://github.com/cfe84/teamsmonkey). Teamsmonkey loads scripts
from `~/.config/teamsmonkey/scripts` on macOS and Linux, and from the
equivalent per-user AppData directory on Windows. Set
`TEAMSMONKEY_SCRIPT_PATH` to use another directory.

The catalogue can be added directly to Teamsmonkey with this repository URL:

<https://github.com/cfe84/teams-user-scripts>

Individual scripts are also available as raw files, for example:

<https://github.com/cfe84/teams-user-scripts/raw/refs/heads/main/userscripts/teams-vimium.user.js>

## Available scripts

- **Teams Vimium navigation**: Vimium-style scrolling, pane selection, find,
  hints, and navigation for Teams.
- **Meeting meter**: Tracks participant time and estimated meeting cost.
- **Teams Focus mode**: Reduces visual emphasis for unread messages and
  notifications while keeping their counts visible.

The `teams-user-extensions.user.js` manager is bundled with Teamsmonkey and is
not downloaded from this repository.
