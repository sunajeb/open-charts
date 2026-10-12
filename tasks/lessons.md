# Lessons

- **Spawn agents from a trusted folder.** Agents inherit the current working directory; a new folder triggers Claude Code's "trust this folder?" prompt in a hidden pane and the agent waits forever. Before spawning, `cd` to an already-trusted folder (e.g. ~/Downloads) and use absolute paths in the prompt. Within a few minutes of spawning, verify each agent is actually working (check its pane or for file output) instead of assuming.
- **Don't oversubscribe the Mac.** ~7 concurrent agents plus other sessions hit "fork failed: Device not configured". Keep to ~5 and start more only as others finish.

- **Build lean; confirm scope before adding surfaces.** The owner wants the first build to do one thing: add every chart type into PowerPoint. AI, other apps, data inputs and template fill were cut after I started on them. Before starting work beyond the core goal, check it's wanted.

- **Warn before driving the user's apps.** Live PowerPoint tests run in the user's own session and a broken add-in pops modal dialogs at them (and blocks the automation). Validate files offline first, warn the user before any live run, keep live attempts to one or two, quit the app after each, and use fixed file names so their add-in list isn't littered.
