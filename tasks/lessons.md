# Lessons

- **Spawn agents from a trusted folder.** Agents inherit the current working directory; a new folder triggers Claude Code's "trust this folder?" prompt in a hidden pane and the agent waits forever. Before spawning, `cd` to an already-trusted folder (e.g. ~/Downloads) and use absolute paths in the prompt. Within a few minutes of spawning, verify each agent is actually working (check its pane or for file output) instead of assuming.
- **Don't oversubscribe the Mac.** ~7 concurrent agents plus other sessions hit "fork failed: Device not configured". Keep to ~5 and start more only as others finish.
