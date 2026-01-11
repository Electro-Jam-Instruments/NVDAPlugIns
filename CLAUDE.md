# NVDA Plugins Repository

Multi-plugin repository for NVDA accessibility addons.

## Plugins

| Plugin | Description | Docs |
|--------|-------------|------|
| `powerpoint-comments/` | PowerPoint comments and notes navigation | `powerpoint-comments/CLAUDE.md` |
| `windows-dictation-silence/` | Auto-silence NVDA during voice typing | `windows-dictation-silence/docs/` |

## Documentation

- `/docs/` - General NVDA addon development knowledge (shared across plugins)
- `/{plugin}/docs/` - Plugin-specific documentation

## Deployment

Use `/deploy` command for releasing plugin versions. See `.claude/commands/deploy.md`.

## Build System

All plugins use standard NVDA scons build:
```bash
cd {plugin-folder}
scons
```
