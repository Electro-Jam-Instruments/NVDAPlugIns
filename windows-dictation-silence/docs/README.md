# Windows Dictation Silence - Developer Documentation

This folder contains technical documentation for the Windows Dictation Silence NVDA addon.

## Documentation Structure

| File | Purpose |
|------|---------|
| `architecture-decisions.md` | Key technical decisions with rationale |
| `implementation-notes.md` | Current implementation details and code structure |
| `user-requirements.md` | User requirements and problem statement |
| `deployment-plan.md` | Deployment and release planning |

## Key Documents

### Architecture
- **[architecture-decisions.md](architecture-decisions.md)** - 5 decisions documented (GlobalPlugin, detection method, keypress hooking, speech mode, window detection)

### Implementation
- **[implementation-notes.md](implementation-notes.md)** - Current v0.0.3 implementation, code structure, state variables

### Requirements
- **[user-requirements.md](user-requirements.md)** - Problem statement and user needs

## When to Use Each Section

| If you need to... | Look at... |
|-------------------|------------|
| Understand why something was built a certain way | `architecture-decisions.md` |
| Understand current code structure | `implementation-notes.md` |
| Understand what problem we're solving | `user-requirements.md` |
