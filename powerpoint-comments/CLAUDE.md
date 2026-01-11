# PowerPoint Comments Plugin - Development Instructions

## Critical Patterns

These patterns are MANDATORY for this plugin. Deviating will cause silent failures.

### AppModule Inheritance
```python
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(AppModule):  # Must use this exact pattern
```
See: `docs/architecture-decisions.md` Decision #1, `docs/history/pitfalls-to-avoid.md` Pitfall #1

### COM Access
```python
import comHelper
ppt = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
```
See: `docs/architecture-decisions.md` Decision #2, `docs/history/pitfalls-to-avoid.md` Pitfall #3

### Event Handlers
- `super().__init__()` - YES (parent has this)
- `super().event_appModule_gainFocus()` - NO (crashes - optional hook)
- Heavy work → delegate to worker thread

See: `docs/history/pitfalls-to-avoid.md` Pitfalls #2, #8

## Documentation

| Need | Location |
|------|----------|
| Why decisions were made | `docs/architecture-decisions.md` |
| What NOT to do | `docs/history/pitfalls-to-avoid.md` |
| UIA focus patterns | `docs/experts/expert-ppt-uia.md` |
| COM automation | `docs/experts/expert-ppt-com.md` |
| Slideshow mode | `docs/reference/slideshow-override.md` |
| Full implementation plan | `docs/MVP_IMPLEMENTATION_PLAN.md` |

## Deployment

Tag format: `powerpoint-comments-vX.X.X-beta` (plugin prefix REQUIRED)

Use `/deploy` command or see `../.claude/commands/deploy.md`
