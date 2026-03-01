# Contributing to NVDA Plugins

Thank you for your interest in contributing! These plugins help make daily workflows more accessible for screen reader users.

## Ways to Contribute

### Report Bugs

Found a problem? [Open an issue](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/issues/new?template=bug_report.md) with:

- NVDA version
- Windows version
- Plugin version
- Steps to reproduce
- What happened vs what you expected

### Suggest Features

Have an idea? [Open a feature request](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/issues/new?template=feature_request.md) describing:

- What problem it solves
- How it would work
- Why it helps accessibility

### Submit Code

1. Fork the repository
2. Create a branch for your change
3. Make your changes
4. Test with NVDA
5. Submit a pull request

## Development Setup

### Requirements

- NVDA 2024.1 or later (for testing)
- Python 3.11 (NVDA's Python version)
- scons (for building)

### Building a Plugin

```bash
cd powerpoint-comments  # or windows-dictation-silence
scons
```

This creates a `.nvda-addon` file you can install for testing.

### Testing Your Changes

1. Build the addon
2. Install in NVDA (double-click the .nvda-addon file)
3. Restart NVDA
4. Test the functionality
5. Check NVDA log (NVDA+F1) for errors

## Code Guidelines

### NVDA Patterns

- Follow NVDA's coding conventions
- Use `comHelper.getActiveObject()` for COM access (not direct `GetActiveObject`)
- Don't block event handlers - delegate heavy work to threads
- See plugin-specific `CLAUDE.md` files for critical patterns

### Commit Messages

- Use clear, descriptive messages
- Start with verb: "Add", "Fix", "Update", "Remove"
- Reference issue numbers when applicable

### Pull Request Checklist

- [ ] Code follows existing patterns
- [ ] Tested with NVDA
- [ ] No errors in NVDA log
- [ ] Updated CHANGELOG if user-facing change

## Questions?

- **Forum:** [community.electro-jam.com](https://community.electro-jam.com)
- **Issues:** [GitHub Issues](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/issues)

## License

By contributing, you agree that your contributions will be licensed under the GNU General Public License v2.0 (GPL-2.0), consistent with NVDA's licensing.
