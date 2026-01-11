# Security Policy

## Supported Versions

| Plugin | Version | Supported |
|--------|---------|-----------|
| PowerPoint Comments | latest beta | Yes |
| Windows Dictation Silence | latest beta | Yes |

Older versions are not actively maintained. Please update to the latest version.

## Reporting a Vulnerability

NVDA addons run with system-level access, so security matters.

### How to Report

**Do not open a public issue for security vulnerabilities.**

Instead, please email: **security@electro-jam.com**

Include:
- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Any suggested fixes (optional)

### What to Expect

- **Acknowledgment:** Within 48 hours
- **Initial assessment:** Within 7 days
- **Resolution timeline:** Depends on severity, typically 30 days for critical issues

### After Reporting

1. We'll confirm receipt and assess the issue
2. We'll work on a fix privately
3. We'll release a patched version
4. We'll credit you in the release notes (unless you prefer anonymity)

## Security Considerations

These plugins:
- Access Microsoft Office via COM automation
- Run within NVDA's process (which has UIAccess privileges)
- Do not transmit data externally
- Do not store sensitive information

## Scope

This policy covers:
- Code in this repository
- Official releases on GitHub Pages

Third-party forks or modifications are not covered.
