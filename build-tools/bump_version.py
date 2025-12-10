#!/usr/bin/env python
"""
Bump version in NVDA addon manifest.ini

Usage:
    python bump_version.py <plugin-name> <new-version>

Example:
    python bump_version.py powerpoint-comments 0.0.2

This script updates the version field in the addon's manifest.ini file.
Version updates are MANUAL - only run when explicitly requested.
"""

import sys
import re
from pathlib import Path


def get_manifest_path(plugin_name: str) -> Path:
    """Get path to manifest.ini for a plugin."""
    script_dir = Path(__file__).parent
    repo_root = script_dir.parent
    manifest_path = repo_root / plugin_name / "addon" / "manifest.ini"
    return manifest_path


def read_manifest(manifest_path: Path) -> str:
    """Read manifest.ini content."""
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest not found: {manifest_path}")
    return manifest_path.read_text(encoding="utf-8")


def update_version(content: str, new_version: str) -> tuple[str, str]:
    """
    Update version in manifest content.
    Returns (updated_content, old_version).
    """
    # Match version line: version = X.X.X
    pattern = r"^(version\s*=\s*)(\d+\.\d+\.\d+)(.*)$"

    match = re.search(pattern, content, re.MULTILINE)
    if not match:
        raise ValueError("Could not find version line in manifest.ini")

    old_version = match.group(2)

    # Replace version
    updated = re.sub(
        pattern,
        f"\\g<1>{new_version}\\g<3>",
        content,
        flags=re.MULTILINE
    )

    return updated, old_version


def validate_version(version: str) -> bool:
    """Validate version format (X.X.X)."""
    pattern = r"^\d+\.\d+\.\d+$"
    return bool(re.match(pattern, version))


def main():
    if len(sys.argv) != 3:
        print("Usage: python bump_version.py <plugin-name> <new-version>")
        print("Example: python bump_version.py powerpoint-comments 0.0.2")
        sys.exit(1)

    plugin_name = sys.argv[1]
    new_version = sys.argv[2]

    # Validate version format
    if not validate_version(new_version):
        print(f"Error: Invalid version format '{new_version}'")
        print("Version must be in format: X.X.X (e.g., 0.0.1, 1.2.3)")
        sys.exit(1)

    # Get manifest path
    manifest_path = get_manifest_path(plugin_name)

    try:
        # Read current manifest
        content = read_manifest(manifest_path)

        # Update version
        updated_content, old_version = update_version(content, new_version)

        if old_version == new_version:
            print(f"Version is already {new_version}")
            sys.exit(0)

        # Write updated manifest
        manifest_path.write_text(updated_content, encoding="utf-8")

        print(f"Updated {plugin_name} version: {old_version} -> {new_version}")
        print(f"File: {manifest_path}")
        print("")
        print("Next steps:")
        print(f"  1. git add {manifest_path.relative_to(manifest_path.parent.parent.parent)}")
        print(f'  2. git commit -m "Bump {plugin_name} to v{new_version}"')
        print(f"  3. git push origin main")
        print(f"  4. git tag {plugin_name}-v{new_version}-beta  # or without -beta for release")
        print(f"  5. git push origin {plugin_name}-v{new_version}-beta")

    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
