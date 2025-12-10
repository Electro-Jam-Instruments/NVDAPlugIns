#!/usr/bin/env python
"""
Build NVDA addon package (.nvda-addon)

Usage:
    python build_addon.py <plugin-name>

Example:
    python build_addon.py powerpoint-comments

Output:
    powerpoint-comments/powerpoint-comments-0.0.1.nvda-addon
"""

import sys
import re
import zipfile
from pathlib import Path


def get_plugin_path(plugin_name: str) -> Path:
    """Get path to plugin directory."""
    script_dir = Path(__file__).parent
    repo_root = script_dir.parent
    return repo_root / plugin_name


def get_version_from_manifest(addon_dir: Path) -> str:
    """Extract version from manifest.ini."""
    manifest_path = addon_dir / "manifest.ini"

    if not manifest_path.exists():
        raise FileNotFoundError(f"manifest.ini not found in {addon_dir}")

    content = manifest_path.read_text(encoding="utf-8")

    match = re.search(r"^version\s*=\s*(\d+\.\d+\.\d+)", content, re.MULTILINE)
    if not match:
        raise ValueError("Could not find version in manifest.ini")

    return match.group(1)


def build_addon(plugin_name: str) -> Path:
    """Build .nvda-addon package."""
    plugin_path = get_plugin_path(plugin_name)
    addon_dir = plugin_path / "addon"

    if not addon_dir.exists():
        raise FileNotFoundError(f"addon directory not found: {addon_dir}")

    # Get version from manifest
    version = get_version_from_manifest(addon_dir)

    # Output path
    output_name = f"{plugin_name}-{version}.nvda-addon"
    output_path = plugin_path / output_name

    # Create zip
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in addon_dir.rglob('*'):
            if file_path.is_file():
                # Skip __pycache__ and .pyc files
                if '__pycache__' in str(file_path) or file_path.suffix == '.pyc':
                    continue

                # Archive name is relative to addon directory
                arcname = file_path.relative_to(addon_dir)
                zf.write(file_path, arcname)

    return output_path


def main():
    if len(sys.argv) != 2:
        print("Usage: python build_addon.py <plugin-name>")
        print("Example: python build_addon.py powerpoint-comments")
        sys.exit(1)

    plugin_name = sys.argv[1]

    try:
        output_path = build_addon(plugin_name)
        print(f"Built: {output_path}")
        print(f"Size: {output_path.stat().st_size} bytes")

    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
