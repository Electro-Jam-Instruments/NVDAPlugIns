#!/usr/bin/env python3
"""Static checks that need no NVDA and no Windows.

Every check here exists because the thing it checks for actually happened, and in each
case the first sign of it was the user's screen reader misbehaving:

  parse      a text-range edit silently deleted a function and broke a try block
  refs       the deleted function was still called from another module
  parity     the two voice engines drifted, giving a TypeError on every keystroke
  version    a tag was cut that did not match buildVars, publishing a mislabelled build

Run:  python tools/check_addons.py
"""

import ast
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: Interchangeable implementations that must stay callable the same way.
PARITY_PAIRS = [
    (
        "spatial-typing-feedback/addon/globalPlugins/spatialTypingFeedback/_winrt.py",
        "WinRTVoice",
        "spatial-typing-feedback/addon/globalPlugins/spatialTypingFeedback/_voice.py",
        "SapiVoice",
    ),
]

failures = []


def fail(check, message):
    failures.append((check, message))
    print("  FAIL  %s" % message)


def addonSources():
    """Every Python file that ships inside an add-on."""
    for plugin in sorted(os.listdir(ROOT)):
        addon = os.path.join(ROOT, plugin, "addon")
        if not os.path.isdir(addon):
            continue
        for dirpath, dirnames, filenames in os.walk(addon):
            dirnames[:] = [d for d in dirnames if d != "__pycache__"]
            for name in sorted(filenames):
                if name.endswith(".py"):
                    yield os.path.join(dirpath, name)


def parseAll():
    """Every shipped module must parse. A broken file is a dead add-on."""
    print("\n[parse] every shipped module compiles")
    trees = {}
    for path in addonSources():
        rel = os.path.relpath(path, ROOT).replace("\\", "/")
        src = io.open(path, encoding="utf-8").read()
        try:
            trees[path] = ast.parse(src, filename=rel)
        except SyntaxError as e:
            fail("parse", "%s line %s: %s" % (rel, e.lineno, e.msg))
    print("  %d modules parsed" % len(trees))
    return trees


def topLevelNames(tree):
    names = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    names.add(target.id)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.add(node.target.id)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                if alias.name != "*":
                    names.add(alias.asname or alias.name.split(".")[0])
        elif isinstance(node, ast.If):
            for sub in ast.walk(node):
                if isinstance(sub, (ast.FunctionDef, ast.ClassDef)):
                    names.add(sub.name)
    return names


def checkReferences(trees):
    """Names imported from a sibling module must exist there.

    This is the check that would have caught a function being deleted by a careless
    edit while its callers stayed behind.
    """
    print("\n[refs] cross-module references resolve")
    byPackage = {}
    for path, tree in trees.items():
        byPackage.setdefault(os.path.dirname(path), {})[
            os.path.splitext(os.path.basename(path))[0]
        ] = tree

    checked = 0
    for directory, modules in byPackage.items():
        provided = {name: topLevelNames(tree) for name, tree in modules.items()}
        for name, tree in modules.items():
            rel = os.path.relpath(os.path.join(directory, name + ".py"), ROOT)
            for node in ast.walk(tree):
                if isinstance(node, ast.ImportFrom) and node.level and node.module:
                    target = node.module.split(".")[0]
                    if target not in provided:
                        continue
                    for alias in node.names:
                        if alias.name == "*":
                            continue
                        checked += 1
                        if alias.name not in provided[target]:
                            fail("refs", "%s imports %r from %s, which does not define it"
                                 % (rel.replace("\\", "/"), alias.name, target))
                if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
                    base = node.value.id
                    if base in provided and base != name:
                        checked += 1
                        if node.attr not in provided[base]:
                            fail("refs", "%s references %s.%s, which does not exist"
                                 % (rel.replace("\\", "/"), base, node.attr))
    print("  %d references checked" % checked)


def publicMethods(path, className):
    tree = ast.parse(io.open(path, encoding="utf-8").read())
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == className:
            return {
                m.name: [a.arg for a in m.args.args]
                for m in node.body
                if isinstance(m, ast.FunctionDef) and not m.name.startswith("_")
            }
    return None


def checkParity():
    """Interchangeable classes must be callable identically.

    A substitute that cannot be called the same way is not a substitute; when these
    drifted, the fallback voice raised a TypeError on every single keystroke.
    """
    print("\n[parity] interchangeable engines share one interface")
    for aPath, aName, bPath, bName in PARITY_PAIRS:
        a = publicMethods(os.path.join(ROOT, aPath), aName)
        b = publicMethods(os.path.join(ROOT, bPath), bName)
        if a is None or b is None:
            fail("parity", "could not find %s or %s" % (aName, bName))
            continue
        for method in sorted(set(a) | set(b)):
            if method not in a:
                fail("parity", "%s defines %s, %s does not" % (bName, method, aName))
            elif method not in b:
                fail("parity", "%s defines %s, %s does not" % (aName, method, bName))
            elif a[method] != b[method]:
                fail("parity", "%s.%s%s does not match %s.%s%s"
                     % (aName, method, tuple(a[method]), bName, method, tuple(b[method])))
        print("  %s / %s: %d methods" % (aName, bName, len(a)))


def checkVersions():
    """Each add-on must declare a version, and its changelog should mention it."""
    print("\n[version] buildVars declares a version the changelog knows about")
    for plugin in sorted(os.listdir(ROOT)):
        buildVars = os.path.join(ROOT, plugin, "buildVars.py")
        if not os.path.isfile(buildVars):
            continue
        tree = ast.parse(io.open(buildVars, encoding="utf-8").read())
        version = None
        for node in ast.walk(tree):
            if isinstance(node, ast.Dict):
                for key, value in zip(node.keys, node.values):
                    if isinstance(key, ast.Constant) and key.value == "addon_version":
                        try:
                            version = ast.literal_eval(value)
                        except Exception:  # noqa: BLE001
                            pass
        if not version:
            fail("version", "%s has no addon_version" % plugin)
            continue
        changelog = os.path.join(ROOT, plugin, "CHANGELOG.md")
        if os.path.isfile(changelog):
            text = io.open(changelog, encoding="utf-8").read()
            if version not in text:
                fail("version", "%s is at %s but CHANGELOG.md never mentions it"
                     % (plugin, version))
                continue
        print("  %-28s %s" % (plugin, version))


def main():
    print("Static checks (no NVDA required)")
    trees = parseAll()
    if not any(f[0] == "parse" for f in failures):
        checkReferences(trees)
    checkParity()
    checkVersions()

    print("\n" + "=" * 60)
    if failures:
        print("FAILED - %d problem(s)" % len(failures))
        for check, message in failures:
            print("  [%s] %s" % (check, message))
        return 1
    print("All static checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
