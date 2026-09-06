# Makes the add-on's modules importable without a running NVDA.
#
# The point of this file is that the arithmetic in this add-on - the pan law, the
# mono-to-stereo upmix, the SSML rate mapping - is exactly where a silent bug hides. None
# of it needs a screen reader to run, so none of it should need one to test.
#
# The package's own __init__.py pulls in most of NVDA, so it is deliberately NOT executed.
# A synthetic package pointed at the source directory lets the leaf modules import each
# other by relative import without dragging that in.

import importlib
import sys
import types
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[1] / "addon" / "globalPlugins" / "spatialTypingFeedback"

PACKAGE = "stf"


class _FakeLog(object):
    """logHandler.log, minus the handler."""

    def __init__(self):
        self.records = []

    def _record(self, level, msg, *a, **k):
        self.records.append((level, str(msg)))

    def __getattr__(self, name):
        return lambda msg, *a, **k: self._record(name, msg)


class _FakeConf(dict):
    def __init__(self):
        super().__init__()
        self.spec = {}
        self["audio"] = {"outputDevice": "default", "soundVolume": 100}
        self["speech"] = {"symbolLevel": 100}


def _installStubs():
    """Minimal stand-ins for the three NVDA modules the leaf modules import."""
    logHandler = types.ModuleType("logHandler")
    logHandler.log = _FakeLog()
    sys.modules["logHandler"] = logHandler

    config = types.ModuleType("config")
    config.conf = _FakeConf()
    sys.modules["config"] = config

    nvwave = types.ModuleType("nvwave")

    class WavePlayer(object):
        """Records what it was asked to do; never opens a device."""

        def __init__(self, **kwargs):
            self.kwargs = kwargs
            self.fed = []
            self.volumes = []

        def feed(self, data, *a, **k):
            self.fed.append(data)

        def setVolume(self, **kwargs):
            self.volumes.append(kwargs)

        def idle(self):
            pass

        def stop(self):
            pass

    nvwave.WavePlayer = WavePlayer
    nvwave.decide_playWaveFile = types.SimpleNamespace(
        register=lambda *a, **k: None, unregister=lambda *a, **k: None,
    )
    sys.modules["nvwave"] = nvwave


def _installPackage():
    """Expose the source directory as a package WITHOUT running its __init__.py."""
    pkg = types.ModuleType(PACKAGE)
    pkg.__path__ = [str(SRC)]
    sys.modules[PACKAGE] = pkg


_installStubs()
_installPackage()


def load(name):
    """Import one of the add-on's leaf modules, e.g. load("_audio")."""
    return importlib.import_module("%s.%s" % (PACKAGE, name))


@pytest.fixture(scope="session")
def audio():
    return load("_audio")


@pytest.fixture(scope="session")
def streams():
    return load("_streams")


@pytest.fixture(scope="session")
def winrt():
    return load("_winrt")


@pytest.fixture(scope="session")
def cfg():
    return load("_config")
