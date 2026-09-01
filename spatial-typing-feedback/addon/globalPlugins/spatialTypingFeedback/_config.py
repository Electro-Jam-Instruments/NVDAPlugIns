# Persisted settings.
#
# _streams.py holds the starting values; this holds what the user tuned them to.
#
# Saved when the user leaves the settings ring, and on teardown - never on a timer. The
# values are committed at a moment the user chose rather than one we picked.
#
# See docs/architecture.md.

import config
from logHandler import log

from ._streams import LAYOUT, PAN_MAX, PAN_MIN, Stream

CONFIG_SECTION = "spatialTypingFeedback"

#: The character voice's rate and pitch are stored ABSOLUTELY, on the same 0-100 scale
#: NVDA uses, not as a percentage of the main voice. It is a voice in its own right; a
#: relative value meant it kept being dragged around whenever the main voice changed.
#:
#: -1 means "not chosen yet": on first run we seed it from the main voice so it starts
#: somewhere sensible, then it is independent for good.
UNSET = -1

_CHAR_PAN, _CHAR_VOL = LAYOUT[Stream.CHARS]
_ERROR_PAN, _ERROR_VOL = LAYOUT[Stream.ERRORS]

CONFIG_SPEC = {
    "enabled": "boolean(default=True)",
    "charPan": "integer(default=%d, min=%d, max=%d)" % (_CHAR_PAN, PAN_MIN, PAN_MAX),
    "charVolume": "integer(default=%d, min=0, max=100)" % _CHAR_VOL,
    "errorPan": "integer(default=%d, min=%d, max=%d)" % (_ERROR_PAN, PAN_MIN, PAN_MAX),
    # -1 means "follow NVDA's own sound volume" rather than an independent level.
    "errorVolume": "integer(default=%d, min=%d, max=100)" % (UNSET, UNSET),
    "charRate": "integer(default=%d, min=%d, max=100)" % (UNSET, UNSET),
    "charPitch": "integer(default=%d, min=%d, max=100)" % (UNSET, UNSET),
    "charVoice": 'string(default="")',
    "charPunctuation": "boolean(default=True)",
    "charRateBoost": "boolean(default=True)",
}

#: Config keys for each stream we own, as (panKey, volumeKey).
STREAM_KEYS = {
    Stream.CHARS: ("charPan", "charVolume"),
    Stream.ERRORS: ("errorPan", "errorVolume"),
}


def initialize():
    """Register our section with NVDA's config so values validate and persist."""
    try:
        config.conf.spec[CONFIG_SECTION] = CONFIG_SPEC
    except Exception:  # noqa: BLE001
        log.error("Spatial Typing Feedback: could not register config spec", exc_info=True)


def _section():
    return config.conf[CONFIG_SECTION]


def get(key, default=None):
    try:
        return _section()[key]
    except Exception:  # noqa: BLE001
        return default


def set(key, value):  # noqa: A001
    """Change a setting. Written to disk by flush(), not here."""
    global _dirty
    try:
        _section()[key] = value
    except Exception:  # noqa: BLE001
        log.error("Spatial Typing Feedback: could not set %s" % key, exc_info=True)
        return
    _dirty = True


#: True when something has changed and not yet reached disk.
_dirty = False


def isDirty():
    return _dirty


def flush():
    """Write settings to disk now, if there is anything to write.

    Called when the user leaves the settings ring, and on teardown. No timer: settings
    are committed at a moment the user chose, not one we picked.
    """
    global _dirty
    if not _dirty:
        return
    try:
        config.conf.save()
        _dirty = False
        log.debug("Spatial Typing Feedback: settings saved")
    except Exception:  # noqa: BLE001
        log.debugWarning("Could not save config", exc_info=True)


def panAndVolume(stream):
    """Current tuned position and level for a stream."""
    panKey, volKey = STREAM_KEYS[stream]
    fallbackPan, fallbackVol = LAYOUT[stream]
    return get(panKey, fallbackPan), get(volKey, fallbackVol)
