# Stream definitions and starting values.
#
# These are HARDCODED ON PURPOSE. See docs/architecture-decisions.md Decision D11:
# the values are a considered starting point, not a result. They stay constants until
# they have been judged by ear and stop moving. Only then do they become persisted
# defaults with a settings panel around them.
#
# To try different values: edit here, then NVDA+control+F3 to reload plugins.

from enum import Enum


class Stream(Enum):
    """The three places sound can come from."""

    MAIN = "main"
    CHARS = "chars"
    ERRORS = "errors"


#: Pan runs -50 (hard left) .. 0 (centre) .. +50 (hard right).
PAN_MIN = -50
PAN_MAX = 50

#: Starting position and level for each stream, as (pan, volume percent).
#:
#: MAIN carries NVDA's own voice and the completed-word echo. They are separate audio
#: paths but one position.
#: CHARS and ERRORS are symmetric: same distance out, same level, so they inform from
#: the periphery without competing with what you are actually listening to.
LAYOUT = {
    Stream.MAIN: (0, 100),
    Stream.CHARS: (35, 50),
    Stream.ERRORS: (-35, 50),
}

#: Which stream each kind of feedback belongs to.
WORD_ECHO_STREAM = Stream.MAIN
CHAR_ECHO_STREAM = Stream.CHARS
ERROR_ALERT_STREAM = Stream.ERRORS


def panAndVolume(stream):
    """Return (pan, volumePercent) for a stream."""
    return LAYOUT[stream]
