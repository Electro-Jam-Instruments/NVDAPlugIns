# The three streams and where they start.
#
# These are starting points, not results. The user's tuned values live in _config.py and
# override everything here; this file only decides what a fresh install sounds like.
#
# See docs/architecture.md.

from enum import Enum


class Stream(Enum):
    """The three places sound can come from."""

    MAIN = "main"
    CHARS = "chars"
    ERRORS = "errors"
    #: NVDA's own words about a marked word - "spelling error", "out of spelling error".
    #: Information about an error rather than part of the sentence, so it belongs with
    #: the other error feedback, but it is speech and wants its own voice.
    ANNOTATIONS = "annotations"


#: Pan runs -50 (hard left) .. 0 (centre) .. +50 (hard right).
PAN_MIN = -50
PAN_MAX = 50

#: Starting position and level for each stream, as (pan, volume percent).
#:
#: MAIN carries NVDA's own voice and the completed-word echo. They are separate audio
#: paths but one position.
#: CHARS and ERRORS are symmetric: same distance out, same level, so they inform from
#: the periphery without competing with what you are actually listening to.
#: How far below the main voice the side streams start, in volume points.
#: Relative rather than absolute so a fresh install lands somewhere sensible whatever
#: the user's main voice is set to - an absolute default that suits one person's ears is
#: inaudible or deafening on someone else's.
SIDE_VOLUME_OFFSET = -15

LAYOUT = {
    Stream.MAIN: (0, 100),
    Stream.CHARS: (40, 50),
    Stream.ERRORS: (-40, 50),
    Stream.ANNOTATIONS: (-40, 50),
}

#: Streams that have a voice of their own and so can be tuned like one.
VOICE_STREAMS = (Stream.CHARS, Stream.ANNOTATIONS)

#: Speed of the character echo, relative to the main voice.
#:
#: 1.0 means exactly the main voice's speed, rate boost and all - the echo is copied
#: from the rate OneCore is actually running, not rebuilt from a percentage.
#: Below 1.0 is slower, above is faster. OneCore's own limits are 0.5x to 6.0x absolute,
#: and the result is clamped to that.
CHAR_RATE_MULTIPLIER = 1.0

#: OneCore's absolute SpeakingRate limits (Options().SpeakingRate).
ONECORE_MIN_SPEAKING_RATE = 0.5
ONECORE_MAX_SPEAKING_RATE = 6.0

#: Which stream each kind of feedback belongs to.
#:
#: Completed words stay with NVDA's MAIN voice rather than the secondary one. The word
#: echo already works there, at the user's own rate and rate boost; the character echo
#: is the only stream that actually needed its own voice and its own position.
WORD_ECHO_USES_MAIN_VOICE = True
WORD_ECHO_STREAM = Stream.MAIN
CHAR_ECHO_STREAM = Stream.CHARS
ERROR_ALERT_STREAM = Stream.ERRORS


def panAndVolume(stream):
    """Return (pan, volumePercent) for a stream."""
    return LAYOUT[stream]
