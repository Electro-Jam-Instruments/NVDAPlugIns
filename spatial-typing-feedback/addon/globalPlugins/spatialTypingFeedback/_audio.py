# Panned audio output.
#
# Everything here exists because nvwave.WavePlayer.setVolume() is per-channel but the
# audio we are given is mono. See docs/architecture-decisions.md Decisions D2, D8, D10.

import array
import math
import wave

import config
import nvwave
from logHandler import log

from ._streams import PAN_MAX, PAN_MIN

#: nvwave.playWaveFile skips this many bytes of RIFF header before the PCM starts.
#: Matches NVDA's own oneCore driver constant.
WAVE_HEADER_LENGTH = 46


def _getWavePlayerClass():
    """NVDA 2026.1 renamed WasapiWavePlayer to WavePlayer. Support both."""
    cls = getattr(nvwave, "WavePlayer", None)
    if cls is None:
        cls = getattr(nvwave, "WasapiWavePlayer", None)
    return cls


def panToChannels(pan, volumePercent=100):
    """Convert a pan position and level into left/right channel gains.

    Constant-power law, not linear balance. Our source is mono duplicated into both
    channels, so the two are fully correlated: a centred signal is up to 6 dB louder
    than the same signal panned hard to one side. Linear balance would make the centre
    stream boom next to the side streams, and it would read as a volume bug rather than
    a pan bug.

    @param pan: -50 (hard left) .. 0 (centre) .. +50 (hard right)
    @param volumePercent: 0..100
    @return: (left, right), each 0.0..1.0
    """
    pan = max(PAN_MIN, min(PAN_MAX, pan))
    theta = (pan - PAN_MIN) / float(PAN_MAX - PAN_MIN) * (math.pi / 2)
    level = max(0.0, min(100.0, volumePercent)) / 100.0
    return math.cos(theta) * level, math.sin(theta) * level


def monoToStereo(monoBytes):
    """Duplicate 16-bit mono PCM into both channels of interleaved stereo.

    Panning a mono player is a contradiction - setVolume raises E_INVALIDARG on
    channel 1 - so the upmix is mandatory before any position can be applied.

    Done with extended slice assignment on array.array, which runs in C. NVDA bundles
    no numpy and Python 3.13 removed audioop, so a per-sample Python loop is the only
    alternative and it is far too slow.
    """
    mono = array.array("h")
    mono.frombytes(monoBytes[: len(monoBytes) - (len(monoBytes) % 2)])
    stereo = array.array("h", bytes(4 * len(mono)))
    stereo[0::2] = mono
    stereo[1::2] = mono
    return stereo.tobytes()


class PannedPlayer(object):
    """A WavePlayer we own, positioned in the stereo field.

    We create it with channels=2 regardless of the source, so setVolume always has two
    channels to work with.
    """

    def __init__(self, samplesPerSec=22050, bitsPerSample=16, purpose=None):
        self._samplesPerSec = samplesPerSec
        self._bitsPerSample = bitsPerSample
        self._purpose = purpose
        self._player = None
        self._left = 1.0
        self._right = 1.0
        self._panSupported = True

    def _ensurePlayer(self, samplesPerSec):
        if self._player is not None and self._samplesPerSec == samplesPerSec:
            return self._player
        cls = _getWavePlayerClass()
        if cls is None:
            log.error("Spatial Typing Feedback: no WavePlayer class found in nvwave")
            return None
        if self._player is not None:
            try:
                self._player.idle()
            except Exception:  # noqa: BLE001
                log.debugWarning("Error idling previous player", exc_info=True)
        self._samplesPerSec = samplesPerSec
        kwargs = {
            "channels": 2,
            "samplesPerSec": samplesPerSec,
            "bitsPerSample": self._bitsPerSample,
            "outputDevice": config.conf["audio"]["outputDevice"],
        }
        if self._purpose is not None:
            kwargs["purpose"] = self._purpose
        try:
            self._player = cls(**kwargs)
        except TypeError:
            # Older/newer signature - fall back to the arguments every version accepts.
            kwargs.pop("purpose", None)
            self._player = cls(**kwargs)
        return self._player

    def setPosition(self, pan, volumePercent=100):
        """Set where this player sits and how loud it is."""
        self._left, self._right = panToChannels(pan, volumePercent)

    def _applyPosition(self):
        """Push our channel gains into the player.

        Re-applied before every feed rather than once at creation: WavePlayer.open()
        and .stop() both call _setVolumeFromConfig(), which calls setVolume(all=...)
        and wipes panning. That only bites players created with AudioPurpose.SOUNDS,
        but re-applying is cheap and removes the whole class of bug.
        """
        if self._player is None or not self._panSupported:
            return
        try:
            self._player.setVolume(left=self._left, right=self._right)
        except OSError:
            # Mono output device: no channel 1 to set. Positioning is meaningless here,
            # so stop trying rather than raising on every utterance.
            log.warning(
                "Spatial Typing Feedback: output device does not support per-channel "
                "volume; positioning disabled",
            )
            self._panSupported = False
        except Exception:  # noqa: BLE001
            log.debugWarning("Error setting channel volume", exc_info=True)

    @property
    def panSupported(self):
        return self._panSupported

    def feedMono(self, monoBytes, samplesPerSec=None):
        """Upmix mono PCM to stereo, position it, and play it."""
        if not monoBytes:
            return
        player = self._ensurePlayer(samplesPerSec or self._samplesPerSec)
        if player is None:
            return
        self._applyPosition()
        try:
            player.feed(monoToStereo(monoBytes))
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error feeding audio", exc_info=True)

    def feedStereo(self, stereoBytes, samplesPerSec=None):
        """Play already-interleaved 16-bit stereo PCM at this player's position."""
        if not stereoBytes:
            return
        player = self._ensurePlayer(samplesPerSec or self._samplesPerSec)
        if player is None:
            return
        self._applyPosition()
        try:
            player.feed(stereoBytes)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error feeding audio", exc_info=True)

    def feedWaveFile(self, path):
        """Play a wave file at this player's position."""
        try:
            with wave.open(path, "r") as f:
                channels = f.getnchannels()
                samplesPerSec = f.getframerate()
                sampleWidth = f.getsampwidth()
                data = f.readframes(f.getnframes())
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: could not read %s" % path, exc_info=True)
            return
        if sampleWidth != 2:
            log.warning(
                "Spatial Typing Feedback: %s is %d-bit, expected 16-bit; not positioning it"
                % (path, sampleWidth * 8),
            )
            return
        if channels == 1:
            self.feedMono(data, samplesPerSec)
            return
        # Already stereo - position it without touching the samples.
        self.feedStereo(data, samplesPerSec)

    def stop(self):
        if self._player is None:
            return
        try:
            self._player.stop()
        except Exception:  # noqa: BLE001
            log.debugWarning("Error stopping player", exc_info=True)

    def terminate(self):
        if self._player is None:
            return
        try:
            self._player.stop()
            self._player.close()
        except Exception:  # noqa: BLE001
            log.debugWarning("Error closing player", exc_info=True)
        self._player = None
