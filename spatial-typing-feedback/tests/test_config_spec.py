# The config spec and the stream layout.
#
# A malformed spec entry does not raise where it is written - NVDA rejects it later and
# the setting silently reverts to a default, which looks like "my settings did not save".

import pytest


class TestConfigSpec:
    def test_every_default_is_inside_its_own_declared_range(self, cfg):
        """A default outside min/max is accepted here and rejected by NVDA at runtime."""
        import re

        for key, spec in cfg.CONFIG_SPEC.items():
            default = re.search(r"default=(-?\d+)", spec)
            low = re.search(r"min=(-?\d+)", spec)
            high = re.search(r"max=(-?\d+)", spec)
            if not (default and low and high):
                continue
            d, lo, hi = int(default.group(1)), int(low.group(1)), int(high.group(1))
            assert lo <= d <= hi, "%s: default %d outside %d..%d" % (key, d, lo, hi)

    def test_pan_settings_span_the_full_stereo_field(self, cfg, streams):
        for key, spec in cfg.CONFIG_SPEC.items():
            if not key.endswith("Pan"):
                continue
            assert "min=%d" % streams.PAN_MIN in spec, key
            assert "max=%d" % streams.PAN_MAX in spec, key

    def test_unset_is_reachable_by_every_setting_that_uses_it(self, cfg):
        """UNSET means "follow the main voice" and must not be a one-way door.

        A default the user can leave but never return to is a trap - it is why the
        minimums here are UNSET rather than 0.
        """
        for key, spec in cfg.CONFIG_SPEC.items():
            if "default=%d" % cfg.UNSET in spec:
                assert "min=%d" % cfg.UNSET in spec, (
                    "%s defaults to UNSET but cannot be set back to it" % key
                )

    def test_every_stream_key_names_a_real_spec_entry(self, cfg):
        for stream, (panKey, volKey) in cfg.STREAM_KEYS.items():
            assert panKey in cfg.CONFIG_SPEC, "%s: %s" % (stream, panKey)
            assert volKey in cfg.CONFIG_SPEC, "%s: %s" % (stream, volKey)

    def test_voiced_streams_have_a_full_set_of_voice_settings(self, cfg):
        for stream in cfg.VOICE_PREFIX:
            for setting in ("Voice", "Rate", "Pitch", "Volume", "RateBoost", "Punctuation"):
                key = cfg.voiceKey(stream, setting)
                assert key in cfg.CONFIG_SPEC, "%s missing %s" % (stream, key)


class TestStreamLayout:
    def test_every_stream_has_a_starting_position(self, streams):
        for stream in streams.Stream:
            assert stream in streams.LAYOUT, stream

    def test_every_starting_pan_is_in_range(self, streams):
        for stream, (pan, volume) in streams.LAYOUT.items():
            assert streams.PAN_MIN <= pan <= streams.PAN_MAX, stream
            assert 0 <= volume <= 100, stream

    def test_the_main_stream_is_centred_and_full_volume(self, streams):
        assert streams.LAYOUT[streams.Stream.MAIN] == (0, 100)

    def test_characters_and_errors_are_symmetric(self, streams):
        """Same distance out, same level - they inform from opposite sides."""
        charPan, charVol = streams.LAYOUT[streams.Stream.CHARS]
        errPan, errVol = streams.LAYOUT[streams.Stream.ERRORS]
        assert charPan == -errPan
        assert charVol == errVol

    def test_spelling_notes_sit_with_the_error_alert(self, streams):
        """They are both information about a marked word, so they share a position."""
        assert streams.LAYOUT[streams.Stream.ANNOTATIONS] == streams.LAYOUT[streams.Stream.ERRORS]

    def test_side_streams_start_quieter_than_the_main_voice(self, streams):
        assert streams.SIDE_VOLUME_OFFSET < 0
        mainVol = streams.LAYOUT[streams.Stream.MAIN][1]
        for stream in (streams.Stream.CHARS, streams.Stream.ERRORS, streams.Stream.ANNOTATIONS):
            assert streams.LAYOUT[stream][1] < mainVol, stream

    def test_every_voiced_stream_is_a_real_stream(self, streams):
        for stream in streams.VOICE_STREAMS:
            assert stream in streams.Stream
