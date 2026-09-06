# The positioning arithmetic.
#
# Every one of these encodes something that was got wrong at least once, or that would
# fail silently rather than loudly. A wrong pan law does not raise - it just makes the
# centre sound louder than the sides, which reads as a volume bug and sends you looking
# in the wrong place entirely.

import array
import math

import pytest


class TestPanLaw:
    """Constant power, never linear balance."""

    def test_centre_is_minus_three_db_not_unity(self, audio):
        left, right = audio.panToChannels(0)
        assert left == pytest.approx(math.sqrt(0.5), abs=1e-9)
        assert right == pytest.approx(math.sqrt(0.5), abs=1e-9)

    def test_hard_left_silences_the_right(self, audio):
        left, right = audio.panToChannels(audio.PAN_MIN)
        assert left == pytest.approx(1.0, abs=1e-9)
        assert right == pytest.approx(0.0, abs=1e-9)

    def test_hard_right_silences_the_left(self, audio):
        left, right = audio.panToChannels(audio.PAN_MAX)
        assert left == pytest.approx(0.0, abs=1e-9)
        assert right == pytest.approx(1.0, abs=1e-9)

    def test_power_is_constant_across_the_whole_sweep(self, audio):
        """The property that makes this a pan law rather than a balance control.

        Our audio is mono duplicated into both channels, so it is fully correlated.
        Under linear balance the centre would be up to 6 dB louder than a hard pan.
        """
        for pan in range(audio.PAN_MIN, audio.PAN_MAX + 1):
            left, right = audio.panToChannels(pan)
            assert left ** 2 + right ** 2 == pytest.approx(1.0, abs=1e-9), (
                "pan %d is not constant power" % pan
            )

    def test_is_symmetric_about_centre(self, audio):
        for offset in range(0, audio.PAN_MAX + 1):
            l1, r1 = audio.panToChannels(-offset)
            l2, r2 = audio.panToChannels(offset)
            assert l1 == pytest.approx(r2, abs=1e-9)
            assert r1 == pytest.approx(l2, abs=1e-9)

    @pytest.mark.parametrize("pan", [-999, -51, 51, 999])
    def test_out_of_range_pan_clamps_rather_than_inverting(self, audio, pan):
        left, right = audio.panToChannels(pan)
        assert 0.0 <= left <= 1.0
        assert 0.0 <= right <= 1.0
        expected = audio.panToChannels(max(audio.PAN_MIN, min(audio.PAN_MAX, pan)))
        assert (left, right) == pytest.approx(expected, abs=1e-9)

    def test_volume_scales_both_channels_and_preserves_position(self, audio):
        full = audio.panToChannels(25, 100)
        half = audio.panToChannels(25, 50)
        assert half[0] == pytest.approx(full[0] * 0.5, abs=1e-9)
        assert half[1] == pytest.approx(full[1] * 0.5, abs=1e-9)

    def test_zero_volume_is_silence(self, audio):
        assert audio.panToChannels(0, 0) == pytest.approx((0.0, 0.0), abs=1e-9)

    @pytest.mark.parametrize("volume", [-10, 150])
    def test_out_of_range_volume_clamps(self, audio, volume):
        left, right = audio.panToChannels(0, volume)
        assert 0.0 <= left <= 1.0
        assert 0.0 <= right <= 1.0


class TestMonoToStereo:
    """The synthesizer returns mono; setVolume on a mono player raises E_INVALIDARG."""

    def _samples(self, data):
        a = array.array("h")
        a.frombytes(data)
        return list(a)

    def test_doubles_the_byte_count(self, audio):
        mono = array.array("h", [1, 2, 3, 4]).tobytes()
        assert len(audio.monoToStereo(mono)) == len(mono) * 2

    def test_duplicates_each_sample_into_both_channels(self, audio):
        mono = array.array("h", [10, -20, 30]).tobytes()
        out = self._samples(audio.monoToStereo(mono))
        assert out == [10, 10, -20, -20, 30, 30]

    def test_preserves_full_scale_samples_without_clipping_or_wrapping(self, audio):
        mono = array.array("h", [32767, -32768, 0]).tobytes()
        out = self._samples(audio.monoToStereo(mono))
        assert out == [32767, 32767, -32768, -32768, 0, 0]

    def test_empty_input_is_empty_output(self, audio):
        assert audio.monoToStereo(b"") == b""
