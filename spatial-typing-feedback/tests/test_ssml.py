# The SSML envelope and the rate mapping behind it.
#
# Rate is the setting that went wrong the most times, in three different ways, and every
# one of them was silent: the audio just came out at the wrong speed. The saturation test
# below is the important one - SSML prosody pins at 200%, so a mapping that emits 352%
# produces exactly the same audio as one that emits 200% and looks like a dead control.

from xml.etree import ElementTree

import pytest


def voice(winrt, rate=100, pitch=0, volume=100, lang="en-US"):
    """A WinRTVoice with no COM behind it - _buildSsml only reads these four fields."""
    v = object.__new__(winrt.WinRTVoice)
    v.name = "test"
    v._prosodyRate = rate
    v._pitchPercent = pitch
    v._volumePercent = volume
    v._lang = lang
    return v


class TestSsmlEscaping:
    """Typed characters go straight into an XML document. All of them."""

    @pytest.mark.parametrize("char,escaped", [
        ("&", "&amp;"),
        ("<", "&lt;"),
        (">", "&gt;"),
        ('"', "&quot;"),
    ])
    def test_markup_characters_are_escaped(self, winrt, char, escaped):
        out = voice(winrt)._buildSsml(char)
        assert escaped in out
        ElementTree.fromstring(out)  # raises if we produced broken XML

    def test_ampersand_is_escaped_before_the_others(self, winrt):
        """Escaping & last would double-escape the entities the others introduce."""
        out = voice(winrt)._buildSsml("&<>")
        assert "&amp;&lt;&gt;" in out
        assert "&amp;lt;" not in out

    @pytest.mark.parametrize("text", [
        "plain", "a & b", "<tag>", 'say "this"', "&amp;", "<<<>>>", "'", "100% & rising",
    ])
    def test_output_is_always_well_formed(self, winrt, text):
        ElementTree.fromstring(voice(winrt)._buildSsml(text))

    def test_the_text_survives_a_round_trip(self, winrt):
        text = 'a & b < c > d "e"'
        root = ElementTree.fromstring(voice(winrt)._buildSsml(text))
        assert "".join(root.itertext()) == text


class TestProsodyAttributes:
    def test_rate_is_always_present(self, winrt):
        assert 'rate="100%"' in voice(winrt, rate=100)._buildSsml("x")

    def test_pitch_is_omitted_at_zero_and_signed_otherwise(self, winrt):
        assert "pitch=" not in voice(winrt, pitch=0)._buildSsml("x")
        assert 'pitch="+25%"' in voice(winrt, pitch=25)._buildSsml("x")
        assert 'pitch="-25%"' in voice(winrt, pitch=-25)._buildSsml("x")

    def test_volume_is_omitted_at_full_scale(self, winrt):
        assert "volume=" not in voice(winrt, volume=100)._buildSsml("x")
        assert 'volume="60%"' in voice(winrt, volume=60)._buildSsml("x")

    def test_language_comes_from_the_selected_voice(self, winrt):
        assert 'xml:lang="de-DE"' in voice(winrt, lang="de-DE")._buildSsml("x")


class TestRateMapping:
    """setRatePercent maps NVDA's 0-100 onto the range SSML prosody actually honours."""

    def rateFor(self, winrt, percent, boost):
        v = voice(winrt)
        v.setRatePercent(percent, rateBoost=boost)
        return v._prosodyRate

    @pytest.mark.parametrize("boost", [True, False])
    def test_endpoints_hit_the_documented_range(self, winrt, boost):
        low = winrt.PROSODY_RATE_MIN if boost else winrt.PROSODY_RATE_PLAIN_MIN
        high = winrt.PROSODY_RATE_MAX if boost else winrt.PROSODY_RATE_PLAIN_MAX
        assert self.rateFor(winrt, 0, boost) == low
        assert self.rateFor(winrt, 100, boost) == high

    @pytest.mark.parametrize("boost", [True, False])
    def test_never_exceeds_the_saturation_point(self, winrt, boost):
        """Above 200% the engine produces byte-identical audio, so the control dies.

        Compared against a literal, not against PROSODY_RATE_MAX. 200 is a measured
        property of the Windows speech engine rather than a value we get to choose, so
        asserting against our own constant would let someone raise the constant and move
        the goalpost with it - which is exactly what a mutation test caught here.
        """
        ENGINE_SATURATES_AT = 200
        assert winrt.PROSODY_RATE_MAX <= ENGINE_SATURATES_AT, (
            "PROSODY_RATE_MAX above %d: the engine ignores the excess and the speed "
            "control goes dead at the top of its range" % ENGINE_SATURATES_AT
        )
        for percent in range(0, 101):
            assert self.rateFor(winrt, percent, boost) <= ENGINE_SATURATES_AT

    @pytest.mark.parametrize("boost", [True, False])
    def test_is_monotonic(self, winrt, boost):
        rates = [self.rateFor(winrt, p, boost) for p in range(0, 101)]
        assert rates == sorted(rates)

    @pytest.mark.parametrize("percent", [-50, 150, 10 ** 6])
    def test_out_of_range_input_clamps(self, winrt, percent):
        rate = self.rateFor(winrt, percent, True)
        assert winrt.PROSODY_RATE_MIN <= rate <= winrt.PROSODY_RATE_MAX

    def test_boost_off_stays_near_normal_speech(self, winrt):
        """The point of turning boost off is fine control, not a second wide range."""
        assert winrt.PROSODY_RATE_PLAIN_MIN > winrt.PROSODY_RATE_MIN
        assert winrt.PROSODY_RATE_PLAIN_MAX < winrt.PROSODY_RATE_MAX

    def test_the_emitted_ssml_carries_the_mapped_rate(self, winrt):
        v = voice(winrt)
        v.setRatePercent(100, rateBoost=True)
        assert 'rate="%d%%"' % winrt.PROSODY_RATE_MAX in v._buildSsml("x")
