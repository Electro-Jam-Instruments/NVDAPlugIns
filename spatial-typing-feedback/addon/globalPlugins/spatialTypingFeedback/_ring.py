# Settings ring integration.
#
# See docs/architecture.md.
#
# One selector plus a set of controls that follow it. Pick a stream, then adjust voice,
# speed, volume and pan for that stream - rather than a separate named slot per stream
# per property, which multiplies badly and means learning several labels for one idea.
#
# Reached with the keys the user already has: control+NVDA+left/right to move between
# settings, up/down to change the value. No new gestures.
#
# It works because the ring is a single swappable object. Every ring script in
# globalCommands goes through globalVars.settingsRing, and synthDriverHandler updates
# the *existing* object on synth change rather than replacing it, so a subclass survives.

import globalVars
from autoSettingsUtils.driverSetting import NumericDriverSetting
from logHandler import log
from synthSettingsRing import SynthSetting, SynthSettingsRing

from . import _config
from ._streams import PAN_MAX, PAN_MIN, Stream

try:
    _
except NameError:
    def _(text):
        return text

#: Order the stream selector cycles through. MAIN is NVDA's own voice, included so one
#: selector covers everything rather than the user having to remember which controls
#: live where.
STREAMS = (Stream.MAIN, Stream.CHARS, Stream.ANNOTATIONS, Stream.ERRORS)


def _streamLabel(stream):
    """The one name for a stream.

    Used both as the selector's value and as the prefix on every slot below it, so there
    is a single word per stream. Two vocabularies meant the selector said "main voice"
    and so did the voice slot - indistinguishable by ear.
    """
    return {
        # Translators: The main NVDA voice, in the settings ring stream selector.
        Stream.MAIN: _("main"),
        # Translators: The typed character stream, in the settings ring.
        Stream.CHARS: _("characters"),
        # Translators: The stream carrying NVDA's spoken notes about a marked word.
        Stream.ANNOTATIONS: _("notes"),
        # Translators: The error alert sound stream, in the settings ring. Named for
        # what it is - a sound - to distinguish it from the spoken notes.
        Stream.ERRORS: _("alert"),
    }.get(stream, str(stream))


def _streamWord(stream):
    return _streamLabel(stream)


def _panLabel(pan):
    """Speak a pan value as a position, not a bare number.

    "minus forty five" is not usable at a boosted rate; "left 45" is.
    """
    if pan == 0:
        # Translators: The centre pan position.
        return _("centre")
    if pan < 0:
        # Translators: A pan position to the left; {amount} is a number.
        return _("left {amount}").format(amount=-pan)
    # Translators: A pan position to the right; {amount} is a number.
    return _("right {amount}").format(amount=pan)


def _notAvailable():
    # Translators: Reported for a control that does not apply to the selected stream.
    return _("not available")


class _PluginSetting(SynthSetting):
    """A ring slot backed by our plugin rather than the synth.

    NVDA's SynthSetting writes to config.conf["speech"][synthName][settingId], which is
    the wrong place for us and is not in that section's spec, so everything that touches
    a value is overridden.

    Slots that do not apply to the selected stream report "not available" and ignore
    changes, rather than disappearing. A control that vanishes and reappears as the
    selector moves is more disorienting than one that says it does not apply.
    """

    #: Streams this slot can act on. None means all of them.
    appliesTo = None

    #: Short label, combined with the stream name: "characters speed".
    shortLabel = ""

    def __init__(self, plugin, setting):
        self._plugin = plugin
        super().__init__(None, setting)

    @property
    def stream(self):
        return self._plugin.selectedStream

    @property
    def applicable(self):
        return self.appliesTo is None or self.stream in self.appliesTo

    @property
    def dynamicName(self):
        """Name the slot after the stream it is currently acting on.

        "Spatial voice" does not say whose voice. "characters voice" does, and it means
        the selected stream is re-stated on every step through the ring rather than
        having to be remembered.
        """
        if not self.shortLabel:
            return self.setting.displayName
        return "%s %s" % (_streamWord(self.stream), self.shortLabel)

    @property
    def announceThroughStream(self):
        """True when this slot should be announced by the stream itself.

        Depends only on which stream is selected, NOT on whether the slot applies to it.
        "not available" is a statement about the character stream, so it belongs in the
        character voice like everything else about that stream - hearing it in the main
        voice is exactly the confusion this routing exists to avoid.
        """
        return self._plugin.streamHasVoice(self.stream)

    def _announce(self, value):
        """Speak the change through the stream being adjusted, not the main voice.

        Tuning a voice while listening to a different voice, in a different place, at a
        different speed, tells you nothing about the one you are changing. The
        announcement has to be spoken BY the thing being tuned - then it doubles as the
        sample.
        """
        if not self.announceThroughStream:
            return value
        spoken = "%s %s" % (self.dynamicName, value)
        try:
            if self._plugin.announceThroughStream(self.stream, spoken):
                # Said it ourselves; keep NVDA's main voice quiet about this one.
                self._plugin.suppressText(spoken)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error announcing change", exc_info=True)
        return value

    def increase(self):
        return self._announce(super().increase() if self.applicable else _notAvailable())

    def decrease(self):
        return self._announce(super().decrease() if self.applicable else _notAvailable())

    def increaseLarge(self):
        return self._announce(
            super().increaseLarge() if self.applicable else _notAvailable()
        )

    def decreaseLarge(self):
        return self._announce(
            super().decreaseLarge() if self.applicable else _notAvailable()
        )

    def first(self):
        return self._announce(super().first() if self.applicable else _notAvailable())

    def last(self):
        return self._announce(super().last() if self.applicable else _notAvailable())

    def _get_reportValue(self):
        if not self.applicable:
            return _notAvailable()
        return self._getReportValue(self.value)


def _numeric(settingId, label, default, minVal, maxVal, normalStep, largeStep):
    return NumericDriverSetting(
        settingId,
        label,
        availableInSettingsRing=True,
        defaultVal=default,
        minVal=minVal,
        maxVal=maxVal,
        minStep=1,
        normalStep=normalStep,
        largeStep=largeStep,
    )


class StreamSetting(_PluginSetting):
    """Selects which stream every slot below acts on."""

    def __init__(self, plugin):
        super().__init__(
            plugin,
            # Translators: Label for the stream selector in the settings ring.
            _numeric("stfStream", _("stream"), 0, 0, len(STREAMS) - 1, 1, 1),
        )

    def _get_value(self):
        return self._plugin.selectedStreamIndex

    def _set_value(self, value):
        self._plugin.selectedStreamIndex = max(0, min(len(STREAMS) - 1, int(value)))

    def _getReportValue(self, val):
        return _streamLabel(STREAMS[max(0, min(len(STREAMS) - 1, int(val)))])


class PanSetting(_PluginSetting):
    """Stereo position of the selected stream."""

    #: NVDA's own voice is not ours to move - we do not own its player.
    appliesTo = (Stream.CHARS, Stream.ANNOTATIONS, Stream.ERRORS)

    def __init__(self, plugin):
        super().__init__(
            plugin,
            # Translators: Label for the pan control in the settings ring.
            _numeric("stfPan", _("Spatial pan"), 0, PAN_MIN, PAN_MAX, 5, 25),
        )
        # Translators: Short label for the pan control, e.g. "characters pan".
        self.shortLabel = _("pan")

    def _get_value(self):
        if not self.applicable:
            return 0
        pan, _vol = _config.panAndVolume(self.stream)
        return pan

    def _set_value(self, value):
        self._plugin.setStreamPan(self.stream, int(value))

    def _getReportValue(self, val):
        return _panLabel(int(val))


class VolumeSetting(_PluginSetting):
    """Level of the selected stream.

    For the main voice this is NVDA's own volume. For the error alert there is one extra
    position below zero meaning "follow NVDA's sound volume", so the default is
    reachable again after overriding it - a setting you can leave but never return to is
    a trap.
    """

    def __init__(self, plugin):
        super().__init__(
            plugin,
            # Translators: Label for the volume control in the settings ring.
            _numeric("stfVolume", _("Spatial volume"), 50, 0, 100, 5, 20),
        )
        # Translators: Short label for the volume control, e.g. "characters volume".
        self.shortLabel = _("volume")

    def _get_min(self):
        # One step below silent, reserved for "follow NVDA".
        return _config.UNSET if self.stream is Stream.ERRORS else 0

    def _set_min(self, value):
        pass

    def _get_value(self):
        if self.stream is Stream.MAIN:
            return self._plugin.getMainVolume()
        if self._plugin.isFollowingNvdaVolume(self.stream):
            return _config.UNSET
        return self._plugin.streamVolume(self.stream)

    def _set_value(self, value):
        if self.stream is Stream.MAIN:
            self._plugin.setMainVolume(int(value))
            return
        self._plugin.setStreamVolume(self.stream, int(value))

    def _getReportValue(self, val):
        if int(val) == _config.UNSET:
            # Say so, or a level that quietly tracks another setting looks like a bug.
            # Translators: Volume setting that tracks NVDA's own sound volume.
            # {value} is the level it currently works out to.
            return _("following NVDA, {value}").format(
                value=self._plugin.nvdaSoundVolume(),
            )
        return str(int(val))


class SpeedSetting(_PluginSetting):
    """Speaking speed of the selected stream, 0-100.

    Absolute for both voices, on the same scale NVDA uses. The character voice does not
    track the main one - it is a voice in its own right, and a relative value meant it
    could never be put where you actually wanted it. The error alert is a sound, so it
    has no speed.
    """

    appliesTo = (Stream.MAIN, Stream.CHARS, Stream.ANNOTATIONS)

    def __init__(self, plugin):
        super().__init__(
            plugin,
            # Translators: Label for the speaking speed control in the settings ring.
            _numeric("stfSpeed", _("Spatial speed"), 50, 0, 100, 5, 25),
        )
        # Translators: Short label for the speed control, e.g. "characters speed".
        self.shortLabel = _("speed")

    def _get_value(self):
        if self.stream is Stream.MAIN:
            return self._plugin.getMainRate()
        return self._plugin.getStreamRate(self.stream)

    def _set_value(self, value):
        if self.stream is Stream.MAIN:
            self._plugin.setMainRate(int(value))
        else:
            self._plugin.setStreamRate(self.stream, int(value))


class PitchSetting(_PluginSetting):
    """Pitch of the selected stream, 0-100."""

    appliesTo = (Stream.MAIN, Stream.CHARS, Stream.ANNOTATIONS)

    def __init__(self, plugin):
        super().__init__(
            plugin,
            # Translators: Label for the pitch control in the settings ring.
            _numeric("stfPitch", _("Spatial pitch"), 50, 0, 100, 5, 25),
        )
        # Translators: Short label for the voice control, e.g. "characters pitch".
        self.shortLabel = _("pitch")

    def _get_value(self):
        if self.stream is Stream.MAIN:
            return self._plugin.getMainPitch()
        return self._plugin.getStreamPitch(self.stream)

    def _set_value(self, value):
        if self.stream is Stream.MAIN:
            self._plugin.setMainPitch(int(value))
        else:
            self._plugin.setStreamPitch(self.stream, int(value))


class VoiceSetting(_PluginSetting):
    """Which voice the selected stream uses."""

    appliesTo = (Stream.MAIN, Stream.CHARS, Stream.ANNOTATIONS)

    def __init__(self, plugin):
        super().__init__(
            plugin,
            # Translators: Label for the voice control in the settings ring.
            _numeric("stfVoice", _("Spatial voice"), 0, 0, 0, 1, 1),
        )
        # Translators: Short label for the voice control, e.g. "characters voice".
        self.shortLabel = _("voice")

    def _voices(self):
        if self.stream is Stream.MAIN:
            return self._plugin.getMainVoices()
        return self._plugin.getStreamVoices(self.stream)

    def _get_max(self):
        return max(0, len(self._voices()) - 1)

    def _set_max(self, value):
        pass

    def _get_value(self):
        if self.stream is Stream.MAIN:
            return self._plugin.getMainVoiceIndex()
        return self._plugin.getStreamVoiceIndex(self.stream)

    def _set_value(self, value):
        voices = self._voices()
        if not voices:
            return
        index = max(0, min(len(voices) - 1, int(value)))
        if self.stream is Stream.MAIN:
            self._plugin.setMainVoiceIndex(index)
        else:
            self._plugin.setStreamVoiceIndex(self.stream, index)

    def _getReportValue(self, val):
        voices = self._voices()
        if not voices:
            return _notAvailable()
        index = max(0, min(len(voices) - 1, int(val)))
        return voices[index][1]


class _PluginToggle(_PluginSetting):
    """An on/off ring slot.

    Built on the numeric base with a 0-1 range rather than NVDA's BooleanSynthSetting,
    so it inherits the same stream-awareness and announcement routing as every other
    slot here.
    """

    def __init__(self, plugin, settingId, label, shortLabel):
        super().__init__(plugin, _numeric(settingId, label, 1, 0, 1, 1, 1))
        self.shortLabel = shortLabel

    def _getReportValue(self, val):
        # Translators: An on/off setting value.
        return _("on") if int(val) else _("off")


class RateBoostSetting(_PluginToggle):
    """Rate boost, for either voice.

    Widens the range the speed control maps onto: 0.5x-1.5x off, 0.5x-6.0x on. Off gives
    fine control around normal speed, on gives reach at the cost of precision. That
    choice is as useful for the character voice as for the main one.
    """

    appliesTo = (Stream.MAIN, Stream.CHARS, Stream.ANNOTATIONS)

    def __init__(self, plugin):
        super().__init__(
            plugin,
            "stfRateBoost",
            # Translators: Label for the rate boost control in the settings ring.
            _("Spatial rate boost"),
            # Translators: Short label for rate boost, e.g. "main rate boost".
            _("rate boost"),
        )

    def _get_value(self):
        if self.stream is Stream.MAIN:
            return 1 if self._plugin.getMainRateBoost() else 0
        return 1 if self._plugin.getStreamRateBoost(self.stream) else 0

    def _set_value(self, value):
        enable = bool(int(value))
        if self.stream is Stream.MAIN:
            self._plugin.setMainRateBoost(enable)
        else:
            self._plugin.setStreamRateBoost(self.stream, enable)


class PunctuationSetting(_PluginToggle):
    """Short pauses at punctuation, per stream.

    Availability is checked against the engine rather than assumed. NVDA 2026.1.1 ships
    a helper DLL with no punctuation functions at all, so on that build there is nothing
    to toggle - and a control that silently refuses to change is worse than one that
    says it cannot.
    """

    appliesTo = (Stream.MAIN, Stream.CHARS, Stream.ANNOTATIONS)

    @property
    def applicable(self):
        if self.stream is Stream.MAIN:
            return self._plugin.mainSupportsPunctuation()
        if self.stream is Stream.CHARS:
            return self._plugin.streamSupportsPunctuation(self.stream)
        return False

    def __init__(self, plugin):
        super().__init__(
            plugin,
            "stfPunctuation",
            # Translators: Label for the punctuation pause control in the settings ring.
            _("Spatial punctuation pauses"),
            # Translators: Short label, e.g. "characters punctuation pauses".
            _("punctuation pauses"),
        )

    def _get_value(self):
        if self.stream is Stream.MAIN:
            return 1 if self._plugin.getMainPunctuation() else 0
        return 1 if self._plugin.getStreamPunctuation(self.stream) else 0

    def _set_value(self, value):
        enable = bool(int(value))
        if self.stream is Stream.MAIN:
            self._plugin.setMainPunctuation(enable)
        else:
            self._plugin.setStreamPunctuation(self.stream, enable)


class SpatialSettingsRing(SynthSettingsRing):
    """NVDA's ring with our slots appended after the synth's own."""

    def __init__(self, synth, extraSettings):
        # Set before super().__init__, which calls updateSupportedSettings.
        self._extraSettings = extraSettings
        super().__init__(synth)

    def _get_currentSettingName(self):
        setting = self._currentPluginSetting()
        if setting is not None:
            return setting.dynamicName
        return super()._get_currentSettingName()

    def _currentPluginSetting(self):
        if self._current is None or not self.settings:
            return None
        if self._current >= len(self.settings):
            return None
        setting = self.settings[self._current]
        return setting if isinstance(setting, _PluginSetting) else None

    def _announceMove(self, name):
        """Speak a slot name and value through the stream it belongs to.

        Stepping through the ring is most of what tuning actually involves, so this
        matters more than announcing value changes. globalCommands says
        "<name> <value>" through the main voice right after next()/previous() returns;
        when the slot belongs to the character stream, that has to come from the
        character voice instead or there is no way to judge what you are adjusting.
        """
        if not name:
            return name
        setting = self._currentPluginSetting()
        if setting is None or not setting.announceThroughStream:
            return name
        try:
            spoken = "%s %s" % (name, self.currentSettingValue)
            if setting._plugin.announceThroughStream(setting.stream, spoken):
                setting._plugin.suppressText(spoken)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error announcing slot move", exc_info=True)
        return name

    def next(self):
        return self._announceMove(super().next())

    def previous(self):
        return self._announceMove(super().previous())

    def updateSupportedSettings(self, synth):
        """Present OUR slots only.

        Deliberately replacing NVDA's own slots rather than appending to them. Appending
        meant Voice, Speed, Pitch and Volume each appeared twice - once from NVDA acting
        on the main synth, once from ours acting on the selected stream - which is
        confusing to hear and worse to use, since the two look identical but do different
        things depending on the selector.

        One selector, one set of controls, covering every stream. NVDA's synth settings
        remain available in its Voice settings dialog, and come straight back if this
        add-on is turned off.
        """
        extra = getattr(self, "_extraSettings", None)
        if not extra:
            super().updateSupportedSettings(synth)
            return
        self.settings = list(extra)
        if self._current is None or self._current >= len(self.settings):
            self._current = 0


def install(plugin):
    """Swap NVDA's ring for ours. Returns the original so it can be put back."""
    original = getattr(globalVars, "settingsRing", None)
    if original is None:
        log.warning("Spatial Typing Feedback: no settings ring to extend yet")
        return None
    try:
        import synthDriverHandler

        synth = synthDriverHandler.getSynth()
        extras = [
            StreamSetting(plugin),
            VoiceSetting(plugin),
            SpeedSetting(plugin),
            RateBoostSetting(plugin),
            PitchSetting(plugin),
            VolumeSetting(plugin),
            PunctuationSetting(plugin),
            PanSetting(plugin),
        ]
        globalVars.settingsRing = SpatialSettingsRing(synth, extras)
        log.info("Spatial Typing Feedback: settings ring slots installed")
        return original
    except Exception:  # noqa: BLE001
        log.error(
            "Spatial Typing Feedback: could not extend the settings ring; the built-in "
            "settings are unaffected",
            exc_info=True,
        )
        globalVars.settingsRing = original
        return None


def uninstall(original):
    if original is None:
        return
    try:
        globalVars.settingsRing = original
    except Exception:  # noqa: BLE001
        log.debugWarning("Error restoring the settings ring", exc_info=True)
