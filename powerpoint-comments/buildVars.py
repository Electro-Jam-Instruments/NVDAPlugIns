# -*- coding: UTF-8 -*-

# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

# Full geance with samples can be found at:
# https://github.com/nvdaaddons/AddonTemplate/blob/master/buildVars.py

# Since some strings in addon_info are translatable,
# we need to define the translatable string function.
import gettext
_ = gettext.gettext

addon_info = {
    # for previously unpublished addons, please follow the community guidelines at:
    # https://bitly.com/NVDAAdd-Ons
    # add-on Name, internal for nvda
    "addon_name": "powerPointComments",
    # Add-on summary, usually the user visible name of the addon.
    # Translators: Summary for this add-on to be shown on installation and add-on information.
    "addon_summary": _("Accessible PowerPoint Comment Navigation"),
    # Add-on description
    # Translators: Long description to be shown for this add-on on add-on information from add-ons manager
    "addon_description": _("""Navigate and read PowerPoint comments with keyboard shortcuts and automatic announcements."""),
    # version
    "addon_version": "0.0.80",
    # Author(s)
    "addon_author": "Electro Jam Instruments <contact@electrojam.com>",
    # URL for the add-on documentation support
    "addon_url": "https://github.com/Electro-Jam-Instruments/NVDAPlugIns",
    # Documentation file name
    "addon_docFileName": None,
    # Minimum NVDA version supported
    "addon_minimumNVDAVersion": "2024.1",
    # Last NVDA version supported/tested
    "addon_lastTestedNVDAVersion": "2025.1",
    # Add-on update channel (default is stable)
    "addon_updateChannel": None,
}

# Define the python files that are the sources of your add-on.
# You can use glob expressions here, they will be expanded.
pythonSources = [
    "addon/appModules/*.py",
]

# Files that contain strings for translation.
i18nSources = pythonSources

# Files that will be ignored when building the nvda-addon file.
excludedFiles = []
