# -*- coding: UTF-8 -*-
# Gettext tool for SCons
# Based on NVDA addon template
# Copyright (C) 2012-2025 NVDA Add-on team contributors
# This file is covered by the GNU General Public License.

"""Gettext tool for SCons - generates .mo files from .po files and .pot files from source."""

from SCons.Action import Action


def exists(env):
    return True


XGETTEXT_COMMON_ARGS = (
    "--msgid-bugs-address='$gettext_package_bugs_address' "
    "--package-name='$gettext_package_name' "
    "--package-version='$gettext_package_version' "
    "--keyword=pgettext:1c,2 "
    "-c -o $TARGET $SOURCES"
)


def generate(env):
    env.SetDefault(gettext_package_bugs_address="nvda-translations@groups.io")
    env.SetDefault(gettext_package_name="")
    env.SetDefault(gettext_package_version="")

    env["BUILDERS"]["gettextMoFile"] = env.Builder(
        action=Action("msgfmt -o $TARGET $SOURCE", "Compiling translation $SOURCE"),
        suffix=".mo",
        src_suffix=".po",
    )

    env["BUILDERS"]["gettextPotFile"] = env.Builder(
        action=Action("xgettext " + XGETTEXT_COMMON_ARGS, "Generating pot file $TARGET"),
        suffix=".pot"
    )

    env["BUILDERS"]["gettextMergePotFile"] = env.Builder(
        action=Action(
            "xgettext " + "--omit-header --no-location " + XGETTEXT_COMMON_ARGS,
            "Generating pot file $TARGET"
        ),
        suffix=".pot",
    )
