# -*- coding: utf-8 -*-
"""Single source of truth for the desktop app version.

Bump this, tag the repo `desktop-v<x.y.z>` and attach the built .exe to the
GitHub Release — installed copies compare against the latest tag and offer to
self-update.
"""

APP_NAME = "Glasses Filler"
ORG_NAME = "Alensa"
__version__ = "1.1.1"

# Repo that hosts the .exe releases (public code repo).
RELEASE_REPO = "bgal-png/Glasses-Import-Filler"
# Release tags for the desktop app are prefixed so they don't clash with any
# future web-app tags in the same repo.
RELEASE_TAG_PREFIX = "desktop-v"
