# -*- coding: utf-8 -*-

import bkt


if bkt.helpers.confirmation("Dev module replaced by devkit feature folder. Remove module from config?"):
    modules = bkt.config.modules or []
    try:
        modules.remove("modules.dev")
        bkt.config.set_smart("modules", modules)
    except ValueError:
        pass
