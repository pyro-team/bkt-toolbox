# -*- coding: utf-8 -*-

def initiate():
    #FIXME: conflict handling needs to be done centrally during import
    import sys
    if "toolbox_widescreen" not in sys.modules:
        import toolbox_powerpoint
    else:
        import logging
        logging.error("conflicting modules detected")

bkt_feature = {
    "name": "PowerPoint Toolbox",
    "relevant_apps": ["Microsoft PowerPoint"],
    "contructor": initiate,
}
