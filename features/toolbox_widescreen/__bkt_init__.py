# -*- coding: utf-8 -*-

def initiate():
    #FIXME: conflict handling needs to be done centrally during import
    import sys
    if "toolbox" not in sys.modules:
        import my_toolbox
    else:
        import logging
        logging.error("conflicting modules detected")

bkt_feature = {
    "name": "PowerPoint Toolbox Widescreen",
    "relevant_apps": ["Microsoft PowerPoint"],
    "contructor": initiate,
}
