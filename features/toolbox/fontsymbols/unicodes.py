# -*- coding: utf-8 -*-

import unicodedata
import logging

def update_search_index(search_engine):
    search_writer = search_engine.writer()

    # Generate a list of all Unicode character names
    unicode_characters = []
    for codepoint in range(0x30000):  # relevant Unicode range
        try:
            char_name = unicodedata.name(chr(codepoint))
            unicode_characters.append((char_name, codepoint))
        except ValueError:
            # Skip characters without a name
            continue
        # except Exception as e:
        #     logging.error(f"Error processing codepoint {codepoint}: {e}")
    
    logging.info(f"Found {len(unicode_characters)} Unicode characters.")

    # unicode_fonts = [
    #     ("Segoe UI", 0x0000, 0x0780)
    #     ("Segoe UI Emoji", 0x1f600, 0x1f650)
    # ]

    # def _get_font(code):
    #     for font, start, end in unicode_fonts:
    #         if start <= code < end:
    #             return font
    #     return "Segoe UI Symbol"

    for label, code in unicode_characters:
        search_writer.add_document(
            module="unicodes",
            fontlabel="Unicode Symbols",
            # fontname=UnicodeSymbols.rendering_font,
            fontname=None,
            unicode=chr(code),
            label=f"{label} (U+{code:04X})",
            keywords=label.lower().split()
        )

    search_writer.commit()
