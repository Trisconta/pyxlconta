""" Word wrapping and Latin-1 (ISO-8859-1) simplification.
"""

import unicodedata

def ascii_7bit(text, what_none="--"):
    if isinstance(text, (list, tuple)):
        return "?"
    if text is None:
        text = what_none
    astr = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    return astr

