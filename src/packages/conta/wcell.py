""" Excel cells abstractions.
"""

class GenCell:
    """ Abstract cell class. """
    def __init__(self, data=None, name="c"):
        self.name = name
        self._data = data

    def get_value(self):
        """ Returns the data """
        assert self._data, self.name
        return self._data


class WCell(GenCell):
    """ WCell class """
    def_empty_cell = "-"

    def __init__(self, data, name="cell"):
        """ Excel cell abstractions. """
        super().__init__(data, name=name)
        self._orig = data
        self._simple, self._value = "", None
        self._str_cache, self._ref = self._cell_linear(data)

    def to_string(self):
        if self._simple:
            return self._simple
        return self._get_string()

    def lower(self):
        """ Lower-case, when applicable! """
        return self._simple.lower()

    def _cell_linear(self, data):
        """ Simple linearization of an Excel cell. """
        if isinstance(data, tuple):
            ref, val = data
        else:
            ref, val = None, data
        if val is None:
            astr = WCell.def_empty_cell
        elif isinstance(val, float):
            astr = f"{val:.2f}"
        else:
            astr = str(val)
        self._simple = astr
        self._value = val
        #astr = f"{ref};val=<{val}>(len:{len(data)})"
        return astr, ref

    def _get_string(self):
        return self._str_cache

    def __str__(self):
        return self._get_string()

    def __repr__(self):
        astr = self._str_cache
        return f"{repr(astr)}"
