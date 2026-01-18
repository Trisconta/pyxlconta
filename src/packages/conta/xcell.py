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


class XCell(GenCell):
    """ XCell class """
    def_empty_cell = "-"

    def __init__(self, data, name="cell"):
        """ Excel cell abstractions. """
        super().__init__(data, name=name)
        self._orig = data
        self._simple, self._value = "", None

    def lower(self):
        """ Lower-case, when applicable! """
        return self._simple.lower()

    def _cell_linear(self):
        """ Simple linearization of Excel cell. """
        if self._value is not None:
            return self._simple
        data = self._data
        if isinstance(data, tuple):
            ref, val = data
        else:
            ref, val = None, data
        astr = XCell.def_empty_cell if val is None else str(val)
        self._simple = astr
        self._value = val
        return astr

    def _get_string(self):
        astr = self._cell_linear()
        return astr

    def __str__(self):
        return self._get_string()

    def __repr__(self):
        astr = self._get_string()
        return f"{repr(astr)}"
