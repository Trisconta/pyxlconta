""" Excel simple creator.

Uses openpyxl !
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from .wcell import GenCell


class ExcelGera(GenCell):
    def __init__(self, dados, outname="out.xlsx", name=""):
        super().__init__(data=dados, name=name if name else outname)
        self.outname = outname

    def generate(self):
        wbk = Workbook()
        ws = wbk.active
        ws.title = "Folha1"
        for row in self._data:
            ws.append(row)
        self._ajustar_larguras(ws)
        wbk.save(self.outname)

    def _ajustar_larguras(self, ws):
        """ Larguras de colunas. """
        pigment = 4
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value) if cell.value is not None else ""
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max_len + pigment
