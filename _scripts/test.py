__author__ = 'Binh'
from odf.opendocument import Spreadsheet
from odf.opendocument import load
from odf.table import TableRow,TableCell
from odf.text import P
doc = load("../texts/text.ods")
d = doc.spreadsheet
rows = d.getElementsByType(TableRow)
for row in rows[:2]:
    cells = row.getElementsByType(TableCell)
    for cell in cells:
        tps = cell.getElementsByType(P)
        if len(tps) > 0:
            for x in tps:
                print x.firstChild