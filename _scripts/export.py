__author__ = 'Binh'

import sys


# to create new ods file
from odf.opendocument import OpenDocumentSpreadsheet
# to load an ods file
from odf.opendocument import load

from odf.style import Style, TableColumnProperties
from odf.table import Table, TableRow, TableColumn, TableCell, CoveredTableCell
from odf.text import P


def make_ods():

    ods = load("text.ods")

    col = Style(name='col', family='table-column')
    col.addElement(TableColumnProperties(columnwidth='1in'))

    # Can define the table name here
    table = Table(name="DefaultSheet")
    table.addElement(TableColumn(numbercolumnsrepeated=3, stylename=col))
    ods.spreadsheet.addElement(table)

    # Add first row with cell spanning columns A-C
    tr = TableRow()
    table.addElement(tr)
    tc = TableCell(numbercolumnsspanned=3)
    tc.addElement(P(text="ABC1"))
    tr.addElement(tc)
    # Uncomment this to more accurately match native file
    # #tc = CoveredTableCell(numbercolumnsrepeated=2)
    # #tr.addElement(tc)

    # Add two more rows with non-spanning cells
    for r in (2, 3):
        tr = TableRow()
        table.addElement(tr)
        for c in ('A', 'B', 'C'):
            tc = TableCell()
            tc.addElement(P(text='%s%d' % (c, r)))
            tr.addElement(tc)

    assert isinstance(ods.save, object)
    ods.save("text-merged-result.ods")

make_ods()