__author__ = 'Binh'

import sys


# to create new ods file
from odf.opendocument import OpenDocumentSpreadsheet
# to load an ods file
from odf.opendocument import load

from odf.style import Style, TableColumnProperties
from odf.table import Table, TableRow, TableColumn, TableCell, CoveredTableCell
from odf.text import P

ods_result = load("../texts/text_simple.ods")

ods_android = load("../texts/text_simple.ods")
book_android = []


def parse_ods(file):
    book = []
    # for each sheet
    for table in file.spreadsheet.getElementsByType(Table):
        name = table.getAttribute('name')
        data = dict()
        # for each row
        for r, row in enumerate(table.getElementsByType(TableRow)):
            c = 0
            # for each cell
            for cell in row.getElementsByType(TableCell):
                # Check if there some merged cells
                repeat = cell.getAttribute("numbercolumnsrepeated")
                if not repeat:
                    repeat = 1

                # Load cell text
                text = unicode("")
                for p in cell.getElementsByType(P):
                    # print p.encode(sys.stdout.encoding, errors='replace')
                    # p.setAttribute('text',"aaa")
                    # print p.childNodes
                    if not len(p.attributes):
                        for t in p.childNodes:
                             p.removeChild(t)
                    assert isinstance(p.childNodes, object)
                    print p.childNodes

                    p.addText("abcss")

                    if not len(p.attributes):
                        for t in p.childNodes:
                            text += unicode(t)
                            print unicode(t).encode(sys.stdout.encoding, errors='replace')

                # Applied text for merged cell
                for rr in range(int(repeat)):
                    if text:
                        data[r, c] = text
                        print ' Sheet : ' + name.encode(sys.stdout.encoding,
                                                        errors='replace') + '  Row %d , Cell %d   :  ' % (
                            r, c) + text.encode(sys.stdout.encoding, errors='replace')
                    c += 1
        book.append((name, data))
    return book

# --------------------------------------------------------------------------------------
def make_ods():

    # Loop to count all sheet (table)
    for table in ods_result.spreadsheet.getElementsByType(Table):
        # table.setAttribute('name',"AAA")
        name = table.getAttribute('name')
        print name

        for r, row in enumerate(table.getElementsByType(TableRow)):
            c = 0
            # for cell in row.getElementsByType(TableCell):
                # print cell.getAttribute('text')




    '''
    # temporary rem

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
    '''

    assert isinstance(ods_result.save, object)
    ods_result.save("text-merged-result.ods")

book_android = parse_ods(ods_android)
make_ods()