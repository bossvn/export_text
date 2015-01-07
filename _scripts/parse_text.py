import sys
import os
import json
from array import array
from odf.opendocument import Spreadsheet
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P

# parse ods
book_ios = []
book_android = []

# define text file
ios_text = sys.argv[1]
android_text = sys.argv[2]

# This function will parse an ods file to a book
def parse_ods(file):
    book = []
    # for each sheet
    for table in load(file).spreadsheet.getElementsByType(Table):
        name = table.getAttribute('name')
        data = dict()
        # for each row
        for r, row in enumerate(table.getElementsByType(TableRow)):
            c = 0
            # for each cell
            for cell in row.getElementsByType(TableCell):
                repeat = cell.getAttribute("numbercolumnsrepeated")
                if not repeat:
                    repeat = 1

                text = unicode("")
                for p in cell.getElementsByType(P):
                    if not len(p.attributes):
                        for t in p.childNodes:
                            text = text + unicode(t)

                for rr in range(int(repeat)):
                    if text:
                        data[r, c] = text
                        print ' Sheet : ' + name.encode(sys.stdout.encoding,
                                                        errors='replace') + '  Row %d , Cell %d   :  ' % (
                            r, c) + text.encode(sys.stdout.encoding, errors='replace')
                    c = c + 1
        book.append((name, data))
    return book

# Start parse text file
book_ios = parse_ods(ios_text)
# ------------------------------------------------
# write utf-8 to file
def write_utf8(file, text):
    text = text.encode('utf-8')
    arr = array("h")
    arr.append(len(text))
    arr.tofile(file)
    file.write(text)

# language list
languages = ['EN', 'FR', 'DE', 'SP', 'IT', 'BR', 'JP', 'KR', 'CN', 'RU', 'TR', 'AR']

# output files
header_file = open('TEXT.h', 'w')
main_lang_files = []
for lang in languages:
    main_lang_files.append(open('text_' + lang + '.lang', 'wb'))

# char maps
char_sets = []
for lang in languages:
    set = []
    for c in xrange(ord(' '), ord('~') + 1):
        set.append(c)
    char_sets.append(set)

#open graph
openGraph = dict()
#  ---------------------------------------------------------------------
# Export text to binary file
# ---------------------------------------------------------------------
# for each sheet
for (sheet, (sheetName, data)) in enumerate(book_ios):

    sheetName = sheetName.strip().upper()

    idCol = -1
    ogCol = -1
    langCol = dict()

    # parse first row
    for (row, col) in data:
        if row == 0:
            val = str(data[row, col]).strip(' :-').upper()
            if val == 'ID':
                idCol = col
            if val == 'OG':
                ogCol = col
            if val in languages:
                i = languages.index(val)
                langCol[i] = col

    # skip if missing id header
    if idCol == -1:
        print 'Skip sheet ' + sheetName + '. Missing "ID" header';
        continue

    # skip if missing any language
    if len(langCol) != len(languages):
        print 'Skip sheet ' + sheetName + '. Missing language headers';
        continue

    # parse id column
    lines = []
    for (row, col) in data:
        if col == idCol and row > 0:
            val = str(data[row, col])
            name = val.strip().upper().replace(' ', '_')
            lines.append((row, name))

            #sort by row index
    lines.sort()

    #output DLC tabs in different files
    lang_files = main_lang_files
    if sheetName.find('DLC') != -1:
        lang_files = []
        for lang in languages:
            lang_files.append(open('text_' + lang + '_' + sheetName + '.lang', 'wb'))

    #begin tab export

    namePrefix = 'TEXT_' + sheetName + '_'

    header_file.write('\n/// ' + sheetName + '\n')
    for file in lang_files:
        write_utf8(file, namePrefix)
        array("i", (sheet, len(lines))).tofile(file)

    #export each text
    for (i, (row, name)) in enumerate(lines):
        #text id (matches id in aurora)
        id = sheet * 1024 + i

        #output to header file
        fullName = namePrefix + name
        header_file.write('#define\t' + fullName + '\t\t(' + str(id) + ')\n')

        #open graph
        ogLabels = []
        if ogCol != -1 and (row, ogCol) in data:
            ogLabels = str(data[row, ogCol]).split(',')
            ogLabels = [x.strip().lower() for x in ogLabels]
            for label in ogLabels:
                if not label in openGraph:
                    openGraph[label] = dict()
                openGraph[label][fullName] = dict()


        #each language
        for i in langCol:
            col = langCol[i]
            text = ''
            if (row, col) in data:
                text = data[row, col]

            if text == '' and i != 0:
                col = langCol[0]
                if (row, col) in data:
                    text = data[row, col]

            #output to lang file
            write_utf8(lang_files[i], name)
            write_utf8(lang_files[i], text)

            #output to char map
            set = char_sets[i]
            for c in text:
                if not ord(c) in set:
                    set.append(ord(c))

            #output to openGraph dict
            for label in ogLabels:
                openGraph[label][fullName][languages[i].encode('UTF-16LE')] = text

    #export arabic special symbols and diacritics afterwards
    set = char_sets[11]
    for c in xrange(0xFE80, 0xFEFF):
        if not c in set:
            set.append(c)
    for c in xrange(0x064B, 0x0658):
        if not c in set:
            set.append(c)
    for c in xrange(0xFDF1, 0xFDF3):
        if not c in set:
            set.append(c)

#export open graph files		
for label in openGraph:
    file = open(label + '_texts' + '.json', 'w')
    file.write(json.dumps(openGraph[label], encoding='UTF-16LE', sort_keys=True, indent=4))  #Pretty
    #file.write(json.dumps(openGraph[label], encoding = 'UTF-16LE', sort_keys = True, separators=(',',':'))) #Compact
    file.close()

#export font map files	
# for i, set in enumerate(char_sets):
# 	set.sort()
# 	file = open('font_' + languages[i] + '_map.cod', 'wb')
# 	for c in set:
# 		file.write(unichr(c).encode('UTF-16LE'))
		