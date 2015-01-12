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
merged_book = []  # contain all text after merge
# This is specific case for About of WAA team
CustomerCareTextSheet = 'ABOUT'
CustomerCareTextID = 'STRING_CC'
# global CustomerCareTextRow = 0


#define text file
ios_text_file = sys.argv[1]
android_text_file = sys.argv[2]

#
# fsheet in UPPER CASE
#
#'---------------------------------------------------------------------------------------------------'
def find_text(book, fsheet, ftext):
    for (sheet_id, (sheetName, data)) in enumerate(book):

        sheetName = sheetName.strip().upper()

        if sheetName != fsheet:  # skip other sheet
            continue

        for (r, c) in data:
            if r == 0 or c != 0:
                continue
            if data.has_key((r, 0)) and data[r, 0] == ftext:
                #Return sheetid and row in sheet
                return (sheet_id, r)
        return (sheet_id, -1)
    return None


#'---------------------------------------------------------------------------------------------------'
# This function will parse an ods file to a book
def parse_ods(file, isOriginalFile):
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
                        if isOriginalFile:
                            if name.strip().upper() == CustomerCareTextSheet and text == CustomerCareTextID:
                                global CustomerCareTextRow
                                CustomerCareTextRow = r
                                print 'CustomerCareTextRow =  %d' % r
                    c = c + 1
        book.append((name, data))
    return book


#'---------------------------------------------------------------------------------------------------'
def merge_text(android_text, ios_text):
    print '-------------------------------MERGING TEXT---------------------------------------'
    book = []
    i_sheet = []
    # Firstly :  Check for modified text and added text
    for (sheet, (sheetName, data)) in enumerate(ios_text):  # Loop all sheet in IOS_TEXT
        print 'SHEET : ' + sheetName + '    sheetid : %d' % sheet
        i_sheet.append(sheetName)
        merged_data = dict()
        max_row = 0
        max_col = 0
        mapping_sheet = -1
        # CustomerCareTextRow = 0
        # --------------------Start checking modified text---------------------------
        for (r, c) in data:
            # Skip check HEADER row
            if max_col < c:
                max_col = c
            if max_row < r:
                max_row = r

            if r == 0:
                # print '----------------  ADDED HEADER ROW'
                merged_data[r, c] = data[r, c]  #DONE
                continue

            # Start checking TEXT_ID
            if data.has_key((r, 0)):

                if sheetName.strip().upper() == CustomerCareTextSheet:
                    # if data[r, 0] == CustomerCareTextID :
                    # CustomerCareTextRow = r
                    # print 'Skip CustormerCare Text, Add later CustomerCareTextRow = %d' %CustomerCareTextRow
                    if r == CustomerCareTextRow:
                        continue

                row_replace = find_text(android_text, sheetName.strip().upper(), data[r, 0])

                if row_replace != None and row_replace[1] != -1:
                    mapping_sheet = row_replace[0]
                    if c == 0:
                        merged_data[r, c] = data[r, c]  # DONE
                    else:
                        r_data = android_text[row_replace[0]][1]
                        r_row = row_replace[1]
                        if r_data.has_key((r_row, c)):
                            merged_data[r, c] = r_data[r_row, c]
                else:
                    merged_data[r, c] = data[r, c]  # DONE

        print 'SHEET : ' + sheetName + '    sheetid : %d max_row = %d max_col = %d ' % (sheet, max_row, max_col)

        # --------------------Start checking added text---------------------------
        if mapping_sheet != -1:
            temp_data = android_text[mapping_sheet][1]
            backup_row = max_row
            for (r, c) in temp_data:
                if r != 0 and temp_data.has_key((r, 0)):
                    row_replace = find_text(ios_text, sheetName.strip().upper(), temp_data[r, 0])
                    if row_replace != None and row_replace[1] == -1:
                        if backup_row + r > max_row:
                            max_row = backup_row + r
                        merged_data[backup_row + r, c] = temp_data[r, c]

        # Re-added customercare text
        if sheetName.strip().upper() == CustomerCareTextSheet:
            for (r, c) in data:
                if r == CustomerCareTextRow:
                    if data.has_key((r, 0)):
                        merged_data[max_row + 1, c] = data[r, c]

        # Add data to merged book
        book.append((sheetName, merged_data))
    # Then check for added SHEET
    print '-------------------------------START ADDED SHEET----------------------------------'
    # print 'IOS ALL SHEET  : '
    # print i_sheet
    for (sheet, (sheetName, data)) in enumerate(android_text):
        if sheetName not in i_sheet:
            print 'START ADDED THIS SHEET TO IOS_TEXT		.' + sheetName + '.'
            book.append((sheetName, data))
    print '-------------------------------MERGING END----------------------------------------'
    return book

#'-------------------------MAIN CODE---------------------------------------------------------------------'
# Start parse text file
book_ios = parse_ods(ios_text_file, 1)
book_android = parse_ods(android_text_file, 0)
merged_book = merge_text(book_android, book_ios)
#------------------------------------------------
#write utf-8 to file
def write_utf8(file, text):
    text = text.encode('utf-8')
    arr = array("h")
    arr.append(len(text))
    arr.tofile(file)
    file.write(text)

#language list
languages = ['EN', 'FR', 'DE', 'SP', 'IT', 'BR', 'JP', 'KR', 'CN', 'RU', 'TR', 'AR']

#output files 
header_file = open('TEXT.h', 'w')
main_lang_files = []
for lang in languages:
    main_lang_files.append(open('text_' + lang + '.lang', 'wb'))

#char maps
char_sets = []
for lang in languages:
    set = []
    for c in xrange(ord(' '), ord('~') + 1):
        set.append(c)
    char_sets.append(set)

#open graph
openGraph = dict()

#for each sheet
for (sheet, (sheetName, data)) in enumerate(merged_book):

    sheetName = sheetName.strip().upper()

    idCol = -1
    ogCol = -1
    langCol = dict()

    #parse first row
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

    #skip if missing id header
    if idCol == -1:
        print 'Skip sheet ' + sheetName + '. Missing "ID" header';
        continue

    #skip if missing any language
    if len(langCol) != len(languages):
        print 'Skip sheet ' + sheetName + '. Missing language headers';
        continue

    #parse id column
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
for i, set in enumerate(char_sets):
    set.sort()
    file = open('font_' + languages[i] + '_map.cod', 'wb')
    for c in set:
        file.write(unichr(c).encode('UTF-16LE'))
		