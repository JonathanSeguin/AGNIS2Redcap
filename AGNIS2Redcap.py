"""
Converts an AGNIS metadata report xml to
a RedCap data instrument (instrument.zip).
"""

import xlwt, xlrd, xlutils
import argparse
import re
import csv
import zipfile
import os

parser = argparse.ArgumentParser()
parser.add_argument("metadata_xls", help="filename")
args = parser.parse_args()

agnis_wb = xlrd.open_workbook(args.metadata_xls)
agnis_ws = agnis_wb.sheet_by_name('Sheet0')
agnis_head = [c.value for c in agnis_ws.row(10)]

head = [u"Variable / Field Name", u"Form Name", u"Section Header", u"Field Type", u"Field Label",
        u"Choices, Calculations, OR Slider Labels", u"Field Note", u"Text Validation Type OR Show Slider Number",
        u"Text Validation Min", u"Text Validation Max", u"Identifier?", u"Branching Logic (Show field only if...)",
        u"Required Field?", u"Custom Alignment", u"Question Number (surveys only)", u"Matrix Group Name",
        u"Matrix Ranking?", u"Field Annotation"]

content = []

form_name = "Form {0}".format(agnis_ws.cell(0, 1).value.split(' Indication')[0])
form_id = form_name.split(" - ")[0].replace(" ", "").replace("Revision", "r")
form_name = re.sub('[^A-Za-z0-9]+', '_', form_name)

form_pid = agnis_ws.cell(6, 1).value
form_ver = agnis_ws.cell(7, 1).value

section_header = u""
nrow = -1
choices = []

# loop through each row of the report xls
for agnis_nrow in range(11, agnis_ws.nrows):

    nmod = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Display Order')).value
    nqst = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Display Order')).value
    mod_pid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Public ID')).value if agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Public ID')).value else False
    mod_ver = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Version')).value if agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Version')).value else False
    q_pid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'CDE Public ID')).value
    q_ver = agnis_ws.cell(agnis_nrow, agnis_head.index(u'CDE Version')).value
    qst_long_name = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Long Name')).value
    data_type = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Data Type')).value
    valid_val = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Valid Value')).value
    val_meaning_text = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Text')).value
    val_meaning_pubid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Public ID')).value
    val_meaning_version = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Version')).value
    display_format = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Display Format')).value
    is_mandatory = "y" if (agnis_ws.cell(agnis_nrow, agnis_head.index(u'Answer is Mandatory')).value == "Yes") else ""
    normalized_curation = (agnis_ws.cell(agnis_nrow, agnis_head.index(u'Normalized Curation')).value == "Yes")
    question_instructions = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Instructions')).value

    if q_pid:
        # question
        nrow += 1
        content.append({k: u'' for k in head})
        content[nrow][u"Form Name"] = form_name

        redcapid = "f{0}r{1}xm{2}r{3}xq{4}r{5}".format(int(form_pid), str(form_ver).replace('.', '_'),
                                                      int(mod_pid), str(mod_ver).replace('.', '_'),
                                                      int(q_pid), str(q_ver).replace('.', '_'))

        content[nrow][u"Variable / Field Name"] = redcapid
        content[nrow][u"Section Header"] = section_header
        content[nrow][u"Field Label"] = qst_long_name.rstrip(":")

        # default to text
        content[nrow][u"Field Type"] = "text"

        if display_format == u"YYYY-MM-DD":
            content[nrow][u"Text Validation Type OR Show Slider Number"] = u"date_ymd"

        content[nrow][u"Field Note"] = question_instructions
        content[nrow][u"Required Field?"] = is_mandatory

        section_header = u""
    else:
        # module header
        section_header = agnis_ws.cell(agnis_nrow, agnis_head.index(u"Module Long Name")).value

    if val_meaning_pubid and val_meaning_text:
        # choice of question
        choices.append("a{0}r{1},{2}".format(int(val_meaning_pubid),
                                             str(val_meaning_version).replace('.', '_'),
                                             valid_val))
    elif nrow != -1:
        # end of a streak of rows enumerating choices for question
        choices = "|".join(choices)
        if choices:
            if section_header:
                content[nrow][u"Choices, Calculations, OR Slider Labels"] = choices
                content[nrow][u"Field Type"] = "radio"
            else:
                content[nrow - 1][u"Choices, Calculations, OR Slider Labels"] = choices
                content[nrow - 1][u"Field Type"] = "radio"
        choices = []

# TODO : see if you can programmatically set the last question (if always the case) to current date with action tag @TODAY

# write csv
out = open('instrument.csv', 'wb')
wr = csv.writer(out, quoting=csv.QUOTE_ALL)
wr.writerow(head)

for r in content:
    row = []
    for k in head:
        row.append(r[k])
    wr.writerow(row)
out.close()

# compress csv
zf = zipfile.ZipFile('instrument.zip', mode='w')
zf.write('instrument.csv')
zf.close()

# remove csv
os.remove('instrument.csv')
