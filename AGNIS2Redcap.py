import xlwt, xlrd, xlutils
import argparse
import re
import csv
import zipfile
import os

parser = argparse.ArgumentParser()
parser.add_argument("metadata_xls", help="filename")

args = parser.parse_args()

# agnis metadata
agnis_wb = xlrd.open_workbook(args.metadata_xls)
agnis_ws = agnis_wb.sheet_by_name('Sheet0')
agnis_head = [c.value for c in agnis_ws.row(10)]

head = [u"Variable / Field Name", u"Form Name", u"Section Header", u"Field Type", u"Field Label", u"Choices, Calculations, OR Slider Labels",
        u"Field Note", u"Text Validation Type OR Show Slider Number", u"Text Validation Min", u"Text Validation Max", u"Identifier?",
        u"Branching Logic (Show field only if...)", u"Required Field?", u"Custom Alignment", u"Question Number (surveys only)",
        u"Matrix Group Name", u"Matrix Ranking?", u"Field Annotation"]

content = []

form_name = agnis_ws.cell(0, 1).value
form_id = form_name.split(" - ")[0].replace(" ", "").replace("Revision", "r")
form_name = re.sub('[^A-Za-z0-9]+', '_', form_name)

form_pid = agnis_ws.cell(6, 1).value
form_ver = agnis_ws.cell(7, 1).value

section_header = u""
nrow = -1
choices = []
for agnis_nrow in range(11, agnis_ws.nrows):

    nmod = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Display Order')).value
    nqst = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Display Order')).value
    if agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Public ID')).value:
        mod_pid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Public ID')).value
    if agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Version')).value:
        mod_ver = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Version')).value
    q_pid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'CDE Public ID')).value
    q_ver = agnis_ws.cell(agnis_nrow, agnis_head.index(u'CDE Version')).value
    qst_long_name = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Long Name')).value
    data_type = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Data Type')).value
    valid_val = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Valid Value')).value
    val_meaning_text = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Text')).value
    val_meaning_pubid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Public ID')).value
    display_format = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Display Format')).value

    if q_pid:
        nrow += 1
        content.append({k: u'' for k in head})
        content[nrow][u"Form Name"] = form_name

        redcapid = "f{0}m{1}q{2}cde{3}rev{4}".format(form_id, int(nmod), int(nqst), int(q_pid), int(q_ver))
        redcapmodxid = "{0}_{1};{2}_{3};{4}_{5};".format(int(form_pid), form_ver, int(mod_pid), mod_ver, int(q_pid), q_ver)

        content[nrow][u"Variable / Field Name"] = redcapid
        content[nrow][u"Field Annotation"] = redcapmodxid
        content[nrow][u"Section Header"] = section_header
        content[nrow][u"Field Label"] = qst_long_name.rstrip(":")

        #if data_type in ["NUMBER", "CHARACTER", "DATE"]:
        content[nrow][u"Field Type"] = "text" # default to text

        if display_format == u"YYYY-MM-DD":
            content[nrow][u"Text Validation Type OR Show Slider Number"] = u"date_ymd"

        section_header = u""

    else:
        section_header = agnis_ws.cell(agnis_nrow, agnis_head.index(u"Module Long Name")).value

    if val_meaning_pubid and val_meaning_text:
        #choices.append(str(int(val_meaning_pubid)) + "," + val_meaning_text)
        choices.append(str(int(val_meaning_pubid)) + "," + valid_val)
    elif nrow != -1:
        choices = "|".join(choices)
        if choices:
            if section_header:
                content[nrow][u"Choices, Calculations, OR Slider Labels"] = choices
                content[nrow][u"Field Type"] = "radio"
            else:
                content[nrow - 1][u"Choices, Calculations, OR Slider Labels"] = choices
                content[nrow - 1][u"Field Type"] = "radio"
        choices = []

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
