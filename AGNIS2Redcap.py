import xlwt, xlrd, xlutils
import argparse
import re
import csv

parser = argparse.ArgumentParser()
parser.add_argument("metadata_xls", help="filename")

args = parser.parse_args()

# agnis metadata
agnis_wb = xlrd.open_workbook(args.metadata_xls)
agnis_ws = agnis_wb.sheet_by_name('Sheet0')
agnis_head = [c.value for c in agnis_ws.row(10)]

# instrument.csv
out = open('instrument.csv', 'wb')
wr = csv.writer(out, quoting=csv.QUOTE_ALL)

head = [u"Variable / Field Name", u"Form Name", u"Section Header", u"Field Type", u"Field Label", u"Choices, Calculations, OR Slider Labels",
        u"Field Note", u"Text Validation Type OR Show Slider Number", u"Text Validation Min", u"Text Validation Max", u"Identifier?",
        u"Branching Logic (Show field only if...)", u"Required Field?", u"Custom Alignment", u"Question Number (surveys only)",
        u"Matrix Group Name", u"Matrix Ranking?", u"Field Annotation"]
wr.writerow(head)

content = []

form_name = agnis_ws.cell(0, 1).value
form_id = form_name.split(" - ")[0].replace(" ", "").replace("Revision", "r")
form_name = re.sub('[^A-Za-z0-9]+', '_', form_name)

section_header = u""
nrow = -1
choices = []
for agnis_nrow in range(11, agnis_ws.nrows):

    nmod = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Module Display Order')).value
    nqst = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Display Order')).value
    cdepid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'CDE Public ID')).value
    qst_long_name = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Question Long Name')).value
    data_type = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Data Type')).value
    valid_val = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Valid Value')).value
    val_meaning_text = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Text')).value
    val_meaning_pubid = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Value Meaning Public ID')).value
    display_format = agnis_ws.cell(agnis_nrow, agnis_head.index(u'Display Format')).value

    if cdepid:
        nrow += 1
        content.append({k: u'' for k in head})
        content[nrow][u"Form Name"] = form_name

        redcapid = "f%sm%dq%dcde%d" % (form_id, int(nmod), int(nqst), int(cdepid))
        content[nrow][u"Variable / Field Name"] = redcapid
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
for r in content:
    row = []
    for k in head:
        row.append(r[k])
    wr.writerow(row)
out.close()

