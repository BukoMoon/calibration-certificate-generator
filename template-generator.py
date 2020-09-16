#! python3
#
# Docs used to create
# https://www.reddit.com/r/learnpython/comments/cjxo0b/docx_adding_text_to_textbox/
# https://docxtpl.readthedocs.io/en/latest/#indices-and-tables
# https://automatetheboringstuff.com/chapter13/
# https://python-docx.readthedocs.io/en/latest/user/quickstart.html

import docx
from docx.shared import Cm, Pt
from os import path
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


# Example variables
pic = 'HLP-Logo-Aus.png'
current_date = '3/9/20'
sn = 'HLP–PTPB005457'
model = 'Pocket Temp Pro'
name = 'L.Adams'

model_text = 'MODEL:  ' + model
heading_text = 'CONFORMANCE CERTIFICATE'
serial_text = 'Serial Number:  ' + sn
line_text = 'This unit has had an operational and calibration'
check_text = 'check on  ' + current_date
next_text = '& meets the specifications as set out in the Manufacturers specification sheet ' + \
            'included with the unit & meets the Australian Food Standards requirements'
cert_text = 'This certificate is valid for 24 months from the above date.'


dir_path = path.dirname(__file__)


class MultipleConformanceCertTemplate:

    def __init__(self, date, serial_number, save_location):
        self.date = date
        self.serial_number = serial_number
        self.save_location = save_location
        self.create_docx()

    def create_docx(self):
        docu = docx.Document()
        sections = docu.sections
        for section in sections:
            section.top_margin = Cm(1)
            section.bottom_margin = Cm(1)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)
        self.create_table(docu, rows=1, cols=1)

        docu.save(self.save_location)

    # TODO Create a table and format it the same as the conformance certificates
    def create_table(self, document, rows=0, cols=0):
        doc = document
        number_of_tables = 4
        # Creating the first table
        for i in range(0, number_of_tables):
            table1 = doc.add_table(rows, cols)
            table1.cell(0, 0).width = Cm(9)
            table1.alignment = WD_TABLE_ALIGNMENT.LEFT
            table1.style = 'TableGrid'
            table1.autofit = False
            table1.width = Cm(9)
            self.table_contents(table1)
            doc.add_paragraph()

    def table_contents(self, table):
        # Adding a paragraph to the table cell
        p = table.rows[0].cells[0].add_paragraph()
        p.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        p.height = Cm(11)
        table.rows[0].height = Cm(9)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Adding runs to the paragraph
        logo = p.add_run()
        r = p.add_run()
        x = p.add_run()

        # Paragraph alignment and Font settings
        r.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        x.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        logo.add_picture(pic, width=Cm(7.2898), height=Cm(1.9304))
        r.bold = True
        r.underline = True
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        x.font.name = 'Arial'
        x.bold = False
        x.underline = False

        # Adding the text to each run
        r.add_text('\r' + model_text + '\r' + heading_text).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        x.add_text('\r' + serial_text + '\r' + '\r' + line_text + '\r' + check_text + '\r' + next_text + '\r' +
                   cert_text + '\r' + 'Signed:' + '\t' + name + '\r').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


if __name__ == '__main__':
    MultipleConformanceCertTemplate(current_date, sn, (path.join(dir_path, 'generated.docx')))
