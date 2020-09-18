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
img = 'HLP-Logo-Aus.png'
set_date = '3/9/20'
sn = 'HLP–PTPB005457'
model = 'Pocket Temp Pro'
name = 'L.Adams'
months = '12'

# Text to be displayed in a conformance certificate
heading_text = '\rMODEL: ' + model + '\rCONFORMANCE CERTIFICATE'
paragraph_text = '\rSerial Number:  ' + sn + '\r\rThis unit has had an operational and calibration' + \
                 '\rcheck on  ' + set_date + '\r& meets the specifications as set out in the Manufacturers' + \
                 'specification sheet included with the unit & meets the Australian Food Standards requirements' + \
                 '\rThis certificate is valid for' + months + 'months from the above date.'

# Text to be displayed for the single conformance certificate
compliance_heading = 'Certificate of Compliance\rModel:  ' + model
# Add valid_text in-between compliance_heading and compliance_text
compliance_text = '\rThe above unit meets the stated specifications as set out in the Manufacturers specification ' + \
                  'sheet included with the unit & meets the Australian Food Standards requirements.\r\r\r\r\r\r' + \
                  'Signed\t' + name + '\t\tDate Issued: ' + set_date

# Text to be displayed in a calibration certificate
heading_text2 = '\rCalibration Certificate'
valid_text = '(This Certificate valid for ' + months + ' Months)'
paragraph_text2 = '\rChecked against NATA calibrated unit – AZ8801\rS/N  ' + sn + 'Calibration Report: 41598-4\r\r' +\
                  'The above unit meets the stated specifications as set out in the Manufacturers specification ' + \
                  'sheet included with the unit & meets the Australian Food Standards requirements.\r\r\r\r' + \
                  'Signed\t' + name + '\t\tDate Issued: ' + set_date

# Initialising the path for where the document will be saved
dir_path = path.dirname(__file__)


class Template(object):

    def __init__(self, date, serial_number, save_location):
        self.date = date
        self.serial_number = serial_number
        self.save_location = save_location
        self.create_docx()

    def create_docx(self):
        doc = docx.Document()

        doc.save(self.save_location)

    @staticmethod
    def margin_size(doc, size):
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(size)
            section.bottom_margin = Cm(size)
            section.left_margin = Cm(size)
            section.right_margin = Cm(size)


class CalibrationCertificate(Template):

    def __init__(self, date, serial_number, save_location):
        super().__init__(date, serial_number, save_location)

    def create_docx(self):
        doc = docx.Document()
        super().margin_size(doc, 3)
        self.docx_contents(doc)
        doc.save(self.save_location)

    @staticmethod
    def docx_contents(doc):
        # TODO Rescale img width and height to fit page
        doc.add_picture(img, width=Cm(17.3), height=Cm(3.95))
        heading = doc.add_paragraph(heading_text2)
        valid = doc.add_paragraph(valid_text)
        text = doc.add_paragraph(paragraph_text2)

        # Font Settings for each paragraph
        heading.font.name = 'Times New Roman'
        valid.font.name = 'Times New Roman'
        text.font.name = 'Times New Roman'
        heading.font.size = Pt(30)
        heading.font.size = Pt(10)
        heading.font.size = Pt(18)


class SingleConformance(Template):

    def __init__(self, date, serial_number, save_location):
        super().__init__(date, serial_number, save_location)

    def create_docx(self):
        doc = docx.Document()
        super().margin_size(doc, 3)
        self.docx_contents(doc)
        doc.save(self.save_location)

    @staticmethod
    def docx_contents(doc):
        doc.add_picture(img, width=Cm(17.3), height=Cm(3.95))
        heading = doc.add_paragraph().add_run().add_text(heading_text2)
        valid = doc.add_paragraph().add_run().add_text(valid_text)
        text = doc.add_paragraph().add_run().add_text(paragraph_text2)

        heading_run = heading.run
        valid_run = valid.run
        text_run = text.run

        # Font Settings for each paragraph
        heading_run.font.name = 'Times New Roman'
        valid_run.font.name = 'Times New Roman'
        text_run.font.name = 'Times New Roman'
        heading_run.font.size = Pt(30)
        valid_run.font.size = Pt(10)
        text_run.font.size = Pt(18)


class MultipleConformance(Template):

    def __init__(self, date, serial_number, save_location):
        super().__init__(date, serial_number, save_location)

    def create_docx(self):
        doc = docx.Document()
        super().margin_size(doc, 1)
        self.create_table(doc, 4, rows=1, cols=1)

        doc.save(self.save_location)

    # TODO Create a table and format it the same as the conformance certificates
    def create_table(self, doc, number_of_tables, rows=0, cols=0):
        # Creating the first table
        for i in range(0, number_of_tables):
            table = doc.add_table(rows, cols)
            table.cell(0, 0).width = Cm(9)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            table.style = 'TableGrid'
            table.autofit = False
            table.width = Cm(9)
            self.table_contents(table)
            doc.add_paragraph()

    @staticmethod
    def table_contents(table):
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

        # Adding image
        logo.add_picture(img, width=Cm(7.2898), height=Cm(1.9304))

        # Paragraph alignment and Font settings
        r.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        x.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        r.bold = True
        r.underline = True
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        x.font.name = 'Arial'
        x.font.size = Pt(11)
        x.bold = False
        x.underline = False

        # Adding the text to each run
        r.add_text(heading_text).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        x.add_text(paragraph_text).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


if __name__ == '__main__':
    # Generated Calibration Certificate must have 'device_model and serial number' for the name.
    MultipleConformance(set_date, sn, (path.join(dir_path, 'generated.docx')))
    CalibrationCertificate(set_date, sn, (path.join(dir_path, 'generated-calibration.docx')))
    SingleConformance(set_date, sn, (path.join(dir_path, 'generated-conformance-single.docx')))
