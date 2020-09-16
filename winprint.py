#! python3
#
# Docs used to create
# https://smallbusiness.chron.com/sending-things-printer-python-58655.html
# http://timgolden.me.uk/pywin32-docs/win32print.html

import os
import sys
import win32print


class GeneratorPrinter(object):

    def __init__(self, filename):
        if sys.platform == 'win32':
            self.printer = win32print.OpenPrinter('Microsoft Print to PDF')
            self.print_docx(filename)
        elif sys.platform == 'win64':
            self.printer = win32print.OpenPrinter('Microsoft Print to PDF')
            self.print_docx(filename)
        else:
            pass

    def print_docx(self, filename):
        job = win32print.StartDocPrinter(self.printer, 1, (filename, None, None))
        print(job)


if __name__ == '__main__':
    print(sys.platform)
    GeneratorPrinter('generated.docx')
