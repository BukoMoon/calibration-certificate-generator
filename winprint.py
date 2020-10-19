#! python3
import sys
import win32api
import win32print


class GeneratorPrinter(object):

    def __init__(self, filename):
        self.filename = filename

    def print_docx(self):
        win32api.ShellExecute(0, "print", self.filename, '/d:"%s"' % win32print.GetDefaultPrinter(), ".", 0)


if __name__ == '__main__':
    print(sys.platform)
    GeneratorPrinter('generated.docx')
