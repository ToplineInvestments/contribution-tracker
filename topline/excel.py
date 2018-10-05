import openpyxl
import re

months = [['JANUARY',  'JAN'],
          ['FEBRUARY', 'FEB'],
          ['MARCH',    'MAR'],
          ['APRIL',    'APR'],
          ['MAY'            ],
          ['JUNE',     'JUN'],
          ['JULY',     'JUL'],
          ['AUGUST',   'AUG'],
          ['SEPTEMBER', 'SEPT', 'SEP'],
          ['OCTOBER',  'OCT'],
          ['NOVEMBER', 'NOV'],
          ['DECEMBER', 'DEC']]


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def format_string(s):
    # return list of string or number groups from input string
    # eg: "AA - AUG17" -> ['AA', 'AUG', '17']
    #     "NOV '17 - OCT '18" -> ['NOV', '17', 'OCT', '18']
    s = s.upper()
    string_list = re.findall(r"\d+|[a-z]+", s, re.I)
    string_list = [int(i) if is_number(i) else i.upper() for i in string_list]
    return string_list


class Excel:

    def __init__(self, filename):
        self.filename = filename
        self.workbook = openpyxl.load_workbook(filename)
        self.sheet_names = None
        self.sheet_list = None
        self.summary_sheet = None

    def get_sheets(self):
        self.sheet_names = self.workbook.get_sheet_names()
        self.sheet_list = []
        for sheet in self.sheet_names:
            if 'summary' in sheet.lower():
                self.summary_sheet = self.workbook[sheet]
            else:
                # format sheet name
                s = format_string(sheet)
                s = [i if type(i) is int else (next((j for j, x in enumerate(months) if i in x), None)) for i in s]
                self.sheet_list.append(s)
        self.sheet_names.remove(self.summary_sheet.title)
        
    def close_workbook(self, overwrite=True, filename=None):
        new_filename = self.filename
        if filename:
            new_filename = filename

        if not overwrite and filename is None:
            print("No filename provided. Discarding changes")
        else:
            print("Saving workbook to file: {}".format(new_filename))
            self.workbook.save(new_filename)
        self.workbook.close()
