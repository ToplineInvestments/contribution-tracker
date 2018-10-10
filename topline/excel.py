import openpyxl
import openpyxl.utils
import re
from pathlib import Path
from topline.db import DB

months = [['JANUARY', 'JAN'],
          ['FEBRUARY', 'FEB'],
          ['MARCH', 'MAR'],
          ['APRIL', 'APR'],
          ['MAY'],
          ['JUNE', 'JUN'],
          ['JULY', 'JUL'],
          ['AUGUST', 'AUG'],
          ['SEPTEMBER', 'SEPT', 'SEP'],
          ['OCTOBER', 'OCT'],
          ['NOVEMBER', 'NOV'],
          ['DECEMBER', 'DEC']]
header_row = 32


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
        self.sheet_names = None
        self.sheet_list = []
        self.header_list = []
        self.summary_sheet = None
        self.workbook = None
        try:
            self.workbook = openpyxl.load_workbook(filename)
        except FileNotFoundError:
            print("File not found: {}".format(Path(filename).absolute()))

    def get_sheets(self):
        self.sheet_names = self.workbook.get_sheet_names()
        for sheet in self.sheet_names:
            if 'summary' in sheet.lower():
                self.summary_sheet = self.workbook[sheet]
            else:
                # format sheet name
                s = format_string(sheet)
                s = [i if type(i) is int else (next((j for j, x in enumerate(months) if i in x), None)) for i in s]
                self.sheet_list.append(s)

                # Format transaction header
                header = self.workbook[sheet][header_row]
                header = [[h.value.month - 1, h.value.year % 2000] if h.is_date else str(h.value) for h in header]
                self.header_list.append(header)
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

    def update_excel(self, transactions):
        print("Processing {} transactions".format(len(transactions)))
        # loop through each new transaction
        transaction_count = 0
        user_list = DB().get_user_ids()
        for trans in transactions:
            # Find transaction reference. Check reference column first then description column
            ref = format_string(trans[2])
            if not ref:
                ref = format_string(trans[1])
            date = format_string(trans[0])

            if len(ref) < 3:
                print("Reference Error - Too few parameters: {}".format(ref))
                continue

            # determine user
            user_id = [(ui, ri) for ui, u in enumerate(user_list) for ri, r in enumerate(ref) if r in u]
            if len(user_id) is not 1:
                print("Error finding user in reference: {}".format(ref))
                continue
            user_id = user_id[0]

            # find month in transaction reference
            month_id = [(mi, ri) for mi, m in enumerate(months) for ri, r in enumerate(ref) if r in m]
            if len(month_id) is not 1:
                month_id = [(mi, 1) for mi, m in enumerate(months) if date[1] in m]
                if len(month_id) is not 1:
                    print("Error finding month in reference: {}".format(ref))
                    continue
            month_id = month_id[0]

            year = [r for i, r in enumerate(ref) if i not in [user_id[1], month_id[1]] and is_number(r)]
            if len(year) == 1:
                year_id = year[0] % 2000
                # print("Error finding year in reference: {}".format(ref))
                # continue
            else:
                year_id = date[2] % 2000

            # determine which sheet to use for current transaction
            sheet_id = [i for i, s in enumerate(self.sheet_list) if
                        (month_id[0] >= s[0] and year_id == s[1]) or (month_id[0] <= s[2] and year_id == s[3])]
            if not sheet_id:
                print("Error finding correct sheet for month {}, year {}".format(month_id[0], year_id))
                continue
            sheet_id = sheet_id[0]

            # Use required sheet
            sheet = self.workbook[self.sheet_names[sheet_id]]
            header = self.header_list[sheet_id]

            # find correct column for contribution month
            try:
                column = header.index([month_id[0], year_id]) + 1
            except ValueError:
                print("Error finding correct column for month {}, year {}".format(month_id[0], year_id))
                continue

            # find correct row for user
            row = header_row + 1 + user_id[0]

            cell_val = sheet.cell(row=row, column=column).value
            if cell_val == 0 or cell_val == '-' or cell_val is None:
                sheet.cell(row=row, column=column, value=float(trans[4].replace(",", "")))
            else:
                print("Cell {}{} contains data! Transaction: {}".format(openpyxl.utils.get_column_letter(column), row,
                                                                        trans))
                continue
            transaction_count += 1
        print("Processed {}/{} transactions".format(transaction_count, len(transactions)))
        print("Excel Update complete. Saving file {}".format(self.filename))
        self.workbook.save(self.filename)
        self.workbook.close()
        return True
