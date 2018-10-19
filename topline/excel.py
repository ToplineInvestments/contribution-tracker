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
        self.users = DB().get_user_ids()
        try:
            self.workbook = openpyxl.load_workbook(filename)
        except FileNotFoundError:
            print("File not found: {}".format(Path(filename).absolute()))

    def add_transaction(self, account_name, account_number, transaction):
        # All income, including contributions, and expenses should be from the cheque account
        # All other accounts are savings and fixed deposit accounts so should only have profit share returns
        success = False
        if 'cheque' in account_name.lower():
            success = self.process_contribution(transaction)
            if not success:
                success = self.process_income_expense(transaction)
        elif 'savings' in account_name.lower() or 'deposit' in account_name.lower():
            success = self.process_roi(account_number, transaction)

        if not success:
            print('Unable to process transaction {}'.format(transaction))
            return False
        return True

    def process_contribution(self, transaction):
        # Find transaction reference. Check reference column first then description column
        ref = format_string(transaction[2]) or format_string(transaction[1])
        if len(ref) < 3:
            print("Reference Error - Too few parameters: {}".format(ref))
            return False

        date = format_string(transaction[0])
        amount = float(transaction[4].replace(",", ""))

        # determine user
        user_ids = [ui for ui, u in enumerate(self.users) for ri, r in enumerate(ref) if r in u]
        if len(user_ids) is not 1:
            print("Error finding user in reference: {}".format(ref))
            return False
        user_id = user_ids[0]

        # find month in transaction reference
        month_ids = [mi for mi, m in enumerate(months) for ri, r in enumerate(ref) if r in m]
        if not month_ids:
            month_ids = [mi for mi, m in enumerate(months) if date[1] in m]
            if len(month_ids) is not 1:
                print("Error finding month in reference: {}".format(ref))
                return False
        month_id = month_ids[0]

        years = [r for i, r in enumerate(ref) if is_number(r)]
        year_id = (years[0] if len(years) == 1 else date[2]) % 2000

        # determine which sheet to use for current transaction
        sheet = self.get_target_sheet(month_id, year_id)
        if not sheet:
            print("Error finding correct sheet for transaction. ref: {}".format(ref))
            return False
        header = self.get_column_headers(sheet)

        # find correct column for contribution month
        try:
            column = header.index([month_id, year_id]) + 1
        except ValueError:
            print("Error finding correct column for month {}, year {}".format(month_id, year_id))
            return False

        # find correct row for user
        user_row = 33
        row = user_row + user_id

        return self.write_to_sheet(sheet, row, column, amount)

    def process_income_expense(self, transaction):
        return False

    def process_roi(self, account_number, transaction):
        return False

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
        self.sheet_names.remove(self.summary_sheet.title)

    def get_column_headers(self, sheet):
        # Format transaction header
        header = sheet[header_row]
        header = [[h.value.month - 1, h.value.year % 2000] if h.is_date else str(h.value) for h in header]
        return header

    def get_target_sheet(self, month, year):
        # determine which sheet to use for current transaction
        sheet_id = [i for i, s in enumerate(self.sheet_list) if
                    (month >= s[0] and year == s[1]) or (month <= s[2] and year == s[3])]
        if not sheet_id:
            return None
        return self.workbook[self.sheet_names[sheet_id[0]]]

    def write_to_sheet(self, sheet, row, col, value, overwrite=False):
        target_cell = sheet.cell(row=row, column=col)
        if target_cell.protection.locked:
            print("Cell {}{} is locked!".format(openpyxl.utils.get_column_letter(col), row))
            return False

        cell_val = target_cell.value
        if cell_val == 0 or cell_val == '-' or cell_val is None or overwrite:
            sheet.cell(row=row, column=col, value=value)
        elif cell_val == value:
            print("Transaction already processed!")
            return False
        else:
            print("Cell {}{} contains data!".format(openpyxl.utils.get_column_letter(col), row))
            return False
        return True

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
