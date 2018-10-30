import openpyxl
import openpyxl.utils
import re
from pathlib import Path
from topline.db import DB
import logging

# imports for openpyxl merge patch
from openpyxl.worksheet import Worksheet
from openpyxl.reader.worksheet import WorkSheetParser
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.merge import MergeCells

logger = logging.getLogger(__name__)

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
user_row = 33
roi_row = 71
income_row = 75
expense_row = 77


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


def patch_worksheet():
    """This monkeypatches Worksheet.merge_cells to remove cell deletion bug
    https://bitbucket.org/openpyxl/openpyxl/issues/365/styling-merged-cells-isnt-working
    """
    # apply patch 1
    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        cr = CellRange(range_string=range_string, min_col=start_column, min_row=start_row,
                       max_col=end_column, max_row=end_row)
        self.merged_cells.add(cr.coord)
        # self._clean_merge_range(cr)

    Worksheet.merge_cells = merge_cells

    # apply patch 2
    def parse_merge(self, element):
        merged = MergeCells.from_tree(element)
        self.ws.merged_cells.ranges = merged.mergeCell
        # for cr in merged.mergeCell:
        #     self.ws._clean_merge_range(cr)

    WorkSheetParser.parse_merge = parse_merge


patch_worksheet()


class Excel:
    def __init__(self, filename):
        self.filename = filename
        self.sheet_names = None
        self.sheet_list = []
        self.header_list = []
        self.summary_sheet = None
        self.workbook = None
        self.users = None
        try:
            self.workbook = openpyxl.load_workbook(filename)
            self.get_sheets()
            self.get_users()
        except FileNotFoundError:
            logger.error("File not found: %s", Path(filename).absolute())

    def add_transaction(self, account_name, account_number, transaction):
        # All income, including contributions, and expenses should be from the cheque account
        # All other accounts are savings and fixed deposit accounts so should only have profit share returns
        logger.info('Processing transaction: %s', transaction)
        success = False
        if 'cheque' in account_name.lower():
            success = self.process_contribution(transaction)
            if not success:
                success = self.process_income_expense(transaction)
        elif 'savings' in account_name.lower() or 'deposit' in account_name.lower():
            success = self.process_roi(account_number, transaction)

        if not success:
            logger.warning('Unable to process transaction')
            return False
        return True

    def process_contribution(self, transaction):
        # Find transaction reference. Check reference column first then description column
        ref = format_string(transaction[2]) or format_string(transaction[1])
        if len(ref) < 3:
            logger.debug("Reference Error - Too few parameters: %s", ref)
            return False

        date = format_string(transaction[0])
        amount = float(transaction[4].replace(",", ""))

        # determine user
        # usernames retrieved from database to account for alternatives. If match is found, index is determined from
        # internal user list to retain spreadsheet order
        db_users = DB().get_user_ids()
        user_ids = [ui for ui, u in enumerate(db_users) for ri, r in enumerate(ref) if r in u]
        if len(user_ids) is not 1:
            logger.debug("Error finding user in reference: %s", ref)
            return False
        user_id = self.users.index(db_users[user_ids[0]][0])

        # find month in transaction reference
        month_ids = [mi for mi, m in enumerate(months) for ri, r in enumerate(ref) if r in m]
        if not month_ids:
            month_ids = [mi for mi, m in enumerate(months) if date[1] in m]
            if len(month_ids) is not 1:
                logger.debug("Error finding month in reference: %s", ref)
                return False
        month_id = month_ids[0]

        years = [r for i, r in enumerate(ref) if is_number(r)]
        year_id = (years[0] if len(years) == 1 else date[2]) % 2000

        # determine which sheet to use for current transaction
        sheet = self.get_target_sheet(month_id, year_id)
        if not sheet:
            logger.debug("Error finding correct sheet for transaction. ref: %s", ref)
            return False
        header = self.get_column_headers(sheet)

        # find correct column for contribution month
        try:
            column = header.index([month_id, year_id]) + 1
        except ValueError:
            logger.debug("Error finding correct column for month %s, year %s", month_id, year_id)
            return False

        # find correct row for user
        row = user_row + user_id

        return self.write_to_sheet(sheet, row, column, amount)

    def process_income_expense(self, transaction):
        # Find transaction reference. Check reference column first then description column
        ref = transaction[2] if format_string(transaction[2]) else transaction[1]
        date = format_string(transaction[0])
        amount = float(transaction[4].replace(",", ""))

        if ref == '#MONTHLY ACCOUNT FEE':
            row = expense_row
        else:
            logger.info('Unknown income or expense!')
            return False

        month_ids = [mi for mi, m in enumerate(months) if date[1] in m]
        if len(month_ids) is not 1:
            logger.debug("Error finding month in date: %s", date)
            return False
        month_id = month_ids[0]

        year_id = date[2] % 2000

        # determine which sheet to use for current transaction
        sheet = self.get_target_sheet(month_id, year_id)
        if not sheet:
            logger.debug("Error finding correct sheet for transaction. ref: %s", ref)
            return False
        header = self.get_column_headers(sheet)

        # find correct column for contribution month
        try:
            column = header.index([month_id, year_id]) + 1
        except ValueError:
            logger.debug("Error finding correct column for month %s, year %s", month_id, year_id)
            return False
        return self.write_to_sheet(sheet, row, column, abs(amount))

    def process_roi(self, account_number, transaction):
        # Find transaction reference.
        ref = transaction[1]
        date = format_string(transaction[0])
        amount = float(transaction[4].replace(",", ""))

        if 'profit share' in ref.lower():
            month_ids = [mi for mi, m in enumerate(months) if date[1] in m]
            if len(month_ids) is not 1:
                logger.debug("Error finding month in date: %s", date)
                return False
            month_id = month_ids[0]

            year_id = date[2] % 2000

            # determine which sheet to use for current transaction
            sheet = self.get_target_sheet(month_id, year_id)
            if not sheet:
                logger.debug("Error finding correct sheet for transaction. ref: %s", ref)
                return False
            header = self.get_column_headers(sheet)

            # find correct column for contribution month
            try:
                column = header.index([month_id, year_id]) + 1
            except ValueError:
                logger.debug("Error finding correct column for month %s, year %s", month_id, year_id)
                return False

            roi_range = sheet['A'+str(roi_row):'A'+str(income_row-1)]
            row = 0
            for r in roi_range:
                if account_number in r[0].value:
                    row = r[0].row
                    break
            if not row:
                logger.debug('Unable to find row for ROI transaction from account: %s', account_number)
                return False

            return self.write_to_sheet(sheet, row, column, amount)
        else:
            logger.debug("Not an ROI transaction!")
            return False

    def get_sheets(self):
        self.sheet_names = self.workbook.sheetnames
        for sheet in self.sheet_names:
            if 'summary' in sheet.lower():
                self.summary_sheet = self.workbook[sheet]
            else:
                # format sheet name
                s = format_string(sheet)
                name = [i if type(i) is int else (next((j for j, x in enumerate(months) if i in x), None)) for i in s]
                self.sheet_list.append(name)
        self.sheet_names.remove(self.summary_sheet.title)

    def get_users(self):
        row = 5
        self.users = []
        while True:
            user = self.summary_sheet.cell(row=row, column=2).value
            if user is None:
                break
            else:
                self.users.append(user)
                row += 1

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
            logger.warning("[%s] %s%s is locked!", sheet.title, openpyxl.utils.get_column_letter(col), row)
            return False

        cell_val = target_cell.value
        if cell_val == 0 or cell_val == '-' or cell_val is None or overwrite:
            sheet.cell(row=row, column=col, value=value)
        elif cell_val == value:
            logger.warning("Transaction already processed!")
            return False
        else:
            logger.warning("[%s] %s%s contains data!", sheet.title, openpyxl.utils.get_column_letter(col), row)
            return False
        return True

    def close_workbook(self, overwrite=True, filename=None):
        new_filename = self.filename
        if filename:
            new_filename = filename

        if not overwrite and filename is None:
            logger.warning("No filename provided. Discarding changes")
        else:
            logger.info("Saving workbook to file: %s", new_filename)
            self.workbook.save(new_filename)
        self.workbook.close()
