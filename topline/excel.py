import openpyxl
import openpyxl.utils
import re
from pathlib import Path
import logging
from topline import MONTHS

# imports for openpyxl merge patch
from openpyxl.worksheet import Worksheet
from openpyxl.reader.worksheet import WorkSheetParser
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.merge import MergeCells

logger = logging.getLogger(__name__)

header_row = 32
user_row = 33
roi_row = 71
income_row = 75
expense_row = 77
account_row = 25
last_account_row = 35


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
    user_ids = None

    def __init__(self, filename):
        self.filename = filename
        self.sheet_names = None
        self.sheet_list = []
        self.header_list = []
        self.summary_sheet = None
        self.workbook = None
        try:
            self.workbook = openpyxl.load_workbook(filename)
            self.get_sheets()
        except FileNotFoundError:
            logger.error("File not found: %s", Path(filename).absolute())
        
        if not Excel.user_ids:
            Excel.user_ids = self.get_user_ids()

        # Force account balances to 0
        logger.info("Clearing all account balances")
        account_range = self.summary_sheet['A' + str(account_row):'C' + str(last_account_row)]
        for r in account_range:
            r[2].value = 0

    def add_transaction(self, transaction):
        # determine which sheet to use for current transaction
        sheet = self.get_target_sheet(transaction.month_id, transaction.year % 2000)
        if not sheet:
            logger.warning("Error finding correct sheet for transaction. Month = %s, Year = %s",
                           transaction.month_id, transaction.year % 2000)
            return False
        header = self.get_column_headers(sheet)

        # find correct column for contribution month
        try:
            column = header.index([transaction.month_id, transaction.year % 2000]) + 1
        except ValueError:
            logger.warning("Error finding correct column: Month %s, Year %s",
                           transaction.month_id, transaction.year % 2000)
            return False

        if transaction.type == 'contribution':
            user_offset = [ui for ui, u in enumerate(Excel.user_ids) if u[1] == transaction.username]
            row = user_row + user_offset[0]
        elif transaction.type == 'expense':
            row = expense_row
        elif transaction.type == 'roi':
            roi_range = sheet['A' + str(roi_row):'A' + str(income_row - 1)]
            row = 0
            for r in roi_range:
                if str(transaction.account) in r[0].value:
                    row = r[0].row
                    break
            if not row:
                logger.warning('Unable to find row for ROI transaction')
                return False
        else:
            return False
        return self.write_to_sheet(sheet, row, column, abs(transaction.amount), add=True)

    def update_account_balances(self, account, balance):
        sheet = self.summary_sheet
        account_range = sheet['A' + str(account_row):'C' + str(last_account_row)]
        cell = [r[0] for r in account_range if r[0].value == account]
        if len(cell) == 1:
            old_balance = sheet.cell(row=cell[0].row, column=3).value
            logger.info("Updating account %s balance: R %.2f -> R %.2f", account, old_balance, balance)
            sheet.cell(row=cell[0].row, column=3).value = balance
        else:
            logger.info("Adding account %s balance: R %.2f", account, balance)
            row = [r for r in account_range if r[0].value is None][0]
            row[0].value = account
            row[2].value = balance

    def get_sheets(self):
        self.sheet_names = self.workbook.sheetnames
        for sheet in self.sheet_names:
            if 'summary' in sheet.lower():
                self.summary_sheet = self.workbook[sheet]
            else:
                # format sheet name
                s = format_string(sheet)
                name = [i if type(i) is int else (next((j for j, x in MONTHS.items() if i in x), None)) for i in s]
                self.sheet_list.append(name)
        self.sheet_names.remove(self.summary_sheet.title)

    def get_user_ids(self):
        row = 5
        user_ids = []
        while row < self.summary_sheet.max_row:
            user_id = self.summary_sheet.cell(row=row, column=2).value
            if user_id is None:
                break
            else:
                user_ids.append([row - 5, user_id])
                row += 1
        return user_ids

    def get_column_headers(self, sheet):
        # Format transaction header
        header = sheet[header_row]
        header = [[h.value.month, h.value.year % 2000] if h.is_date else str(h.value) for h in header]
        return header

    def get_target_sheet(self, month, year):
        # determine which sheet to use for current transaction
        sheet_id = [i for i, s in enumerate(self.sheet_list) if
                    (month >= s[0] and year == s[1]) or (month <= s[2] and year == s[3])]
        if not sheet_id:
            return None
        return self.workbook[self.sheet_names[sheet_id[0]]]

    def write_to_sheet(self, sheet, row, col, value, overwrite=False, add=False):
        target_cell = sheet.cell(row=row, column=col)
        if target_cell.protection.locked:
            logger.warning("[%s] %s%s is locked!", sheet.title, openpyxl.utils.get_column_letter(col), row)
            return False

        cell_val = target_cell.value
        if cell_val == 0 or cell_val == '-' or cell_val is None or overwrite:
            sheet.cell(row=row, column=col, value=value)
            logger.info("Transaction written to sheet [%s] %s%s.",
                        sheet.title, openpyxl.utils.get_column_letter(col), row)
        elif add:
            logger.info("[%s] %s%s contains data! Adding to existing value!",
                        sheet.title, openpyxl.utils.get_column_letter(col), row)
            sheet.cell(row=row, column=col, value=cell_val + value)
        else:
            logger.warning("[%s] %s%s contains data! Transaction not written to sheet!",
                           sheet.title, openpyxl.utils.get_column_letter(col), row)
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
