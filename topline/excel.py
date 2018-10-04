import openpyxl


class Excel:

    def __init__(self, filename):
        self.filename = filename
        self.workbook = openpyxl.load_workbook(filename)
        self.sheet_names = None
        self.sheet_list = None
        self.summary_sheet = None
