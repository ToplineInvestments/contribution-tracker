import re
import logging
from datetime import datetime
from topline import MONTHS


logger = logging.getLogger(__name__)


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
    string_list = re.findall(r"\d+|[a-z.]+", s, re.I)
    string_list = [int(i) if is_number(i) else i.replace('.', '').upper() for i in string_list]
    return string_list


def date_check(month1, year1, month2, year2):
    diff = (year2 - year1) * 12 + month2 - month1
    offset = int((12 * round(diff / 12)) / 12)
    if offset > 0:
        logger.debug("Date updated: %s/%s -> %s/%s", month1, year1, month1, year1 + offset)
    return month1, year1 + offset


class Transaction:
    usernames = None
    accounts = None

    def __init__(self, transaction, account):
        self.account = account
        self.transaction = transaction
        self.transaction_id = None
        self.date = datetime.strptime(transaction[0], '%d %b %Y').date()
        self.description = transaction[1]
        self.reference = transaction[2]
        self.amount = float(transaction[4].replace(",", ""))
        self.username = None
        self.user_id = None
        self.month = None
        self.month_id = None
        self.year = None
        self.type = None

    def transaction_type(self):
        account_name = [acc[1] for acc in Transaction.accounts if acc[0] == self.account][0]
        if 'cheque' in account_name.lower():
            if self.username is not None:
                self.type = 'contribution'
            else:
                if 'MONTHLY ACCOUNT FEE' in self.description.upper():
                    self.type = 'expense'
                else:
                    self.type = 'unknown'
        elif 'savings' in account_name.lower() or 'deposit' in account_name.lower():
            if 'profit share' in self.description.lower():
                self.type = 'roi'
            else:
                self.type = 'unknown'

        if self.type == 'expense' or self.type == 'roi':
            self.username = "TIG"
            self.user_id = [u[0] for u in Transaction.usernames if u[1] == "TIG"][0]
        return self.type

    def process_transaction(self):
        ref = format_string(self.reference) or format_string(self.description)
        self.get_user(ref)
        self.transaction_type()

        if self.type == 'unknown':
            logger.info("Unknown transaction")
            return False

        if self.type == 'contribution':
            # find month in transaction reference
            self.month_id = [mi for mi, m in MONTHS.items() for ri, r in enumerate(ref) if r in m][0]
            years = [r for i, r in enumerate(ref) if is_number(r)]
            self.year = (years[0] if len(years) == 1 else self.date.year) % 2000
        else:
            self.year = self.date.year % 2000
        self.year += 2000
        if not self.month_id:
            self.month_id = self.date.month

        self.month_id, self.year = date_check(self.month_id, self.year, self.date.month, self.date.year)

        if self.type == 'contribution':
            self.month = MONTHS[self.month_id][0].capitalize()

        return True

    def get_user(self, ref):
        if len(ref) < 3:
            logger.debug("Reference Error - Too few parameters: %s", ref)
            return False

        # determine user
        # usernames retrieved from database to account for alternatives. If match is found, index is determined from
        # internal user list to retain spreadsheet order
        uid = [ui for ui, u in enumerate(Transaction.usernames) for ri, r in enumerate(ref) if r in u[1:]]
        if len(uid) is not 1:
            logger.debug("Error finding user in reference: %s", self.reference)
            return False
        self.user_id = Transaction.usernames[uid[0]][0]
        self.username = Transaction.usernames[uid[0]][1]
        return True
