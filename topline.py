import configparser
import logging
from logging.config import fileConfig
from datetime import datetime, date
from shutil import copy2
from pathlib import Path
import zipfile
from topline.scraper import FNB
from topline.excel import Excel
from topline.db import DB
from topline.transaction import Transaction
from topline.gmail import Gmail
from topline import MONTHS

now = datetime.now()

config = configparser.ConfigParser()
if not config.read('config.ini'):
    print('No config.ini file found')
    raise SystemExit(0)

fileConfig(config, disable_existing_loggers=False)
logger = logging.getLogger()

path_to_driver = None

try:
    username = config['SCRAPER']['USERNAME']
    password = config['SCRAPER']['PASSWORD']
    url = config['SCRAPER']['URL']
    driver = config['SCRAPER']['DRIVER']
    if driver == 'chrome':
        path_to_driver = config['SCRAPER']['DRIVER_PATH']
    db_file = config['DB']['FILENAME']
    excel_file = config['EXCEL']['FILENAME']
except KeyError:
    logger.error("Error getting config from config.ini file")
    raise SystemExit(0)

db = DB(db_file)
if not db.connection:
    logger.error('DB error')
    raise SystemExit(0)

fnb = FNB(driver=driver, headless=False, driver_path=path_to_driver)

if fnb.driver:
    fnb.driver.get(url)
    if fnb.login(username, password):
        fnb.get_accounts(get_transactions=True)
        logger.debug("transactions = %s", fnb.accounts)
        fnb.logout()
    else:
        fnb.driver.quit()

db_usernames = db.get_usernames()
if not db_usernames:
    logger.error("No users found in database.")
    raise SystemExit(0)
cur_accounts = list(fnb.accounts.keys())
db_accounts = db.get_accounts()
db_account_numbers = [account[0] for account in db_accounts]

for acc in list(set(cur_accounts).union(db_account_numbers)):
    if acc in cur_accounts and acc in db_account_numbers:
        db.update_account(acc, fnb.accounts[acc]['name'], float(fnb.accounts[acc]['balance']))
    elif acc in cur_accounts:
        db.add_account(acc, fnb.accounts[acc]['name'], float(fnb.accounts[acc]['balance']))
    elif acc in db_account_numbers:
        db.remove_account(acc)

Transaction.usernames = db_usernames
Transaction.accounts = db.get_accounts()

excel = Excel(excel_file)
if excel.workbook:
    for account in fnb.accounts:
        logger.info("Processing transactions in account: %s - %s", account, fnb.accounts[account]['name'])
        count = 0
        db_count = 0
        excel_count = 0
        transaction_count = len(fnb.accounts[account]['transactions'])
        if transaction_count > 0:
            for trans in reversed(fnb.accounts[account]['transactions']):
                count += 1
                logger.debug("Processing transaction %s/%s: date = %s, desc = %s, ref = %s, amount = %s.",
                             count, transaction_count, trans[0], trans[1], trans[2], trans[4])
                t = Transaction(trans, account)
                if db.check_transaction(t.account, t.date, t.description, t.reference, t.amount):
                    continue
                contrib = t.process_transaction()
                t.transaction_id = db.add_transaction(t.account, t.date, t.description, t.reference, t.amount,
                                                      t.user_id, t.month, t.year)
                if contrib and t.transaction_id:
                    db.update_user(t.user_id, t.username, t.date, t.amount, t.transaction_id)
                    db_count += 1
                    if excel.add_transaction(t):
                        excel_count += 1
            logger.info("Processed %s transactions in account: %s. %s added to database, %s written to excel",
                        transaction_count, fnb.accounts[account]['name'], db_count, excel_count)
            excel.update_account_balances(account, float(fnb.accounts[account]['balance']))
        else:
            logger.info('No transactions for account: %s', fnb.accounts[account]['name'])
    excel.close_workbook(overwrite=True)

wb_backup = 'TOPLINE TRACKING SHEET - {}.xlsx'.format(now.strftime('%B %Y'))
copy2(excel_file, wb_backup)

user_details = db.get_users()
gmail = Gmail()
gmail.authenticate()
for user in user_details:
    if user[4] != "TIG":
        logger.debug("User details: %s", user)
        msg = (f"Dear {user[1]} {user[2]},\n\n"
               f"Please find attached the tracking sheet for {MONTHS[now.month][0].capitalize()} {now.year}.\n\n")
        if user[9] is not None:
            next_date = date(year=user[5].year + user[5].month // 12, month=user[5].month % 12 + 1, day=5)
            if next_date < date(year=now.year + now.month // 12, month=now.month % 12 + 1, day=5):
                next_date = date(year=now.year + now.month // 12, month=now.month % 12 + 1, day=5)
            msg += (f"Your last contribution of R {user[6]:.2f} was received on {user[5].strftime('%d %B, %Y')} "
                    f"for {user[7]} {user[8]}.\n"
                    f"Your total contribution to date is R {user[9]:.2f} for a total share of {user[10]:.2f}%\n"
                    f"Your next contribution is due by {next_date.strftime('%d %B, %Y')}. "
                    f"The reference should be: {user[4]}-{next_date.strftime('%b-%y').upper()}.\n\n")

        msg += (f"Please ensure that all details contained in this email and the tracking sheet are correct.\n"
                f"If any errors are found, please contact thassan743@gmail.com.\n\n"
                f"Thank you.\n"
                f"The Topline Automated Contribution Tracker")

        logger.info("Sending email to %s %s: %s", user[1], user[2], user[3])
        message = gmail.create_message(user[3], 'Tracking - {}'.format(now.strftime('%B %Y')),
                                       msg, wb_backup)
        gmail.send_message('me', message)

logfile = [h.baseFilename for h in logger.handlers if type(h) == logging.FileHandler][0]
logfile = Path(logfile).name
logger.info("Sending logfile {}".format(logfile))
message = gmail.create_message('thassan743@gmail.com', 'Logfile - {}'.format(now.strftime('%B %Y')),
                               "Logfile for {}".format(now.strftime('%B %Y')), logfile)
gmail.send_message('me', message)

db.close_db()

logger.info("Making backup of database and excel workbook files.")
backup = Path('backup')
if not backup.exists():
    backup.mkdir()
zip_path = backup.joinpath('backup_{}.zip'.format(now.strftime('%Y%m%d_%H%M%S')))
logger.info("Backup path: %s", zip_path)
with zipfile.ZipFile(zip_path, 'w') as myzip:
    myzip.write(db_file)
    myzip.write(logfile)
    myzip.write(wb_backup)

logger.info("Done")
