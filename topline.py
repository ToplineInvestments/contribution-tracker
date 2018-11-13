from topline.scraper import FNB
from topline.excel import Excel
from topline.db import DB
from topline.transaction import Transaction
import configparser
import logging
from logging.config import fileConfig

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
        if 'transactions' in fnb.accounts[account]:
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
    excel.close_workbook(overwrite=False, filename='master_update.xlsx')

db.close_db()
logger.info("Done")
