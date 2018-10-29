from topline.scraper import FNB
from topline.excel import Excel
from topline.db import DB
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

DB = DB(db_file)
if not DB.connection:
    logger.error('DB error')
    raise SystemExit(0)

FNB = FNB(driver=driver, headless=False, driver_path=path_to_driver)

if FNB.driver:
    FNB.driver.get(url)
    if FNB.login(username, password):
        FNB.get_accounts(get_transactions=True)
        logger.debug("transactions = %s", FNB.accounts)
        FNB.logout()
    else:
        FNB.driver.quit()

excel = Excel(excel_file)
if excel.workbook:
    for account in FNB.accounts:
        logger.info("Processing transactions in account: %s", FNB.accounts[account]['name'])
        count = 0
        if 'transactions' in FNB.accounts[account]:
            for trans in FNB.accounts[account]['transactions']:
                if excel.add_transaction(FNB.accounts[account]['name'], FNB.accounts[account]['acc_num'], trans):
                    count += 1
            logger.debug("Processed %d/%d transactions in account: %s", count,
                         len(FNB.accounts[account]['transactions']), FNB.accounts[account]['name'])
        else:
            logger.debug('No transactions for account: %s', FNB.accounts[account]['name'])
    excel.close_workbook(overwrite=False, filename='master_update.xlsx')

logger.info("Done")
