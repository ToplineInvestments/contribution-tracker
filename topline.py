from topline.scraper import FNB
from topline.excel import Excel
from topline.db import DB
import configparser

config = configparser.ConfigParser()
if not config.read('config.ini'):
    print('No config.ini file found')
    raise SystemExit(0)

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
    print("Error getting config from config.ini file")
    raise SystemExit(0)

DB = DB(db_file)
if not DB.connection:
    print('DB error')
    raise SystemExit(0)

FNB = FNB(driver=driver, headless=False, driver_path=path_to_driver)

if FNB.driver:
    FNB.driver.get(url)
    if FNB.login(username, password):
        FNB.get_accounts(get_transactions=True)
        FNB.logout()
    else:
        FNB.driver.quit()

excel = Excel(excel_file)
if excel.workbook:
    for account in FNB.accounts:
        print("Processing transactions in account: {}".format(FNB.accounts[account]['name']))
        count = 0
        if 'transactions' in FNB.accounts[account]:
            for trans in FNB.accounts[account]['transactions']:
                if excel.add_transaction(FNB.accounts[account]['name'], FNB.accounts[account]['acc_num'], trans):
                    count += 1
            print("Processed {}/{} transactions in account: {}".format(count,
                                                                       len(FNB.accounts[account]['transactions']),
                                                                       FNB.accounts[account]['name']))
        else:
            print('No transactions for account: {}'.format(FNB.accounts[account]['name']))
    excel.close_workbook(overwrite=False, filename='master_update.xlsx')

print("Done")
