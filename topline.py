from topline.scraper import FNB
from topline.excel import Excel
from topline.db import DB
import configparser

config = configparser.ConfigParser()
if not config.read('config.ini'):
    print('No config.ini file found')
    raise SystemExit(0)

DB = DB(config['DB']['FILENAME'])
if not DB.connection:
    print('DB error')
    raise SystemExit(0)

path_to_driver = None

try:
    username = config['SCRAPER']['USERNAME']
    password = config['SCRAPER']['PASSWORD']
    url = config['SCRAPER']['URL']
    driver = config['SCRAPER']['DRIVER']
    if driver == 'chrome':
        path_to_driver = config['SCRAPER']['DRIVER_PATH']
except KeyError:
    print("Error getting config from config.ini file")
    raise SystemExit(0)

FNB = FNB(driver=driver, headless=False, driver_path=path_to_driver)

if FNB.driver:
    FNB.driver.get(url)
    if FNB.login(username, password):
        FNB.get_accounts(get_transactions=True)
        FNB.logout()
    else:
        FNB.driver.quit()

excel = Excel(config['EXCEL']['FILENAME'])
if excel.workbook:
    excel.get_sheets()
    for account in FNB.accounts:
        excel.update_excel(FNB.accounts[account]['transactions'])
    excel.close_workbook()
