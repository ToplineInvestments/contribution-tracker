from topline.scraper import FNB
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
