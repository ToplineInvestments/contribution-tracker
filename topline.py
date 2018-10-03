from topline.scraper import FNB
import config

path_to_driver = None

try:
    username = config.fnb_username
    password = config.fnb_password
    url = config.fnb_url
    driver = config.driver
    if driver == 'chrome':
        path_to_driver = config.path_to_driver
except AttributeError:
    print("Error getting config from config.py file")
    raise SystemExit(0)

FNB = FNB(driver=driver, headless=False, driver_path=path_to_driver)

if FNB.driver:
    FNB.driver.get(url)
    if FNB.login(username, password):
        FNB.get_accounts(get_transactions=True)
        FNB.logout()
    else:
        FNB.driver.quit()
