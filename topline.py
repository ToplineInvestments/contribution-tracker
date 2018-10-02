from topline.scraper import FNB
import config

try:
    username = config.fnb_username
    password = config.fnb_password
    url = config.fnb_url
    driver = config.driver
    if driver == 'chrome':
        path_to_driver = config.path_to_driver
except AttributeError:
    print("Error getting config from config.py file")
    # End

FNB = FNB(driver=driver, headless=False, driver_path=path_to_driver)

if FNB.driver:
    FNB.driver.get(url)
    FNB.login(username, password)
    if FNB.open_account('Cheque'):
        cheque_acc = FNB.get_transactions()
        print(cheque_acc)

    FNB.logout()
