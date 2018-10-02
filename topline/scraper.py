from selenium import webdriver
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime


class Scraper:

    def __init__(self, driver, headless=False, driver_path=None):
        driver = driver.lower()
        self.driver = None
        if driver == 'firefox':
            self.driver = self.init_firefox_driver(headless)
        elif driver == 'chrome':
            self.driver = self.init_chrome_driver(headless, driver_path)
        else:
            print("Invalid or unsupported driver. Drivers supported: Chrome, Firefox")

    @staticmethod
    def init_firefox_driver(headless):
        options = FirefoxOptions()
        if headless:
            options.add_argument('--headless')
        try:
            driver = webdriver.Firefox(firefox_options=options)
        except WebDriverException:
            print("No Firefox driver found")
            return False
        return driver

    @staticmethod
    def init_chrome_driver(headless, path):
        options = ChromeOptions()
        if headless:
            options.add_argument('--headless')
        if not path:
            print("Need path to chromedriver.exe")
            return False
        try:
            driver = webdriver.Chrome(path, chrome_options=options)
        except WebDriverException:
            print("chromedriver not found at:", path)
            return False
        return driver

    def take_screenshot(self, filename=None):
        if not filename:
            filename = datetime.now().strftime('%Y%m%d_%H%M%S_Capture.png')
        return self.driver.save_screenshot(filename)


class FNB(Scraper):
    def __init__(self, driver, headless=False, driver_path=None):
        super().__init__(driver, headless=headless, driver_path=driver_path)
        self.accounts = {}

    def login(self, username, password):
        user_field = self.driver.find_element_by_xpath("//input[@id='user']")
        pass_field = self.driver.find_element_by_xpath("//input[@id='pass']")

        user_field.send_keys(username)
        pass_field.send_keys(password)

        self.driver.find_element_by_xpath("//input[@id='OBSubmit']").click()

        try:
            footer = self.driver.find_element_by_xpath("//div[@id='footerButtonGroup']")
            button = footer.find_element_by_tag_name('a')
            print("Pop Up visible, Pressing button")
            button.click()
        except NoSuchElementException:
            print("No popup")

        try:
            print("Waiting for load")
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "// *[ @ id = 'newsLanding'] / div[1]")))
        except TimeoutException:
            print("Timeout")
            return False

        print("Ready for account")
        return True

    def logout(self):
        logout_button = self.driver.find_element_by_xpath("//*[@id='headerButton_1']")
        logout_button.click()
        self.driver.quit()
        self.driver = None
        return True

    def open_account(self, name):
        # Make sure we're on the 'My Bank Accounts' tab
        my_accounts = self.driver.find_element_by_xpath("// *[ @ id = 'topTabs'] / span[2] / span")
        # check my_accounts.text = 'My Bank\nAccounts' before clicking
        my_accounts.click()

        try:
            print("Waiting for accounts")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='accountsTable']")))
        except TimeoutException:
            return False

        accounts_table = self.driver.find_element_by_id("accountsTable_tableContent")
        accounts = accounts_table.find_elements_by_xpath('//*[contains(@id,"nickname")]')
        if not (self.click_tab(accounts, name)):
            print("Multiple accounts found containing: " + name)
            return False
        print("Account opened")
        return True

    def get_transactions(self):
        print("Get transactions")
        # Make sure we're on the Transaction History Tab
        header = self.driver.find_element_by_xpath('//*[@id="subTabsScrollable"]')
        tabs = header.find_elements_by_xpath('//*[contains(@class,"subTabText")]')
        if not (self.click_tab(tabs, 'transaction')):
            return False

        transaction_table = self.driver.find_element_by_xpath("// *[ @ id = 'transactionHistoryTables_tableContent']")

        # Change to find all class contains tableGroup then loop through all tableGroup for class contains tableCell
        # Below should work for cheque account and savings
        # Doesn't work for credit card account but topline doesn't have a credit card so this might still work
        # Need to check layout of investment accounts
        dates = [d.text for d in transaction_table.find_elements_by_xpath('//*[@id="counter"]')]
        descriptions = [desc.text for desc in transaction_table.find_elements_by_xpath('//*[@id="shortDescription"]')]
        references = [ref.text for ref in transaction_table.find_elements_by_xpath('//*[@id="reference"]')]
        fees = [f.text for f in transaction_table.find_elements_by_xpath('//*[@id="serviceFee"]')]
        amounts = [amt.text for amt in transaction_table.find_elements_by_xpath('//*[contains(@id,"amount")]')]
        balances = [bal.text for bal in transaction_table.find_elements_by_xpath('//*[@id="ledgerBalance"]')]

        transactions = list(zip(dates, descriptions, references, fees, amounts, balances))
        return transactions

    def wait_for_loader(self):
        loader = self.driver.find_element_by_xpath('//*[@id="loaderOverlay"]')
        while loader.get_attribute('data-visible').lower() == 'true':
            pass

    def click_tab(self, tabs, tab_name):
        tab_names = [tab.text for tab in tabs]
        idx = [i for i, name in enumerate(tab_names) if tab_name.lower() in name.lower()]
        if len(idx) == 1:
            tabs[idx[0]].click()
            self.wait_for_loader()
            return True
        return False
