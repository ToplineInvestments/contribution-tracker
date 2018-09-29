from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
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
