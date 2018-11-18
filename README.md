# contribution-tracker
Automated contribution tracker for Topline Investment Group

### Dependencies
```sh
pip install selenium
pip install openpyxl
pip install --upgrade google-api-python-client oauth2client
```

#### Supported WebDrivers
* [ChromeDriver](http://chromedriver.chromium.org/) (Tested on v2.37 - Not working)
* Firefox [geckodriver](https://github.com/mozilla/geckodriver/releases)
(Tested on v0.23.0 with Firefox v63.0.3 - Working)

#### Gmail API setup
Follow the instructions found [here](https://developers.google.com/gmail/api/quickstart/python)
to enable the Gmail API and download the credentials file.
