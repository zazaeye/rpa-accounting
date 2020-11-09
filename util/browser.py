import time
import os
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException


class BrowserHelper(webdriver.Chrome):
    # noinspection PyMissingConstructor
    def __init__(self, chrome_driver_path, download_folder):
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_experimental_option(
            'prefs',
            {
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_settings": {"images": 2},
                "download.default_directory": download_folder
            }
        )
        self.download_folder = download_folder
        self.browser = webdriver.Chrome(chrome_driver_path, options=chrome_options)
        time.sleep(1)
        # initial logger
        self._logger = logging.getLogger(__name__)

    def quit(self):
        self._logger.debug("Quit the browser.")
        self.browser.quit()

    def download_invoice(self, invoice_link):
        self._logger.debug(f"Start to download invoice from the link '{invoice_link}'.")
        self.browser.get(invoice_link)
        time.sleep(3)
        WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable((By.ID, "print_btn"))
        ).click()
        time.sleep(3)
        WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable((By.ID, "hand_open_inv"))
        ).click()
        self._logger.debug(f"Finished downloading invoice from the link '{invoice_link}'.")

    def login_neticrm(self, login_url, account_name, password):
        self._logger.debug(f"Start to login netiCRM through the link '{login_url}'.")
        self.browser.get(login_url)
        time.sleep(1)
        WebDriverWait(self.browser, 10).until(
            EC.visibility_of_element_located((By.ID, "edit-name--2"))
        ).send_keys(account_name)
        time.sleep(1)
        WebDriverWait(self.browser, 10).until(
            EC.visibility_of_element_located((By.ID, "edit-pass--2"))
        ).send_keys(password)
        time.sleep(1)
        self.browser.find_element_by_id("edit-submit--2").click()
        time.sleep(1)
        if self.browser.current_url != login_url:
            self._logger.debug("Login to netiCRM success.")
        else:
            self._logger.error("Login to netiCRM FAILED! Maybe due to the CAPTCHA is presented.")
            raise RuntimeError("FAILED to login into NetiCRM platfrom.")

    def search_donation_by_date(self, start_date, end_date):
        self._logger.debug(f"Start to search donation in netiCRM between '{start_date}' and '{end_date}'.")
        self.browser.get("https://zazaeye.neticrm.tw/civicrm/contribute/search?reset=1")
        # fill in start_date
        self._logger.debug(f"Start to fill in the start_date '{start_date}'.")
        WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable((By.ID, "contribution_date_low"))
        ).click()
        time.sleep(1)
        Select(self.browser.find_element_by_class_name("ui-datepicker-year")) \
            .select_by_value(str(start_date.year))
        time.sleep(1)
        Select(self.browser.find_element_by_class_name("ui-datepicker-month")) \
            .select_by_value(str(start_date.month - 1))  # 5 means "六月"
        time.sleep(1)
        self.browser.find_element_by_xpath(
            f"//td[@data-handler='selectDay']/a[text()='{start_date.day}']") \
            .click()
        time.sleep(1)
        # fill in end_date
        self._logger.debug(f"Start to fill in the end_date '{end_date}'.")
        WebDriverWait(self.browser, 10).until(
            EC.element_to_be_clickable((By.ID, "contribution_date_high"))
        ).click()
        time.sleep(1)
        Select(self.browser.find_element_by_class_name("ui-datepicker-year")) \
            .select_by_value(str(end_date.year))
        time.sleep(1)
        Select(self.browser.find_element_by_class_name("ui-datepicker-month")) \
            .select_by_value(str(end_date.month - 1))  # 5 means "六月"
        time.sleep(1)
        self.browser.find_element_by_xpath(
            f"//td[@data-handler='selectDay']/a[text()='{end_date.day}']") \
            .click()
        time.sleep(1)
        # select only succeed
        self.browser.find_element_by_id('contribution_status_id[1]').click()
        time.sleep(1)
        # submit
        self.browser.find_element_by_id("_qf_Search_refresh").click()
        time.sleep(1)
        try:
            # get result table
            tbody = self.browser.find_element_by_class_name("selector")
            table_list = tbody.find_elements_by_xpath("./tbody/tr")
            self._logger.debug(f"Finished searching Neti donations, '{len(table_list)}' donations are found")
            return table_list
        except NoSuchElementException:
            self._logger.debug(f"Finished searching Neti donations, 'NO' donations are found")
            return []

    def get_latest_download_file_path(self):
        self._logger.debug(f"Try to get the latest downloaded file.")
        time.sleep(5)
        while True:
            files = os.listdir(self.download_folder)
            paths = [os.path.join(self.download_folder, file_name) for file_name in files]
            final_file_path = max(paths, key=os.path.getctime)
            self._logger.debug(f"Found the latest created file {final_file_path}.")
            if final_file_path.endswith(".crdownload"):
                self._logger.debug(f"File '{final_file_path}' is still downloading, please wait.")
            else:
                self._logger.debug(f"Found the latest created and downloaded file {final_file_path}.")
                break
        return final_file_path


    def _get_latest_download_file(self):
        # reference: https://stackoverflow.com/questions/34548041/selenium-give-file-name-when-downloading
        self.browser.execute_script("window.open()")
        time.sleep(1)
        # switch to new tab
        self.browser.switch_to.window(self.browser.window_handles[-1])
        time.sleep(1)
        # navigate to chrome downloads
        self.browser.get("chrome://downloads")
        time.sleep(1)
        lastest_download_file_name = self.browser.execute_script(
            "return document"
            ".querySelector('downloads-manager')"
            ".shadowRoot"
            ".querySelector('downloads-item[aria-rowindex=\"1\"]')"
            ".shadowRoot"
            ".querySelector('#file-link')"
            ".text"
        )
        time.sleep(1)
        self.browser.close()
        time.sleep(1)
        self.browser.switch_to.window(self.browser.window_handles[0])
        return lastest_download_file_name
