import re
import yaml
import logging
import argparse
from logging.config import fileConfig
from datetime import datetime, timedelta, timezone
from util.google import GmailService, DriveService, SheetsServcie
from util.browser import BrowserHelper


# Todo list:
#   *. change donation info from mailbox instead of netiCRM (due to CAPTCHA issue)
#   *. Think of how to update token if needed otherwise might encounter some issue in the future


class DateParseAction(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        if isinstance(values, str):
            setattr(namespace, self.dest, datetime.strptime(values, "%Y-%m-%d").date())
        elif isinstance(values, datetime.date):
            setattr(namespace, self.dest, values)


class ZAZARobot(object):
    def __init__(self, start_date, end_date, config_file_path):
        # read config in
        with open(config_file_path, 'r') as file:
            self._config = yaml.safe_load(file)
        # Create logger and set for format
        self._logger = logging.getLogger(__name__)
        # initial services
        self._logger.debug("Start to initial all the Google services")
        self.email_service = GmailService(self._config["GOOGLE_TOKEN_NAME"], self._config["GOOGLE_SCOPES"],
                                          self._config["GOOGLE_CREDENTIALS_JSON"])
        self.drive_service = DriveService(self._config["GOOGLE_TOKEN_NAME"], self._config["GOOGLE_SCOPES"],
                                          self._config["GOOGLE_CREDENTIALS_JSON"], self._config["CERTIFICATE_FOLDER"])
        self.sheets_service = SheetsServcie(self._config["GOOGLE_TOKEN_NAME"], self._config["GOOGLE_SCOPES"],
                                            self._config["GOOGLE_CREDENTIALS_JSON"], self._config["SHEET_ID"],
                                            self._config["SHEET_RANGE"])
        self.start_date = start_date
        self.end_date = end_date

    def crawl_newebpay_invoice(self):
        self._logger.info("Start to crawl Newebpay invoice")
        # get invoices
        invoice_result = self.email_service.get_gmail_search_result(
            self.email_service.build_gamil_search_query(
                subject="藍新金流電子發票開立通知",
                start_date=self.start_date,
                end_date=self.end_date,
            )
        )
        if "messages" not in invoice_result:
            self._logger.info("No email result for Newebpay invoice")
            return
        # initial browser
        browser_helper = BrowserHelper(self._config["CHROME_DRIVER_PATH"], self._config["CHROME_DOWNLOAD_FOLDER"])
        for email in invoice_result["messages"]:
            # parse email
            email_html = self.email_service.parse_email_content_from_id(email["id"])
            # get invoice_link and download invoice
            invoice_link = email_html.xpath('//*[text()="發票明細"]/@href')[0]
            browser_helper.download_invoice(invoice_link)
            invoice_file_path = browser_helper.get_latest_download_file_path()
            # upload invoice
            date = datetime.strptime(
                re.findall(r'\d{4}-\d{2}-\d{2}',
                           email_html.xpath('//*[contains(text(),"開立日期")]')[0].getnext().text
                           )[0],
                "%Y-%m-%d")
            last_month = date + timedelta(days=-20)
            purpose = f'藍新 {last_month.year} 年 {last_month.month} 月手續費'
            uploaded_file = self.drive_service.pdf_upload(
                upload_name="{0} 發票.pdf".format(purpose),
                file_path=invoice_file_path
            )
            # process all the info
            self.sheets_service.add_row(
                date=date,
                purpose=purpose,
                from_account="O_________1-1-5: 應收款項：應收未收之一切款項。",
                to_account="O_________5-2-12: 其他辦公費。",
                amount=int(re.findall(r"\d+", email_html.xpath("//*[contains(text(),'發票金額')]")[0].getnext().text)[0]),
                certificate_type="發票、收據",
                certificate_upload=f"https://drive.google.com/open?id={uploaded_file.get('id')}",
                verification=False,
                certificate_collected=False
            )
        # close browser
        browser_helper.quit()
        self._logger.info(f"Finished crawling Newebpay invoice, '{len(invoice_result['messages'])}' invoices found")

    def crawl_transfer_result(self):
        self._logger.info("Start to crawl transfer result")
        # get transfer notes
        transferred_result = self.email_service.get_gmail_search_result(
            self.email_service.build_gamil_search_query(
                subject="提領到帳",
                start_date=self.start_date,
                end_date=self.end_date,
            )
        )
        if "messages" not in transferred_result:
            self._logger.info("No email for account transfer")
            return
        for email_id in transferred_result["messages"]:
            # parse email
            email = self.email_service.get_message_by_id(email_id["id"])
            email_html = self.email_service.parse_email_content_from_id(email_id["id"])
            # process all the info
            email_date = datetime.fromtimestamp(int(email["internalDate"]) / 1000)
            self.sheets_service.add_row(
                date=email_date,
                purpose=f"藍新 {email_date.year} 年 {email_date.month} 月 {email_date.day} 日提領到帳戶",
                from_account="O_________1-1-5: 應收款項：應收未收之一切款項。",
                to_account="O_________1-1-2-1: 華南銀行存款",
                amount=int(
                    re.findall(
                        r'\$[\d,]+',
                        email_html.xpath('//*[contains(text(),"提領藍新金流帳戶")]/text()')[0])[0] \
                        .replace('$', '').replace(',', '')),
                certificate_type="現金、票據、證券等之收付移轉單據",
                verification=False,
                certificate_collected=False
            )
        self._logger.info(f"Finished crawling transfer result, '{len(transferred_result['messages'])}' transfer found")

    def crawl_neti_result(self):
        self._logger.info("Start to crawl neti donation result")
        # initial browser
        browser_helper = BrowserHelper(self._config["CHROME_DRIVER_PATH"], self._config["CHROME_DOWNLOAD_FOLDER"])
        # get donation result for neticrm
        browser_helper.login_neticrm(
            self._config["NETI_LOGIN_URL"],
            self._config["NETI_ACCOUNT_NAME"],
            self._config["NETI_ACCOUNT_PASWD"])
        result = browser_helper.search_donation_by_date(start_date=self.start_date, end_date=self.end_date)
        if len(result) == 0:
            self._logger.info("No donation in this period")
            return
        for row in result:
            # download receipt
            receipt_link = row.find_element_by_link_text('收據').get_attribute('href')
            browser_helper.browser.get(receipt_link)
            receipt_file_path = browser_helper.get_latest_download_file_path()
            # set upload info
            donator_name = row.find_element_by_class_name('crm-search-display_name').text
            donate_id = row.find_element_by_class_name('crm-contribution-trxn-id').text
            repeated_or_not = '定期捐款' if '_' in donate_id else '單筆捐款'
            purpose = f'{donator_name}-{repeated_or_not}-{donate_id}'
            # upload receipt
            uploaded_file = self.drive_service.pdf_upload(
                upload_name="{0} 收據.pdf".format(purpose),
                file_path=receipt_file_path
            )
            # process all the info
            donate_date = datetime.strptime(
                row.find_element_by_class_name('crm-contribution-receive_date').text[:10],
                "%Y-%m-%d"
            )
            donate_price = int(row.find_element_by_class_name('nowrap').text.replace('NT$ ', '').replace(',', ''))
            self.sheets_service.add_row(
                date=donate_date,
                purpose=purpose,
                from_account="O_____4-9: 其他收入：不屬於上列之各項收入。",
                to_account="O_________1-1-5: 應收款項：應收未收之一切款項。",
                amount=donate_price,
                certificate_type="發票、收據",
                certificate_upload=f"https://drive.google.com/open?id={uploaded_file.get('id')}",
                verification=False,
                certificate_collected=False
            )
        # close browser
        browser_helper.quit()
        self._logger.info(f"Finished crawling Neti donations, '{len(result)}' donations are found")




if __name__ == '__main__':
    # initial logging
    fileConfig("logging_config.ini")
    # Define all arguments
    parser = argparse.ArgumentParser(description="A tool for helping gather accounting info into sheet.")
    parser.add_argument("--start_date",
                        help="specify observe start date. Format: 'YYYY-mm-dd' ",
                        action=DateParseAction,
                        default=(datetime.now(timezone(timedelta(hours=8))) + timedelta(days=-1)).date())
    parser.add_argument("--end_date",
                        help="specify observe end date. Format: 'YYYY-mm-dd' ",
                        action=DateParseAction,
                        default=(datetime.now(timezone(timedelta(hours=8))) + timedelta(days=-1)).date())
    parser.add_argument("--config",
                        help="path to configuration file",
                        default="./config.yml")
    parser.add_argument("-a", "--crawl_all", help="crawl all result", action="store_true")
    parser.add_argument("--crawl_newebpay_invoice", help="crawl Newebpay inovice from mail", action="store_true")
    parser.add_argument("--crawl_transfer_result", help="crawl Newebpay transfer result", action="store_true")
    parser.add_argument("--crawl_neti_result", help="crawl NetiCRM donation result", action="store_true")
    args = parser.parse_args()
    # valid the start_date and end_date
    assert args.start_date <= args.end_date, "start_date should smaller than end_date."
    # Initial profiler instance
    zaza_robot = ZAZARobot(
        start_date=args.start_date,
        end_date=args.end_date,
        config_file_path=args.config
    )

    # run crawler
    if args.crawl_all:
        zaza_robot.crawl_newebpay_invoice()
        zaza_robot.crawl_transfer_result()
        zaza_robot.crawl_neti_result()
    else:
        if args.crawl_newebpay_invoice:
            zaza_robot.crawl_newebpay_invoice()
        if args.crawl_transfer_result:
            zaza_robot.crawl_transfer_result()
        if args.crawl_neti_result:
            zaza_robot.crawl_neti_result()
