import os
import pickle
import base64
import logging
from datetime import datetime, timedelta, timezone
from lxml import html
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload


class BaseService(object):
    def __init__(self, token_name, scopes, credential_path='./credentials.json'):
        # Get credential or refresh credential and save it
        creds = None
        if os.path.exists(token_name):
            with open(token_name, 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log.py in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credential_path, scopes)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_name, 'wb') as token:
                pickle.dump(creds, token)
        self.creds = creds
        self._logger = logging.getLogger(__name__)


class GmailService(BaseService):
    def __init__(self, token_name, scopes, credential_path):
        super().__init__(token_name, scopes, credential_path)
        self.service = build("gmail", 'v1', credentials=self.creds, cache_discovery=False)

    @staticmethod
    def build_gamil_search_query(search_text=None, subject=None,
                                  start_date=None, end_date=None,
                                  relative_days=None):
        query_string = ""
        if search_text:
            query_string += "{0} ".format(search_text)
        if subject:
            query_string += "subject:({0}) ".format(subject)
        if start_date or end_date:
            if start_date:
                query_string += "after:{0} ".format(start_date.strftime("%Y/%m/%d"))
            if end_date:
                query_string += "before:{0} ".format((end_date + timedelta(days=1)).strftime("%Y/%m/%d"))
        elif relative_days:
            start_date = datetime.today() - timedelta(days=relative_days)
            end_date = datetime.today()
            query_string += "after:{0} before:{1}".format(start_date.strftime("%Y/%m/%d"),
                                                          end_date.strftime("%Y/%m/%d"))
        return query_string

    def get_gmail_search_result(self, q, page_token=''):
        self._logger.debug(f"Send the gmail search query '{q}'.")
        search_result = self.service.users().messages().list(
            userId="me", q=q, pageToken=page_token, maxResults=1000
        ).execute()
        if "messages" in search_result:
            self._logger.debug(f"Found '{len(search_result['messages'])}' email result.")
        else:
            self._logger.debug(f"No email result found.")
        return search_result

    def get_message_by_id(self, email_id):
        self._logger.debug(f"Get the email message by id: '{email_id}'.")
        return self.service.users().messages().get(userId='me', id=email_id).execute()

    def parse_email_content_from_id(self, email_id):
        self._logger.debug(f"Parse the email content by id: '{email_id}'.")
        email_return = self.service.users().messages().get(userId='me', id=email_id).execute()
        html_content = email_return['payload']['parts'][1]['body']['data']
        string_content = base64.urlsafe_b64decode(html_content).decode('utf-8')
        return html.fromstring(f'<html>{string_content}</html>')


class DriveService(BaseService):
    def __init__(self, token_name, scopes, credential_path,  upload_folder):
        super().__init__(token_name, scopes, credential_path)
        self.service = build("drive", 'v3', credentials=self.creds, cache_discovery=False)
        self.upload_folder = upload_folder

    def pdf_upload(self, upload_name, file_path):
        self._logger.debug(f"Start to upload '{file_path}' to the folder as '{upload_name}'.")
        metadata = {
            "name": upload_name,
            "parents": [self.upload_folder],
        }
        media_pdf = MediaFileUpload(file_path, mimetype="application/pdf")
        uploaded_file = self.service.files().create(
            body=metadata,
            media_body=media_pdf,
            supportsAllDrives=True).execute()
        self._logger.debug(f"Finish uploading '{file_path}' to the folder as '{upload_name}'.")
        return uploaded_file


class SheetsServcie(BaseService):
    def __init__(self, token_name, scopes, credential_path, sheet_id, sheet_range):
        super().__init__(token_name, scopes, credential_path)
        self.service = build("sheets", 'v4', credentials=self.creds, cache_discovery=False)
        self.sheet_id = sheet_id
        self.sheet_range = sheet_range
        self.upload_rows = []

    def add_row(
            self, date: datetime.date,  purpose: str, amount: int, from_account: str, to_account: str,
            certificate_type: str, verification: bool, certificate_collected: bool, certificate_upload="",
            email="rpa_user@zazaeye.org"):
        self._logger.debug(f"Start to add a new row in the sheet '{self.sheet_id}'.")
        self._logger.debug(f"Purpose is: '{purpose}', Amount is: '{amount}'.")
        check_list = [
            (amount, int),
            (date, datetime),
            (verification, bool),
            (certificate_collected, bool)
        ]
        for check_item in check_list:
            if not isinstance(check_item[0], check_item[1]):
                raise RuntimeError("Input `{0}` should be `{1}` type".format(
                    [k for k, v in locals().items() if v == check_item[0]][0],
                    check_item[1])
                )
        values = [
            datetime.now(timezone(timedelta(hours=8))).strftime("%Y/%-m/%-d %p %I:%M:%S"),  # 時間戳記
            date.strftime("%Y/%-m/%-d"),  # 消費日期
            purpose,  # 費用目的
            from_account,  # From 帳戶/科目/項目
            to_account,  # To 帳戶/科目/項目
            amount,  # 金額
            certificate_type,  # 憑證類型
            certificate_upload,  # 憑證上傳
            email,  # 電子郵件地址
            verification,  # 驗證通過
            certificate_collected,  # 憑證歸檔
        ]

        append_body = {
            "range": self.sheet_range,
            "majorDimension": "ROWS",
            "values": [values],
        }
        self.service.spreadsheets().values().append(
            spreadsheetId=self.sheet_id,
            range=self.sheet_range,
            body=append_body,
            valueInputOption="USER_ENTERED"
        ).execute()
        self._logger.debug(f"Finished adding a new row to the sheet '{self.sheet_id}'.")
