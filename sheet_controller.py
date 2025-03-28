import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

class SheetController:
    def __init__(self, sheet_name):
        credentials = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        self.client = gspread.authorize(credentials)
        self.sheet_name = sheet_name



    def _connect(self):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/spreadsheets']
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.creds_file, scope)
        return gspread.authorize(creds)


    def get_sheet(self, worksheet_name):
        sheet = self.client.open(self.sheet_name)
        return sheet.worksheet(worksheet_name)
    
    
    def verify_user(self, student_id, password):
        worksheet = self.get_sheet('학생DB')
        records = worksheet.get_all_records()
        for record in records:
            if record['학번'] == student_id and record['비밀번호'] == password:
                return record
        return None

    def increment_usage(self, student_id):
        worksheet = self.get_sheet('학생DB')
        cell = worksheet.find(student_id)
        usage_col = worksheet.find('사용횟수').col
        current_usage = int(worksheet.cell(cell.row, usage_col).value or 0)
        worksheet.update_cell(cell.row, usage_col, current_usage + 1)

    def get_prompts(self, prompt_type='전반', subject=None):
        worksheet = self.get_sheet('프롬프트')
        records = worksheet.get_all_records()
        prompts = []
        for record in records:
            if record['종류'] == prompt_type:
                if subject and record['교과명'] != subject:
                    continue
                prompts.append(record['시스템프롬프트'])
        return prompts

    def log_chat(self, student_id, question, answer, date):
        worksheet = self.get_sheet('채팅로그')
        worksheet.append_row([student_id, question, answer, date])
