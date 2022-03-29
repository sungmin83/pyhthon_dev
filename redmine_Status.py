import requests
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials

class Crawler:
    def __init__(self):
        self.main_url = 'https://projects.rsupport.com'
        self.login_url = 'https://projects.rsupport.com/login'

        # 7.0.4
        self.total_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5180'
        self.complete_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5184'
        self.resolved_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5448'
        self.unResolved_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5449'
        self.stop_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5185'
        self.pause_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5183'

        # 7.0.4_QA
        self.qa_total_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5174'
        self.qa_complete_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5177'
        self.qa_resolved_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5446'
        self.qa_unResolved_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5447'
        self.qa_stop_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5179'
        self.qa_pause_url = 'https://projects.rsupport.com/projects/rc7/issues?query_id=5178'

        # 로그인 계정
        self.login_data = {
            'username':'smyoo',
            'password':'dream0917@'
        }

        self.status_url = [self.total_url, self.complete_url, self.resolved_url, self.unResolved_url, self.stop_url, self.pause_url]
        self.qa_status_url = [self.qa_total_url, self.qa_complete_url, self.qa_resolved_url, self.qa_unResolved_url,
                           self.qa_stop_url, self.qa_pause_url]


        self.scope = {
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        }

        self.json_file_name = 'J:\pyhthon_dev\spreadsheet-301501-1278dbf13d02.json'

        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(self.json_file_name, self.scope)
        self.gc = gspread.authorize(self.credentials)

        self.spreadsheet_url = 'https://docs.google.com/spreadsheets/d/11w84Mg9GZeoSCwVGPiwD5akcOL6oUZDAmV3WtVPIFZE/edit#gid=745116986'
        return

    def main(self):
        with requests.session() as s:
            main_req = s.get(self.main_url)
            html = main_req.text
            obj = BeautifulSoup(html, 'html.parser')
            auth_token = obj.find('input',{'name':'authenticity_token'})['value']
            self.login_data = {**self.login_data, **{'authenticity_token':auth_token}}

            login_req = s.post(self.login_url, self.login_data)

            # 7.0.4
            status_req = []
            status_obj = []
            for i in range(len(self.status_url)):
                status_req.insert(i, s.get(self.status_url[i]))
                status_obj.insert(i, BeautifulSoup(status_req[i].text, 'html.parser'))

            issue_status = []
            empty = 0
            for i in range(len(status_obj)):
                try:
                    issue_status.insert(i, status_obj[i].find('span', {'class': 'items'}).text)
                except AttributeError as e:
                    issue_status.insert(i, empty)

            # 7.0.4_QA
            qa_status_req = []
            qa_status_obj = []
            for i in range(len(self.qa_status_url)):
                qa_status_req.insert(i, s.get(self.qa_status_url[i]))
                qa_status_obj.insert(i, BeautifulSoup(qa_status_req[i].text, 'html.parser'))

            qa_issue_status = []
            empty = 0
            for i in range(len(qa_status_obj)):
                try:
                    qa_issue_status.insert(i, qa_status_obj[i].find('span', {'class': 'items'}).text)
                except AttributeError as e:
                    qa_issue_status.insert(i, empty)



            # 스프레스시트 문서 가져오기
            doc = self.gc.open_by_url(self.spreadsheet_url)

            # 시트 선택하기
            worksheet = doc.worksheet('Ref')

            # 7.0.4 데이터 삽입
            for i in range(len(issue_status)):
                worksheet.update_cell(i+55, 4, issue_status[i])

            # 7.0.4_QA 데이터 삽입
            for i in range(len(qa_issue_status)):
                worksheet.update_cell(i+63, 4, qa_issue_status[i])
        return
crawler = Crawler()
crawler.main()