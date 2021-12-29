import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import datetime
import pandas as pd
import numpy as np

class StandardInput:
    def __init__(self) -> None:
        pass
    def Data(self, ):
        pass
    
class GSpreadAuthentication():
    def __init__(self, json_keyfile_name: str=None, spreadsheet_key: str=None) -> None:
        self.dt_now = datetime.datetime.now()
        intyyyymm=self.dt_now.year*100+self.dt_now.month+1 if self.dt_now.month!=12 else (self.dt_now.year+1)*100+1
        self.yyyymm=str(intyyyymm)
        
        scope=['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        if json_keyfile_name is not None:
            credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, scope)
            SPREADSHEET_KEY = spreadsheet_key
        else:
            private_key= "-----BEGIN PRIVATE KEY-----{key}-----END PRIVATE KEY-----\n".format(key=os.environ['SHEET_PRIVATE_KEY_STR'].replace('\\n', '\n'))
            credential = {
                            "type": "service_account",
                            "project_id": os.environ['SHEET_PROJECT_ID'],
                            "private_key_id": os.environ['SHEET_PRIVATE_KEY_ID'],
                            "private_key": private_key,
                            "client_email": os.environ['SHEET_CLIENT_EMAIL'],
                            "client_id": os.environ['SHEET_CLIENT_ID'],
                            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                            "token_uri": "https://oauth2.googleapis.com/token",
                            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                            "client_x509_cert_url":  os.environ['SHEET_CLIENT_X509_CERT_URL']
                        }
            credentials =  ServiceAccountCredentials.from_json_keyfile_dict(credential, scope)
            SPREADSHEET_KEY = os.environ['SHEET_SPREADSHEET_KEY_KG']
        self.spreadsheet =gspread.authorize(credentials)  
        self.workbook= self.spreadsheet.open_by_key(SPREADSHEET_KEY)

    def InputSchedule(self):
        #Open worksheet of schedule table of next month.
        worksheetname=self.yyyymm+" 予定表"
        worksheet = self.workbook.worksheet(worksheetname)

        cell_list = worksheet.get_all_values()
        df=pd.DataFrame(cell_list)
        df_input=df.drop(df.index[[0, 1]]).set_axis(df.iloc[1, :].tolist(), axis=1).reset_index(drop=True)
        index_date = [i for i, col in enumerate(df_input.columns) if '日付' in col]
        df_kagisime_input = df_input.iloc[:, index_date[0]:index_date[1]-1]
        df_gomisute_input = df_input.iloc[:, index_date[1]:index_date[2] if len(index_date) > 2 else None]

        df_kagisime=df_kagisime_input.copy()
        df_gomisute=df_gomisute_input.copy()
        for index in range(2, len(df_gomisute.columns.values)):
            if index < len(df_kagisime.columns.values):
                df_kagisime.iloc[:, index]=df_kagisime.iloc[:, index].apply(lambda x: x if x!='' else np.nan)
            df_gomisute.iloc[:, index]=df_gomisute.iloc[:, index].apply(lambda x: x if x!='' else np.nan)

        #Not assignable: False，Assignable: True
        df_gomisute.iloc[:, 2:]=df_gomisute.iloc[:, 2:].isnull()
        df_kagisime.iloc[:, 2:]=df_kagisime.iloc[:, 2:].isnull()

        df_gomisute['参加可能人数']=df_gomisute.iloc[:, 2:].sum(axis=1).astype(int)
        df_kagisime['参加可能人数']=df_kagisime.iloc[:, 2:].sum(axis=1).astype(int)
        df_gomisute['必要人数']=[4 if df_gomisute['参加可能人数'][i]>4 else df_gomisute['参加可能人数'][i] for i in df_gomisute.index.values]
        df_kagisime['必要人数']=[2 if df_kagisime['参加可能人数'][i]>2 else df_kagisime['参加可能人数'][i] for i in df_kagisime.index.values]
        self.df_kagisime=df_kagisime
        self.df_gomisute=df_gomisute

    def GetMailingList(self):
        #Make mailing list from spreadsheet.
        MailingListWorksheet = self.workbook.worksheet('MailingList')
        cell_list = MailingListWorksheet.get_all_values()
        df=pd.DataFrame(cell_list[1:][:], columns=cell_list[0])
        self.to_mails=[(mail, name) for mail, name in zip(df['Mail'], df['Name'])]
        self.FName=list(df.FName)

    def GetDecidedMailTemplate(self, recipient: str=None) -> dict:
        recipient='皆様' if recipient is None else recipient+'様'
        outworksheetname=self.yyyymm+" 配置"
        intyyyymm=int(self.yyyymm)
        outworksheet = self.workbook.worksheet(outworksheetname)
        shifttable_url='https://docs.google.com/spreadsheets/d/'+os.environ['SHEET_SPREADSHEET_KEY_KG']+'/edit#gid='+str(outworksheet.id)
        Sender=tuple(os.environ['FROM_MAIL'].split())
        nextintyyyymm=intyyyymm+1 if ((intyyyymm%100) % 12) !=0 else intyyyymm+100-11
        NextMonthScheduleWorksheet=self.workbook.worksheet(str(nextintyyyymm)+" 予定表")
        NextMonthScheduleURL='https://docs.google.com/spreadsheets/d/'+os.environ['SHEET_SPREADSHEET_KEY_KG']+'/edit#gid='+str(NextMonthScheduleWorksheet.id)

        MailTempWorksheet = self.workbook.worksheet('MailTemp')
        MemberDecisionMailTempCell=MailTempWorksheet.find('MemberDecisionMailTemp')
        MemberDecisionMail_content_temp=MailTempWorksheet.cell(MemberDecisionMailTempCell.row, MemberDecisionMailTempCell.col+1).value
        plaintextcontent=MemberDecisionMail_content_temp.replace('Recipient', recipient).replace('DecidedShiftTableSheetURL', shifttable_url).replace('Sender', Sender[1])
        plaintextcontent=plaintextcontent.replace('ScheduleSheetURL', NextMonthScheduleURL).replace('MonthAfterLater', str((nextintyyyymm%100)%12 if (nextintyyyymm%100)%12!=0 else 12))
        plaintextcontent=plaintextcontent.replace('NextMonth', str(self.dt_now.month+1 if self.dt_now.month!=12 else 1)).replace('Date', '24')

        MemberDecisionMailTitleTempCell=MailTempWorksheet.find('MemberDecisionMailTempTitle')
        MemberDecisionMail_title_temp=MailTempWorksheet.cell(MemberDecisionMailTitleTempCell.row, MemberDecisionMailTitleTempCell.col+1).value
        MemberDecisionMail_title=MemberDecisionMail_title_temp.replace('NextMonth', str(self.dt_now.month+1 if self.dt_now.month!=12 else 1))
        MemberDecisionMailContent={
            'Title': MemberDecisionMail_title,
            'PlainTextContent': plaintextcontent
        }

        return MemberDecisionMailContent
    
    def get_remind_mail_template(self, recipient=None):
        recipient='皆様' if recipient is None else recipient
        worksheetname=self.yyyymm+" 予定表"
        worksheet = self.workbook.worksheet(worksheetname)
        schedule_url='https://docs.google.com/spreadsheets/d/'+os.environ['SHEET_SPREADSHEET_KEY_KG']+'/edit#gid='+str(worksheet.id)
        #Select sheet named "MailTemp" and get templates of remind mail.
        MailTempWorksheet = self.workbook.worksheet('MailTemp')
        RemindMailTempCell=MailTempWorksheet.find('RemindMailTemp')
        remindmail_content_temp=MailTempWorksheet.cell(RemindMailTempCell.row, RemindMailTempCell.col+1).value
        RemindMailTitleTempCell=MailTempWorksheet.find('RemindMailTitleTemp')
        remindmail_Title_temp=MailTempWorksheet.cell(RemindMailTitleTempCell.row, RemindMailTitleTempCell.col+1).value

        Sender=tuple(os.environ['FROM_MAIL'].split())
        NextMonth=self.dt_now.month+1 if self.dt_now.month!=12 else 1
        Subject=remindmail_Title_temp.replace('NextMonth', str(NextMonth))
        Content=remindmail_content_temp.replace('ScheduleSheetURL', schedule_url).replace('Sender', Sender[1]).replace('NextMonth', str(NextMonth)).replace('Recipient', recipient)
        RemindMailContent={
            'Title': Subject,
            'PlainTextContent': Content
        }

        return RemindMailContent

    def update_spreadsheet(self, df_output: pd.DataFrame) -> None:
        """Update shift table of spreadsheet directly.
        Validate Google drive and spreadsheet API and set authentication info in advance. See "https://developers.google.com/workspace/guides/create-project".
        """
        #Open worksheet of next month shift table.
        outworksheetname=self.yyyymm+" 配置"
        outworksheet = self.workbook.worksheet(outworksheetname)

        #Find cell named '鍵閉め' and set cell range to copy df_output table.
        cell = outworksheet.find('鍵閉め')
        start_cell=num2alpha(cell.col)+str(cell.row+1)
        end_cell=num2alpha(cell.col+len(df_output.columns.values)-1)+str(cell.row+len(df_output.index.values))

        outworksheet.update(start_cell+':'+end_cell, df_output.values.tolist())

class InputviaSpreadSheet(StandardInput):
    def __init__(self) -> None:
        super().__init__()


def num2alpha(num):
    """Convert number to alphabet. 
    Use for convert R1C1 style to A1 style in spreadsheet.
    """
    if num<=26:
        return chr(64+num)
    elif num%26==0:
        return num2alpha(num//26-1)+chr(90)
    else:
        return num2alpha(num//26)+chr(64+num%26)
        
