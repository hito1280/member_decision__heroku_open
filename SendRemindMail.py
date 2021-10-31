import os
from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import manipulate_spreadsheet

def main(local=False):
    if local:
        load_dotenv(verbose=True)
        dotenv_path='./.env'
        load_dotenv(dotenv_path)
        print(os.environ['EXAMPLE'])

    workbook=manipulate_spreadsheet.gaccount_auth()

    mail_content=manipulate_spreadsheet.get_remind_mail_template(workbook)
    #Make mailing list from spreadsheet.
    to_mails, _ = manipulate_spreadsheet.get_mail_list(workbook)

    message = Mail(
        from_email=tuple(os.environ['FROM_MAIL'].split()),
        to_emails=to_mails,
        subject=mail_content['Title'],
        plain_text_content=mail_content['PlainTextContent'])
    try:
        sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
    except Exception as e:
        print(e.message)

 
if __name__ == '__main__':
    path_input="./example/GSS_test - 202111 予定表.csv"#Type path of schedule table csv file downloaded from spreadsheet unless direct_in=True.
    dir_output="./example"#Type path of output file(.ical, 予定表.csv).
    main()
