import pandas as pd
from mip import Model, minimize, xsum
import os
import datetime
from icalendar import Calendar, Event
import pytz
import base64
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
from dotenv import load_dotenv
import copy
import manipulate_spreadsheet



def main(path_input = None, dir_output = None, direct_in=False, local=False, Auto_Mail=False):
    if direct_in:
        if local:
            json_keyfile_name="hoge/key_file_name.json"#Google API秘密鍵のパスを入力
            spreadsheet_key="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"#https://docs.google.com/spreadsheets/d/xxx/....のxxx部分を入力
            workbook=manipulate_spreadsheet.gaccount_auth_local(json_keyfile_name, spreadsheet_key) 
        else:
            workbook=manipulate_spreadsheet.gaccount_auth()

        df_kagisime, df_gomisute, yyyymm=manipulate_spreadsheet.df_direct_input(workbook)
    else:
        path_input=input("Please enter path of schedule csv file:\n") if path_input is None else path_input
        
        df_kagisime, df_gomisute=df_input(path_input)
        fname = os.path.basename(path_input)
        key = ' - '
        yyyymm = fname[fname.find(key)+len(key):fname.find(' 予定表')]

    df_output, _=member_decision_mip(df_kagisime=df_kagisime, df_gomisute=df_gomisute)

    dir_output="./example" if dir_output is None else dir_output
    if not (os.path.isdir(dir_output)):
        os.mkdir(dir_output)

    if direct_in:
        manipulate_spreadsheet.update_spreadsheet(workbook=workbook, df_output=df_output)

    to_mails, FName=manipulate_spreadsheet.get_mail_list(workbook)
    mail_content_temp=manipulate_spreadsheet.get_decided_mail_template(workbook, recipient='recipient')
    mail_content=copy.deepcopy(mail_content_temp)
    print('mail_content', mail_content)
    print('mail_content_temp', mail_content_temp, '\n\n')
    # .icsファイルを各メンバーごとに作成
    for to_mail, name in zip(to_mails, FName):
        member=to_mail[1]
        encoded_file=make_ical(df_output, dir_output, yyyymm, member, local) # ゴミ捨てに登録されている全員のicsファイルを作成
        mail_content['PlainTextContent']=mail_content_temp['PlainTextContent'].replace('recipient', member)
        icsfilename=yyyymm+'_'+name
        
        if Auto_Mail:
            send_mail(encoded_file, to_mail, mail_content, icsfilename)
    

    # 出力用ファイルの作成，出力
    if local:
        df_output.to_csv(os.path.join(dir_output, yyyymm + ' 配置.csv'), encoding = 'utf_8_sig')

def df_input(input_path=None):
    """Original code: https://github.com/yu9824/member_decision main.py main function. 
    Partially changed for menber_decision_mip function.
    """
    df=pd.read_csv(input_path, skiprows=1)
    index_Unnamed = [i for i, col in enumerate(df.columns) if 'Unnamed: ' in col]
    df_kagisime_input = df.iloc[:, index_Unnamed[0]+1:index_Unnamed[1]]
    df_gomisute_input = df.iloc[:, index_Unnamed[1]+1:index_Unnamed[2] if len(index_Unnamed) > 2 else None]

    #割り当て不可→False，割り当て可能→True
    df_gomisute_input.iloc[:, 2:]=df_gomisute_input.iloc[:, 2:].isnull()
    df_kagisime_input.iloc[:, 2:]=df_kagisime_input.iloc[:, 2:].isnull() 
    #csv読み込みの際のカラム名重複回避の.1を除去
    df_gomisute=df_gomisute_input.rename(columns=lambda s: s.strip('.1'))
    df_kagisime=df_kagisime_input.copy()

    df_gomisute['参加可能人数']=df_gomisute.iloc[:, 2:].sum(axis=1).astype(int)
    df_kagisime['参加可能人数']=df_kagisime.iloc[:, 2:].sum(axis=1).astype(int)
    df_gomisute['必要人数']=[4 if df_gomisute['参加可能人数'][i]>4 else df_gomisute['参加可能人数'][i] for i in df_gomisute.index.values]
    df_kagisime['必要人数']=[2 if df_kagisime['参加可能人数'][i]>2 else df_kagisime['参加可能人数'][i] for i in df_kagisime.index.values]

    return df_kagisime, df_gomisute

def member_decision_mip(df_kagisime, df_gomisute):
    """Decide shift table by solving MIP using Python-MIP.
    See Python-MIP documentation(https://docs.python-mip.com/en/latest/index.html).
    """
    #Constant
    N_days=df_gomisute.shape[0]
    L_gactivedays=[i for i, v in enumerate(df_gomisute['参加可能人数']) if v!=0]
    L_kactivedays=[i for i, v in enumerate(df_kagisime['参加可能人数']) if v!=0]
    N_gactivedays=len(L_gactivedays)
    N_kactivedays=len(L_kactivedays)
    N_gomisute_members=df_gomisute.shape[1]-4
    N_kagisime_members=df_kagisime.shape[1]-4
    L_ksplited=[L_kactivedays[idx:idx + 5] for idx in range(0,N_kactivedays, 5)]
    N_weeks=len(L_ksplited)

    m=Model()

    #Variable for optimization.
    V_kshift_table = m.add_var_tensor((N_days,N_kagisime_members), name='V_kshift_table', var_type='INTEGER', lb=0, ub=1)
    z_kequal_person = m.add_var_tensor((N_kagisime_members, ), name='z_kequal_person', var_type='INTEGER')
    z_kequal_week =m.add_var_tensor((N_weeks,N_kagisime_members), name='z_kequal_week', var_type='INTEGER')

    V_gshift_table = m.add_var_tensor((N_days,N_gomisute_members), name='V_gshift_table', var_type='INTEGER', lb=0, ub=1)
    z_gequal_person = m.add_var_tensor((N_gomisute_members, ), name='z_gequal_person', var_type='INTEGER')

    z_sameday_kg = m.add_var_tensor((N_gactivedays, N_kagisime_members), name='z_kequal_person', var_type='INTEGER')

    C_equal_person=1000
    Cl_equal_week=[100 for i in range(N_kagisime_members)]
    Cl_sameday_kg=[10 for i in range(N_kagisime_members)]

    # 目的関数
    m.objective = minimize(
                        C_equal_person*xsum(z_kequal_person)
                        +C_equal_person*xsum(z_gequal_person)
                        +xsum(Cl_equal_week[i]*xsum(z_kequal_week[:, i]) for i in range(N_kagisime_members))
                        +xsum(Cl_sameday_kg[i]*xsum(z_sameday_kg[:, i])for i in range(N_kagisime_members))
                        )

    #制約条件：
    for i,r in df_kagisime.iloc[:, 2:].iterrows():
        #入れない日には入らない（入れない日=0(False)）, 必要人数を満たす
        for j in range(N_kagisime_members):
            m += V_kshift_table[i][j] <=r[j]
        m += xsum(V_kshift_table[i])==r['必要人数']

    for i,r in df_gomisute.iloc[:, 2:].iterrows():
        for j in range(N_gomisute_members):
            m += V_gshift_table[i][j] <=r[j]
        m += xsum(V_gshift_table[i])==r['必要人数']
    #絶対値に変換するための制限
    for i in range(N_kagisime_members):
        m += (xsum(V_kshift_table[:, i]) - (N_kactivedays*2)//N_kagisime_members) >=-z_kequal_person[i]
        m += (xsum(V_kshift_table[:, i]) - (N_kactivedays*2)//N_kagisime_members) <=z_kequal_person[i]
        for j, l_weekday in enumerate(L_ksplited):
            m += (xsum(V_kshift_table[l_weekday, i]) - (len(l_weekday)*2)//N_kagisime_members) >=-z_kequal_week[j, i]
            m += (xsum(V_kshift_table[l_weekday, i]) - (len(l_weekday)*2)//N_kagisime_members) <=z_kequal_week[j, i]
        #差をとって絶対値にする->最小化：(k, g)->z (1, 0), (0, 1)->1, (1, 1), (0, 0)->0, (2, 0)->2, (2, 1)->1
        for j, v in enumerate(L_gactivedays):
            m += (V_kshift_table[v, i]-V_gshift_table[v, i]) >=-z_sameday_kg[j, i]
            m += (V_kshift_table[v, i]-V_gshift_table[v, i])<=z_sameday_kg[j, i]
    for i in range(N_gomisute_members):
        m += (xsum(V_gshift_table[:, i]) - (N_gactivedays*4)//N_gomisute_members) >=-z_gequal_person[i]
        m += (xsum(V_gshift_table[:, i]) - (N_gactivedays*4)//N_gomisute_members) <=z_gequal_person[i]

    m.optimize()

    kagisime_shift_table=(V_kshift_table).astype(float).astype(int)
    gomisute_shift_table=(V_gshift_table).astype(float).astype(int)

    df_kagisime['Result'] = [', '.join(j for i,j in zip(r,df_kagisime.iloc[:, 2:2+N_kagisime_members].columns) if i==1) for r in kagisime_shift_table]
    df_gomisute['Result'] = [', '.join(j for i,j in zip(r,df_gomisute.iloc[:, 2:2+N_gomisute_members].columns) if i==1) for r in gomisute_shift_table]
    print('目的関数', m.objective_value)
    print(df_kagisime[['日付','曜日','Result']])
    print(df_gomisute[['日付','曜日','Result']])
    
    df_output=pd.DataFrame({'鍵閉め':list(df_kagisime['Result']), 'ゴミ捨て': list(df_gomisute['Result'])}, index=list(df_kagisime['日付']))
    L_gomisute_members=df_gomisute.iloc[:, 2:2+N_gomisute_members].columns

    return df_output, L_gomisute_members

def make_ical(df, dir_output, filename, member, local):
    """Make ical file of his shift table. Same as https://github.com/yu9824/member_decision main.py make_ical function.
    """
    # カレンダーオブジェクトの生成
    cal = Calendar()

    # カレンダーに必須の項目
    cal.add('prodid', 'hito1280')
    cal.add('version', '2.0')

    # タイムゾーン
    tokyo = pytz.timezone('Asia/Tokyo')

    for name, series in df.iteritems():
        series_ = series[series.str.contains(member)]
        if name == '鍵閉め':
            start_td = datetime.timedelta(hours = 17, minutes = 45)   # 17時間45分
        elif name == 'ゴミ捨て':
            start_td = datetime.timedelta(hours = 12, minutes = 30)   # 12時間30分
        else:
            continue
        need_td = datetime.timedelta(hours = 1)

        for date, cell in zip(series_.index, series_):
            # 予定の開始時間と終了時間を変数として得る．
            start_time = datetime.datetime.strptime(date, '%Y/%m/%d') + start_td
            end_time = start_time + need_td

            # Eventオブジェクトの生成
            event = Event()

            # 必要情報
            event.add('summary', name)  # 予定名
            event.add('dtstart', tokyo.localize(start_time))
            event.add('dtend', tokyo.localize(end_time))
            event.add('description', cell)  # 誰とやるかを説明欄に記述
            event.add('created', tokyo.localize(datetime.datetime.now()))    # いつ作ったのか．

            # カレンダーに追加
            cal.add_component(event)

    # カレンダーのファイルへの書き出し
    encoded_file=base64.b64encode(cal.to_ical()).decode()
    if local:
        with open(os.path.join(dir_output, filename +member + '.ics'), mode = 'wb') as f:
            f.write(cal.to_ical())
    return encoded_file

def send_mail(encoded_file, to_mail, mail_content, icsfilename):

    attachedfile=Attachment(
        FileContent(encoded_file),
        FileName(icsfilename+'.ics'),
        FileType('application/ics'),
        Disposition('attachment')
    )


    message = Mail(
        from_email=tuple(os.environ['FROM_MAIL'].split()),
        to_emails=to_mail,
        subject=mail_content['Title'],
        plain_text_content=mail_content['PlainTextContent'])
    message.attachment=attachedfile
    try:
        sg = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
    except Exception as e:
        print(e.message)

if __name__ == '__main__':
    load_dotenv(verbose=True)
    dotenv_path='./.env'
    load_dotenv(dotenv_path)
    path_input="./example/GSS_test - 202111 予定表.csv"#Type path of schedule table csv file downloaded from spreadsheet unless direct_in=True.
    dir_output="./example"#Type path of output file(.ical, 予定表.csv).
    main(path_input=path_input, dir_output = dir_output, direct_in=True, local=False, Auto_Mail=True)