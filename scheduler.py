import datetime
from dotenv import load_dotenv

import auto_men_dec_mip_main
import SendRemindMail

dt_now = datetime.datetime.now()

if dt_now.day==20:
    SendRemindMail.main(local=False)
elif dt_now.day==25:
    auto_men_dec_mip_main.main(path_input = None, dir_output = None, direct_in=True, local=False, Auto_Mail=True)

""" load_dotenv(verbose=True)
dotenv_path='./.env'
load_dotenv(dotenv_path)
 """
