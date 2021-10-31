# member_decision_heroku
当番表を決定し、決定したシフトをメールで共有するプログラムです．

## Feature
herokuのschedulerと組み合わせることで毎月決まった日にリマインドメールと当番表決定を実行します．  
### member decision
当番表の決定は個人の予定がある日には割り振らないなどの制約条件のもと、担当者の偏りなどを最小化するようにしています．  
予定表の読み込み，シフト表の共有はスプレッドシートにて行い，各個人の担当日をまとめたicalファイルを作成，メールで共有します．  
具体的な当番表の決定は，制約条件として  
* 個人の予定がある日には割り振らない  
* 毎日2人/4人を割り振る  
（加えて目的関数で絶対値を用いるための制約条件）  
目的関数として  
* 各担当者の担当回数を平均に近づける（ペナルティ1000 pt.）  
* 週ごとの担当日数が偏らないようにする（ペナルティ100 pt., 担当者ごとに設定可能）  
* ごみ捨てをする日はなるべく鍵閉めもする（ペナルティ10 pt., 担当者ごとに設定可能）  
とした整数混合線形最適化問題を解くことで行っています．   

### program
* scheduler.py  
毎月決まった日にSendRemindMail.pyとauto_mem_dec_mip_main.pyを実行  
* SendRemindMail.py  
スプレッドシートのテンプレートをもとにリマインドメールを送る．  
* auto_mem_dec_mip_main.py  
スプレッドシートに記入した個人の予定表をもとにシフトを決定．  
* manipulate spreadsheet.py  
スプレッドシートを操作する部分を抜き出したもの．  

## How to use
###　accounts requirement
以下のサービスを使います．
* heroku  
無料版（以上）を登録．定期実行のためにHeroku Schedulerアドイン（Standardでok）を追加してください．  
* google spread sheet  
個人の予定表と決まったシフト表，メールテンプレートを記載したシートを作成してください．  
APIが利用できるように設定してください．  
* SendGrid  
メールを送るために使用します．  

### heroku environment variables requirement
Herokuにdeployする際にgitを使うため，誤ってAPIキーを公開することを防ぐためにAPIキーは環境変数に格納します．  
環境変数に以下を設定してください．  
* google spreadsheet APIの認証用  
SHEET_PROJECT_ID=<"project_id" in json file>  
SHEET_PRIVATE_KEY_ID=<"private_key_id" in json file>  
SHEET_PRIVATE_KEY=<"private_key" in json file. Convert \n>  
SHEET_CLIENT_EMAIL=<"client_email" in json file>  
SHEET_CLIENT_ID=<"client_id" in json file>  
SHEET_CLIENT_X509_CERT_URL=<"client_x509_cert_url" in json file>  
*"auth_uri", "token_uri" and "auth_provider_x509_cert_url" are common.  
* アクティブなスプレッドシートを設定
SHEET_SPREADSHEET_KEY_KG=<"xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" in spread sheet URL of https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx/....>  
* To activate sendgrid api.  
SENDGRID_API_KEY="SendGrid API key"  
* Sender info.  
FROM_MAIL="e-mail address"+"space"+"Name" (e.g. : example@example.com anonymous)  
* If you load enviroment variables from .env(set local=True), you must add below.  
EXAMPLE=Successfully loaded environment variables from .env file.  

### Deploy 
Herokuへのdeployは他の記事を参考にしてください．  

