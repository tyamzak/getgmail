from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


class GmailAPI:

    def __init__(self):
        # If modifying these scopes, delete the file token.json.
        self._SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
        # 認証情報ファイルを定義
        self.credentialPath = "client_secret_379796874485-ftuhgauo2nt1ob5ff9badcjl3r2fg7ju.apps.googleusercontent.com.json"

    def ConnectGmail(self):
        store = file.Storage('token.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(
                self.credentialPath, self._SCOPES)
            creds = tools.run_flow(flow, store)
        service = build('gmail', 'v1', http=creds.authorize(Http()))

        return service

    def GetLabelIDFromName(self, service, labelname: str = 'plazaans') -> str:
        """_summary_

        Args:
            service (gmail service): _made by "service = build('gmail', 'v1', http=creds.authorize(Http()))"_
            labelname (str): _label name which ID you want to know_

        Returns:
            str: _label ID_
        """
        uss = service.users
        LabelNamesIDs = service.users().labels().list(userId='me').execute()
        for label in LabelNamesIDs['labels']:
            if labelname == label['name']:
                labelID = label['id']
                break
        return labelID

    def GetMessageList(self, DateFrom, DateTo, labelname):

        # APIに接続
        service = self.ConnectGmail()

        MessageList = []

        query = ''
        # 検索用クエリを指定する
        if DateFrom != None and DateFrom != "":
            query += 'after:' + DateFrom + ' '
        if DateTo != None and DateTo != "":
            query += 'before:' + DateTo + ' '
        # if MessageFrom != None and MessageFrom !="":
        #     query += 'From:' + MessageFrom + ' '

        # labelIds = ['Label_3546802866937988889']

        labelIDs = self.GetLabelIDFromName(service, labelname)

        # メールIDの一覧を取得する(最大100件)
        messageIDlist = service.users().messages().list(
            userId='me', maxResults=100, q=query).execute()
        # 該当するメールが存在しない場合は、処理中断
        if messageIDlist['resultSizeEstimate'] == 0:
            print("Message is not found")
            return MessageList
        # メッセージIDを元に、メールの詳細情報を取得
        for message in messageIDlist['messages']:
            row = {}
            row['ID'] = message['id']
            MessageDetail = service.users().messages().get(
                userId='me', id=message['id']).execute()
            for header in MessageDetail['payload']['headers']:
                # 日付、送信元、件名を取得する
                if header['name'] == 'Date':
                    row['Date'] = header['value']
                elif header['name'] == 'From':
                    row['From'] = header['value']
                elif header['name'] == 'Subject':
                    row['Subject'] = header['value']
            MessageList.append(row)
        return MessageList


if __name__ == '__main__':
    test = GmailAPI()
    # パラメータは、任意の値を指定する
    messages = test.GetMessageList(
        DateFrom='2022-04-18',
        DateTo='2022-05-19',
        labelname='plazaans'
    )
    # 結果を出力
    for message in messages:
        print(message)
