from __future__ import print_function
from email.mime import message
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


class GmailAPI:

    def __init__(self):
        # If modifying these scopes, delete the file token.json.
        self._SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
        # 認証情報ファイルを定義
        self.credentialPath = ""

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
        print('メール情報取得中１')
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
        request = service.users().messages().list(userId='me', maxResults=100, q=query)
        # メールIDの一覧を取得する(最大100件)
        response = request.execute()

        messageIDlist = response

        # 該当するメールが存在しない場合は、処理中断、存在した場合は無くなるまで繰り返す
        if messageIDlist['resultSizeEstimate'] == 0:
            print("Message is not found")
            return MessageList
        else:
            # requestがNoneになるまで繰り返す
            while request is not None:

                request = service.users().messages().list_next(request, response)

                # Attribute Error occur when list_next returns None
                try:
                    response = request.execute()
                except AttributeError:
                    break

                if response['resultSizeEstimate'] == 0:
                    break
                else:
                    for mes in response['messages']:
                        messageIDlist['messages'].append(mes)
        print('メール詳細情報取得中')
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
            print(
                f"メール詳細情報取得中....{len(MessageList)} / {len(messageIDlist['messages'])}")
        return MessageList


def ResultPlot(MessageList, CountNames):
    print('プロット中')
    # メッセージリストをDataFrameに変換する
    import pandas as pd
    import matplotlib.pyplot as plt

    df = pd.DataFrame(MessageList)
    # 描画に必要な部分のみ残す
    del df['From'], df['ID']
    for name in CountNames:
        df[name] = df['Subject'].str.count(name)
    # カメラ名　時間帯　カウント数
    mini = min(df['Date'])
    maxi = max(df['Date'])

    df['Date'] = pd.to_datetime(df['Date'])
    df = df.set_index('Date')

    df = df.resample('H').sum()
    df = df.fillna(0)

    # N行 × 1列のグラフを設定 NはCountNamesの数
    row = len(CountNames)
    col = 1

    plt.rcParams['figure.figsize'] = (16.0, 10)

    for i in range(0, row):

        # N番目のグラフのプロット

        print(CountNames[i])
        print(df[CountNames[i]])
        plt.figure()
        plt.plot(df[CountNames[i]], label=CountNames[i])
        plt.legend()
        plt.savefig(f'image{CountNames[i]}.png')

    filename = 'result.csv'
    df.to_csv(filename)

    return


if __name__ == '__main__':
    test = GmailAPI()
    print('1')
    # パラメータは、任意の値を指定する
    messages = test.GetMessageList(
        DateFrom='2022-05-19',
        DateTo='2022-05-24',
        labelname='plazaans'
    )

    # 結果を描画
    # カウントしたい名前のリストを作成する
    CountNames = ['D' + str(x) for x in range(1, 30)]
    if messages:
        ResultPlot(messages, CountNames)
    else:
        print('該当メッセージ無し')
