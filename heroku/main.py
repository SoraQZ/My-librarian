from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
# from webdriver_manager.chrome import ChromeDriverManager  # pc上での動作確認時にコメントアウト解除
from time import sleep
import pandas as pd
import requests
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Heroku上のChrome Driverを指定(※デプロイするときはコメントを外す)
driver_path = '/app/.chromedriver/bin/chromedriver'  # ★変更ポイント：コメントアウトを解除した

# Headless Chromeをあらゆる環境で起動させるオプション
options = Options()
options.add_argument('--disable-gpu')
options.add_argument('--disable-extensions')
options.add_argument('--proxy-server="direct://"')
options.add_argument('--proxy-bypass-list=*')
options.add_argument('--start-maximized')
options.add_argument('--headless')

# 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# 認証情報設定
# ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    'KEY.json', scope)

# OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

# 共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = 'SPREADSHEET_KEY'

# 共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

df = pd.DataFrame(worksheet.get_all_values())

# 一番上の行をカラムに設定
df.columns = df.iloc[0]

# browser = webdriver.Chrome(ChromeDriverManager().install())  # pc上での動作確認時はコメントアウト外すとわかりやすい
browser = webdriver.Chrome(executable_path=driver_path,
                           chrome_options=options)  # pc上での動作確認時はコメントアウトする
sleep(2.8)

# 待機時間の設定（暗黙的待機）
browser.implicitly_wait(10)

# ['ISBN','蔵書有無','蔵書数','貸出可能数','貸出数','予約件数','予約率','本の紹介']
LIB_INFO = []

# 図書館のIDとPASS
ID = 'ID'
PASS = 'PASS'

# 現在のページの書籍を予約　(書籍詳細から動作確認済み)


def BOOKING_THIS_PAGE():
    try:
        # 予約かごに追加
        browser.find_element_by_class_name('bookshelf').click()
        sleep(3.38)

        # urlが変化したか？
        cur_url = browser.current_url

        if cur_url == 'https://ilisod002.apsel.jp/kawagoe/login':
            # ID入力,pass入力
            browser.find_element_by_id('loginName').send_keys(ID)
            sleep(1.23)
            browser.find_element_by_id('loginPassword').send_keys(PASS)
            sleep(1.47)

            # ログインボタンのクリック
            browser.find_element_by_id('loginSubmit').click()
            sleep(3.23)

        # 予約かごに追加をクリック
        # https://ilisod002.apsel.jp/kawagoe/item-details
        browser.find_element_by_class_name('bookshelf').click()
        sleep(3.38)

        # 予約かごをクリック
        browser.find_element_by_class_name('bottomButtonInnerBox').click()
        sleep(3.48)

        # 予約確認ボタンのクリック (2回実行で次のページ)
        for i in range(2):
            browser.find_element(By.XPATH, '//button[text()="予約"]').click()
            sleep(3.12)
        # """
        # 最終確認 送信ボタンのクリック
        browser.find_element(By.XPATH, '//button[text()="送信"]').click()
        sleep(2.89)
        print('予約完了')
        # """
    except NoSuchElementException:
        print('予約エラー　---------------BOOKING_THIS_PAGE:NoSuchElementException---------------')
        pass

# ISBNで川越図書館から書籍情報を取得:
#lib_info = ['ISBN','蔵書有無','蔵書数','貸出可能数','貸出数','予約件数','予約率','本の紹介']


def GET_LIB_INFO(isbn):
    # lib_infoの初期値を設定
    lib_info = [isbn, '-', '-', '-', '-', '-', '-', '-']
    try:
        p = re.compile('[0123456789]+')
        if p.fullmatch(str(isbn)):
            browser.get('https://ilisod002.apsel.jp/kawagoe/advanced-search')
            sleep(3.26)

            # 単独検索をクリック
            browser.find_element_by_link_text('単独検索').click()
            sleep(1.72)
            # print('単独検索をクリック')

            # 検索ボックスにISBNを入力
            browser.find_element_by_name('txtUniqCode').send_keys(isbn)
            sleep(1.36)
            #print('検索ボックスにISBNを入力', isbn)

            # 検索ボタンをクリック
            browser.find_element_by_xpath('//*[@id="main"]/form/div/div/button').click()
            sleep(3.14)
            # print('検索ボタンをクリック')

            # 検索結果 次のurlかチェック https://ilisod002.apsel.jp/kawagoe/search-results
            cur_url = browser.current_url
            # print(cur_url)

            # 検索結果の確認
            if cur_url == 'https://ilisod002.apsel.jp/kawagoe/search-results':
                # 資料が該当しました。
                #elem_result = browser.find_element_by_css_selector('.explain.done')
                # print('資料が該当しました。')

                # 詳細ページに移動
                bookTitle = browser.find_element_by_class_name('bookTitle')
                bookTitle.find_element_by_tag_name('span').click()
                sleep(2.54)
                # print('詳細ページに移動')

                # 詳細を取得
                title = browser.find_element_by_tag_name('h1').text
                book_intoro = browser.find_element_by_id('summaryText').text
                elems = browser.find_element_by_class_name('booksNumberInformation')
                info = [int(re.sub(r'\D', '', i.text))
                        for i in elems.find_elements_by_tag_name('td')]
                booking_rate = round(info[3] / info[0], 3)
                # print('詳細を取得')

                # 詳細をlib_infoに格納
                lib_info = [isbn, '蔵書あり', info[0], info[1],
                            info[2], info[3], booking_rate, book_intoro]
                # print('詳細をlib_infoに格納')

                # """
                # 本1冊当たりの予約数が4冊以上で予約
                if booking_rate >= 5 and my_book_info != '予約中':
                    BOOKING_THIS_PAGE()
                # """
    except NoSuchElementException:
        print('---------------GET_LIB_INFO:NoSuchElementException---------------')
        pass

    return(lib_info)


try:
    for i in df.iloc:
        if i['蔵書有無'] == '蔵書あり':
            my_book_info = i['読書状況']
            print('*'*50)

            print(i['タイトル'])

            ISBN = str(i['ISBN'])
            INFO = GET_LIB_INFO(str(i['ISBN']))
            LIB_INFO.append(INFO)
            if not INFO[1:6] == ['-', '-', '-', '-', '-']:
                df.loc[df['ISBN'].isin([ISBN]), '蔵書有無':'本の紹介'] = INFO[1:]
                # print('1行dfの更新')
            print(i['ブクログ順'], INFO[1:6])
    browser.close()
    print('dfの更新完了　for文')

except:
    browser.close()
    print('dfの更新エラー　for文')


def toAlpha(num):
    if num <= 26:
        return chr(64+num)
    elif num % 26 == 0:
        return toAlpha(num//26-1)+chr(90)
    else:
        return toAlpha(num//26)+chr(64+num % 26)


col_lastnum = len(df.columns)  # DataFrameの列数
row_lastnum = len(df.index)   # DataFrameの行数

cell_list = worksheet.range('A1:'+toAlpha(col_lastnum)+str(row_lastnum))
for cell in cell_list:
    val = df.iloc[cell.row-1][cell.col-1]
    cell.value = val
worksheet.update_cells(cell_list)
print('シートの更新完了')
