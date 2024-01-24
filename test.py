from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import shutil
import os
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
import re
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
import pandas as pd
import time
import ssl
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.schedulers.background import BlockingScheduler

ssl._create_default_https_context = ssl._create_unverified_context

breakfast = ''
lunch = ''
dinner = ''


def read_menu() : 
    # 파일명
    file_name = '/Users/dkkim/Downloads/download/myfile.xlsx'

    # Daraframe형식으로 엑셀 파일 읽기
    weekday = datetime.today().weekday()
    
    #df = pd.read_excel(file_name, usecols=[2,*range(4, 9)])
    if(weekday > 4) :
        return None

    df = pd.read_excel(file_name, usecols=[2,weekday + 4], header=3)
    # 데이터 프레임 출력
    #print(df)
    #print(df.iloc[0,0])
    prev = df.iloc[0,0]
    index = 0
    length = len(df)
    breakfast = ''
    lunch = ''
    dinner = ''
    sum = ''
    flag = 0
    while index < length :
        current = df.iloc[index,0] 
        current_menu = df.iloc[index,1]
        if(current == prev) :
            sum += '\n'
    #        print(prev)
        elif(pd.isna(current)) :
            prev = prev
        else :
            prev = current
            sum += '\n'
    #        print(prev)
        if(pd.isna(current_menu)) :
            if(flag == 0) :
                breakfast = sum
                sum = ''
                flag = 1
            elif(flag == 1) :
                lunch = sum
                sum = ''
                flag = 2
        else :
            sum += '\n'
            sum += current_menu 
    #        print(current_menu) 
        index += 1
    dinner = sum
    return breakfast, lunch, dinner
    




def download_menu(download_path):
# 브라우저 꺼짐 방지 옵션
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"/Users/dkkim/Downloads/download",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
    })
    chrome_options.add_experimental_option("detach", True)

    dir = download_path
    driver = webdriver.Chrome(options=chrome_options)
    if os.path.isdir(dir) == False:
        os.mkdir(dir)

    # 웹페이지 해당 주소 이동
    driver.get("https://talk.tmaxsoft.com/login.do")

    user_input_id = driver.find_element(By.ID, 'id')
    user_input_pw = driver.find_element(By.ID, 'pass')

    user_input_id.send_keys("")
    user_input_pw.send_keys("")

    driver.find_element(By.ID, 'loginPage_btn').click()

    driver.get("https://talk.tmaxsoft.com/front/bbs/findBoardList.do?boardKind=BBS20140826008&bbsGroupCd=TM0007&curPageBbsDiv=TOTAL&menuLevel=2&srchMenuNo=TM0007&toggleMenuNo=TM0012")
    driver.find_element(By.ID, 'listT0').click()

    for files in os.listdir(dir):
        path = os.path.join(dir, files)
        try:
            shutil.rmtree(path)
        except OSError:
            os.remove(path)

    driver.find_element(By.XPATH, '//*[@id="detailView0"]/td/div/div/ul/li[2]/div/a').click()


def rename_excel(dir):
    try: # 가장 최근에 다운받은 파일 이름을 변경
        for files in os.listdir(dir):
            path = os.path.join(dir,files )
            shutil.move(os.path.join(dir, files),os.path.join(dir,r"myfile.xlsx"))
            new_file_name = dir + "\\" + r"myfile.xlsx"
            print(f'{new_file_name}')
            break
    except Exception as e:
        print(e)
        print('파일명 변경 실패')



app = App(token='xoxb-293903869666-6507268915922-sR4PpdDr87e2GWzyFzTSeWYF')


@app.message('아침') 
def notice_breakfast(message, say):
    breakfast, lunch, dinner = read_menu()
    say(f'<@{message['user']}>\n 오늘 아침 메뉴는 \n {breakfast} ')

@app.message("점심") 
def notice_lunch(message, say):
    breakfast, lunch, dinner = read_menu()
    say(f'<@{message['user']}>\n 오늘 점심 메뉴는 \n {lunch} ')

@app.message("저녁") 
def notice_lunch(message, say):
    breakfast, lunch, dinner = read_menu()
    say(f'<@{message['user']}>\n 오늘 저녁 메뉴는 \n {dinner} ')




scheduler = BackgroundScheduler()

# 정기적으로 메시지를 보내는 함수
def send_periodic_message():
    # 여기에 정기적으로 실행할 작업을 추가
    channel_id = "testslackbot"
    breakfast, lunch, dinner = read_menu()
    app.client.chat_postMessage(channel=channel_id, text=lunch)

# 1분마다 실행되도록 스케줄러에 등록
scheduler.add_job(send_periodic_message, 'cron', hour='10', minute='01', id='bob')




if __name__ == '__main__':
    scheduler.start()
    SocketModeHandler(app, 'xapp-1-A06EGLK7ESK-6507258272674-f55df3ae26ddb95f991926b39f85b90be3b20c3f96b04fdadd8ff713520aa4d7').start()



