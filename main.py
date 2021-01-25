from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from datetime import date
import os
from discord_webhooks import DiscordWebhooks

from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth

global download_tries
download_tries=0

webhook_url = ' MY WEBHOOK URL'

URL1 = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=f3d888b7-3d8f-4fa2-a770-c63acb01dc1c&&client-request-id=c5729da2-52ee-48c6-8d3e-75f40730e24a&x-client-SKU=Js&x-client-Ver=1.0.9&nonce=0f957d75-9105-467b-850b-0e6a9c2b7ce4&domain_hint="
URL2 = "https://teams.microsoft.com/_#/school//?ctx=teamsGrid"
subjects = ['B2_PHY_1701_F_20-21',"F2 SLOT/ TT 415/ STS1201",'Calculus for Engineers Slot:E2+TE2    Lab: L11+L12','L19+L20 EC Lab FALL 2020','FALL 2020 D2 slot','BEEE Workshop','FALL 20-21 TE-II ENG 1902','Electric circuits']
subject={'B2_PHY_1701_F_20-21':'PHYSICS',"F2 SLOT/ TT 415/ STS1201":'SOFT SKILLS','Calculus for Engineers Slot:E2+TE2    Lab: L11+L12': "MATHS",'L19+L20 EC Lab FALL 2020': 'CHEMISTRY LAB','FALL 2020 D2 slot':'CHEMISTRY THEORY','BEEE Workshop':"BEE WORKSHOP",'FALL 20-21 TE-II ENG 1902':"ENGLISH 1902",'Electric circuits':"ELECTRIC CIRCUITS"}
#'B2_PHY_1701_F_20-21'



global i
i=0
def send_msg(problem,sub):

    WEBHOOK_URL = webhook_url

    webhook = DiscordWebhooks(WEBHOOK_URL)
    # Attaches a footer
    webhook.set_footer(text='')

    webhook.set_content(title= sub, description=problem)

    webhook.send()

    print("Sent message to discord")


def authorize():

    # gauth = GoogleAuth()
    # gauth.LocalWebserverAuth()
    # global drive
    # drive = GoogleDrive(gauth)

    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile("mycreds.txt")
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("mycreds.txt")
    global drive
    drive = GoogleDrive(gauth)

def start_browser():

    global driver
    o = webdriver.ChromeOptions()
    o.add_argument("user-data-dir=N:\\google\\User Data")  # Path to your chrome profile

    # o.add_argument('--headless')
    # o.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=o)
    driver.minimize_window()
    driver.get(URL2)


def login():
    start_browser()
    global driver
    driver.get(URL1)
    print("logging in")
    emailField = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "i0116")))
    #     emailField = driver.find_element_by_xpath('//*[@id="i0116"]')
    emailField.click()
    emailField.send_keys(['MY EMAIL'])

    driver.find_element_by_xpath('//*[@id="idSIButton9"]').click()

    password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "i0118")))
    #     password = driver.find_element_by_xpath('//*[@id="i0118"]')
    password.send_keys('MY PASSWORD')
    #     driver.find_element_by_xpath('//*[@id="idSIButton9"]').click()
    #     driver.find_element_by_xpath('//*[@id="idSIButton9"]').click()
    time.sleep(2)
    sign = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "inline-block")))
    sign.click()
    sign = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "inline-block")))
    sign.click()



def remove_video(date_of_video=(date.today()).strftime("%B %d, %Y")):

    while True:
        print("inside removing loop")
        response = drive.ListFile({"q": 'trashed=false'}).GetList()
        for folder in response:
            if folder['title'] == date_of_video:
                os.remove(r'N:\\Downloads\\video.mp4')
                print("removed")
                return True

def searching():
    a=time.time()
    while True:
        if time.time()-a > 1800:
            print(" File not found after searching for 30mins :( ")
            return False

        if os.path.exists("N:\\Downloads\\video.mp4"):
            print("Video downloaded and found :)")
            return True

def upload_video(subject_name, date_of_video=(date.today()).strftime("%B %d, %Y")):
    authorize()
    print(1)
    response = drive.ListFile({"q": "mimeType='application/vnd.google-apps.folder' and trashed=false"}).GetList()

    for folder in response:
        if folder['title'] == subject[subject_name]:
            a = folder
    # z=(date.today()).strftime("%B %d, %Y")
            file = drive.CreateFile({
                'title': date_of_video ,
                'parents': [{
                    'kind': 'drive#fileLink',
                    'id': a['id']
                }]
                    })

            file.SetContentFile(r'N:\\Downloads\\video.mp4')
            file.Upload()
            print("uploading :)")

def chat_collapser(subject_name):

        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.collapsed-indicators")))
        element = (driver.find_elements_by_css_selector('div.collapsed-indicators'))[-1]  # pressing collapse
        webdriver.ActionChains(driver).click(element).perform()
        webdriver.ActionChains(driver).click(element).perform()


def video_downloader(subject_name,date_of_video):

    global download_tries
    try:
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.recording-thumbnail")))
        time.sleep(1)
        element = (driver.find_elements_by_css_selector('div.recording-thumbnail'))[-1]  # pressing download button
        webdriver.ActionChains(driver).click(element).perform()
        webdriver.ActionChains(driver).click(element).perform()
    except:
        try:
            print("exce")
            element = (driver.find_elements_by_css_selector('div.collapsed-indicators'))[-1]  # pressing collapse
            webdriver.ActionChains(driver).click(element).perform()
            element = (driver.find_elements_by_css_selector('div.recording-thumbnail'))[-1]  # pressing download button
            webdriver.ActionChains(driver).click(element).perform()
            webdriver.ActionChains(driver).click(element).perform()
        except:
            element = (driver.find_elements_by_css_selector('div.collapsed-indicators'))[-1]  # pressing collapse
            webdriver.ActionChains(driver).click(element).perform()
            webdriver.ActionChains(driver).click(element).perform()
            element = (driver.find_elements_by_css_selector('div.recording-thumbnail'))[-1]  # pressing download button
            webdriver.ActionChains(driver).click(element).perform()
            webdriver.ActionChains(driver).click(element).perform()




    if download_tries == 5:
        send_msg("Download keeps failing, even after 5 tries : |",subject_name)
        return False
    if searching():
        upload_video(subject_name, date_of_video)
        remove_video(date_of_video)
        return True
    else:
        print("Video not downloaded, trying again")
        video_downloader(subject_name,date_of_video)
        download_tries += 1


def run(date_of_video = (date.today()).strftime("%B %d, %Y")):
    global flag
    global driver
    flag=0

    start_browser()
    try:
        element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.use-app-lnk")))
        webdriver.ActionChains(driver).click(element).perform()
    except:
        print("no app add shown :)")
    print('logged in')

    team_card_checker = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "team-card")))
    l=len(driver.find_elements_by_class_name('team-card'))

    for x in range(l-2):
        driver.get(URL2)
        team_card_waiter = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "team-card")))

        driver.find_elements_by_class_name('team-card')[x].click()

        for y in subjects:
            time.sleep(1)
            if y in driver.page_source:
                print("Checking if",y,'had class today')
                subject_name=y
                time.sleep(1)

                if 'sanitizeStringAsHtml">Today</div>' in driver.page_source:

                    print(subject_name,"had a class today")

                    chat_collapser(subject_name)
                    video_downloader(subject_name,date_of_video)


                else: print("No classes found today for", subject_name)

    send_msg("Status","Program ran and completed!  :)")
    driver.quit()




try:
    run()
except:
    send_msg("Problem", "Something went wrong :(")
