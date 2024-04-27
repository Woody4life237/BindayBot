from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
import time
from datetime import datetime, timedelta;
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import requests
import urllib.request
import zipfile
import os
import shutil

def updateChromeDrivers():    
    print('Chrome drivers updating...')
    url = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
    r = requests.get(url= url)
    json = r.json()
    drivers = json.get("channels").get("Stable").get("downloads").get("chromedriver")
    
    for driver in drivers:
        if(driver.get("platform") == "win64"):
            url = driver.get("url")
            filehandle, _ = urllib.request.urlretrieve(url)
            zip_file_object = zipfile.ZipFile(filehandle, 'r')        
            for member in zip_file_object.namelist():
                filename = os.path.basename(member)
                # skip directories
                if not filename or not filename == "chromedriver.exe":
                    continue
                # copy file (taken from zipfile's extract)
                source = zip_file_object.open(member)
                target = open(os.path.join("./", filename), "wb")
                with source, target:
                    shutil.copyfileobj(source, target)
    print('Chrome drivers updated.')
    
def createEvent(name, eventDate): 
    print('Creating event (' + name + ').')    
    currYear = datetime.now().year        
    eventDate = eventDate.replace(year=currYear)
    eventDate = eventDate - timedelta(days=1)
    if eventDate < datetime.now():
        eventDate = eventDate.replace(year=currYear + 1)
    eventDate = eventDate.replace(hour=20, minute=0, second=0)
    
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/calendar.events']
    
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=54433)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    try:
        service = build('calendar', 'v3', credentials=creds)
    
        calendar_list = service.calendarList().list().execute()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == "Kieran and Ellen":
                calid = calendar_list_entry['id']        
                
                checkdatestart = eventDate - timedelta(hours=2)
                checkdateend = eventDate + timedelta(hours=2)
                
                events = service.events().list(calendarId=calid, timeMin=checkdatestart.strftime('%Y-%m-%dT%H:%M:%SZ'), timeMax=checkdateend.strftime('%Y-%m-%dT%H:%M:%SZ')).execute()
                
                eventRequired = True
                for event in events['items']:
                  if event['summary'] == name:
                      eventRequired = False
                if eventRequired:
                    event = {
                        'summary': name,
                        'start': {
                          'dateTime': eventDate.strftime('%Y-%m-%dT%H:%M:%S'),
                          'timeZone': 'Europe/London',
                        },
                        'end': {
                          'dateTime': eventDate.strftime('%Y-%m-%dT%H:%M:%S'),
                          'timeZone': 'Europe/London',
                        },
                        'reminders': {
                          'useDefault': False,
                          'overrides': [
                            {'method': 'popup', 'minutes': 10}
                          ],
                        },
                      }
                      
                    event = service.events().insert(calendarId=calid, body=event).execute()
                    print('Event created: %s' % (event.get('htmlLink')))   
                else:
                    print('Event already found')
                    
                checkdatestart = eventDate - timedelta(days=365)
                checkdateend = eventDate - timedelta(days=1)
                events = service.events().list(calendarId=calid, timeMin=checkdatestart.strftime('%Y-%m-%dT%H:%M:%SZ'), timeMax=checkdateend.strftime('%Y-%m-%dT%H:%M:%SZ')).execute()
                
                print('Deleting all %s events: %s - %s' % (name, checkdatestart, checkdateend))
                
                for event in events['items']:
                    if event['summary'] == name:
                        print('Deleting %s' % name + " " + event['start']['dateTime'])
                        service.events().delete(calendarId=calid, eventId=event['id']).execute()
                        
                
                checkdatestart = eventDate + timedelta(days=1)
                checkdateend = eventDate + timedelta(days=365)
                events = service.events().list(calendarId=calid, timeMin=checkdatestart.strftime('%Y-%m-%dT%H:%M:%SZ'), timeMax=checkdateend.strftime('%Y-%m-%dT%H:%M:%SZ')).execute()
                 
                print('Deleting all %s events: %s - %s' % (name, checkdatestart, checkdateend))
                
                for event in events['items']:
                    if event['summary'] == name:
                        print('Deleting %s' % name + " " + event['start']['dateTime'])
                        service.events().delete(calendarId=calid, eventId=event['id']).execute()
                        
    except Exception as error:
        print('An error occurred: %s' % error)
        
service = Service()
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")

print('Checking drivers.')
try:
    # Instantiate a webdriver
    driver = webdriver.Chrome(options=options, service=service)
except: 
    updateChromeDrivers()
    driver = webdriver.Chrome(options=options, service=service)    
print('Driver check complete.')

print('Opening web browser.')
# Load the HTML page
driver.get("https://forms.north-norfolk.gov.uk/outreach/BinCollectionDays.ofml")

print('Entering details.')
try:
    driver.find_element(By.ID, "F_Address_subform:Postcode").send_keys("NR12 8GG");
    driver.find_element(By.ID, "BA_Address_subform:Search_button").click();
    
    time.sleep(1)
    
    driver.find_element(By.ID, "F_Address_subform:Id").find_elements(By.TAG_NAME, "option")[3].click()
    
    time.sleep(1)
    print('Finding dates.')
    greybin = driver.find_element(By.XPATH, "/html/body/div/div/form/div/div[6]/div[4]/div[6]/fieldset/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/ul/li/strong[1]/span").find_element(By.XPATH, "./../..")
    greenbin = driver.find_element(By.XPATH, "/html/body/div/div/form/div/div[6]/div[4]/div[6]/fieldset/div/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/ul/li/strong[1]/span").find_element(By.XPATH, "./../..")
    
    greybindate = driver.find_element(By.XPATH, '//*[@id="Search_result_details_cps_hd"]/div/div/div/div/div[1]/div[2]/ul/li/strong[3]');
    greenbindate = driver.find_element(By.XPATH, '//*[@id="Search_result_details_cps_hd"]/div/div/div/div/div[2]/div[2]/ul/li/strong[3]');
    
    date = datetime.now().strptime(greybindate.text , '%A %d %B')
    if date:
        print('Found date for Grey bin')
        createEvent("Put Out Grey Bin", date)
        
    date = datetime.now().strptime(greenbindate.text , '%A %d %B')
    if date:
        print('Found date for Green bin')
        createEvent("Put Out Green Bin", date)
        
    driver.close()
except NoSuchElementException:
    print("Website not loading as expected")