# Load needed packages:
# -------------------------------------------------------------------------------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# for waits:
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# for printing current time when crashes
import time
import datetime
# for htmls:
from bs4 import BeautifulSoup
# for loading the page : start only after loading
from selenium.webdriver.chrome.options import Options
# import random for time
import random
# for loading sound directly in python (not working?)
from playsound import playsound
# for openning sound file from computer
import os
alarm = "C:\\Users\\meita\\Documents\\auto reservation\\fire.wav"
# for opening the web browser
PATH = "C:\\Program Files (x86)\\chromedriver.exe"
driver = webdriver.Chrome(PATH)
# for voicing errors 
from win32com.client import Dispatch
speak = Dispatch("SAPI.SpVoice")

# Define functions:
# -------------------------------------------------------------------------------------------------------------------------------------

def check_badgateway():
    # check for http Errors
    print ("\nCHECKING GATEWAY\n")
    # check if h1 is an error
    obj = WebDriverWait(driver, 90).until(
        EC.presence_of_element_located((By.TAG_NAME, 'h1')))
    
    print(obj.text)
    txt = obj.text
    if "502" in txt or "503" in txt:
        print("\nDETECTED: Bad Gateway\n")
        time.sleep(random.randint(15, 120))
        # refresh window
        driver.refresh()
        # re-run check on the refreshed page
        print("\nREFRESHED PAGE\n")
        time.sleep(random.randint(15, 60))
        check_badgateway()
       
    else:
        print("NO GATEWAY PROBLEMS")
        
def cookies():
    # RUN ONLY ONCE
    # accept cookies
    element = WebDriverWait(driver, 90).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Accepter')))
    element.click()
    print("cockies")

def loop_reserve():
    print("SCANNING: FIRST PAGE")
    # check for http Errors
    check_badgateway()

    # click checkbox to accept conditions
    element = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="condition"]')))
    element.click()
    print("CLICKED: CHECKBOX")

    # accept and move the the next page
    link = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="submit_Booking"]/input[1]')))
    link.click()
    print("CLICKED: NEXT PAGE")

    # check if the page didn't crash during move
    # check for http Errors:
    check_badgateway()
    
    # if there are no places - restart
    noGood = "Il n\'existe plus de plage horaire libre pour votre demande de rendez-vous. Veuillez recommencer ultérieurement."
    soup = driver.page_source
    
    if soup.find(noGood):
        print("\nNO TIME SLOTS AVAILABE\n")
        # accept and move the the next page
        link = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="submit_Booking"]/input')))
        time.sleep(random.randint(3, 20))
        link.click()
        print("\nCLICKED BACK TO FIRST PAGE\n")
        loop_reserve() 
    else:
        i = 0 
        while i < 500:
            speak.Speak("APPOINTMENT"*3)
            playsound(alarm)  
            print("\nFOUND - FOUND - FOUND\n"*5)
            i = i + 1
            time.sleep(10)
        
def book():
    try:
        driver.get("http://pprdv.interieur.gouv.fr/booking/create/953")
        check_badgateway()
        cookies()
        loop_reserve()
    finally:
        #speak.Speak(" CODE CRASHED "*3)
        print("\nCODE CRASHED\n")
        print(datetime.datetime.now(),"\n")
# Let's find an appointment
# -------------------------------------------------------------------------------------------------------------------------------------
book() 