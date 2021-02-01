# Load needed packages:
# -------------------------------------------------------------------------------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.common.keys import Keys # for interacting with page (clicks, etc.)
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
# for openning sound file from computer -------------------> download an alarm sound of your choosing and change the path
import os
alarm = "C:\\Users\\meita\\Documents\\auto reservation\\fire.wav"
# for opening the web browser -------------------> download the 'chromedriver' that matches your chrome version and change path
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
    element.click() # mouse click to accept cookies and close the pop-up bar
    print("cockies")

def loop_reserve():
    print("SCANNING: FIRST PAGE") # looking at the main page with all the details about the meeting you want to set
    # check for http Errors
    check_badgateway()

    # click checkbox to accept conditions - buttom left of page
    element = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="condition"]')))
    element.click()
    print("CLICKED: CHECKBOX")

    # accept and move the the next page - where you would book if there were vacant time slots
    link = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="submit_Booking"]/input[1]')))
    link.click()
    print("CLICKED: NEXT PAGE")

    # check if the page didn't crash during move
    # check for http Errors:
    check_badgateway()
    
    # if there are no vacant time slots - start over
    noGood = "Il n\'existe plus de plage horaire libre pour votre demande de rendez-vous. Veuillez recommencer ultérieurement." # the text that appears one there are no openings  
    soup = driver.page_source
    
    if soup.find(noGood):
        print("\nNO TIME SLOTS AVAILABE\n")
        # accept and move the the next page
        link = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="submit_Booking"]/input')))
        time.sleep(random.randint(3, 20)) # be kind and wait - to not overload the server
        link.click()
        print("\nCLICKED BACK TO FIRST PAGE\n")
        loop_reserve() 
    else: # if found booking option
        i = 0 
        while i < 500:
            speak.Speak("APPOINTMENT"*3) # sound indication  
            playsound(alarm)  
            print("\nFOUND - FOUND - FOUND\n"*5) # text indication
            i = i + 1
            time.sleep(10) # give option to shut down code while typing in the computer (the noise is delibertly annoying)
        
def book(): # the Main function tying together all sub functions
    try:
        driver.get("http://pprdv.interieur.gouv.fr/booking/create/953") # -------------------> change to the link to the depertment and service that are intresting for you
        check_badgateway()
        cookies()
        loop_reserve()
    finally:
        speak.Speak(" CODE CRASHED "*3) # sound indication
        print("\nCODE CRASHED\n") # text indication
        print(datetime.datetime.now(),"\n") # know when the code crashed / finished
# Let's find an appointment
# -------------------------------------------------------------------------------------------------------------------------------------
book() # runs the main function
