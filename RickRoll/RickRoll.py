from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# RickRoll = "https://www.youtube.com/watch?v=oHg5SJYRHA0&ab_channel=cotter548"


def process_link(choice, link):
    """Disguises Rick Roll Video into a shady or legit URL

    Args:
        choice (int): Link Converter Website Primary Key
        link (string): Link Converter Website
    """
    print("\nPlease wait :D", "\n")

    if (choice == 1):
        driver.get(link)
        url_box = driver.find_element_by_name("u")
        url_box.send_keys(
            "https://www.youtube.com/watch?v=oHg5SJYRHA0&ab_channel=cotter548")
        url_box.send_keys(Keys.RETURN)

        dis_link = driver.find_element_by_id("shortenurl")

        if (dis_link):
            print("This is your new link (Already copied to your clipboard!)", "\n")
            print(dis_link.get_attribute("value"))
            r_link = "https://www." + str(dis_link.get_attribute("value"))
            df = pd.DataFrame([r_link])
            df.to_clipboard(index=False, header=False)
    elif (choice == 2):
        driver.get(link)
        url_box = driver.find_element_by_id("myUrl")
        url_box.send_keys(
            "https://www.youtube.com/watch?v=oHg5SJYRHA0&ab_channel=cotter548")
        url_box.send_keys(Keys.RETURN)

        dis_link = driver.find_element_by_css_selector(
            "#output a:nth-child(2)")

        if (dis_link):
            print("This is your new link (Already copied to your clipboard!)", "\n")
            print(dis_link.get_attribute("href"))
            r_link = "https://www." + str(dis_link.get_attribute("href"))
            df = pd.DataFrame([r_link])
            df.to_clipboard(index=False, header=False)

# Preprocessing


ShortURL = "https://www.shorturl.at/"
ShadyURL = "http://shadyurl.com/"
PATH = "C:/Users/seann/Documents/Python Automation/chromedriver.exe"
op = webdriver.ChromeOptions()
op.add_argument('headless')
driver = webdriver.Chrome(PATH, options=op)
correct_choice = False

while not correct_choice:
    print("1 - Short URL")
    print("2 - Shady URL")
    choice = int(input("Enter your choice: "))

    if (choice == 1):
        process_link(1, ShortURL)
        correct_choice = True
    elif (choice == 2):
        process_link(2, ShadyURL)
        correct_choice = True
    else:
        print("Invalid Input! Please Try Again")

time.sleep(3)

driver.quit()
