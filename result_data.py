'''
Points to Know:
1. This is my personal project & it is designed to work on my college's website only.
2. If you copy & paste it, it won't work.
3. If you want to do something similar, you can take inspiration from this project.
4. Every library & package used in this project is free & open source.
5. All documentations about these libraries are available online for free.
'''

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import numpy as np


class Student:
    roll_no = list(range(start_number, end_number))  # Specify roll no. range.
    student_name = []


class Marks(Student):
    student_percentage = []
    student_CGPA = []
    student_SGPA = []


def enterData(driver, roll):
    # Location of Roll no. input box on website
    roll_box = driver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    roll_box.send_keys(roll)  # Enters roll number

    # Location of Submit button on website
    submit_button = driver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    submit_button.click()  # Clicks submit button
    return

def storeData(driver, s):
    # Location of Name box on website
    name_box = driver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    s.student_name.append(name_box.text)  # Reads & stores name

    # Location of SGPA box on website
    sgpa_box = driver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    s.student_SGPA.append(float(sgpa_box.text))  # Reads & stores SGPC

    # Location of CGPA box on website
    cgpa_box = driver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    s.student_CGPA.append(float(cgpa_box.text))  # Reads & stores CGPA

    # Location of Percentage box on website
    percentage_box = driver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    s.student_percentage.append(percentage_box.text)  # Reads & stores Percentage
    return


def make_df(s) -> pd.DataFrame:
    # Creating a Dictionary of all student's data
    student_df =  {"Name": s.student_name,
                   "PRN No.": s.roll_no,
                   "CGPA": s.student_CGPA,
                   "SGPA": s.student_SGPA,
                   "Percentage": s.student_percentage
                   }
    # Returning DataFrame of the dictionary
    return pd.DataFrame(student_df)


# main()
s = Marks  # object of class marks

webdriver = webdriver.Chrome()  # choosing webdriver
webdriver.get("Website_address")  # Getting Result website

captcha = input("Enter Captcha : ")
# if there is no captcha, good
# in case of our college's result website, i was able to do it with single captcha
# if captcha changes every time, you can use captcha solving api's (Like 2captcha), but you'll have to pay to use them.
# Or you can enter every captcha yourself (Good luck with that)

# Location of captcha box
captcha_box = webdriver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
captcha_box.send_keys(str(captcha))  # Entering captcha

for roll in s.roll_no:
    # Entering & storing data of each student
    enterData(webdriver,roll)
    storeData(webdriver, s)

    webdriver.back()  # going to back page

    roll_box = webdriver.find_element(By.XPATH, "Xpath_of_element_copied_from_chrome")
    roll_box.clear()  # claering roll no. box

    time.sleep(0.2)  # 0.2 sec delay before next entry

# storing the DataFrame returned by make_df()
student_result_data = make_df(s)
# Initializing DataFrame index to 1
student_result_data.index = np.arange(1,len(student_result_data) + 1)

# Writing DataFrame to Excel
student_result_data.to_excel("your_filename")