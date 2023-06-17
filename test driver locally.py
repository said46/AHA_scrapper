from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import InvalidSelectorException
import ctypes
import time

def message_box(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

#  Styles:
#  0 : OK *** #  1 : OK | Cancel *** 2 : Abort | Retry | Ignore *** 3 : Yes | No | Cancel ***  4 : Yes | No *** 5 : Retry | Cancel *** 6 : Cancel | Try Again | Continue
 
 
# edgeBrowser = webdriver.Edge(r"msedgedriver.exe")
edgeBrowser: WebDriver = webdriver.Chrome()
edgeBrowser.maximize_window()
 
path = "C:/Users/alexander.sverchkov/PyProjects/Get Tag Links from AHA/test_html/asp/tree full.html"
edgeBrowser.get(path)

doc_number = ''

element_xpath = "//a[contains(.,'LOOP')]"

try:
    elems = edgeBrowser.find_elements(By.XPATH, element_xpath)
    for e in elems:
        print(e.text)
        message_box('element text', e.text, 0)
    message_box('len of elems', str(len(elems)), 0)
except NoSuchElementException:
    message_box('Error', 'NoSuchElementException: ' + element_xpath, 0)
except InvalidSelectorException:
    message_box('Error', 'InvalidSelectorException: ' + element_xpath, 0)

message_box('Ok?', 'Ok!', 0)

