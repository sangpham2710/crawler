from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import psutil
import time
import openpyxl
import csv

timeout = 30
numOfStudents = 508

for process in psutil.process_iter():
    if (process.name() == "chromedriver.exe"):
        process.terminate()

driver = webdriver.Chrome(executable_path=r"C:/chromedriver.exe")
driver.get("https://www.ivyprephcm.edu.vn/traketqua/")
with open("tmp.csv", "w", newline="", encoding="utf-8") as file:
    csvWriter = csv.writer(file)
    csvWriter.writerow(["id", "name", "listening", "reading",
                        "writing", "speaking", "overall"])
    for i in range(1, numOfStudents + 1):
        idx = str(i).zfill(3)
        # check presence
        try:
            element_present = EC.presence_of_element_located(
                (By.NAME, "ivy_numberid"))
            WebDriverWait(driver, timeout).until(element_present)
        except TimeoutException:
            print("Timed out waiting for page to load")
        inputBox = driver.find_element_by_name("ivy_numberid")
        inputBox.send_keys(idx)
        submitButton = driver.find_element_by_id("submit_test_result")
        submitButton.click()

        loaderBox = driver.find_element_by_class_name("loader-box")

        while loaderBox.get_attribute("style") == "display: block;":
            pass
        N = driver.find_element_by_css_selector(
            "h2.result.result-fullname").get_attribute("innerHTML")
        L = driver.find_element_by_css_selector(
            "td.result.result-listening").get_attribute("innerHTML")
        R = driver.find_element_by_css_selector(
            "td.result.result-reading").get_attribute("innerHTML")
        W = driver.find_element_by_css_selector(
            "td.result.result-writing").get_attribute("innerHTML")
        S = driver.find_element_by_css_selector(
            "td.result.result-speaking").get_attribute("innerHTML")
        O = driver.find_element_by_css_selector(
            "div.result.result-overall").get_attribute("innerHTML")
        print(N, i, L, R, W, S, O)
        csvWriter.writerow([str(i), str(N), str(
            L), str(R), str(W), str(S), str(O)])

driver.quit()
