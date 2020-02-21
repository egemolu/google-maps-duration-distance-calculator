from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import xlsxwriter
import csv


def search(start, end, worksheet, counter,driver):
    startPoint = driver.find_element_by_xpath(
        '/html/body/jsl/div[3]/div[9]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div/input')

    # Clear the textbox
    startPoint.clear()
    # Entering the starting point.
    startPoint.send_keys(start)

    sleep(1)

    endPoint = driver.find_element_by_xpath(
        '/html/body/jsl/div[3]/div[9]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[2]/div[2]/div/div/input')

    # Clear the textbox
    endPoint.clear()
    # Entering the end point.
    endPoint.send_keys(end)

    sleep(1)

    endPoint.send_keys(Keys.ENTER)

    # Need to wait until results will be shown. Otherwise, Driver cannot find elements by xpath.
    sleep(5)

    duration = driver.find_element_by_xpath(
        '/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[5]/div[1]/div[2]/div[1]/div[1]/div[1]/span[1]')
    distance = driver.find_element_by_xpath(
        '/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[5]/div[1]/div[2]/div[1]/div[1]/div[2]/div')

    # Excel Column Letters.
    startColumn = 'A'
    endColumn = 'B'
    durationColumn = 'C'
    distanceColumn = 'D'

    print("Total Time " + str(duration.text) + " And Total Distance " + str(distance.text))

    worksheet.write("{}{}".format(startColumn, counter), start)
    worksheet.write("{}{}".format(endColumn, counter), end)
    worksheet.write("{}{}".format(durationColumn, counter), duration.text)
    worksheet.write("{}{}".format(distanceColumn, counter), distance.text)

    sleep(2)


def read_csv_file(file_name, worksheet,driver):
    with open(file_name, 'r') as file:
        reader = csv.reader(file)

        # Row index that address will be written in the excel file.
        rowCounter = 2

        for row in reader:
            search(row[0], row[1], worksheet, rowCounter,driver)
            rowCounter += 1;


def main():
    # Initialize WebDriver. You can choose different browsers like Firefox, Mozilla.
    driver = webdriver.Chrome()

    # Open the Google Maps website.
    driver.get("https://www.google.com/maps/dir///@41.035154,29.2669434,15z/data=!4m2!4m1!3e0")

    # Enter the name of excel file that you want to create.
    # It must be in that format "name.xlsx".
    workbook = xlsxwriter.Workbook('results.xlsx')

    # You need to create new worksheet.
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Start Point')
    worksheet.write('B1', 'End Point')
    worksheet.write('C1', 'Duration')
    worksheet.write('D1', 'Distance')

    # Enter the name of your csv file as a first parameter of read_csv_file method.
    # If csv file does not in the same directory, you need to enter it's location like /Downloads/xx/abc.csv
    read_csv_file('test.csv', worksheet,driver)

    # You need to close your workbook. Otherwise, it will not be saved.
    workbook.close()


main()
