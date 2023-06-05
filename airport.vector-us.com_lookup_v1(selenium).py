#! python3
#airport.vector-us.com_Lookup.py
# command line or clipboard

#BEFORE USE:
#1)close the excel workbook before running
#2)make sure workbook is in the same folder as the program file
#3)make sure pieces of data being taken from the excel workbook are values (not formulas)

from selenium import webdriver
from selenium.webdriver.common.by import By
import webbrowser, sys, pyperclip, openpyxl, time

start_time = time.time()
workbook = openpyxl.load_workbook(input("Enter workbook name (incl .xlsx extension):\n")) #Requests_20220207.xlsx
worksheet = workbook[input("Enter worksheet name:\n")] #Request_Ops, Request_Complaints
s, e, column = [input("Enter the starting row:\n"), input("Enter the ending row:\n"), input("Enter the column letter :\n")]
column = column.upper()
s = int(s) - 1
e = int(e)

# initiate
driver = webdriver.Chrome()  # initiate a driver, chrome
driver.get('https://airport.vector-us.com/login.aspx?ReturnUrl=%2fAirport%2fAircraftSearch.aspx')  # go to the url
# locate the login form
username_field = driver.find_element(By.NAME, "Login2$UserName")  # get the username field
password_field = driver.find_element(By.NAME, "Login2$Password")  # get the password field
login_field = driver.find_element(By.NAME, "Login2$LoginButton")
# log in
username_field.send_keys("Adam Scholten")  # enter username
password_field.send_keys("DWCatc2010**")  # enter password
login_field.click()  # click login
time.sleep(3)


print("Results:\n")
while s<e:
    s = s + 1
    j = str(s)

    cellnum = column + j
    cellvalue = worksheet[cellnum].value
    cellvalue = str(cellvalue)

    # pyperclip.copy(cellvalue)
    #
    # if len(sys.argv) > 1:
    #     #get address from command line.
    #     aircraftnum = ''.join(sys.argv[1:])
    # else:
    #     #get address from clipboard.
    #     aircraftnum = pyperclip.paste()

    url = 'https://airport.vector-us.com/Airport/AircraftDetailsPage.aspx?aircraftnumber=' + cellvalue

    # #opens the aircraft lookup
    # webbrowser.open(url)

    driver.get(url)
    mtow_element = driver.find_element(By.CSS_SELECTOR,"span[ID='ContentPlaceHolder1_mtw']")
    mtow_element.click()
    mtow = mtow_element.text
    print(mtow)

print ("Total run time", time.time() - start_time, "seconds")




