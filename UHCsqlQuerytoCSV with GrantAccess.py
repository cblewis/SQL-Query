from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.alert import Alert
import pyodbc
import csv

#Masterfully crafted by Cameron Lewis
#userNamme = f"dai\\cameron.lewis"
userNamme = input("Enter your DAI or ODS account name, including the DAI\\ or ODSDAI\\: ")
passWurd = getpass.getpass(prompt='What is your ODS Grant Access Password?: ')
print('"""""""""""\n Please wait while your access is granted. You should receive your usual notification email if this is successful.\n"""""""""""\n')

firefox_options = Options()
firefox_options.add_argument("--headless")
driver = webdriver.Firefox(options=firefox_options)
driver.get("http://odssysmonweb.odsdai.netdai.com/grantaccess")
wait = WebDriverWait(driver, 10)
WebDriverWait(driver, 1).until(ec.alert_is_present())
alert = Alert(driver)
#alert.send_keys(f"dai\\cameron.lewis{Keys.TAB}{input('What is your ODS Grant Access Password?')}{Keys.ENTER}")
alert.send_keys(f"{userNamme}{Keys.TAB}{passWurd}{Keys.ENTER}")
actions = ActionChains(driver)


#Server:
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[1]/div/div/button"))).click()
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[1]/div/div/div/div[1]/input"))).send_keys("ODSUNITEDDBLS")
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[1]/div/div/div/ul/li[253]/a"))).send_keys(Keys.ENTER)
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[1]/div/div/div/ul/li[253]/a"))).send_keys(Keys.TAB)

#Products:
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[2]/div/div"))).click()
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[2]/div/div/div/div[1]/input"))).send_keys("OTHER")
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[2]/div/div/div/ul/li[16]/a"))).send_keys(Keys.ENTER, Keys.TAB)

#Reason:
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[3]/div/div/button"))).click()
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[3]/div/div/div/div[1]/input"))).send_keys("PRODUCTION SUPPORT")
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[3]/div/div/div/ul/li[9]/a"))).send_keys(Keys.ENTER, Keys.TAB)

#Duration:
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[4]/div/div/button"))).click()
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[4]/div/div/div/ul/li[4]/a"))).click()

#Access Level:
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[5]/div/div[2]/button"))).click()
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[5]/div/div[2]/div/ul/li[2]/a"))).click()

#Comment:
commentbox = driver.find_element_by_id("Comment")
#commentbox = wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[7]/div")))
commentbox.click()
commentbox.send_keys("Checking UHC Jobs")

#Submit:
wait.until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[1]/div[2]/div/form/div[9]/div/button"))).click()

driver.save_screenshot("grantAccessODS.png")

#Query SQL:
connection = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=ODSUNITEDDBLS;"
                      "Database=XLHEALTH;"
                      "Trusted_Connection=yes;")

cursor = connection.cursor()
cursor.execute('''
SELECT m.companyid, JN = i.JobNumber, Start_ImpID = MIN(i.ImportFileID), End_ImpID = MAX(i.ImportFileID), MemberCount=count (*)
FROM Members m
INNER JOIN Member_Raws r ON r.MemberRecID = m.MemberRecID
INNER JOIN ImportFile i ON i.ImportFileID = m.ImportFileID AND i.ImportDate >= convert(varchar(10),CAST(CONVERT(varchar,GetDate(),23) as DateTime),112)
INNER JOIN ImportFile oi ON oi.ImportFileID = m.oImportFileID AND oi.ImportDate >= convert(varchar(10),CAST(CONVERT(varchar,GetDate(),23) as DateTime),112)
WHERE oi.ImportFileID <> 0
GROUP BY m.companyid, i.JobNumber, i.ImportFileID, i.ImportDate
ORDER BY i.jobnumber, m.companyid
''')
result = cursor.fetchall()

column_names = [i[0] for i in cursor.description]
fp = open("UHCDailyJobs.csv", "w")
myFile = csv.writer(fp, lineterminator = '\n')
myFile.writerow(column_names+["Status"])
myFile.writerows(result)
fp.close()
#print('result = %r' % (result))
