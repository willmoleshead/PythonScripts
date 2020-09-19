from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from datetime import date
import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


driver = webdriver.Firefox()
driver.set_window_size(1804, 2000)

driver.get('https://skiworld.ontigerbay.co.uk/admin/')
driver.find_element_by_id("Username").clear()
driver.find_element_by_id("Username").send_keys("pl2@skiworld.co.uk")
driver.find_element_by_id("Password").clear()
driver.find_element_by_id("Password").send_keys("P@55word")
driver.find_element_by_name("button").click()

conn = pyodbc.connect('Driver={SQL Server};''Server=SKW-TBSQL;''Database=BIA_REPORTS_2019v2;''Trusted_Connection=yes;')
cursor = conn.cursor()

rows = cursor.execute("""SELECT ID, [COMP ID] FROM dbo.NEW_CancelledComponentsWithCosts""")
for row in rows:
    ReservationID = str(row[0])
    ReservationComponentID = str(row[1])
    URL = "https://skiworld.ontigerbay.co.uk/admin/reservations/accounts/fComponentDetails.aspx?ReservationID=" + ReservationID + "&ReservationComponentID=" + ReservationComponentID
    driver.get(URL)
    wait = WebDriverWait(driver, 60)
    element = wait.until(ec.visibility_of_element_located((By.XPATH, '//*[@id="cph_main_cph_main_btnSummary"]')))
    MoveOn = False
    attempts = 0
    while not MoveOn:
        try:
            driver.find_element_by_id("cph_main_cph_main_rptCostRows_lnkDelete_0").click()
            alert_obj = driver.switch_to.alert
            alert_obj.accept()
            import time
            time.sleep(2)
        except Exception:
            MoveOn = True
    print("Finished - ", URL)
print("Finsihed - Components")

rows = cursor.execute("""SELECT ID, [Supp ID] FROM dbo.NEW_CancelledSupplementWithCosts""")
for row in rows:
    ReservationID = str(row[0])
    ReservationComponentID = str(row[1])
    URL = "https://skiworld.ontigerbay.co.uk/admin/reservations/accounts/fComponentDetails.aspx?ReservationID=" + ReservationID + "&ReservationComponentID=" + ReservationComponentID
    driver.get(URL)
    wait = WebDriverWait(driver, 60)
    element = wait.until(ec.visibility_of_element_located((By.XPATH, '//*[@id="cph_main_cph_main_btnSummary"]')))
    MoveOn = False
    attempts = 0
    while not MoveOn:
        try:
            driver.find_element_by_id("cph_main_cph_main_rptCostRows_lnkDelete_0").click()
            alert_obj = driver.switch_to.alert
            alert_obj.accept()
            import time
            time.sleep(2)
        except Exception:
            MoveOn = True
    print("Finished - ", URL)
print("Finsihed - Supplements")


driver.quit()

recipients = ['wmoleshead@skiworld.co.uk', 'dpaulley@skiworld.co.uk']
fromaddr = 'products@skiworld.co.uk'
msg = MIMEMultipart('alternative')
msg['From'] = fromaddr
msg['To'] = ", ".join(recipients)
today = date.today()
d1 = today.strftime("%d/%m/%Y")
subject = "Removing Costs " + d1
msg['Subject'] = subject
text = "Removing Costs"
html = """<p>Hi all,</p>
<p>Removing costs from cancelled components has been run.</p>
<p>Thanks,</p>
<p>The Product Team</p>
"""
part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')
msg.attach(part1)
msg.attach(part2)
s = smtplib.SMTP('webmail.skiworld.co.uk', 587)
s.starttls()
s.login('skiworld-hq/products', "123sk!w0rld")
text = msg.as_string()
s.sendmail(fromaddr, recipients, msg.as_string())
s.quit()
