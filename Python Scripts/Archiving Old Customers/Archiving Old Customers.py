import pyodbc
from selenium import webdriver
import time
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec


def ConnectRunStoredProcedure(name, storedprocedure):
    # This function connects to SKW SQL and excutes a stored procedure.
    # The varibles are -
    # Name - the names of the process
    # Storedprocedure - a SQL query to execute a stored procedure e.g. 'EXEC dbo.FlatFileAvailabilityUpdate'
    try:
        start = int(round(time.time() * 1000))
        conn = pyodbc.connect('Driver={SQL Server};' 'Server=SKW-TBSQL;''Database=BIA_REPORTS_2019v2;' 'Trusted_Connection=yes;')
        cursor = conn.cursor()
        cmd_prod_executesp = storedprocedure
        conn.autocommit = True
        cursor.execute(cmd_prod_executesp)
        conn.close()
        end = int(round(time.time() * 1000))
        duration = (end - start) / 1000
        endstatement = name + ' - completed successfully in ' + str(duration) + 'secs'
    except Exception as e:
        endstatement = name + ' - failed with the following error; ' + str(e)
    return endstatement


HTMLResults = ''
ArchiveCustomersSQLUpdate = ConnectRunStoredProcedure('Archiving Customers SQL Update', 'EXEC dbo.ArchiveCustomers')
print(ArchiveCustomersSQLUpdate)
if 'failed' in ArchiveCustomersSQLUpdate:
    HTMLResults += """<p style="color:red">""" + ArchiveCustomersSQLUpdate + """</p> """
else:
    HTMLResults += """<p style="color:green">""" + ArchiveCustomersSQLUpdate + """</p> """


conn = pyodbc.connect('Driver={SQL Server};' 'Server=SKW-TBSQL;''Database=BIA_REPORTS_2019v2;' 'Trusted_Connection=yes;')
cursor = conn.cursor()

cursor.execute('SELECT COUNT(*) FROM BIA_REPORTS_2019v2.dbo.customerstoarchive')
rowcount = cursor.fetchone()
print(str(rowcount[0]), 'customers to archive.')
rowcountvalue = str(rowcount[0])

if rowcount[0] < 1:
    recipients = ['wmoleshead@skiworld.co.uk', 'dpaulley@skiworld.co.uk']
    fromaddr = 'products@skiworld.co.uk'
    msg = MIMEMultipart('alternative')
    msg['From'] = fromaddr
    msg['To'] = ", ".join(recipients)
    subject = "Customer Archiving Processed - " + datetime.today().strftime('%d-%m-%Y @ %H:%M')
    msg['Subject'] = subject
    text = "Customer Archiving Processed"
    html = """<p>Hi all,</p>
    <p>Customer Archiving has been processed.</p><p>""" + HTMLResults + """</p>
    <p>No customers were due to be archived.</p>
    <p>Thanks,</p>
    <p>The C&P Team</p>"""
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
else:
    options = Options()
    options.headless = True
    profile = FirefoxProfile()
    profile.set_preference('browser.cache.disk.enable', False)
    profile.set_preference('browser.cache.memory.enable', False)
    profile.set_preference('browser.cache.offline.enable', False)
    # profile.set_preference('network.cookie.cookieBehavior', 2)

    driver = webdriver.Firefox(firefox_profile=profile, options=options)
    driver.get('https://skiworld.ontigerbay.co.uk/admin/')
    driver.find_element_by_id("Username").clear()
    driver.find_element_by_id("Username").send_keys("pl2@skiworld.co.uk")
    driver.find_element_by_id("Password").clear()
    driver.find_element_by_id("Password").send_keys("P@55word")
    driver.find_element_by_name("button").click()
    driver.set_window_size(1804, 8000)

    EmailsSent = ''
    cursor = conn.cursor()
    rows = cursor.execute('SELECT * FROM BIA_REPORTS_2019v2.dbo.customerstoarchive')

    count = 0

    for row in rows:
        if count < 50:
            URL = "https://skiworld.ontigerbay.co.uk/admin/sales/customers/customer/fArchive.aspx?CustomerID=" + str(row[0])
            Email = str(row[0]) + ' - ' + str(row[1])
            driver.get(URL)
            wait = WebDriverWait(driver, 60)
            element = wait.until(ec.visibility_of_element_located((By.XPATH, '/html/body/form/div[4]/div[2]/div/div[1]/a[1]')))
            try:
                driver.find_element_by_id('cph_main_cph_main_cphMain_btnArchive').click()
                wait = WebDriverWait(driver, 60)
                element = wait.until(ec.visibility_of_element_located((By.XPATH, '//*[@id="panMessage"]')))
                EmailsSent += """<tr><td style="color:green">""" + Email + """</td></tr>"""
                print("Finished - ", Email)
                time.sleep(3)
            except Exception as e:
                print(Email, 'failed - ', e)
                EmailsSent += """<tr><td style="color:red">""" + Email + """ Failed</td></tr>"""
                pass
            count += 1
        else:
            pass
    conn.close()

    recipients = ['wmoleshead@skiworld.co.uk', 'dpaulley@skiworld.co.uk']
    fromaddr = 'products@skiworld.co.uk'
    msg = MIMEMultipart('alternative')
    msg['From'] = fromaddr
    msg['To'] = ", ".join(recipients)
    subject = "Customer Archiving Processed - " + datetime.today().strftime('%d-%m-%Y @ %H:%M')
    msg['Subject'] = subject
    text = "Customer Archiving Processed"
    html = """<p>Hi all,</p>
    <p>Customer Archiving has been processed.</p><p>""" + HTMLResults + """</p>
    <p>The following """ + str(count) + """ customers have been processed. Green = processed. Red = issue, check Tigerbay.</p><p><table>""" + EmailsSent + """</table></p>
    <p>Thanks,</p>
    <p>The C&P Team</p>"""
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
    driver.quit()
