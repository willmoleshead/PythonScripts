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

print('Starting Email Queue')


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


UpdateSQL = ConnectRunStoredProcedure('Update Email Queue', 'EXEC dbo.EmailQueue')

conn = pyodbc.connect('Driver={SQL Server};' 'Server=SKW-TBSQL;''Database=BIA_REPORTS_2019v2;' 'Trusted_Connection=yes;')
cursor = conn.cursor()

cursor.execute('SELECT COUNT(*) FROM BIA_REPORTS_2019v2.base.emailqueue')
rowcount = cursor.fetchone()
print(str(rowcount[0]), 'emails in queue.')
rowcountvalue = str(rowcount[0])

if rowcount[0] < 1:
    recipients = ['wmoleshead@skiworld.co.uk', 'dpaulley@skiworld.co.uk', 'ops@skiworld.co.uk', 'cwilloughby@skiworld.co.uk']
    fromaddr = 'ops@skiworld.co.uk'
    msg = MIMEMultipart('alternative')
    msg['From'] = fromaddr
    msg['To'] = ", ".join(recipients)
    subject = "Email Queue Processed - " + datetime.today().strftime('%d-%m-%Y @ %H:%M')
    msg['Subject'] = subject
    text = "Email Queue Processed"
    html = """<p>Hi all,</p>
    <p>The email queue has processed.</p>
    <p>No emails were in the queue.</p>
    <p>Thanks,</p>
    <p>The Ops Team</p>"""
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    msg.attach(part1)
    msg.attach(part2)
    s = smtplib.SMTP('webmail.skiworld.co.uk', 587)
    s.starttls()
    s.login('skiworld-hq/ops', "m0rh0lT")
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
    rows = cursor.execute('SELECT * FROM BIA_REPORTS_2019v2.base.emailqueue')

    for row in rows:
        URL = "https://skiworld.ontigerbay.co.uk/admin/utilities/emailer/fView.aspx?EmailID=" + str(row[0])
        Email = str(row[1]) + ' - ' + str(row[2])
        driver.get(URL)
        wait = WebDriverWait(driver, 60)
        element = wait.until(ec.visibility_of_element_located((By.XPATH, '//*[@id="cph_main_cph_main_btnSend"]')))
        driver.find_element_by_xpath('//*[@id="cph_main_cph_main_btnSend"]').click()
        try:
            wait = WebDriverWait(driver, 60)
            element = wait.until(ec.visibility_of_element_located((By.XPATH, '//*[@id="panMessage"]')))
            EmailsSent += """<tr><td style="color:green">""" + Email + """</td></tr>"""
            print("Finished - ", Email, end='', flush=True)
            time.sleep(1)
            print("...Waiting 20 secs")
            time.sleep(19)
        except Exception as e:
            print(Email, 'failed - ', e)
            EmailsSent += """<tr><td style="color:red">""" + Email + """ Failed</td></tr>"""
            pass
    conn.close()

    recipients = ['wmoleshead@skiworld.co.uk', 'dpaulley@skiworld.co.uk', 'ops@skiworld.co.uk', 'cwilloughby@skiworld.co.uk']
    fromaddr = 'ops@skiworld.co.uk'
    msg = MIMEMultipart('alternative')
    msg['From'] = fromaddr
    msg['To'] = ", ".join(recipients)
    subject = "Email Queue Processed - " + datetime.today().strftime('%d-%m-%Y @ %H:%M')
    msg['Subject'] = subject
    text = "Email Queue Processed"
    html = """<p>Hi all,</p>
    <p>The email queue has processed.</p>
    <p>The following """ + rowcountvalue + """ emails have been processed. Green = sent. Red = issue, check email queue.</p><p><table>""" + EmailsSent + """</table></p>
    <p>Thanks,</p>
    <p>The Ops Team</p>"""
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    msg.attach(part1)
    msg.attach(part2)
    s = smtplib.SMTP('webmail.skiworld.co.uk', 587)
    s.starttls()
    s.login('skiworld-hq/ops', "m0rh0lT")
    text = msg.as_string()
    s.sendmail(fromaddr, recipients, msg.as_string())
    s.quit()
    driver.quit()
