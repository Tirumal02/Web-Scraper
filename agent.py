import time
import smtplib
import pandas as pd
from tabulate import tabulate
from selenium import webdriver
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from email.mime.multipart import MIMEMultipart
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class web_scraper:
    def __init__(self):
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--start-maximized")
        options.add_experimental_option("prefs", {
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_setting_values.popups": 2
        })
        service = Service(r"C:\Users\acer\chromedriver-win64\chromedriver.exe")
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 10)

    def is_session_valid(self):
        try:
            _ = self.driver.title
            return True
        except:
            return False

    def ScrapeTgsrtcData(self, origin, destination, xl_date):
        self.origin = origin
        self.destination = destination
        self.xl_date = xl_date

        if not self.is_session_valid():
            raise Exception("WebDriver session is closed.")

        try:
            self.driver.get("https://www.tgsrtcbus.in")
            time.sleep(3)
            org = self.wait.until(EC.element_to_be_clickable((By.ID, "rc_select_0")))
            org.send_keys(self.origin)
            org.send_keys(Keys.ENTER)
            time.sleep(1)
            dest = self.wait.until(EC.element_to_be_clickable((By.ID, "rc_select_1")))
            dest.send_keys(self.destination)
            dest.send_keys(Keys.ENTER)
            time.sleep(1)

            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ant-picker-content")))
            self.TgsrtcPickDate()

            search = self.driver.find_element(By.ID, "gt-search")
            search.send_keys(Keys.ENTER)
            time.sleep(5)

            self.services = self.driver.find_element(By.CLASS_NAME, "child2")
            time.sleep(2)
            df = self.text_to_csv(self.services.text)
            return df
        except Exception as e:
            raise e

    def TgsrtcPickDate(self):
        try:
            while True:
                # Wait until month and year elements are non-empty
                web_month_driver = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ant-picker-month-btn")))
                web_year_driver = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ant-picker-year-btn")))

                web_month_text = web_month_driver.text.strip()
                web_year_text = web_year_driver.text.strip()

                if not web_month_text or not web_year_text:
                    time.sleep(0.5)
                    continue

                web_month = datetime.strptime(web_month_text, "%b").month
                web_year = int(web_year_text)

                if self.xl_date.year == web_year:
                    if self.xl_date.month == web_month:
                        date_elem = self.driver.find_element(By.XPATH, f"//table[@class='ant-picker-content']//td[@title='{self.xl_date}']")
                        date_elem.click()
                        break
                    elif self.xl_date.month < web_month:
                        self.driver.find_element(By.CLASS_NAME, "ant-picker-header-prev-btn").click()
                    else:
                        self.driver.find_element(By.CLASS_NAME, "ant-picker-header-next-btn").click()
                else:
                    self.driver.find_element(By.CLASS_NAME, "ant-picker-header-super-next-btn").click()
        except Exception as e:
            raise e


    def text_to_csv(self, text):
        service_nos = []
        lines = text.strip().split('\n')
        for line in lines:
            if line.isnumeric():
                service_nos.append(line)
        df = pd.DataFrame({'Service No': service_nos})
        return df

    def QuitDriver(self):
        try:
            self.driver.quit()
        except:
            pass

class SendEmail:
    def __init__(self, missing_df, extra_df, start_time, end_time):
        if (missing_df is not None and not missing_df.empty) or (extra_df is not None and not extra_df.empty):
            self.body = ""
            if missing_df is not None and not missing_df.empty:
                self.body += "MISSING SERVICES:\n\n" + tabulate(missing_df,missing_df.columns,tablefmt="html") + "\n\n"
                print(tabulate(missing_df.iloc[1:],missing_df.iloc[0].to_list(),tablefmt="html"))
            else:
                self.body += "MISSING SERVICES:\n\nNone\n\n"

            if extra_df is not None and not extra_df.empty:
                self.body += "EXTRA SERVICES:\n\n" + tabulate(extra_df,extra_df.columns,tablefmt="html")
            else:
                self.body += "EXTRA SERVICES:\n\nNone"
        else:
            self.body = "NO MISSING SERVICES"

        self.body += f"\n\n Start Time: {start_time} \n End Time: {end_time} \n Total Duration: {end_time - start_time}"
        self.sender_email = "tgsrtcaiagent@gmail.com"
        self.receiver_email1 = "ame3it@gmail.com"
        self.receiver_email2 = "oprsho@gmail.com"
        self.receiver_email3 = "tirumalmirdoddi@gmail.com"
        self.sub = f"SUBJECT: MISSING AND EXTRA SERVICES DATA FROM '{(datetime.today() + timedelta(days=1)).date()}' TO '{(datetime.today() + timedelta(days=3)).date()}'"
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587
        self.pwd = "hhfzailvzxxyeind"

    def email(self):
        message = MIMEMultipart()
        message["From"] = self.sender_email
        message["To"] = ", ".join([self.receiver_email1, self.receiver_email2, self.receiver_email3])
        message["Subject"] = self.sub
        message.attach(MIMEText(self.body, "html"))

        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.sender_email, self.pwd)
            server.sendmail(self.sender_email, self.receiver_email1, message.as_string())
            server.sendmail(self.sender_email, self.receiver_email2, message.as_string())
            server.sendmail(self.sender_email, self.receiver_email3, message.as_string())
            server.quit()
            print(self.body)
            print("✅ Email sent successfully.")
        except Exception as e:
            print("❌ Failed to send email:", e)

def MissingServices(web_df, original_df, org, dest,date):
    web_df['Source'] = org
    web_df['Destination'] = dest
    web_df['Service No'] = web_df['Service No'].astype(int)
    filtered_df = original_df[(original_df['Source']==org) & (original_df['Destination']==dest)]
    missing_df = filtered_df[~filtered_df['Service No'].isin(web_df['Service No'])]
    extra_df = web_df[~web_df['Service No'].isin(original_df['Service No'])]
    extra_df = extra_df.drop_duplicates(subset='Service No')
    if not missing_df.empty:
        missing_df['Date'] = date

    if not extra_df.empty:
        extra_df['Date'] = date

    return missing_df, extra_df

def main():
    original_df = pd.read_excel(r"C:\Users\acer\Desktop\AI_AGENT\Service_data.xlsx")
    original_df['Service No'] = original_df['Service No'].astype(int)
    org_dest_df = original_df[['Source', 'Destination']].drop_duplicates()
    org_list = org_dest_df['Source'].to_list()
    dest_list = org_dest_df['Destination'].to_list()
    start_date = datetime.today()
    date_list = [start_date + timedelta(days=i) for i in range(1, 4)]

    missing_dfs = []
    extra_dfs = []

    data_obj = web_scraper()
    start_time = datetime.now()
    for i in date_list:
        date = i.date()
        for origin, destination in zip(org_list, dest_list):
            retry = True
            while retry :
                try:
                    print(f"🔍 Checking {origin} → {destination} for {date}")
                    if date > datetime.today().date():
                        Tgsrtc_df = data_obj.ScrapeTgsrtcData(origin, destination, date)
                        missing_df, extra_df = MissingServices(Tgsrtc_df, original_df,origin, destination, date)
                        if not missing_df.empty:
                            missing_dfs.append(missing_df)
                            print("MISSING SERVICES:\n", missing_df.to_string(index=False))
                        else:
                            print("MISSING SERVICES:\nEMPTY")
                        
                        if not extra_df.empty:
                            print("EXTRA SERVICES:\n", extra_df.to_string(index=False))
                            extra_dfs.append(extra_df)
                        else:
                            print("MISSING SERVICES:\nEMPTY")
                    else:
                        raise Exception("Invalid Date!")
                    retry = False
                except Exception as e:
                    print(f"❌ Error for {origin} → {destination} ({date}): {e}")
                    time.sleep(2)
                    print("🔁 Restarting browser...")
                    try:
                        data_obj.QuitDriver()
                    except:
                        pass
                    data_obj = web_scraper()

    try:
        data_obj.QuitDriver()
    except:
        pass

    end_time = datetime.now()
    print(f"START Time : {start_time}")
    print(f"End_Time: {end_time}")
    print(f"Total Duration: {end_time - start_time}")
    final_missing_df = pd.concat(missing_dfs, ignore_index=True) if missing_dfs else pd.DataFrame()
    final_extra_df = pd.concat(extra_dfs, ignore_index=True) if extra_dfs else pd.DataFrame()

    send_mail = SendEmail(final_missing_df, final_extra_df, start_time, end_time)
    send_mail.email()


if __name__ == "__main__":
    main()
