import win32timezone
import win32com.client
import datetime
import time
from bs4 import BeautifulSoup
import os
from SnowDaemon.config import *

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from SnowDaemon.config import mail_box
from SnowDaemon.web.Snow.my_snow import my_snow

outlook = win32com.client.Dispatch("Outlook.Application")

# Set the initial timestamp to the current time
num_days = 1
#lastDayDateTime = datetime.now() - timedelta(days=int(num_days))
last_email_time = datetime.datetime.now() - datetime.timedelta(days=int(num_days))

while True:
    try:
        # Connect to Outlook and access shared mailbox
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # Get a reference to the shared mailbox folder
        #shared_mailbox = outlook.Session.GetSharedDefaultFolder(mail_box, 6)
        # shared_mailbox = outlook.Folders(mail_box).GetDefaultFolder("6")
        # get the account object
        for account_name in outlook.Accounts:
            if mail_box in account_name.DisplayName:
                account_obj = account_name
                print("Found...." + str(mail_box))
                break

        # get the default inbox folder for the account
        shared_mailbox = account_obj.DeliveryStore.GetDefaultFolder(6)

        while(True):
            # Calculate the start and end times for the time frame (last_email_time to now)
            start_time = last_email_time
            end_time = datetime.datetime.now()
            # Use the Items.Restrict() method to get only the items within the time frame
            filter_str = "[ReceivedTime] >= '" + start_time.strftime('%m/%d/%Y %H:%M %p') + "' AND [ReceivedTime] <= '" + end_time.strftime('%m/%d/%Y %H:%M %p') + "'"
            items = shared_mailbox.Items.Restrict(filter_str)

            # Loop through the items in the time frame and print the subject
            print("Reviewing mails...." + str(len(items)))
            for message in items:
                if message.Class == 43 and message.Subject == "Acción requerida: Solicitud de acceso pendiente de aprobación / Action Required: Access request awaiting approval":
                    html_body = BeautifulSoup(message.HTMLBody, "html.parser")
                    paragraphs = html_body.findAll('p')
                    id_request = str(paragraphs[3].text).split(": ")[-1].strip()
                    id_user = str(paragraphs[4].text).split(": ")[-1].strip()
                    action_request = str(paragraphs[5].text).split(": ")[-1].strip()

                    print("email to Snow...." + str(id_request))
                    # Try to acquire a lock for this email
                    lock_file_path = f"email_locks/{id_request}.lock"
                    try:
                        lock_file = open(dwnload_path + lock_file_path, 'x')
                        message.SaveAs(dwnload_path + f"email_locks/{id_request}.msg")
                        # If lock is acquired, process the email
                        mSnw = my_snow()
                        id_req = mSnw.do_Send_Snow(id_request, id_user, action_request, dwnload_path + f"email_locks/{id_request}.msg")
                        print("Remedy generated: " + str(id_req))
                        # Clear email saved
                        try:
                            if os.path.exists(dwnload_path + f"email_locks/{id_request}.msg"): os.remove(dwnload_path + f"email_locks/{id_request}.msg")
                        except Exception as e:
                            pass
                        del mSnw
                        # Release the lock
                        lock_file.close()
                        # os.remove(lock_file_path)
                    except FileExistsError:
                        # Lock file already exists, another instance of the program is processing this email
                        pass

            # If we processed any new emails, update the last_email_time to the most recent one
            if items:
                last_email_time = max([email.ReceivedTime for email in items])

            # Pause the program for 5 minutes before running again
            print("Sleeping...")
            time.sleep(sleep_time)

    except Exception as e:
        print(f"Error: {e}")

