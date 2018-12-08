# Documentation for Microsoft API - https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitemclass.body?view=outlook-pia

from outlook import Outlook
import time
import sys

# Creating an object of the Outlook class.
outlook = Outlook()
while True:
    try:
        for subfolder in outlook.folders():
            print('Checking subfolder : '+subfolder)
            mails = outlook.fetchMails(subfolder)
            # If new mails are received, forward them to the personal email.
            if mails:
                outlook.sendMail(mails)
        
        print('Pausing for 30 seconds.')
        time.sleep(30)

    except KeyboardInterrupt as e:
        outlook.close()
        sys.exit()
