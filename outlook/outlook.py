"""
The following file contains the implementation for the Outlook class.

"""
import datetime
import logging
import configparser
import win32com.client

class Outlook(object):
    """
    __Version__ = 1.0

    @author = dev.dibyo@gmail.com 
    
    __init__()          ----> Params : None.
                              Description : Initializes all the configuration and logging.
    
    __totalMessages()   ----> Params : None.
                              Description : Returns the count of emails received in a logging period.
    
    __formatDate()      ----> Params : date = PywinDateType object.
                              Description : Returns the formatted date.

    __extractDetails()  ----> Params : messages = List of mails.
                              Description : Return the newly received mails as a key-value pair.
    

    fetchMails()        ----> Params : folder = Sub-folder where the search takes place.
                              Description : Returns a dictionary of emails received when system is logged-off.

    sendMail()          ----> Params : messages = Dictionary of mails received.
                              Description : Forwards all the new emails to personal email.
    
    folders()           ----> Params : None
                              Description : Returns list of subfolders.
    
    close()             ----> Params : None
                              Description : Terminates Logging.

    """
    def __init__(self):
        # Read config values from the config.ini file.
        self.subfolders = []
        self.config = configparser.ConfigParser()
        self.config.read('../config/config.ini')
        # Set the log file path.
        self.logfile=self.config['Default']['Logfile']
        # Set the forwarding mail address.
        self.to = self.config['Default']['Email']
        # Fetch all the subfolders to keep track of.
        for key,value in self.config.items('Folders'):
            self.subfolders.append(value)
        # Initialize outlook API.
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.mapi = self.outlook.GetNamespace("MAPI")
        # Set the checkpoint to 5 mins - This will fetch mails sent 5 mins before the system is locked.
        self.checkpoint= datetime.datetime.now() - datetime.timedelta(minutes = 5)
        # Counter to keep track of messages received during a logging period.
        self.countMessages=0 
        # Logging File config.
        logging.basicConfig(filename=self.logfile,format='* %(message)s',level=logging.INFO)
        logging.info('-------------------------------------  Logging started at : '\
                     + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S') \
                     +' ------------------------------')
    
    def __totalMessages(self):
        """
        Returns the count of emails received.
        """
        return self.countMessages
    
    def __formatDate(self,date):
        """
        Returns the formatted date.
        """
        self.formattedDate = date[:-6]
        return self.formattedDate
    
    def __extractDetails(self, messages):
        """
        Returns the details of the received emails as a key-value pair.
        """
        self.messageDetails = {}
        for message in messages:
            self.receivedTime = message.ReceivedTime
            self.sender=message.SenderName
            if message.MessageClass == 'IPM.Schedule.Meeting.Request':
                self.appointment = message.GetAssociatedAppointment(0)
                self.text = ' sent you a meeting request.\n' \
                            +'  Location: '+str(self.appointment.Location) \
                            + ' Start: '+str(self.appointment.Start.strftime('%H:%M')) \
                            +' End: '+str(self.appointment.End.strftime('%H:%M'))        
            else :
                self.text = ' sent you a mail.'
            
            self.messageDetails[ self.__formatDate(str(self.receivedTime)) ] = str(self.sender) + self.text
            # Log the details of the new mails received.
            logging.info(self.__formatDate(str(self.receivedTime))+' '+ str(self.sender) + self.text)
            self.countMessages += 1
        return self.messageDetails

    def fetchMails(self,folder):
        """
        Returns a dictionary of new emails.
        Key : received Time
        Value : Mail details.
        """
        # Navigate to the subfolder of interest.
        self.folder = self.mapi.Folders['Directory'].Folders['Sub-directory'].Folders[folder] 
        self.messages = self.folder.Items
        # Sort the subfolder based on the ReceivedTime attribute
        self.messages.Sort("[ReceivedTime]", True)
        # Filter those messages that were received since the last checkpoint i.e before the  program went to sleep.
        self.messages = self.messages.Restrict("[ReceivedTime] >= '" +self.checkpoint.strftime('%H:%M %p')+"'")

        if self.messages:
            self.newMessages = self.__extractDetails(self.messages) 
            # Advance the checkpoint to the current time.
            self.checkpoint = datetime.datetime.now() + datetime.timedelta(minutes = 1)
            return self.newMessages


    def sendMail(self,messages):
        """
        Forwards the mails to personal Email.
        """
        self.messages = messages
        self.mail = self.outlook.CreateItem(0)
        self.mail.To = self.to
        self.mail.Subject = 'Outlook Notification'
        self.messageBody = 'You have received the following messages:\n\n'

        for timestamp,body in self.messages.items():
            self.messageBody +="* " + timestamp +' -- '+body+'\n'
        
        self.mail.Body = self.messageBody
        self.mail.Send()
    
    def folders(self):
        """
        Returns a list of subfolders currently tracking.
        """
        return self.subfolders

    def close(self):
        """
        Terminates the program on Log-on.
        """
        if self.__totalMessages() == 0:
            logging.info('* No emails received.')
        logging.info('-------------------------------------  Logging ended at : ' \
                    + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S') +'  --------------------------------')
