from win10toast import ToastNotifier #library to send notification
import time
import win32com.client as win32 #library to send email
import time
#pip install win10toast
#pip install pywin32

# Automatic notification and email sender
#first let's create the integration with the email application, here I use Outlook
outlook = win32.Dispatch('outlook.application')



# function for create an email
email = outlook.CreateItem(0)


# here you configure your email information
email.To = "email to send"
email.Subject= "subject"
email.HTMLBody = f"""

<p>text</p>

"""
email.Send() #send the email

#here you can create a function to send the email, and call it when you want to send the email

#here I create a function to send the email every 5 seconds

def send_email():
    email.Send()
    time.sleep(5)
    send_email()
    
    
send_email()

#now you can create a function to send the notification, and call it when you want to send the notification

toaster= ToastNotifier() #create the integration with the notification


def notification():
    toaster.show_toast ("Notificação", #title  
                   "", #message
                    threaded=True, #if you want to send the notification in a thread
                    duration=5) #duration of the notification


notification()
