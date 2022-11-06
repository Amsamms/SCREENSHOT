import os
import pyautogui as pg
import win32com.client as client

# Zoom in or out the screen
pg.hotkey('ctrlleft','+') # zoom in
pg.hotkey('ctrlleft','-') # zoom out

# take screenshot
destination=r'E:\Data learning\selenium' # destination folder
os.chdir(destination) # changing directory to destination folder
pg.screenshot("demo.png") # records this image in the destination folder, filename is demo.png

# send the image via mail

to_list = 'asabri@eprom-midor.com.eg; afahim@eprom-midor.com.eg' # or you can write the name as in address book
cc_list = 'asabri@eprom-midor.com.eg; afahim@eprom-midor.com.eg' # or you can write the name as in address book

outlook = client.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = to_list
mail.CC = cc_list
mail.Subject = 'Daily monitoring data' # mail subject
mail.Body = f'Dear Coleagues,\n\n\nkindly find attached  screenshot\n\n\n\n\n\n\nBest Regards.' #mail body

# To attach a file to the email (optional):
attachment  = os.path.join(destination,"demo.png") # the path of demo.png
mail.Attachments.Add(attachment)
mail.Send()

