import win32com.client

ol = win32com.client.Dispatch("outlook.application")

ol_mail_item = 0x0 # size of the new email

new_mail = ol.CreateItem(ol_mail_item)

new_mail.Subject = "Postnord"
new_mail.To = "xyz@gmail.com"
# new_mail.CC = "xyz@gmail.com"
new_mail.Body = "Hej! \n\n" \
               "" \
               "Det finns Postnord för att hämta I repan. 😊 \n\n" \
               "" \
               "Med vänlig hälsning\nBekir"

# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# new_mail.Attachments.Add(attach)
# To display the mail before sending it
new_mail.Display()

#newmail.Send()