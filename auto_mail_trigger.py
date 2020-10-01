import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "rajashekhar.pbr@accenture.com"
message.CC = "a.a.gaurav@accenture.com"
# message.BCC = "a.a.gaurav@accenture.com"
message.Subject = "Test mail"
message.Body = "Hey this is test mail using python prog"
# message.SentOnBehalfOfName = "rajashekhar.pbr@accenture.com"
attachment1 = "C:\\Users\\r.peddaboddu\\Desktop\\schedError.PNG"
message.Attachments.Add(attachment1)
message.send()


