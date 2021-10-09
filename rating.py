
import win32com.client

outlook = win32com.client.Dispatch('outlook.application')
write = outlook.CreateItem(0)

write.To = 'exemple1@mail.com; exemple2@mail.com; exemple3@mail.com'
write.Subject = 'Rate the trip'
write.Body = 'Please go to (web adresse) and please rate your driver to help us.'
write.Send()
