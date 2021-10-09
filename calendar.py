import win32com.client
from datetime import datetime
import pytz
timezone = pytz.timezone('Europe/Paris')

start_time = timezone.localize(datetime(2021, 10, 10, 18)) 
#add 2 hours, because problem by europe time zone
subject = 'Schweizer super League - car pooling dernier test2.0'
#change for your own meeting
duration = 180
location = 'Stadion Wankdorf - 71 Papiermühlestrasse - 3014 Bern - Switzerland' 
#change for your own meeting

recipient1 = 'exemple1@mail.com'
recipient2 = 'exemple2@mail.com'
recipient3 = 'exemple3@mail.com'
sender = 'exemple@mail.com'

outlook = win32com.client.Dispatch('outlook.application')
appt = outlook.CreateItem(1) 

appt.Start = start_time
appt.Duration = duration
appt.Location = location
appt.Subject = subject

appt.MeetingStatus = 1 
appt.Recipients.Add(recipient1)
appt.Recipients.Add(recipient2)
appt.Recipients.Add(recipient3)
appt.Organizer = sender
appt.ReminderMinutesBeforeStart = 15
appt.ResponseRequested = True
appt.Save()
appt.Send()

print("Message send to")
print(recipient1)
print(recipient2)
print(recipient3)