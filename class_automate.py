import csv
from time import sleep
from datetime import date, time
import calendar
import datetime
import webbrowser

chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
now = datetime.datetime.now()
time = float(str(now.hour)  + "." + str(now.minute))
reader = csv.reader(open('Class_schedule - Sheet1.csv', 'r'))
classes = {}
for row in reader:
   k, v = row
   classes[k] = v

my_date = date.today()
today = calendar.day_name[my_date.weekday()]



def shedule():
   flag = 0
   if today == "Monday":
      if time >= 8.55 and time < 9.10 :
         flag = 1
         link = ((classes.get("Monday_class1")))
      if time >= 9.43 and time < 9.58 :
         flag = 1
         link = ((classes.get("Monday_class2")))
      if time >= 10.27 and time < 10.42 :
         flag = 1
         link = ((classes.get("Monday_class3")))
      if time >= 12.13 and time < 12.28 :
         flag = 1
         link = ((classes.get("Monday_class4")))
      if time >= 13.32 and time < 13.47 :
         flag = 1
         link = ((classes.get("Monday_class5")))
      if time > 13.47:
         flag = 0

   elif today == "Tuesday":
      if time >= 8.55 and time < 9.10 :
         flag = 1
         link = ((classes.get("Tuesday_class1")))
      if time >= 9.43 and time < 9.58 :
         flag = 1
         link = ((classes.get("Tuesday_class2")))
      if time >= 11.26 and time < 11.38 :
         flag = 1
         link = ((classes.get("Tuesday_class3")))
      if time >= 12.13 and time < 12.28 :
         flag = 1
         link = ((classes.get("Tuesday_class")))
      if time >= 14.16 and time < 14.31 :
         flag = 1
         link = ((classes.get("Tuesday_class")))
      if time > 14.31 :
         flag = 0

   elif today == "Wednesday":
      if time >= 8.55 and time < 9.10 :
         flag = 1
         link = ((classes.get("Wednesday_class1")))
      if time >= 9.43 and time < 9.58 :
         flag = 1
         link = ((classes.get("Wednesday_class2")))
      if time >= 10.27 and time < 10.42 :
         flag = 1
         link = ((classes.get("Wednesday_class3")))
      if time >= 11.23 and time < 11.38 :
         flag = 1
         link = ((classes.get("Wednesday_class4")))
      if time >= 13.32 and time < 13.47 :
         flag = 1
         link = ((classes.get("Wednesday_class5")))
      if time > 13.47 :
         flag = 0

   elif today == "Thursday":
      if time >= 9.43 and time < 9.58 :
         flag = 1
         link = ((classes.get("Thursday_class1")))
      if time >= 10.27 and time < 10.42 :
         flag = 1
         link = ((classes.get("Thursday_class2")))
      if time >= 11.23 and time < 11.38 :
         flag = 1
         link = ((classes.get("Thursday_class3")))
      if time > 11.38 :
         flag = 0

   elif today == "Friday":
      if time >= 8.55 and time < 9.10 :
         flag = 1
         link = ((classes.get("Friday_class1")))
      if time >= 10.27 and time < 10.42 :
         flag = 1
         link = ((classes.get("Friday_class2")))
      if time >= 11.23 and time < 11.38 :
         flag = 1
         link = ((classes.get("Friday_class3")))
      if time >= 13.32 and time < 13.47 :
         flag = 1
         link = ((classes.get("Friday_class4")))
      if time >= 14.17 and time < 14.32 :
         flag = 1
         link = ((classes.get("Friday_class5")))
      if time > 14.32 :
         flag = 0

   #check if class is running or not
   if flag == 0:
      return None
   if flag > 0:
      return link



start_time = [8.55, 9.43,10.27,11.26,12.13,13.32,14.16] # start time each of the class


def automate(): #checks time with start time then run the script
   now = datetime.datetime.now()
   time1 = float(str(now.hour)  + "." + str(now.minute))
   y = time1
   if y in start_time:
      if shedule() is not None :
         url = " " + shedule()
         webbrowser.get(chrome_path).open(url)
      else:
         pass
   else:      
      pass
        

while True:
   automate()
   sleep(35)