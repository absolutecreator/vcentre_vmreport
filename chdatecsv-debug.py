import os
import time
import datetime

fileLocation = "/opt/vmware-report/csv/vmware_report_30.03.2023.csv"
year = 2023
month = 3
day = 30
hour = 19
minute = 50
second = 0

date = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=second)
modTime = time.mktime(date.timetuple())

os.utime(fileLocation, (modTime, modTime))