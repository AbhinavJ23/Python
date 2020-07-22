#
# Example file for working with date information
#
from datetime import date
from datetime import time
from datetime import datetime

def main():
  ## DATE OBJECTS
  # Get today's date from the simple today() method from the date class
  today = date.today()
  print("Today's date is ", today)


  # print out the date's individual components
  print("Date components are as follows :", today.day, today.month, today.year)
  
  # retrieve today's weekday (0=Monday, 6=Sunday)
  print("Today's Weekday: ", today.weekday())
  days = ["Monday", "Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
  print("Day is ", days[today.weekday()])
  
  ## DATETIME OBJECTS
  # Get today's date from the datetime class
  today = datetime.now()
  print("Now is : ",today)

  # Get the current time
  print(today.time())
  print(datetime.time(datetime.now()))
  
  
if __name__ == "__main__":
  main()
  