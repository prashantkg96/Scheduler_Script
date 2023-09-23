# Scheduler_Script
This script is used to schedule runs of other python based scripts.
Make sure other scripts scheduled is inside the function called `main_func`. Fill in the excel file with other script folder path, other script name and time of execution (optional) and run the schedulerScript.py.
The scheduler script can be used to schedule other scripts to run on
  1. Adhoc basis - When the scheduler has nothing to run for the day or it has more than 30 minutes for the next script to start, it will continuously cycle through the scripts that are in entered in `AdHoc` sheet
  2. Sunday to Saturday - Scripts mentioned here will run based on current day. 
  3. Monthly basis - Scripts mentioned here will run is the current day of the month is matched. Input in the `Monthly` sheet can be the day number or `Last Day Less n` where n = integer. Note `Last Day Less` is case sensitive so make sure you type is as mentioned

## Commands
`python schedulerScript.py skipmonthlyrun`
This will skip any monthly runs scheduled for the current day and jump to any scripts mentioned in current week day schedule and then run adhoc scripts. Next day it will run all the schedules mentioned in `Monthly` sheet, `Sunday to Saturday` sheet and `AdHoc` sheet if the criteria of current day is matched.

`python schedulerScript.py skipdailyrun`
This will skip any daily runs scheduled for the current day and jump to any scripts mentioned in day of the month(Monthly) schedule and then run adhoc scripts. Next day it will run all the schedules mentioned in `Monthly` sheet, `Sunday to Saturday` sheet and `AdHoc` sheet if the criteria of current day is matched. 

`python schedulerScript.py skipdailyrun n` where n is integer
This will skip for `n` scripts in the current weekday day of the schedule. Next day it will run all the schedules mentioned in `Monthly` sheet, `Sunday to Saturday` sheet and `AdHoc` sheet if the criteria of current day is matched. 

`python schedulerScript.py skipmonthlyrun skipdailyrun`
This will skip both the current weekday schedule and current day of the month schedule and jump directly to adhoc scripts schedule and will keep cycling though the adhoc schedule till the next day. Next day it will run all the schedules mentioned in `Monthly` sheet, `Sunday to Saturday` sheet and `AdHoc` sheet if the criteria of current day is matched. 
