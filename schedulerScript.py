
import pandas as pd
import time
from datetime import date, datetime, timedelta
from datetime import time as timegen
import calendar
import os
import importlib.util     
from tqdm import tqdm
import math
import sys
import traceback
from calendar import monthrange


def callForeignScriptMainFun(currentPath, scriptAddress, scriptName): 
    # passing the file name and path as argument
    spec = importlib.util.spec_from_file_location("mod", scriptAddress + "\\" + scriptName)   
    
    # importing the module as foo
    test_py = importlib.util.module_from_spec(spec)       
    spec.loader.exec_module(test_py)
    
    # calling the hello function of mod.py
    try:
        test_py.main_func(scriptAddress)
    except Exception as e:
        print("\nError ⇩ ⇩ ⇩ ⇩ ⇩ ⇩ ⇩")
        print(e)
        traceback.print_exception(*sys.exc_info())
        time.sleep(5)
        print("Error ⇧ ⇧ ⇧ ⇧ ⇧ ⇧ ⇧\n")
        print(f"{scriptName} execution skipped!!!!\n")

    
    
def getCurrDay():
    curr_date = date.today() + timedelta(0)
    curr_day = str(calendar.day_name[curr_date.weekday()])
    return curr_day

def sleepTillTommorow():
    sleepTimeVal = (datetime.combine((datetime.now() + timedelta(1)).date(), timegen(0,0,0))- datetime.now())
    print("Sleeping for ",sleepTimeVal)
    sleep_sec = math.ceil(sleepTimeVal.total_seconds()/60)
    for slTime in tqdm(range(sleep_sec)):
        time.sleep(60)
    return True
    
def getTimeDeltaSeconds(sTime):  #Takes time from excel dataframe and return "excel time" - "current time" in seconds. DO NOT SEND DATES, ONLY TIME 
    fun_sTime = datetime.combine(date.today(), sTime)
    current_time = datetime.now()  #.time()
    timeDeltaVal = (fun_sTime - current_time)
    timeDeltaVal = timeDeltaVal.total_seconds()
    return timeDeltaVal

def runAdhocScripts(currentPath, currRunday, adhocCounter, sTime = ""):
    thisRunDay = getCurrDay()
    deltaTimeSeconds = 0
    if sTime == "":
        deltaTimeSeconds = (datetime.combine((datetime.now() + timedelta(1)).date(), timegen(0,0,0))- datetime.now())
        deltaTimeSeconds = deltaTimeSeconds.total_seconds()
    else:
        deltaTimeSeconds = getTimeDeltaSeconds(sTime)
    
    if (deltaTimeSeconds/60) < 30 and sTime == "":
        sleepTillTommorow()
    elif (deltaTimeSeconds/60) < 30 and sTime != "":
        return adhocCounter
    
    if currRunday != thisRunDay:
        return adhocCounter
    print("\n\n")
    print(f"{adhocCounter}.----->>>>>>>>>> Running AdHoc Scripts <<<<<<<<<<-----")
    adHoc_DF = pd.read_excel(currentPath + "\\Scheduler_Sheet.xlsx", sheet_name="AdHoc").fillna(0)
    for thisAdhoc in range(len(adHoc_DF)):
        adHocScriptPath = str(adHoc_DF['Path'][thisAdhoc]).strip()
        adHocScriptName = str(adHoc_DF['Script Name'][thisAdhoc]).strip()
        adHocScriptRunStat = str(adHoc_DF['Run (Y/N)'][thisAdhoc]).strip()
        if adHocScriptRunStat == "N":  # iterate to next item
            continue 
        
        AD_sTime = datetime.now()
        print(f"--->>> {adHocScriptName} <<<---")
        print(f"Current Run Day: {currRunday}; Actual Run Day: {thisRunDay}; Adhoc Cycle: {adhocCounter}; Script Serial: {thisAdhoc+1}; Execution Start Time: {AD_sTime}")
        callForeignScriptMainFun(currentPath, adHocScriptPath, adHocScriptName)
        AD_eTime = datetime.now()
        AD_totalScriptTimeDelta = (AD_eTime - AD_sTime)
        print(f"Execution Finish Time: {AD_eTime}; Total Runtime: {AD_totalScriptTimeDelta}")
        print("\n")
    adhocCounter  = adhocCounter  +1
    adhocCounter  = runAdhocScripts(currentPath, currRunday, adhocCounter, sTime)



def main_func():
    argsForSS = sys.argv
    skipDailyFlag = False
    skipCounter = 0
    skipMonthlyRunFlag = False
    
    if "skipdailyrun" in [x.lower() for x in argsForSS if type(x) == str]:
        skipDailyFlag = True
        skipCounter = 9999
        try:
            skipCounter = int(argsForSS[2])
        except:
            pass
    
    if "skipmonthlyrun" in [x.lower() for x in argsForSS if type(x) == str]:
        skipMonthlyRunFlag = True
            
    
    prev_RunDay = ""
    prev_current_date_of_month = -1
    adhocCounter = 1
    currentPath = os.getcwd()
    
    current_date_of_month = int(datetime.today().strftime('%d'))
    last_day_less_n = monthrange(int(datetime.today().strftime('%Y')), int(datetime.today().strftime('%m')))[1] - current_date_of_month
    last_day_less_n_string = ["Last Day Less " + str(last_day_less_n) if last_day_less_n >0 else "Last Day"][0]
    while True:
        currRunDay = getCurrDay()
        if prev_RunDay != currRunDay:
            prev_RunDay = currRunDay
            adhocCounter = 1
            if skipDailyFlag == True and skipCounter >= 9999:
                print(f"Skipping all daily run scipts for today ({currRunDay})")
            else:
                print("\n\n\n\n")
                print(f"------------------------------------------------------___{currRunDay}___------------------------------------------------------")
                print(f"---------------------------------------________________________________________________---------------------------------------")
                schedulerSheet_DF = pd.read_excel(currentPath + "\\Scheduler_Sheet.xlsx", sheet_name=currRunDay).fillna(0)

                for i in range(len(schedulerSheet_DF)):
                    
                    sTime = schedulerSheet_DF['Time'][i]
                    sPath = str(schedulerSheet_DF['Path'][i]).strip()
                    sName = str(schedulerSheet_DF['Script Name'][i]).strip()
                    sRunStat = str(schedulerSheet_DF['Run (Y/N)'][i]).strip()
                    
                    if sRunStat == "N":  # iterate to next item
                        continue
                    
                    if skipCounter >= i+1:  #iterate to next item based on skipCounter provided by user.... skip counter will count the skip of scripts with "Y" as run status only
                        continue
                    
                    if sTime == 0:
                        sTime = datetime.now().time()
                    while True:
                        this_timeDeltaVal = getTimeDeltaSeconds(sTime)/60   # Function returns seconds, converted to minutes
                        if this_timeDeltaVal > 30:
                            adhocCounter  = runAdhocScripts(currentPath, currRunDay, adhocCounter, sTime)
                        elif this_timeDeltaVal > 0 and this_timeDeltaVal <= 30:
                            print(f"Sleeping for {this_timeDeltaVal}")
                            sleep_sec = math.ceil(this_timeDeltaVal)
                            for slTime in tqdm(range(sleep_sec)):
                                time.sleep(60)
                        else:
                            break
                    
                    print("\n\n")
                    print(f"{i+1}. ------------>>> {sName} <<<------------")
                    print(f"    Scheduled start time =",sTime)
                    scriptStartTime = datetime.now() 
                    print(f"    Executing {sName} script at: {scriptStartTime}")
                    print("\n")
                    callForeignScriptMainFun(currentPath, sPath, sName)
                    print("------------------------\n")
                    scriptEndTime = datetime.now()
                    print(f"    Execution complete at: {scriptEndTime}")
                    totalScriptTimeDelta = (scriptEndTime- scriptStartTime)
                    print(f"    Total run time: {totalScriptTimeDelta}")
            
            skipDailyFlag = False
            skipCounter = 0 
            
        elif prev_current_date_of_month != current_date_of_month:
            prev_current_date_of_month = current_date_of_month
            if skipMonthlyRunFlag == True:
                print("Skipping all daily run scipts for today, Day: ({current_date_of_month})")
            else:
                
                schedulerSheet_DF = pd.read_excel(currentPath + "\\Scheduler_Sheet.xlsx", sheet_name="Monthly").fillna(0)
                schedulerSheet_DF = schedulerSheet_DF[(schedulerSheet_DF['Day of Month'].values == current_date_of_month)|(schedulerSheet_DF['Day of Month'].values == last_day_less_n_string)]
                schedulerSheet_DF.reset_index(drop=True, inplace=True)
                if len(schedulerSheet_DF)>0:
                    print("\n\n\n\n")
                    print(f"------------------------------------------------------___Day: {current_date_of_month}___------------------------------------------------------")
                    print(f"---------------------------------------________________________________________________---------------------------------------")
                    for i in range(len(schedulerSheet_DF)):
                        
                        sTime = schedulerSheet_DF['Time'][i]
                        sPath = str(schedulerSheet_DF['Path'][i]).strip()
                        sName = str(schedulerSheet_DF['Script Name'][i]).strip()
                        sRunStat = str(schedulerSheet_DF['Run (Y/N)'][i]).strip()
                        
                        if sRunStat == "N":  # iterate to next item
                            continue
                        
                        if sTime == 0:
                            sTime = datetime.now().time()
                        while True:
                            this_timeDeltaVal = getTimeDeltaSeconds(sTime)/60   # Function returns seconds, converted to minutes
                            if this_timeDeltaVal > 30:
                                adhocCounter  = runAdhocScripts(currentPath, currRunDay, adhocCounter, sTime)
                            elif this_timeDeltaVal > 0 and this_timeDeltaVal <= 30:
                                print(f"Sleeping for {this_timeDeltaVal}")
                                sleep_sec = math.ceil(this_timeDeltaVal)
                                for slTime in tqdm(range(sleep_sec)):
                                    time.sleep(60)
                            else:
                                break
                        
                        print("\n\n")
                        print(f"{i+1}. ------------>>> {sName} <<<------------")
                        print(f"    Scheduled start time =",sTime)
                        scriptStartTime = datetime.now() 
                        print(f"    Executing {sName} script at: {scriptStartTime}")
                        print("\n")
                        callForeignScriptMainFun(currentPath, sPath, sName)
                        print("------------------------\n")
                        scriptEndTime = datetime.now()
                        print(f"    Execution complete at: {scriptEndTime}")
                        totalScriptTimeDelta = (scriptEndTime- scriptStartTime)
                        print(f"    Total run time: {totalScriptTimeDelta}")                   
                    
            skipMonthlyRunFlag = False
                
        else:
            adhocCounter  = runAdhocScripts(currentPath, currRunDay, adhocCounter)



if __name__ == "__main__":
    print("Starting Scheduler Script...\n")
    main_func()


