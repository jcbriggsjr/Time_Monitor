# -*- coding: utf-8 -*-
"""
Created on Tue Mar 17 11:30:50 2020

@author: JBriggs
"""

import pandas as pd
import send_email as email
import time
import pyodbc as db

wippath = "G:\\3 - Production Departments\\4 - Grinding\\0 - Department Documents\\4 - Programs & Software\\1 - Operating Software\\Setup_Tracker\\Data\\setups_in_progress.csv"
latepath = ".\\latewip.csv"
warnpath = ".\\warnwip.csv"
twohourpath = ".\\twohoursetup.csv"

joe = 'jbriggs@catiglass.com;'
penny = 'penny.quandt@catiglass.com;'
slawek = 'slawek@catiglass.com;'
send_to = joe
joephone = '8159006943@txt.att.net'
text_to = joephone

def getJobData(jobnum):
    jobnum = "'" + jobnum.upper() + "'"    
    with open('sql_access.txt', 'r') as file:
        access = file.readlines()
    

    job_data = {}
    conn = db.connect('Driver={SQL Server};'
                      'Server=SVR-APP\\SQLEXPRESS;'
                      + access[0] +
                      'Trusted_Connect=yes;')
    
    command = ("SELECT Parts.CustName, Parts.CustPartNum "
                   "FROM QssCatiJobTrack.dbo.Jobs, QssCatiJobTrack.dbo.Parts "
                   "WHERE Jobs.JobNum LIKE " + jobnum + " "
                   "AND Jobs.PartID = Parts.PartID ")    
    
    cursor = conn.cursor()
    
    cursor.execute(command)
    
    for row in cursor:
        job_data[jobnum] = [row[0],row[1]]
    cursor.close()
    conn.close()
    return job_data

def formatMessage(jlist,machine,et,operator):
    #this returns a string containing dict data in a more human readable format
    msg = ""
    
    for k in sorted(jlist):
        job = k.strip("'")
        cust = jlist[k][0]
        p_num = jlist[k][1]        
        msg = ("Job " + job + " for " + cust + ", part # " + p_num + 
               " at " + machine + " is currently at " + str(et) + " elapsed time. Tech is " + operator + ".")
    return msg

def checkWip():
    wip = pd.read_csv(wippath)
    now = pd.Timestamp.now()
    wip['timestamp'] = wip['Start Date'].astype(str) + "T" + wip['Start Time'].astype(str)    
    wip['timestamp'] = pd.to_datetime(wip['timestamp']) 
    wip['elapsed'] = now - wip['timestamp']
    warning = wip.elapsed.dt.total_seconds() >= 2700
    overtime = wip.elapsed.dt.total_seconds() > 3600
    two_hours = wip.elapsed.dt.total_seconds() > 7200
    
    return wip[overtime], wip[warning], wip[two_hours]


def manageOT():
    late_wip, warn_wip, two_hours = checkWip()
    
    column_names = ['Job Number','Machine','Setup Tech','Start Date','Start Time','timestamp','elapsed']
    
    while True:
        try:
            prev_late = pd.read_csv(latepath)
        except FileNotFoundError as fnf:
            print(fnf, 'Creating missing file.')
            df = pd.DataFrame(columns = column_names)
            df.to_csv('latewip.csv')
            continue
            
        try:
            prev_warn = pd.read_csv(warnpath)
        except FileNotFoundError as fnf:
            print(fnf, 'Creating missing file.')
            df = pd.DataFrame(columns = column_names)
            df.to_csv('warnwip.csv')
            continue
            
        try:
            two_late = pd.read_csv(twohourpath)
        except FileNotFoundError as fnf:
            print(fnf, 'Creating missing file.')
            df = pd.DataFrame(columns = column_names)
            df.to_csv('twohoursetup.csv')
            continue
        break
    
    late_wip.to_csv('latewip.csv', index=False)
    warn_wip.to_csv('warnwip.csv', index=False)
    two_hours.to_csv('twohoursetup.csv', index=False)
    
    jac = prev_late['Job Number'].unique().astype(str)
    wjac = prev_warn['Job Number'].unique().astype(str)
    tjac = two_late['Job Number'].unique().astype(str)
    
    for row in late_wip.itertuples(index=False):
        jobnum = str(getattr(row, '_0'))        
        machine = getattr(row, 'Machine')
        operator = getattr(row, '_2')
        et = getattr(row,'elapsed')
        if jobnum not in jac:            
            # need to send email
            subject = 'Grinding setup over 1 hour: Job # ' + jobnum
            job_list = getJobData(jobnum)
            msg = formatMessage(job_list,machine,et,operator)
            email.mailMessage(send_to, subject, msg)
            email.txtMessage(text_to, subject, msg)
            
    for row in warn_wip.itertuples(index=False):
        jobnum = str(getattr(row, '_0'))        
        machine = getattr(row, 'Machine')
        operator = getattr(row, '_2')
        et = getattr(row,'elapsed')
        if jobnum not in wjac:            
            # need to send email
            subject = 'Grinding setup 45 minutes warning: Job # ' + jobnum
            job_list = getJobData(jobnum)
            msg = formatMessage(job_list,machine,et,operator)
            email.mailMessage(send_to, subject, msg)
            email.txtMessage(text_to, subject, msg)
            
    for row in two_hours.itertuples(index=False):
        jobnum = str(getattr(row, '_0'))        
        machine = getattr(row, 'Machine')
        operator = getattr(row, '_2')
        et = getattr(row,'elapsed')
        if jobnum not in tjac:            
            # need to send email
            subject = 'Grinding setup over 2 hours: Job # ' + jobnum
            job_list = getJobData(jobnum)
            msg = formatMessage(job_list,machine,et,operator)
            email.mailMessage(send_to, subject, msg)
            email.txtMessage(text_to, subject, msg)
            
            
            
if __name__ == "__main__":
    while True:
        checkWip()
        manageOT()
        time.sleep(10)