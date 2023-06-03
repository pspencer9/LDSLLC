import requests
import pandas as pd
import urllib3
from requests_ntlm import HttpNtlmAuth
from config import USERNAME, PASSWORD, JIRA_USERNAME, JIRA_API_KEY
import jira.client
from jira.client import JIRA
from datetime import datetime, timedelta
import sys


###################
##GET JIRA DETAILS
##OUTPUT HOURS TO EXCEL
###################

def main(argv):  
    debug = 1
    if sys.argv[1] == '1':                      ##argv[1] is a debug parameter
        debug = True
        print('Debug Mode')
    elif sys.argv[1] == '0':
        debug = False
           
    start_date = sys.argv[2]                    ##argv[2] is the start date
    end_date = sys.argv[3]                      ##argv[3] is the end date
        
    GenInvoiceOrTimesheet = 'Invoice'
    if len(sys.argv) == 5:
        GenInvoiceOrTimesheet = sys.argv[4]     ##argv[4] is a invoice vs timesheet parameter
        
        
    ##Variable Declaration
    hourly_rate_str = '$135'
    hourly_rate = 135
    
    start_date_date = datetime.strptime(start_date,"%m/%d/%y")
    end_date_date = datetime.strptime(end_date,"%m/%d/%y")
    start_date_str = start_date_date.strftime("%Y/%m/%d")
    end_date_str = end_date_date.strftime("%Y/%m/%d")
      
    api_key = JIRA_API_KEY
    username = JIRA_USERNAME
    options = { 'server': 'https://ccchsd.atlassian.net' }
    
    
    
    ##retrieve jira issues (issues is a jira.client object)
    jira = JIRA(options, basic_auth=(username, api_key))
    
    #this search string below will return all logs that had an update between start_date_str and end_date_str. It
    #will not restrict the updates to those that were done during that time period.
    search_str = 'worklogAuthor=\'' + username + '\' && worklogDate >= "' + start_date_str + '" && worklogDate <= "' + end_date_str + '"'
    issues = jira.search_issues(search_str, fields='summary,worklog,labels') ##worklogDate filtering won't filter out individual time updates, just worklogs that do or do not have an update as of a date
    
    
    jira_log_label = []
    jira_log_date = []
    jira_log_week = []
    jira_log_month = []
    jira_log_week_start = []
    jira_log_hours = []
    
    for issue_name in issues:
        issue = jira.issue(issue_name)
        summary = issue.fields.summary
        labels = issue.fields.labels
        label = ''
        if labels:
            label = labels[0]
            if label != '6911COVID':
                label = ''
            elif label == '6911COVID':
                label = label + ' - '
                
        i=0
        #to get around jira.fields.worklog not returning more than 20 worklogs, use jira.worklogs workaround
        worklogs = jira.worklogs(issue_name.id)
        for worklog in worklogs:
            update_date = datetime.strptime(worklog.started[0:10],'%Y-%m-%d') ##use worklog.started instead of worklog.updated
            author = worklog.author.displayName
            #only loop over my updates that occurred during the time period
            if update_date >= start_date_date and update_date <= end_date_date and author == 'Brendan Lamarre':
                update_date_week_start = update_date - timedelta(days=update_date.weekday())
                update_date_week_end = update_date_week_start + timedelta(days=6)
                update_date_str = worklog.started[0:10]
                
                #if a week extends outside of the input date range, truncate the 
                if update_date_week_start < start_date_date:
                    update_date_week_start = start_date_date
                if update_date_week_end > end_date_date:
                    update_date_week_end = end_date_date
                    
                
                timeSpentInHours = worklog.timeSpentSeconds/3600
                
                jira_log_label.append(label+summary)
                jira_log_date.append(update_date_str)
                jira_log_week.append(update_date_week_start.strftime("%d %b, %Y") + ' - '+ update_date_week_end.strftime("%d %b, %Y"))
                jira_log_week_start.append(update_date_week_start)
                jira_log_month.append(start_date_date.strftime("%d %b, %Y") + ' - ' + end_date_date.strftime("%d %b, %Y"))
                jira_log_hours.append(timeSpentInHours)
                
         
    if GenInvoiceOrTimesheet == 'Invoice':
        d = {'Time Frame': jira_log_month, 'Hours': jira_log_hours}
        df = pd.DataFrame(data = d)
        df2 = df.groupby('Time Frame').sum('Hours').reset_index()
        df2['Total'] = df2.Hours * hourly_rate
        df2['Hours']=df2["Hours"].map(str) ##convert to string so that the next step doesn't act up
        df2['Description'] = 'BI Consulting Services (@ ' + hourly_rate_str + '/hour)'
        
        df2 = df2.iloc[:,[0,3,1,2]] ##reorder columns
    else:
        d = {'Date': jira_log_date, 'Hours': jira_log_hours}
        df = pd.DataFrame(data = d)
        df2 = df.groupby('Date').sum('Hours').reset_index()
        df2['Hours']=df2["Hours"].map(str) ##convert to string so that the next step doesn't act up
        df2['Description'] = 'BI Consulting Services'
        df2 = df2.iloc[:,[0,2,1]] ##reorder columns
            
    ##log the day of the successful load
    if not debug:        
        ##save off the timesheet
        if GenInvoiceOrTimesheet != 'Invoice':
            df2.sort_values('Date').to_excel('timesheet.xlsx')
            print('Timesheet generated')
        else:
            df2.to_excel('invoice.xlsx')
            print('Invoice generated')
 
        
if __name__ == "__main__":
   main(sys.argv[1:])