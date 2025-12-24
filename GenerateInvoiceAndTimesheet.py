import requests
import pandas as pd
from datetime import datetime, timedelta
import sys
from requests.auth import HTTPBasicAuth
from config import JIRA_USERNAME, JIRA_API_KEY

BASE_URL = "https://ccchsd.atlassian.net"
JIRA_API_SEARCH = f"{BASE_URL}/rest/api/3/search/jql"
JIRA_API_ISSUE = f"{BASE_URL}/rest/api/3/issue"

def get_issues(jql, auth):
    """Query Jira using the /search/jql endpoint."""
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    payload = {"jql": jql, "maxResults": 100, "fields": ["summary", "worklog", "labels"]}
    response = requests.post(JIRA_API_SEARCH, headers=headers, json=payload, auth=auth)
    response.raise_for_status()
    return response.json()

def get_all_worklogs(issue_id, auth):
    """Retrieve all worklogs for a given issue."""
    url = f"{JIRA_API_ISSUE}/{issue_id}/worklog"
    headers = {"Accept": "application/json"}
    response = requests.get(url, headers=headers, auth=auth)
    response.raise_for_status()
    return response.json().get("worklogs", [])

def main(argv):
    if len(argv) < 3:
        print("Usage: python GenerateInvoiceAndTimesheet.py <debug_flag> <start_date> <end_date> [Invoice|Timesheet]")
        sys.exit(1)

    debug = argv[0] == '1'
    start_date = argv[1]
    end_date = argv[2]
    GenInvoiceOrTimesheet = argv[3] if len(argv) == 4 else "Invoice"

    hourly_rate = 145
    hourly_rate_str = "$145"

    start_date_date = datetime.strptime(start_date, "%m/%d/%y")
    end_date_date = datetime.strptime(end_date, "%m/%d/%y")
    start_date_str = start_date_date.strftime("%Y/%m/%d")
    end_date_str = end_date_date.strftime("%Y/%m/%d")

    auth = HTTPBasicAuth(JIRA_USERNAME, JIRA_API_KEY)

    jql = (
        f'worklogAuthor = "{JIRA_USERNAME}" '
        f'AND worklogDate >= "{start_date_str}" '
        f'AND worklogDate <= "{end_date_str}"'
    )

    issues_data = get_issues(jql, auth)
    issues = issues_data.get("issues", [])

    jira_log_label, jira_log_date, jira_log_week, jira_log_month = [], [], [], []
    jira_log_week_start, jira_log_hours = [], []

    for issue in issues:
        issue_id = issue["id"]
        fields = issue["fields"]
        summary = fields.get("summary", "")
        labels = fields.get("labels", [])
        label = ""

        if labels and labels[0] == "6911COVID":
            label = labels[0] + " - "

        worklogs = get_all_worklogs(issue_id, auth)

        for worklog in worklogs:
            update_date = datetime.strptime(worklog["started"][:10], "%Y-%m-%d")
            author = worklog["author"]["displayName"]

            if start_date_date <= update_date <= end_date_date and author == "Brendan Lamarre":
                week_start = update_date - timedelta(days=update_date.weekday())
                week_end = week_start + timedelta(days=6)

                # Clamp week range to the given date bounds
                week_start = max(week_start, start_date_date)
                week_end = min(week_end, end_date_date)

                hours = worklog["timeSpentSeconds"] / 3600
                jira_log_label.append(label + summary)
                jira_log_date.append(update_date.strftime("%Y-%m-%d"))
                jira_log_week.append(f"{week_start:%d %b, %Y} - {week_end:%d %b, %Y}")
                jira_log_week_start.append(week_start)
                jira_log_month.append(f"{start_date_date:%d %b, %Y} - {end_date_date:%d %b, %Y}")
                jira_log_hours.append(hours)

    if GenInvoiceOrTimesheet == "Invoice":
        df = pd.DataFrame({"Time Frame": jira_log_month, "Hours": jira_log_hours})
        df2 = df.groupby("Time Frame").sum("Hours").reset_index()
        df2["Total"] = df2.Hours * hourly_rate
        df2["Hours"] = df2["Hours"].astype(str)
        df2["Description"] = f"BI Consulting Services (@ {hourly_rate_str}/hour)"
        df2 = df2[["Time Frame", "Description", "Hours", "Total"]]
    else:
        df = pd.DataFrame({"Date": jira_log_date, "Hours": jira_log_hours})
        df2 = df.groupby("Date").sum("Hours").reset_index()
        df2["Hours"] = df2["Hours"].astype(str)
        df2["Description"] = "BI Consulting Services"
        df2 = df2[["Date", "Description", "Hours"]]

    if not debug:
        if GenInvoiceOrTimesheet.lower() == "invoice":
            df2.to_excel("invoice.xlsx", index=False)
            print("Invoice generated.")
        else:
            df2.sort_values("Date").to_excel("timesheet.xlsx", index=False)
            print("Timesheet generated.")

if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except Exception:
        import traceback
        traceback.print_exc()
