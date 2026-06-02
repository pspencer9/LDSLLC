import win32com.client as client
import os
import shutil
import requests
import pandas as pd
from datetime import datetime, timedelta
import sys
from requests.auth import HTTPBasicAuth
import openpyxl
from docx import Document
from num2words import num2words

from config import JIRA_USERNAME, JIRA_API_KEY, BILLING_NAME, BILLING_TASK, INVOICE_ABREVIATION, HOURLY_RATE

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
        if len(argv) == 1 and argv[0] == "lastMonth":
            today = datetime.today()
            first_day_last_month = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
            last_day_last_month = today.replace(day=1) - timedelta(days=1)
            argv = ["0", first_day_last_month.strftime("%Y-%m-%d"), last_day_last_month.strftime("%Y-%m-%d"), "all"]
            print("Auto-calculated last month date range:", argv[1], "to", argv[2])
        else:
            print("Usage: python GenerateInvoiceAndTimesheet.py <debug_flag> <start_date> <end_date> [Invoice|Timesheet]")
            sys.exit(1)

    debug = argv[0] == '1'
    start_date = argv[1]
    end_date = argv[2]
    GenInvoiceOrTimesheet = argv[3] if len(argv) == 4 else "Invoice"
    GenInvoice = GenInvoiceOrTimesheet.lower() in ["invoice", "all"]
    GenTimesheet = GenInvoiceOrTimesheet.lower() in ["timesheet", "all"]
    GenDemand = GenInvoiceOrTimesheet.lower() in ["demand", "all"]


    hourly_rate = HOURLY_RATE
    hourly_rate_str = f"${HOURLY_RATE}"

    print("Parameter values:", debug, start_date, end_date, GenInvoiceOrTimesheet)
    start_date_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date_date = datetime.strptime(end_date, "%Y-%m-%d")
    start_date_str = start_date
    end_date_str = end_date
    year_month = end_date_str[:7]

    auth = HTTPBasicAuth(JIRA_USERNAME, JIRA_API_KEY)

    jql = (
        f'worklogAuthor = "{JIRA_USERNAME}" '
        f'AND worklogDate >= "{start_date_str}" '
        f'AND worklogDate <= "{end_date_str}"'
    )

    issues_data = get_issues(jql, auth)
    issues = issues_data.get("issues", [])

    print("Total issues retrieved:", len(issues))

    jira_log_label, jira_log_date, jira_log_week, jira_log_month = [], [], [], []
    jira_log_week_start, jira_log_hours = [], []

    for issue in issues:
        issue_id = issue["id"]
        fields = issue["fields"]
        summary = fields.get("summary", "")
        labels = fields.get("labels", [])
        label = ""

        worklogs = get_all_worklogs(issue_id, auth)

        for worklog in worklogs:
            update_date = datetime.strptime(worklog["started"][:10], "%Y-%m-%d")
            author = worklog["author"]["displayName"]

            if start_date_date <= update_date <= end_date_date and author == BILLING_NAME:
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

    # Create comprehensive dataframe with all worklog data
    df_all = pd.DataFrame({
        "Label": jira_log_label,
        "Date": jira_log_date,
        "Week": jira_log_week,
        "Time Frame": jira_log_month,
        "Hours": jira_log_hours
    })

    print("All worklog data preview:")
    print(df_all.head())
    print(f"\nTotal worklog entries: {len(df_all)}")
    print(f"Total hours: {df_all['Hours'].sum():.1f}")
    print(f"# of Days worked: {df_all['Date'].nunique()} of {(end_date_date - start_date_date).days + 1}")
    print("Average hours per day:", df_all.groupby("Date")["Hours"].sum().mean())

    if GenInvoice:
        df2 = df_all.groupby("Time Frame").sum(numeric_only=True).reset_index()
        df2["Total"] = df2.Hours * hourly_rate
        df2["Hours"] = df2["Hours"].astype(str)
        df2["Description"] = f"{BILLING_TASK} (@ {hourly_rate_str}/hour)"
        df2 = df2[["Time Frame", "Description", "Hours", "Total"]]
    if GenTimesheet:
        df3 = df_all.groupby("Date").sum(numeric_only=True).reset_index()
        df3["Hours"] = df3["Hours"].astype(str)
        df3["Description"] = f"{BILLING_TASK}"
        df3 = df3[["Date", "Description", "Hours"]]

    if not debug:
        print("Debug mode off - generating files...")
        #Create a new folder with the name INVOICE_ABREVIATION and year_month if it doesn't exist
        invoiceName = INVOICE_ABREVIATION+year_month
        invoiceFileName = invoiceName + "_Invoice.xlsx"
        timesheetFileName = invoiceName + "_Timesheet.xlsx"
        demandNameFileName = "D15 Demand form - " + year_month
        folder_name = f"../Invoices/{invoiceName}"
        os.makedirs(folder_name, exist_ok=True)
        timesheetPdfName = os.path.abspath(os.path.join(folder_name, f"{invoiceName}_Timesheet.pdf"))


        #Copy template file to the new folder, and rename it to with invoiceName, and timesheetName respectively
        inv_template_path = "../Templates/Invoice - Template.xlsx"
        timesheet_template_path = "../Templates/Timesheet - Template.xlsx"
        demand_template_path = "../Templates/Demand - Template.docx"
        if os.path.exists(inv_template_path):
            new_template_path = f"{folder_name}/{invoiceFileName}"
            openpyxl.load_workbook(inv_template_path).save(new_template_path)
        else:
            print("Template file not found. Please ensure InvoiceTemplate.xlsx exists in the parent directory.")

        if os.path.exists(timesheet_template_path):
            new_template_path = f"{folder_name}/{timesheetFileName}"
            openpyxl.load_workbook(timesheet_template_path).save(new_template_path)
        else:
            print("Template file not found. Please ensure TimesheetTemplate.xlsx exists in the parent directory.")

        if os.path.exists(demand_template_path):
            new_template_path = f"{folder_name}/{demandNameFileName}.docx"
            shutil.copy2(demand_template_path, new_template_path)
        else:
            print("Template file not found. Please ensure DemandTemplate.docx exists in the parent directory.")

        #Fill out info for invoice in new file
        if GenInvoice:
            invoice_path = f"{folder_name}/{invoiceFileName}"
            invoice_wb = openpyxl.load_workbook(invoice_path)
            invoice_ws = invoice_wb.active
            invoice_ws["D3"] = invoiceName
            invoice_ws["D5"] = today.strftime("%Y-%m-%d")
            # Populate the table body with df2 data
            for i, row_data in enumerate(df2.itertuples(index=False), start=14):
                invoice_ws[f'B{i}'] = row_data[0]  # Time Frame
                invoice_ws[f'C{i}'] = row_data[1]  # Description
                invoice_ws[f'D{i}'] = row_data[2]  # Hours
                invoice_ws[f'E{i}'] = row_data[3]  # Total
            invoice_wb.save(invoice_path)
            print("Invoice details filled out.")

        #Fill out info for timesheet in new file
        if GenTimesheet:
            timesheet_path = f"{folder_name}/{timesheetFileName}"
            timesheet_wb = openpyxl.load_workbook(timesheet_path)
            timesheet_ws = timesheet_wb.active
            timesheet_ws["D4"] = end_date_str
            timesheet_ws["D5"] = today.strftime("%Y-%m-%d")
            table1_range = timesheet_ws.tables["Table1"].ref
            table_last_row = int(table1_range.split(":")[1][1:])

            # Populate the table body with df3 data
            for i, row_data in enumerate(df3.itertuples(index=False), start=14):
                timesheet_ws[f'B{i}'] = row_data[0]  # Date
                timesheet_ws[f'C{i}'] = row_data[1]  # Description
                timesheet_ws[f'D{i}'] = row_data[2]  # Hours

            num_deleted_rows = 0
            # Shorten table starting from 42 going upwards, deleting rows that are empty until we hit a row with data, then stop
            for row_idx in range(table_last_row-1, 14, -1):
                # Check if all cells in the row are empty
                if timesheet_ws[f'D{row_idx}'].value is None:
                    timesheet_ws.delete_rows(row_idx)
                    num_deleted_rows += 1

            # Update the table reference to account for deleted rows + totals rows in range
            timesheet_ws.tables["Table1"].ref = table1_range.replace(f"{table_last_row}", f"{table_last_row - num_deleted_rows}")

            timesheet_path_abs = os.path.abspath(timesheet_path)
            timesheet_wb.save(timesheet_path_abs)
            print("Timesheet details filled out.")
            excel = None
            wb = None
            try:
                excel = client.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(timesheet_path_abs)
                # 0 = PDF format
                wb.ActiveSheet.ExportAsFixedFormat(0, timesheetPdfName)
                print("Timesheet PDF generated.")
            except Exception as e:
                print("Error during Timesheet PDF generation:", e)
            finally:
                if wb is not None:
                    try:
                        wb.Close(False)
                    except Exception:
                        pass
                if excel is not None:
                    try:
                        excel.Quit()
                    except Exception:
                        pass
                try:
                    del wb
                except NameError:
                    pass
                try:
                    del excel
                except NameError:
                    pass

            # Move the excel document to a folder named "Non-PDF Files" within the same directory if PDF generation was successful  
            if os.path.exists(timesheet_path_abs) and os.path.exists(timesheetPdfName):
                non_pdf_folder = os.path.join(folder_name, "Non-PDF Files")
                os.makedirs(non_pdf_folder, exist_ok=True)
                shutil.move(timesheet_path_abs, os.path.join(non_pdf_folder, f"{timesheetFileName}"))

        if GenDemand:
            #Update parts of word document for demand form
            demand_path = f"{folder_name}/{demandNameFileName}.docx"
            document = Document(demand_path)

            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    if "DATE " in run.text:
                        run.text = f"DATE {today.strftime('%m-%d-%Y')}"
                    if "99999.00" in run.text:
                        run.text = f"{df2['Total'].sum():.2f}"
                    if "****" in run.text:
                        run.text = num2words(df2['Total'].sum(), to='currency', lang='en').replace("euro", "dollars").replace("cents", "cents")

            # Setup string date writing
            write_count = 0
            todaySTR = str({today.strftime('%m-%d-%Y')})[2:-2]
            write_to = len(todaySTR)

            # Update specific cells from the template with the total amount and date
            for table in document.tables:
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        start_write = False
                        # Update billed amount
                        if cell.text == "$99999.00":
                            cell.text = f"${df2['Total'].sum():.2f}"
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if start_write and write_count <= write_to:
                                    run.text = ""
                                if start_write and write_count < write_to:
                                    if run.text is not None:
                                        run.text = todaySTR[write_count]
                                        write_count += 1
                                if "DATE " in run.text and write_count < write_to:
                                    start_write = True

            document.save(demand_path)
            print("Demand form details filled out. Please confirm the spelled out dollar amount.")

            # Generate PDF versions of invoice and demand form
            try:
                from docx2pdf import convert
                convert(demand_path)
                print("Demand form PDF generated.")
                # Move the word document to a folder named "Non-PDF Files" within the same directory
                non_pdf_folder = os.path.join(folder_name, "Non-PDF Files")
                os.makedirs(non_pdf_folder, exist_ok=True)
                shutil.move(demand_path, os.path.join(non_pdf_folder, f"{demandNameFileName}.docx"))
            except ImportError:
                print("docx2pdf not installed. Skipping PDF generation for demand form. Please install docx2pdf to enable this feature.")


if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except Exception:
        import traceback
        traceback.print_exc()
