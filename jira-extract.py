from jira import JIRA
import argparse
import os
import sys
from openpyxl import Workbook
import re
import requests

from dateutil.parser import parse

# For windows
# $env:PASSWORD_JIRA='YourPassword'
# python .\jira-extract.py -sp 24
# for debug purpose : import pdb;pdb.set_trace()
# issue_ids_in.fields.__dict__ to have a struct
#pp issue_ids_in.changelog.histories
def log(str):
    print(str)

#tdc board = 217
TDC_JIRA_BOARD_ID = 217
EXCEL_FILE_NAME = "jira-report"
JIRA_URL = "https://jira.talendforge.org/"
TDC_JIRA_SPRINT_PAGINATION = 30

if 'USER_JIRA' not in os.environ:
    log("CONFIG: USER_JIRA environment variable set to default")
USER_LOGIN = os.getenv('USER_JIRA','eguillossou')
PATH_EXCEL_FILE = "c:\\Users\\{}\\".format(USER_LOGIN)

if 'PASSWORD_JIRA' in os.environ:
    USER_PASSWORD = os.environ['PASSWORD_JIRA']
else:
    log("CONFIG: Missing environment variable PASSWORD_JIRA.")
    sys.exit(1)

def arguments():
    parser = argparse.ArgumentParser(description='Launch extraction and process.')
    parser.add_argument('--sp','-sp', type=int,help='Sprint number selection (default active sprint)')
    return parser

def get_active_sprint(handler_jira):
    active_sprint = handler_jira.sprints(TDC_JIRA_BOARD_ID,extended=False, startAt=0, maxResults=1, state="active")
    if(len(active_sprint) ==0):
        active_sprint = handler_jira.sprints(TDC_JIRA_BOARD_ID,extended=False, startAt=0, maxResults=1, state="future")
        print("Future sprint {} selected".format(active_sprint[0].name))
    else:
        print("Active sprint {}".format(active_sprint[0].name))
    return active_sprint[0]

def get_selected_sprint_number(sprint_details_in):
    return int(re.search(r"\d+",str(sprint_details_in)).group())

def get_sprints_list(handler_jira):
    return(handler_jira.sprints(TDC_JIRA_BOARD_ID,extended=True, startAt=TDC_JIRA_SPRINT_PAGINATION))

def get_selected_sprint(sp_nb_in, sprints_list_in):
    return(sprints_list_in[sp_nb_in])

def get_sprint_details(handler_jira, arg_sp , sprints_list_in):
    if(arg_sp is not None):
        # out=[sprints_list_in[i] for i in range(0,len(sprints_list_in)) if sprints_list_in[i].name == 'TDC Sprint {}'.format(arg_sp)]
        for i in range(0,len(sprints_list_in)):
            if sprints_list_in[i].name == 'TDC Sprint {}'.format(arg_sp):
                return(sprints_list_in[i])
        # return(out)
    else:
        return(get_active_sprint(handler_jira))

def get_sprint_start_date(sprint_details_in, field_11070):
    regexp=re.compile('name={}[ ]([0-9]+)[,]startDate=([^,]+)'.format(get_selected_sprint_number(sprint_details_in)))
    
    return re.findall(regexp,str(field_11070))[0].replace("startDate=","")

def construct_jql_query(sp_nb, handler_jira):
    jql_qry='project = TDC AND issuetype in (Bug, "New Feature", "Work Item") AND Sprint = "TDC Sprint {}" ORDER BY labels ASC, RANK'.format(sp_nb)
    return jql_qry

def pad_or_truncate(some_list, target_len):
    return some_list[:target_len] + [0]*(target_len - len(some_list))

def if_already_started(sprints_list_in,sprint_number_in):
    result = re.search(r"\d+",sprints_list_in.split(',')[0])
    if result is not None and int(result.group(0)) < sprint_number_in:
        return(True)
    else:
        return(False)

def if_added_after_started(issue_ids_in, issuelist_added_to_sprint_in):
    return issue_ids_in.key in issuelist_added_to_sprint_in.keys()

def get_reports_from_jira(jira_talend_in, sprint_id_in, user_in, password_in):
    url = '{}rest/greenhopper/1.0/rapid/charts/sprintreport?rapidViewId=217&sprintId={}'.format(jira_talend_in, sprint_id_in)
    headers = {'content-type': 'application/json'}
    r = requests.get(url, headers, auth=(user_in,password_in))
    return(r.json()['contents']['issueKeysAddedDuringSprint'])

def parse_sprints(field_11070):
    return ', '.join(re.findall(r"name=[^,]+",str(field_11070) )).replace("name=","")

def construct_datas(header_list_in, values_issues_in):
    issues_according_to_header_list = [pad_or_truncate(values_issues_in[idx],len(header_list_in)) for idx in range(len(values_issues_in)) ]

    return issues_according_to_header_list

def fillIT(issue_ids_in, sprint_details_in, sprint_number_in, issuelist_added_to_sprint_in):
    customfield_11070 = issue_ids_in.fields.customfield_11070
    sprint_list = parse_sprints(customfield_11070)

    return [issue_ids_in.fields.issuetype.name, 
            issue_ids_in.key,
            issue_ids_in.fields.summary,
            issue_ids_in.fields.customfield_10150,
            issue_ids_in.fields.status.name,
            issue_ids_in.fields.priority.name,
            sprint_list, 
            if_already_started(sprint_list, sprint_number_in),
            if_added_after_started(issue_ids_in, issuelist_added_to_sprint_in),
            ', '.join(issue_ids_in.fields.labels),
            issue_ids_in.fields.customfield_11071]

def fill_cell(ws_in, line_in, col_in, value_in):
    ws_in.cell(row=line_in, column=col_in).value = value_in
    # log("fill rowidx=1, colidy={}, value={}".format(col_in, value_in))

def fill_headers(ws_in, header_list_in):
    [fill_cell(ws_in, 1, idx+1, header_list_in[idx]) for idx in range(len(header_list_in))]

def fill_values(ws_in, lineidx_in, issues_line_in, header_list_in):
    [fill_cell(ws_in, lineidx_in+1, linecol+1, issues_line_in[lineidx_in][linecol]) for linecol in range(len(header_list_in))]

def fill_headers_and_values(ws_in, header_list_in, lineidx_in, issues_line_in):
    if(lineidx_in == 0):
        fill_headers(ws_in, header_list_in)
    else:
        fill_values(ws_in, lineidx_in, issues_line_in, header_list_in)


def save_excel_file(wb_in, sprint_number_in):
    my_file = "{}{}{}.xlsx".format(PATH_EXCEL_FILE, EXCEL_FILE_NAME, sprint_number_in)
    log("Saving file to : "+my_file)
    try:
        wb_in.save(my_file)
    except IOError:
        print("File already opened. Closed it first.")

if __name__ == '__main__':


    jira_options={'server': JIRA_URL ,'agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=(USER_LOGIN , USER_PASSWORD))

    sprint_list=get_sprints_list(jira)
    sprint_details=get_sprint_details(jira, arguments().parse_args().sp, sprint_list)
    sprint_number=get_selected_sprint_number(sprint_details)
    
    issuelist_added_to_sprint = get_reports_from_jira(JIRA_URL, sprint_details.id, USER_LOGIN , USER_PASSWORD)

    issues=jira.search_issues(construct_jql_query(sprint_number,jira), startAt=0, maxResults=500, validate_query=True, fields=None, expand="changelog", json_result=None)

    wb = Workbook()
    ws = wb.active
    ws.title ="Data"

    #Column to fill
    # Issue type	Issue key	Summary	Custom field (Story Points)	Status	Priority	Sprint	Already started before	Added after started
    header_list = ["Issue type", "Issue key", "Summary", "Custom field (Story Points)", "Status", "Priority", "Sprint", "Already started before", "Added after started", "Labels", "Epic Link"]
    
    values_issues = [ fillIT(issue_ids, sprint_details, sprint_number, issuelist_added_to_sprint) for issue_ids in issues ]
    issues_line = construct_datas(header_list, values_issues)

    [fill_headers_and_values(ws, header_list, lineidx, issues_line) for lineidx in range(len(issues))]

    print("Sprint number : {}".format(sprint_number))
    save_excel_file(wb, sprint_number)