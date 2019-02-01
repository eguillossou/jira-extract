from jira import JIRA
import argparse
import os
import sys
from openpyxl import Workbook
import re

from dateutil.parser import parse

# For windows
# $env:PASSWORD_JIRA='YourPassword'
# python .\jira-extract.py -sp 24
# for debug purpose : import pdb;pdb.set_trace()
# issue_ids_in.fields.__dict__ to have a struct
#pp issue_ids_in.changelog.histories

#tdc board = 217
TDC_JIRA_BOARD_ID = 217
PATH_EXCEL_FILE = "c:\\Users\\eguillossou\\"
EXCEL_FILE_NAME = "jira-report"

def log(str):
    print(str)

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
    return(handler_jira.sprints(TDC_JIRA_BOARD_ID,extended=True, startAt=0))

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

def is_issue_by_name_added(item_in_history, name_detail):
    # fromString for field "sprint" in payload is empty and toString is empty
    if(item_in_history.fromString is None and item_in_history.toString is None):
        return(False)

    # fromString is empty and toString is not empty
    # if "TDC Sprint X" is present within toString field, issue has been added to sprint
    if(item_in_history.fromString is None and item_in_history.toString is not None):
        if(name_detail not in item_in_history.toString):
            return(False)
        else:
            return(True)

    # fromString is not empty and toString is empty
    # if "TDC Sprint X" is present within fromString field but not in toString, issue has been removed from sprint
    if(item_in_history.fromString is not None and item_in_history.toString is None):
        return(False)

    # fromString is not empty and toString is not empty at this step
    # if "TDC Sprint X" is present within fromString field but not in toString, issue has been removed from sprint
    if(name_detail in item_in_history.fromString and name_detail not in item_in_history.toString):
        return(False)

    # if "TDC Sprint X" present in toString, issue is new or still present in sprint
    if(name_detail in item_in_history.toString):
        return(True)    

    return(False)

def if_added_after_started(issue_ids_in, sprint_details_in, customfield_11070):
    starting_sprint_date = "2019-01-15T02:31:56.000-0600"
    # if(check_sprint_status() is Active)
    starting_sprint_date = sprint_details_in.startDate
    # get_sprint_start_date(sprint_details_in,customfield_11070)
    if(starting_sprint_date == '<null>'):
        return(False)
    
    log("key:"+issue_ids_in.key)
    if(str(issue_ids_in.key) == "TDC-1417"): 
        last_status = False
        for history in issue_ids_in.changelog.histories:
            if(history.created is not None and parse(history.created) < parse(starting_sprint_date)):
                l=""
                # log("{} jira ticket added to current sprint after sprint was started : {}".format(str(issue_ids_in.key), starting_sprint_date))
                # return(False)
            for item in history.items:
                # if item.field == 'status':
                #     print('Date:' + history.created + ' From:' + item.fromString + ' To:' + item.toString)
                if item.field == "Sprint":
                    last_status = is_issue_by_name_added(item,sprint_details_in.name)
                    if(last_status and history.created is not None and parse(history.created) < parse(starting_sprint_date) ):
                        print("key:{} added".format(issue_ids_in.key))
            print("history created: "+history.created+" sprint starting date: "+starting_sprint_date+ " last_status: "+str(last_status))

    return(False)

def parse_sprints(field_11070):
    return ', '.join(re.findall(r"name=[^,]+",str(field_11070) )).replace("name=","")

def construct_datas(header_list_in, values_issues_in):
    issues_according_to_header_list = [pad_or_truncate(values_issues_in[idx],len(header_list_in)) for idx in range(len(values_issues_in)) ]

    return issues_according_to_header_list

def fillIT(issue_ids_in, sprint_details_in, sprint_number_in):
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
            if_added_after_started(issue_ids_in, sprint_details_in, customfield_11070)]

def fill_cell(ws_in, line_in, col_in, value_in):
    ws_in.cell(row=line_in, column=col_in).value = value_in

def fill_headers(ws_in, header_list_in):
    [fill_cell(ws_in, 1, idx+1, header_list_in[idx]) for idx in range(len(header_list_in))]
    # log("fill rowidx=1, colidy={}, value={}".format(column_nb_in+1, header_list_in[column_nb_in]))

def fill_values(ws_in, lineidx_in, issues_line_in, header_list_in):
    [fill_cell(ws_in, lineidx_in+1, linecol+1, issues_line_in[lineidx_in][linecol]) for linecol in range(len(header_list_in))]

def fill_headers_and_values(ws_in, header_list_in, lineidx_in, issues_line_in):
    if(lineidx_in == 0):
        fill_headers(ws_in, header_list_in)
    else:
        fill_values(ws_in, lineidx_in, issues_line_in, header_list_in)


def save_excel_file(wb_in, sprint_number_in):
    wb_in.save("{}{}{}.xlsx".format(PATH_EXCEL_FILE, EXCEL_FILE_NAME, sprint_number))

if __name__ == '__main__':

    if('PASSWORD_JIRA' not in os.environ):
        log("Missing environment variable PASSWORD_JIRA.")
        sys.exit(1)

    jira_options={'server': 'https://jira.talendforge.org/','agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=('eguillossou',os.environ['PASSWORD_JIRA']))

    sprint_list=get_sprints_list(jira)
    sprint_details=get_sprint_details(jira, arguments().parse_args().sp, sprint_list)
    sprint_number=get_selected_sprint_number(sprint_details)
    
    issues=jira.search_issues(construct_jql_query(sprint_number,jira), startAt=0, maxResults=500, validate_query=True, fields=None, expand="changelog", json_result=None)

    wb = Workbook()
    ws = wb.active
    ws.title ="Data"

    #Column to fill
    # Issue type	Issue key	Summary	Custom field (Story Points)	Status	Priority	Sprint	Already started before	Added after started
    header_list = ["Issue type", "Issue key", "Summary", "Custom field (Story Points)", "Status", "Priority", "Sprint", "Already started before", "Added after started"]
    
    values_issues = [ fillIT(issue_ids, sprint_details, sprint_number) for issue_ids in issues ]
    issues_line = construct_datas(header_list, values_issues)

    [fill_headers_and_values(ws, header_list, lineidx, issues_line) for lineidx in range(len(issues))]

    print("Sprint number : {}".format(sprint_number))
    save_excel_file(wb, sprint_number)