from jira import JIRA
import argparse
import os
import sys
from openpyxl import Workbook
import re

# For windows
# $env:PASSWORD_JIRA='YourPassword'

#tdc board = 217
TDC_JIRA_BOARD_ID = 217

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
        print("Active sprint {} selected".format(active_sprint[0].name))

    return active_sprint[0].name

def get_selected_sprint_number(arg_sp, handler_jira):
    if(arg_sp is not None):
        return arg_sp
    else:
        return re.findall(r"\d+",get_active_sprint(handler_jira))[0]

def construct_jql_query(sp_nb, handler_jira):
    jql_qry='project = TDC AND issuetype in (Bug, "New Feature", "Work Item") AND Sprint = "TDC Sprint {}" ORDER BY labels ASC, RANK'.format(sp_nb)

    return jql_qry

def fill_header(ws_in, header_list_in):
    [ expression(ws_in, title, header_list_in.index(title)) for title in header_list_in]

def pad_or_truncate(some_list, target_len):
    return some_list[:target_len] + [0]*(target_len - len(some_list))

def construct_datas(issues_in, header_list_in):
    values_issues = [
        [issue_ids.fields.issuetype.name, 
        issue_ids.key, 
        issue_ids.fields.summary, 
        issue_ids.fields.customfield_10150, 
        issue_ids.fields.status.name, 
        issue_ids.fields.priority.name,
        parse_sprints(issue_ids.fields.customfield_11070)] for issue_ids in issues_in ]
    issues_according_to_header_list = [pad_or_truncate(values_issues[idx],len(header_list_in)) for idx in range(len(values_issues)) ]

    return issues_according_to_header_list

def expression(ws_in, header_list_in, column_nb_in):
    ws_in.cell(row=1, column=column_nb_in+1).value = header_list_in[column_nb_in]
    # log("fill rowidx=1, colidy={}, value={}".format(column_nb_in+1, header_list_in[column_nb_in]))

# def getColumn(lst, col):
#     return [i[col] for i in lst]

def fill_datas(ws_in, header_list_in, issues_in):
    issues_line = construct_datas(issues_in, header_list_in)
    #add one line +1 for headers:
    for lineidx in range(len(issues_in)):
            if(lineidx == 0):
                [(expression(ws_in, header_list_in, idx)) for idx in range(len(header_list_in))]
            else:
                for linecol in range(len(header_list_in)):
                    #excel cells starting at indexes 1,1
                    ws_in.cell(row=lineidx+1, column=linecol+1).value = issues_line[lineidx][linecol]
                    # log("fill rowidx excel={}, colidy excel={}, value={}".format(lineidx+1,linecol+1, issues_line[lineidx][linecol]))


def parse_sprints(field_11070):
    return ', '.join(re.findall(r"name=[^,]+",str(field_11070) )).replace("name=","")

if __name__ == '__main__':

    if('PASSWORD_JIRA' not in os.environ):
        log("Missing environment variable PASSWORD_JIRA.")
        sys.exit(1)

    jira_options={'server': 'https://jira.talendforge.org/','agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=('eguillossou',os.environ['PASSWORD_JIRA']))
    sp_nb =  get_selected_sprint_number(arguments().parse_args().sp, jira)
    issues=jira.search_issues(construct_jql_query(sp_nb,jira), startAt=0, maxResults=500, validate_query=True, fields=None, expand=None, json_result=None)

    wb = Workbook()
    ws1 = wb.create_sheet("Data")

    #Column to fill
    # Issue type	Issue key	Summary	Custom field (Story Points)	Status	Priority	Sprint	Already started before	Added after started
    header_list = ["Issue type", "Issue key", "Summary", "Custom field (Story Points)", "Status", "Priority", "Sprint", "Already started before", "Added after started"]
    
    fill_datas(ws1, header_list, issues)
    print("Sprint number : {}".format(sp_nb))
    wb.save("c:\\Users\\eguillossou\\jira-report{}.xlsx".format(sp_nb))