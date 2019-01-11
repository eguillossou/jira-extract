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

def construct_datas(issues_in, header_list_in):
    len(header_list)
    matrix = [title for title in header_list_in]
    values_issues = [
        [issue_ids.fields.issuetype.name, 
        issue_ids.key, 
        issue_ids.fields.summary, 
        issue_ids.fields.customfield_10150, 
        issue_ids.fields.status.name, 
        issue_ids.fields.priority.name,
        issue_ids.fields.customfield_11070] for issue_ids in issues ]

    # values_issues = [
    #     log([issue_ids.fields.issuetype.name, 
    #     issue_ids.key, 
    #     issue_ids.fields.summary, 
    #     issue_ids.fields.customfield_10150, 
    #     issue_ids.fields.status.name, 
    #     issue_ids.fields.priority.name,
    #     issue_ids.fields.customfield_11070]) for issue_ids in issues ]

    matrix = [values_issues for title in header_list_in]

    return matrix

def expression(ws_in, title_in, column_nb_in):
    ws_in.cell(row=1, column=column_nb_in+1).value = title_in

def getColumn(lst, col):
    return [i[col] for i in lst]

def fill_datas(ws_in, matrix_in):
    for idx in range(len(matrix_in)):
        for idy in range(len(matrix_in[0])):
            ws_in.cell(row=idx, column=idy).value = matrix_in[idx][idy]
            print("fill rowidx={}, colidy={}, value={}".format(idx, idy, matrix_in[idx][idy]))

    # [ expression(ws_in, title, header_list_in.index(title)) for idx in len(matrix_in[0])]

if __name__ == '__main__':

    if('PASSWORD_JIRA' not in os.environ):
        log("Missing environment variable PASSWORD_JIRA.")
        sys.exit(1)

    jira_options={'server': 'https://jira.talendforge.org/','agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=('eguillossou',os.environ['PASSWORD_JIRA']))
    sp_nb =  get_selected_sprint_number(arguments().parse_args().sp, jira)
    issues=jira.search_issues(construct_jql_query(sp_nb,jira), startAt=0, maxResults=500, validate_query=True, fields=None, expand=None, json_result=None)

    print("\nissues\n")

    #Column to fill
    # Issue type	Issue key	Summary	Custom field (Story Points)	Status	Priority	Sprint	Already started before	Added after started

    
    for issue_ids in issues:
        log(issue_ids.fields.issuetype.name)
        log(issue_ids.key)
        log(issue_ids.fields.summary)
        log(issue_ids.fields.customfield_10150)
        log(issue_ids.fields.status.name)
        log(issue_ids.fields.priority.name)
        log(', '.join(re.findall(r"name=[^,]+",str(issue_ids.fields.customfield_11070) )).replace("name=",""))
        # log(str(issue_ids.fields.customfield_11070[0]))

        # import pdb;pdb.set_trace()
        # for issue in issue_ids:
        #     log(issue.fields)

    wb = Workbook()
    ws = wb.active
    ws1 = wb.create_sheet("Data")
    header_list = ["Issue type", "Issue key", "Summary", "Custom field (Story Points)", "Status", "Priority", "Sprint", "Already started before", "Added after started"]
    
    # fill_header(ws1,header_list)
    # fill_datas(ws1, construct_datas(issues, header_list))
    # ws1.append([1, 2, 3])
    print(sp_nb)
    # wb.save("c:\\Users\\eguillossou\\jira-report{}.xlsx".format(sp_nb))