from jira import JIRA
import argparse
import os
import sys

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
        print("future sprint {} selected".format(active_sprint[0].name))
    else:
        print("active sprint {} selected".format(active_sprint[0].name))

    return active_sprint[0].name

def construct_jql_query(arg_sp, handler_jira):
    jql_qry = 'project = TDC AND issuetype in (Bug, "New Feature", "Work Item") AND Sprint = "TDC Sprint 20" ORDER BY labels ASC, RANK'
    if(arg_sp is not None):
        print("Sprint \"TDC Sprint {}\" selected".format(arg_sp))
        jql_qry='project = TDC AND issuetype in (Bug, "New Feature", "Work Item") AND Sprint = "TDC Sprint {}" ORDER BY labels ASC, RANK'.format(arg_sp)
    else:
        jql_qry='project = TDC AND issuetype in (Bug, "New Feature", "Work Item") AND Sprint = "{}" ORDER BY labels ASC, RANK'.format(get_active_sprint(handler_jira))

    return jql_qry

if __name__ == '__main__':

    if('PASSWORD_JIRA' not in os.environ):
        log("Missing environment variable PASSWORD_JIRA.")
        sys.exit(1)

    jira_options={'server': 'https://jira.talendforge.org/','agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=('eguillossou',os.environ['PASSWORD_JIRA']))
    issues=jira.search_issues(construct_jql_query(arguments().parse_args().sp,jira), startAt=0, maxResults=500, validate_query=True, fields=None, expand=None, json_result=None)

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
        log(issue_ids.fields.customfield_11070)

        # import pdb;pdb.set_trace()
        # for issue in issue_ids:
        #     log(issue.fields)
