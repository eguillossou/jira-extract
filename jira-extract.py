from jira import JIRA
import argparse
import os
import sys

# For windows
# $env:PASSWORD_JIRA='YourPassword'

def log(str):
    print(str)

def arguments():
    parser = argparse.ArgumentParser(description='Launch extraction and process.')
    parser.add_argument('--sp','-sp', help='Sprint selection (default active sprint)')
    return parser

def fool():
    return

if __name__ == '__main__':

    if('PASSWORD_JIRA' not in os.environ):
        log("Missing environment variable PASSWORD_JIRA.")
        sys.exit(1)

    args = arguments().parse_args()

    jira_options={'server': 'https://jira.talendforge.org/','agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=('eguillossou',os.environ['PASSWORD_JIRA']))

    jql_str='project = TDC AND issuetype in (Bug, "New Feature", "Work Item") AND Sprint = "{}" ORDER BY labels ASC, RANK'.format("TDC Sprint 24")
    issues=jira.search_issues(jql_str, startAt=0, maxResults=500, validate_query=True, fields=None, expand=None, json_result=None)

    for issue_names in issues:
        log('issue id - {}'.format(issue_names))