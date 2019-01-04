from jira import JIRA
import argparse

def arguments():
    return 1

if __name__ == '__main__':
    jira_options={'server': 'https://jira.talendforge.org/','agile_rest_path': 'agile'}
    jira=JIRA(options=jira_options,basic_auth=('eguillossou',''))