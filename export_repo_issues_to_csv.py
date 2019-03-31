"""
Exports Issues from a repository to an Excel file
Uses basic authentication (Github API Token and Zenhub API Token)
to retrieve Issues from a repository that token has access to.
Supports Github API v3 and ZenHubs current working API.
Derived from https://gist.github.com/Kebiled/7b035d7518fdfd50d07e2a285aff3977
"""

#!/usr/bin/env python
import argparse
import os
import time
import requests
import markdown
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.styles import Font

class AttrDict(dict):
    """A dictionary where you can use dot notation for accessing elements."""
    def __getattr__(self, key):
        if key not in self:
            raise AttributeError(key)
        return self[key]

    def __setattr__(self, key, value):
        self[key] = value

    def __delattr__(self, key):
        del self[key]

def get_epics(repo_id):
    """
    Get the epic(s) on an issue and concatenates them into
    one field.
    :param repo_name: the name of the github repo
    :out epic_sum: a concantenated field with all epics
    """
    zen_url = f'https://api.zenhub.io/p1/repositories/{repo_id}/epics/?{ACCESS_TOKEN}'
    zen_response = requests.get(zen_url)
    if not zen_response.status_code == 200:
        raise Exception(zen_response.status_code)
    r_json = zen_response.json()
    return r_json

def get_epic_issues(repo_id, epic_issue_id):
    """
    Get the epic(s) on an issue and concatenates them into
    one field.
    :param repo_name: the name of the github repo
    :out epic_sum: a concantenated field with all epics
    """
    zen_url = f'https://api.zenhub.io/p1/repositories/'
    epic_url = f'{zen_url}{repo_id}/epics/{epic_issue_id}?{ACCESS_TOKEN}'
    zen_response = requests.get(epic_url)
    if not zen_response.status_code == 200:
        raise Exception(zen_response.status_code)
    r_json = zen_response.json()
    return r_json

def create_epic_dict(repo_id):
    """
    Create a dictionary for looking up epics by issue number
    :param repo_id: the repo_id of the zenhub repo
    :out issue_epics: a dictionary of issues with the epics it is under
    """
    response = get_epics(repo_id)
    issue_epics = dict()

    for epic_issue in response['epic_issues'] if response['epic_issues'] else []:
        epic_issue = epic_issue['issue_number']
        epic_issues = get_epic_issues(repo_id, epic_issue)
        for issue in epic_issues['issues'] if epic_issues['issues'] else []:
            issue_number = issue['issue_number']
            if issue_number in issue_epics:
                temp = dict()
                epic = issue_epics[issue_number].epic_issue
                epic.append(epic_issue)
                temp[issue_number] = AttrDict(
                    issue_number=issue_number,
                    epic_issue=epic
                )
                issue_epics.update(temp)
            else:
                issue_epics[issue_number] = AttrDict(
                    issue_number=issue_number,
                    epic_issue=[epic_issue]
                )
    print('waiting after creating the dictionary')
    time.sleep(45)
    return issue_epics

def get_comments(repo_name, issue_id):
    """
    Get the comments on an issue and concatenates them into
    one field.
    :param repo_name: the name of the github repo
    :param issue_id: the issue number
    :out comment_sum: a concantenated field with all comments
    """
    comments_for_issue_url = f'https://api.github.com/repos/{repo_name}/issues/{issue_id}/comments'
    git_response = requests.get(comments_for_issue_url, auth=AUTH)
    r_json = git_response.json()
    comment_sum = ''
    for comment in r_json:
        c_login = comment.get("user", dict()).get('login', "")
        comment_sum = '@'+c_login+' - '+comment_sum + str(comment['body'])
    return comment_sum

def write_issues(git_response, repo_name, repo_id, issues):
    """
    Writes issues to an Excel file
    :param git_response: the response for the github call
    :param repo_name: the name of the github repo
    :param repo_id: the id for the repo used to reference the ZenHub fields
    """
    if not git_response.status_code == 200:
        raise Exception(git_response.status_code)

    r_json = git_response.json()
    zen_url = f'https://api.zenhub.io/p1/repositories/{repo_id}/issues/'
    for issue in r_json:
        issue_number = str(issue['number'])
        issue_url = f'{zen_url}{issue_number}?{ACCESS_TOKEN}'
        zen_response = requests.get(issue_url)
        if not zen_response.status_code == 200:
            raise Exception(zen_response.status_code)
        zen_r = zen_response.json()
        #call here to get all comments
        comments = ''
        if issue['comments'] > 0:
            comments = get_comments(repo_name, issue_number)


        global ISSUES
        ISSUES += 1
        s_assignee_list = ''
        s_priority = ''
        s_pipeline = ''
        s_labels = ''
        s_epics = ''
        for i in issue['assignees'] if issue['assignees'] else []:
            s_assignee_list += i['login'] + ','
        #add output of the payload for records not found
        s_pipeline = zen_r.get("pipeline", dict()).get('name', "")
        try:
            epics = issues[issue['number']]['epic_issue']
            for epic in epics if epics else []:
                s_epics += str(epic) + ','
        except KeyError:
            s_epics = ''
        for label in issue['labels'] if issue['labels'] else []:
            rules = [
                'Low' in label['name'],
                'Medium' in label['name'],
                'High' in label['name'],
            ]
            s_labels += label['name'] + ','
            if any(rules):
                s_priority = label['name']
        if (not s_priority and s_pipeline == 'Closed'):
            s_priority = 'High'
        estimate_value = zen_r.get("estimate", dict()).get('value', "")
        if HTMLFLAG == 1:
            userstory = markdown.markdown(issue['body'])
        else:
            userstory = issue['body']
        rowvalues = [repo_name, issue['number'], issue['title'],
                     userstory, s_pipeline, issue['user']['login'], issue['created_at'],
                     issue['milestone']['title'] if issue['milestone']
                     else "", issue['milestone']['due_on'] if issue['milestone'] else "",
                     s_assignee_list[:-1], estimate_value, s_priority, s_labels, comments, s_epics[:-1]]
        for i in range(len(rowvalues)):
            WS.cell(column=(i+1), row=1+ISSUES, value=rowvalues[i])
        print(f'{ISSUES}')
        #Wait added for the ZenHub api rate limit of 100 requests per minute,
        #wait after the rate limit - 1 issues have been processed
        if ISSUES%(ZENHUB_API_RATE_LIMIT-1) == 0:
            print(f'{ISSUES} issues processed')
            time.sleep(45)

def get_issues(repo_data, issues_dict):
    """
    Get an issue attributes
    :param repo_data: the environment variable with the repo_name
    and the ZenHub id for the repository
    """
    repo_name = repo_data[0]
    repo_id = repo_data[1]
    issues_for_repo_url = f'https://api.github.com/repos/{repo_name}/issues?state=all'
    issue_response = requests.get(issues_for_repo_url, auth=AUTH)
    write_issues(issue_response, repo_name, repo_id, issues_dict)
    # more pages? examine the 'link' header returned
    if 'link' in issue_response.headers:
        pages = dict(
            [(rel[6:-1], url[url.index('<') + 1:-1]) for url, rel in
             [link.split(';') for link in
              issue_response.headers['link'].split(',')]])
        while 'last' in pages and 'next' in pages:
            pages = dict(
                [(rel[6:-1], url[url.index('<') + 1:-1]) for url, rel in
                 [link.split(';') for link in
                  issue_response.headers['link'].split(',')]])
            issue_response = requests.get(pages['next'], auth=AUTH)
            write_issues(issue_response, repo_name, repo_id, issues)
            if pages['next'] == pages['last']:
                break

PARSER = argparse.ArgumentParser()
PARSER.add_argument('--file_name', help='file_name=filename.txt')
PARSER.add_argument('--repo_list', nargs='+', help='repo_list owner/repo zenhub-id')
PARSER.add_argument('--html', default=0, type=int, help='html=1')
ARGS = PARSER.parse_args()

REPO_LIST = ARGS.repo_list
AUTH = ('token', os.environ['AUTH'])
ACCESS_TOKEN = os.environ['ACCESS_TOKEN']
ZENHUB_API_RATE_LIMIT = 101

TXTOUT = open('data.json', 'w')
ISSUES = 0
FILENAME = ARGS.file_name
HTMLFLAG = ARGS.html

FILEOUTPUT = Workbook()

WS = FILEOUTPUT.create_sheet(title="Data")
SH = FILEOUTPUT['Sheet']
FILEOUTPUT.remove(SH)

HEADERS = ['Repository', 'Issue Number', 'Issue Title', 'User Story', 'Pipeline', 'Issue Author',
           'Created At', 'Milestone', 'Milestone End Date', 'Assigned To',
           'Estimate Value', 'Priority', 'Labels', 'Comments', 'Epics']
for h in range(len(HEADERS)):
    WS.cell(column=(h+1), row=1, value=HEADERS[h])
    WS.cell(column=(h+1), row=1).font = Font(bold=True)

#get the epic dictionary
issues = create_epic_dict(143072984)

#for repo_data in REPO_LIST:
get_issues(REPO_LIST, issues)
FILEOUTPUT.save(filename=FILENAME)
