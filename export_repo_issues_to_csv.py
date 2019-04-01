"""
Exports Issues from a repository to an Excel file
Uses basic authentication (Github API Token and Zenhub API Token)
to retrieve Issues from a repository that token has access to.
Supports Github API v3 and ZenHubs current working API.
Derived from https://gist.github.com/Kebiled/7b035d7518fdfd50d07e2a285aff3977
"""
# pylint: disable=W0622
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

def throttle_zenhub(issue_cnt):
    """
    Wait added for the ZenHub api rate limit of 100 requests per minute,
    wait after the rate limit - 1 issues have been processed
    :param issue_cnt: the current issue count
    """
    if issue_cnt%(ZENHUB_API_RATE_LIMIT-1) == 0:
        print(f'{issue_cnt} issues processed')
        time.sleep(45)

def get_assignees(issue):
    """
    Convert the assignees for an issue to a comma-separated string
    :param issue: the current issue
    :returns s_assignees: the concatenated list of assignees
    """
    s_assignee_list = ''
    for assignee in issue['assignees'] if issue['assignees'] else []:
        s_assignee_list += assignee['login'] + ','
    return s_assignee_list

def get_epics_string(issues, issue):
    """
    Convert the epics for an issue to a comma-separated string
    :param issues: a dictionary mapping of the issues to epics
    :param issue: the current issue
    :returns s_epics: the concatenated list of epic ids
    """
    s_epics = ''
    try:
        epics = issues[issue['number']]['epic_issue']
        for epic in epics if epics else []:
            s_epics += str(epic) + ','
    except KeyError:
        s_epics = ''
    return s_epics

def get_zenhubresponse(repo_id, issue_number):
    """
    Gets the ZenHub data for the issue
    :param repo_id: the id for the repo used to reference the ZenHub fields
    :param issue_number: the specific issue number
    :returns zen_r: a json object of the issue
    """
    zen_url = f'https://api.zenhub.io/p1/repositories/{repo_id}/issues/'
    issue_url = f'{zen_url}{issue_number}?{ACCESS_TOKEN}'
    zen_response = requests.get(issue_url)
    if not zen_response.status_code == 200:
        raise Exception(zen_response.status_code)
    zen_r = zen_response.json()
    s_pipeline = zen_r.get("pipeline", dict()).get('name', "")
    estimate_value = zen_r.get("estimate", dict()).get('value', "")
    return s_pipeline, estimate_value

def get_labels_string(issue):
    """
    concatenates labels into a string
    :param issue: the current issue
    :returns s_labels: the string of comma-separated labels
    :returns s_priority: the priority for the issue
    """
    s_priority = ''
    s_labels = ''
    for label in issue['labels'] if issue['labels'] else []:
        rules = [
            'Low' in label['name'],
            'Medium' in label['name'],
            'High' in label['name'],
        ]
        s_labels += label['name'] + ','
        if any(rules):
            s_priority = label['name']
    return s_labels, s_priority

def write_row(issue, repo_name, repo_id, userstory, s_assignee_list,
              s_priority, s_labels, s_epics, issue_cnt):
    """
    Writes rows to an Excel file
    """
    issue_number = str(issue['number'])
    s_pipeline, estimate_value = get_zenhubresponse(repo_id, issue_number)
    comments = ''
    if issue['comments'] > 0:
        comments = get_comments(repo_name, issue_number)
    if (not s_priority and s_pipeline == 'Closed'):
        s_priority = 'High'
    rowvalues = [repo_name, issue['number'], issue['title'],
                 userstory, s_pipeline, issue['user']['login'], issue['created_at'],
                 issue['milestone']['title'] if issue['milestone']
                 else "", issue['milestone']['due_on'] if issue['milestone'] else "",
                 s_assignee_list[:-1], estimate_value, s_priority, s_labels,
                 comments, s_epics[:-1]]
    for i in range(len(rowvalues)):
        WS.cell(column=(i+1), row=1+issue_cnt, value=rowvalues[i])

def write_issues(r_json, repo_name, repo_id, issues, issue_cnt):
    """
    Writes issues to an Excel file
    :param git_response: the response for the github call
    :param repo_name: the name of the github repo
    :param repo_id: the id for the repo used to reference the ZenHub fields
    :param issues: dictionary of the mapping for issues to epics
    :param issue_cnt: counter for the starting issue
    :return issue_cnt: the count of issues processed
    """
    for issue in r_json:
        issue_cnt += 1
        s_assignee_list = get_assignees(issue)
        s_epics = get_epics_string(issues, issue)
        s_labels, s_priority = get_labels_string(issue)
        if HTMLFLAG == 1:
            userstory = markdown.markdown(issue['body'])
        else:
            userstory = issue['body']
        write_row(issue, repo_name, repo_id, userstory, s_assignee_list,
                  s_priority, s_labels, s_epics, issue_cnt)
        print(f'{issue_cnt}')
        throttle_zenhub(issue_cnt)
    return issue_cnt

def get_pages(issue_response):
    """
    gets additional pages information
    :param issue_response: the response for the github call for the issue
    :returns pages_dict: a dictionary object for the pages link
    """
    pages_dict = dict(
        [(rel[6:-1], url[url.index('<') + 1:-1]) for url, rel in
         [link.split(';') for link in
          issue_response.headers['link'].split(',')]])
    return pages_dict

def get_nextpage_response(pages):
    """
    gets the next page information
    :param pages: a dictionary object for the pages link
    :returns issue_response.json(): a json document of the next page link
    """
    issue_response = requests.get(pages['next'], auth=AUTH)
    if not issue_response.status_code == 200:
        raise Exception(issue_response.status_code)
    return issue_response.json()

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
    if not issue_response.status_code == 200:
        raise Exception(issue_response.status_code)
    response = issue_response.json()
    issue_count = write_issues(response, repo_name, repo_id, issues_dict, 0)
    # more pages? examine the 'link' header returned
    if 'link' in issue_response.headers:
        pages = get_pages(issue_response)
        while 'last' in pages and 'next' in pages:
            pages = dict(
                [(rel[6:-1], url[url.index('<') + 1:-1]) for url, rel in
                 [link.split(';') for link in
                  issue_response.headers['link'].split(',')]])
            issue_response = requests.get(pages['next'], auth=AUTH)
            response = issue_response.json()
            issue_count = write_issues(response, repo_name, repo_id, issues_dict, issue_count)
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
ISSUES = create_epic_dict(REPO_LIST[1])

#for repo_data in REPO_LIST:
get_issues(REPO_LIST, ISSUES)
FILEOUTPUT.save(filename=FILENAME)
