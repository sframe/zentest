#!/usr/bin/env python

"""
Exports Issues from a repository to an Excel file
Uses basic authentication (Github API Token and Zenhub API Token)
to retrieve Issues from a repository that token has access to.
Supports Github API v3 and ZenHubs current working API.
Derived from https://gist.github.com/Kebiled/7b035d7518fdfd50d07e2a285aff3977
"""
import argparse
import os
import time
import datetime
import requests
import markdown
from retrying import retry
from openpyxl import Workbook
from openpyxl.styles import Font

GITHUB_TOKEN = os.environ['GITHUB_TOKEN']
ZENHUB_TOKEN = os.environ['ZENHUB_TOKEN']
ZENHUB_API_RATE_LIMIT = 51

GITHUB_HEADERS = {
    'Accept': 'application/vnd.github.v3+json',
    'Authorization': f'token {GITHUB_TOKEN}',
}

ZENHUB_HEADERS = {'X-Authentication-Token': ZENHUB_TOKEN}
HTMLFLAG = 0

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
    :param repo_id: the id of the github repo
    :out r_json: the response from the call
    """
    zen_url = f'https://api.zenhub.io/p1/repositories/{repo_id}/epics/'
    zen_response = requests.get(zen_url, headers=ZENHUB_HEADERS)
    if not zen_response.status_code == 200:
        raise Exception(zen_response.json())
    r_json = zen_response.json()
    return r_json

def get_dependencies(repo_id):
    """
    Get the dependencies on all issues
    :param repo_id: the id of the github repo
    :out r_json: the response from the call
    """
    zen_url = f'https://api.zenhub.io/p1/repositories/{repo_id}/dependencies/'
    zen_response = requests.get(zen_url, headers=ZENHUB_HEADERS)
    if not zen_response.status_code == 200:
        raise Exception(zen_response.json())
    r_json = zen_response.json()
    return r_json

def create_blocked_items(repo_id, issues):
    """
    Create a dictionary for looking up dependenceies by issue number
    :param repo_id: the repo_id of the zenhub repo
    :out issue_epics: a dictionary of the issue dependencies
    """
    response = get_dependencies(repo_id)
    blocked_items = dict()
    closed = dict()

    #find all closed issues
    for issue in issues:
        if issue['state'] == 'closed':
            closed[issue['number']] = AttrDict(
                issue_number=issue['number']
            )

    for dependency in response['dependencies'] if response['dependencies'] else []:
        issue_number = dependency['blocked']['issue_number']
        #if the issue is closed or the dependency is closed, ignore it
        if issue_number in closed or dependency['blocking']['issue_number'] in closed:
            print(f'Skipping {issue_number} dependency...')
            continue
        if issue_number in blocked_items:
            temp = dict()
            depends = blocked_items[issue_number].blocked_by
            depends.append(dependency['blocking']['issue_number'])
            temp[issue_number] = AttrDict(
                issue_number=issue_number,
                blocked_by=depends
            )
            blocked_items.update(temp)
        else:
            blocked_items[issue_number] = AttrDict(
                issue_number=dependency['blocked']['issue_number'],
                blocked_by=[dependency['blocking']['issue_number']]
            )
    return blocked_items

def get_epic_issues(repo_id, epic_issue_id):
    """
    Get the epic(s) on an issue and concatenates them into
    one field.
    :param repo_name: the name of the github repo
    :out epic_sum: a concantenated field with all epics
    """
    zen_url = 'https://api.zenhub.io/p1/repositories/'
    epic_url = f'{zen_url}{repo_id}/epics/{epic_issue_id}'
    zen_response = get_zenresponse(epic_url)
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
    git_response = requests.get(comments_for_issue_url, headers=GITHUB_HEADERS)
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


@retry(wait_exponential_multiplier=1000, wait_exponential_max=10000)
def get_zenresponse(issue_url):
    """
    Gets the ZenHub data for the issue
    :param issue_url: the specific url for the issue number
    :returns zen_response: the response for the issue
    """
    zen_response = requests.get(issue_url, headers=ZENHUB_HEADERS)
    if not zen_response.status_code == 200:
        raise Exception(zen_response.json())
    return zen_response


def get_zenhubresponse(repo_id, issue_number):
    """
    Gets the ZenHub data for the issue
    :param repo_id: the id for the repo used to reference the ZenHub fields
    :param issue_number: the specific issue number
    :returns zen_r: a json object of the issue
    """
    zen_url = f'https://api.zenhub.io/p1/repositories/{repo_id}/issues/'
    issue_url = f'{zen_url}{issue_number}'
    zen_response = get_zenresponse(issue_url)
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

def calculate_status(issue):
    """
    Calculates the traffic light status on an issue
    """
    status = 'Green'

    if issue['state'] == 'closed':
        milestone = '2200-01-01T23:59:59Z'
    elif issue['milestone']:
        milestone = issue['milestone'].get('due_on', '2200-01-01T23:59:59Z')
    else:
        milestone = '2200-01-01T23:59:59Z'

    red_rules = [
        datetime.datetime.strptime(milestone, '%Y-%m-%dT%H:%M:%SZ')
        < datetime.datetime.now(),
        issue['blocked'],
    ]

    for label in issue['labels'] if issue['labels'] else []:
        yellow_rules = [
            'yellow' in label['name'],
        ]
        if any(yellow_rules):
            status = 'Yellow'

    if any(red_rules):
        status = 'Red'

    return status

def write_row(issue, row, worksheet):
    """
    Writes rows to an Excel file
    """
    issue_number = str(issue['number'])
    row['s_pipeline'], row['estimate_value'] = get_zenhubresponse(row['repo_id'], issue_number)
    status = calculate_status(issue)
    if row['s_state'] == 'closed':
        row['s_pipeline'] = 'Closed'
    row['comments'] = ''
    if issue['comments'] > 0:
        row['comments'] = get_comments(row['repo_name'], issue_number)
    if (not row['s_priority'] and row['s_pipeline'] == 'Closed'):
        row['s_priority'] = 'High'
    rowvalues = [row['repo_name'], issue['number'], issue['title'],
                 row['userstory'], row['s_pipeline'], issue['user']['login'], issue['created_at'],
                 issue['milestone']['title'] if issue['milestone']
                 else "", issue['milestone']['due_on'] if issue['milestone'] else "",
                 row['s_assignee_list'][:-1], row['estimate_value'],
                 row['s_priority'], row['s_labels'],
                 row['comments'], row['s_epics'][:-1], status,
                 issue['blocked'], str(issue['blocked_by']).strip('[]')]
    for i, value in enumerate(rowvalues):
        worksheet.cell(column=(i+1),
                       row=1+row['issue_cnt'],
                       value=value)


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

    filename = f"{repo_name.split('/')[1]}.xlsx"
    fileoutput = Workbook()

    worksheet = fileoutput.create_sheet(title="Data")
    defaultsheet = fileoutput['Sheet']
    fileoutput.remove(defaultsheet)

    headers = ['Repository', 'Issue Number', 'Issue Title', 'User Story', 'Pipeline',
               'Issue Author', 'Created At', 'Milestone', 'Milestone End Date',
               'Assigned To', 'Estimate Value', 'Priority', 'Labels', 'Comments',
               'Epics', 'Status', 'Blocked', 'Blocked By']
    for i, header in enumerate(headers):
        worksheet.cell(column=(i+1), row=1, value=header)
        worksheet.cell(column=(i+1), row=1).font = Font(bold=True)

    for issue in r_json:
        issue_cnt += 1
        s_labels, s_priority = get_labels_string(issue)
        if HTMLFLAG == 1:
            userstory = markdown.markdown(issue['body'])
        else:
            userstory = issue['body']
        row = dict(repo_name=repo_name,
                   repo_id=repo_id,
                   userstory=userstory,
                   s_assignee_list=get_assignees(issue),
                   s_priority=s_priority,
                   s_labels=s_labels,
                   s_epics=get_epics_string(issues, issue),
                   s_state=issue['state'],
                   issue_cnt=issue_cnt,)
        write_row(issue, row, worksheet)
        print(f'issue count: {issue_cnt}')
        throttle_zenhub(issue_cnt)

    fileoutput.save(filename=filename)
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
    issue_response = requests.get(pages['next'], headers=GITHUB_HEADERS)
    if not issue_response.status_code == 200:
        print(issue_response.json())
        raise Exception(issue_response.status_code)
    return issue_response.json()

def get_github_issues(repo_name, repo_id, state='all', since=None):
    """
    Get github issues
    :param repo_name: the name of the GitHub repository
    :param state: the state of the issue, all by default
    :param since: the date to search after in the formate of %Y-%m-%d
    :returns issues: the dictionary of issues
    """
    issues_for_repo_url = f'https://api.github.com/repos/{repo_name}/issues?state={state}'
    issues = []
    if since:
        since_date = datetime.datetime.strptime(since, '%Y-%m-%d').strftime("%Y-%m-%dT%H:%M:%SZ")
        print(f'Filtering since {since_date}...')
        issues_for_repo_url = f'{issues_for_repo_url}&since={since_date}'
        print(f'Request {issues_for_repo_url}...')
    issue_response = requests.get(issues_for_repo_url, headers=GITHUB_HEADERS)
    if not issue_response.status_code == 200:
        raise Exception(issue_response.json())
    response = issue_response.json()
    issues += response
    # more pages? examine the 'link' header returned
    if 'link' in issue_response.headers:
        pages = get_pages(issue_response)
        while 'last' in pages and 'next' in pages:
            pages = get_pages(issue_response)
            issue_response = requests.get(pages['next'], headers=GITHUB_HEADERS)
            response = issue_response.json()
            issues += response
            if pages['next'] == pages['last']:
                blocked = create_blocked_items(repo_id, issues)
                for issue in issues:
                    if issue['number'] in blocked:
                        issue.update({"blocked": "blocked",
                                      "blocked_by": blocked[issue['number']]['blocked_by'],})
                    else:
                        issue.update({"blocked": "",
                                      "blocked_by": "",})
                    issues[issues.index(issue)] = issue
                return issues
    blocked = create_blocked_items(repo_id, issues)
    for issue in issues:
        if issue['number'] in blocked:
            issue.update({"blocked": "blocked",
                          "blocked_by": blocked[issue['number']]['blocked_by'],})
        else:
            issue.update({"blocked": "",
                          "blocked_by": "",})
        issues[issues.index(issue)] = issue
    return issues

def get_issues(repo_data, issues_dict, state='all', since=None):
    """
    Get an issue attributes
    :param repo_data: the environment variable with the repo_name
    and the ZenHub id for the repository
    """
    repo_name = repo_data[0]
    repo_id = repo_data[1]
    issues = get_github_issues(repo_name, repo_id, state=state, since=since)

    write_issues(issues, repo_name, repo_id, issues_dict, 0)

def main():
    """The real main function..."""

    parser = argparse.ArgumentParser()
    parser.add_argument('--file_name', help='file_name=filename.txt')
    parser.add_argument('--repo_list', nargs='+', help='repo_list owner/repo zenhub-id')
    parser.add_argument('--html', default=0, type=int, help='html=1')
    parser.add_argument('--since', default=None, help='since date in the format of 2018-01-01')
    args = parser.parse_args()

    repo_list = args.repo_list

    #txtout = open('data.json', 'w')
    issues = 0
    #filename = args.file_name
    #htmlflag = args.html
    since = args.since

    #get the epic dictionary
    issues = create_epic_dict(repo_list[1])

    #for repo_data in REPO_LIST:
    get_issues(repo_data=repo_list, issues_dict=issues, state='all', since=since)

if __name__ == '__main__':
    main()
