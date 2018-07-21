#!/usr/bin/env python
import csv
import json
import requests
import time


"""
Exports Issues from a list of repositories to individual CSV files
Uses basic authentication (Github API Token and Zenhub API Token)
to retrieve Issues from a repository that token has access to.
Supports Github API v3 and ZenHubs current working API.
Derived from https://gist.github.com/Kebiled/7b035d7518fdfd50d07e2a285aff3977
"""


def write_issues(r, csvout, repo_name, repo_ID):
    if not r.status_code == 200:
        raise Exception(r.status_code)
    r_json = r.json()
    for issue in r_json:
        print (repo_name + ' issue Number: ' + str(issue['number']))
        zenhub_issue_url = 'https://api.zenhub.io/p1/repositories/' + \
            str(repo_ID) + '/issues/' + str(issue['number']) + ACCESS_TOKEN
        zen_r = requests.get(zenhub_issue_url).json()
        global Payload


        if 'pull_request' not in issue:
            global ISSUES
            ISSUES += 1
            sAssigneeList = ''
            sTag = ''
            sPhase = ''
            sEscDefect = ''
            sPipeline = ''
            for i in issue['assignees'] if issue['assignees'] else []:
                sAssigneeList += i['login'] + ','
            for x in issue['labels'] if issue['labels'] else []:
                if "Phase" in x['name']:
                    sPhase = x['name']
                if "EscDefect" in x['name']:
                    sEscDefect = x['name']
            #add output of the payload for records not found
            sPipeline = zen_r.get("pipeline", dict()).get('name', "")
            lEstimateValue = zen_r.get("estimate", dict()).get('value', "")

            csvout.writerow([repo_name, issue['number'], issue['title'], sPipeline, issue['user']['login'], issue['created_at'],
                             issue['milestone']['title'] if issue['milestone'] else "",issue['milestone']['due_on'] if issue['milestone'] else "",
                             sAssigneeList[:-1], lEstimateValue, sPhase, sEscDefect] )
#Wait added for the ZenHub api rate limit of 100 requests per minute
            if ISSUES%(ZENHUB_API_RATE_LIMIT-1) == 0:
                time.sleep(60)
        else:
            print ('You have skipped %s Pull Requests' % ISSUES)


def get_issues(repo_data):
    repo_name = repo_data[0]
    repo_ID = repo_data[1]
    issues_for_repo_url = 'https://api.github.com/repos/%s/issues?state=all' % repo_name
    r = requests.get(issues_for_repo_url, auth=AUTH)
    write_issues(r, FILEOUTPUT, repo_name, repo_ID)
    # more pages? examine the 'link' header returned
    if 'link' in r.headers:
        pages = dict(
            [(rel[6:-1], url[url.index('<') + 1:-1]) for url, rel in
             [link.split(';') for link in
              r.headers['link'].split(',')]])
        while 'last' in pages and 'next' in pages:
            pages = dict(
                [(rel[6:-1], url[url.index('<') + 1:-1]) for url, rel in
                 [link.split(';') for link in
                  r.headers['link'].split(',')]])
            r = requests.get(pages['next'], auth=AUTH)
            write_issues(r, FILEOUTPUT, repo_name, repo_ID)
            if pages['next'] == pages['last']:
                break


PAYLOAD = ""

REPO_LIST = [("*GITHUB REPO*", "*ZENHUB REPOID*")]

AUTH = ('token', '*GITHUB ACCESS TOKEN*')
ACCESS_TOKEN = '?access_token=*ZENHUb ACCESS TOKEN*'
ZENHUB_API_RATE_LIMIT = 100

TXTOUT = open('data.json', 'w')
ISSUES = 0
FILENAME = '*Filename*'
OPENFILE = open(FILENAME, 'w', newline='')
FILEOUTPUT = csv.writer(OPENFILE,dialect='excel',delimiter='|')
FILEOUTPUT.writerow(['Repository', 'Issue Number', 'Issue Title', 'Pipeline', 'Issue Author',
	'Created At', 'Milestone', 'Milestone End Date', 'Assigned To', 'Estimate Value', 'Label','Escaped Defect'])
for repo_data in REPO_LIST:
    get_issues(repo_data)
json.dump(PAYLOAD, open('data.json', 'w'), indent=4)
TXTOUT.close()
OPENFILE.close()
