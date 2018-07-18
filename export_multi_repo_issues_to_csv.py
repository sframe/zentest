#!/usr/bin/env python
import csv
import json
import requests


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
            sCategory = ''
            sPriority = ''
            for i in issue['assignees'] if issue['assignees'] else []:
                sAssigneeList += i['login'] + ','
            for x in issue['labels'] if issue['labels'] else []:
                if "Phase" in x['name']:
                    sCategory = x['name']
                if "Tag" in x['name']:
                    sTag = x['name']
                if "Priority" in x['name']:
                    sPriority = x['name']
            lEstimateValue = zen_r.get('estimate', dict()).get('value', "")
            sPipeline = zen_r.get('pipeline', dict()).get('name', "")

            csvout.writerow([repo_name, issue['number'], issue['title'], sPipeline, issue['user']['login'], issue['created_at'],
                             issue['milestone']['title'] if issue['milestone'] else "",issue['milestone']['due_on'] if issue['milestone'] else "",
                             sAssigneeList[:-1], lEstimateValue, sCategory] )
        else:
            print ('You have skipped %s Pull Requests' % ISSUES)


def get_issues(repo_data):
    repo_name = repo_data[0]
    repo_ID = repo_data[1]
    issues_for_repo_url = 'https://api.github.com/repos/%s/issues' % repo_name
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

    #FILEOUTPUT.writerow(['Total', ISSUES])


PAYLOAD = ""
REPO_LIST = [("st-cbp/olb", "136191097")]

AUTH = ('token', '6a8b86b543f0a094dc7c9a677de746b2df23c780')
ACCESS_TOKEN = '?access_token=4ad09624c096ad3f6d763f763a8386e788ee593bfeaab6a8904940e1cfca525716848f346deeb01e'

TXTOUT = open('data.json', 'w')
ISSUES = 0
FILENAME = 'olb.txt'
OPENFILE = open(FILENAME, 'w', newline='')
FILEOUTPUT = csv.writer(OPENFILE,dialect='excel',delimiter='|')
FILEOUTPUT.writerow(['Repository', 'Issue Number', 'Issue Title', 'Pipeline', 'Issue Author',
	'Created At', 'Milestone', 'Milestone End Date', 'Assigned To', 'Estimate Value', 'Label'])
for repo_data in REPO_LIST:
    get_issues(repo_data)
json.dump(PAYLOAD, open('data.json', 'w'), indent=4)
TXTOUT.close()
OPENFILE.close()