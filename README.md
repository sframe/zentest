# zentest
This Python script exports github issues with the ZenHub additional attributes to an Excel file.

# Installation
To use this, you will need Python 3 and then pip install from the requirements.txt:

```bash
python3 -m venv venv
. venv/bin/activate
pip3 install -r requirements.txt
```

Get an API token for your GitHub account under Settings->Developer settings->Personal Access Tokens
 - I am not positive all the rights you need on the token, but at a minimum, you need the repo rights
Get a ZenHub API token under https://app.zenhub.com/dashboard/tokens for organization with the repository you want to pull

Set these values as environment variables.

For Windows, use the following format:

```
setx GITHUB_REPO XXXX/XXXX
setx GITHUB_TOKEN XXXX
setx ZENHUB_REPO_ID XXXX
setx ZENHUB_TOKEN XXXXXXX
```

For Bash, use the following format:

```
export GITHUB_REPO=XXXX/XXXX
export GITHUB_TOKEN=XXXX
export ZENHUB_REPO_ID=XXXX
export ZENHUB_TOKEN=XXXXXXX
```

The command line parameters are as follows:

`--filename` The name of the Excel file to create

`--repo github_owner/github_repo`

You can find you zenhub-id by going to ZenHub id by looking in the url when you go to Zenhub.com for the repository, repos=<zenhub-repo-id>

`--html` This is either 0 or 1, most of the time you will want 0 which leaves the format in Markdown

`--state` The state of the issues to export. Can be open, closed, or all. The default is all

Example:

```sh
python ./export_repo_issues_to_csv.py \
  --file_name ~/Desktop/$(date +"%Y-%m-%d")_zenhub_export.xlsx \
  --repo_list $GITHUB_REPO $ZENHUB_REPO_ID \
  --html 0
  --state closed
```
