The script's main purpose is to reconcile accounts and contacts between Anchor, Northstar and Salesforce.

<h1>Usage</h1>
<h2>Inputs</h2>

* `--anchor-file` - path to Anchor Excel workbook, which should contain columns named:
  * `Salesforce ID` - Salesforce ID
  * `Company` - account name
  * `Name` - contact name
  * `Email`	- contact e-mail
  * `License Key` - license key
  * `Status` - account status
* `--northstar-file` - path to Northstar Excel workbook, which should contain columns named:
  * `license key` - license key
  * `user role` - user role
* `--salesforce-file` - path to Salesforce Excel workbook, which should contain columns named:
  * `Account 18 digit Id` - Salesforce ID
  * `Account Name` - account name
  * `Billing Country` - billing country
  * `Brand ID` - brand ID
  * `Current Products` - products attached to the account 
  * `First Name` - contact first name
  * `Last Name` - contact last name
  * `Email` - contact e-mail
  * `TPS License Information` - license key
* `--account-name-match-ratio-threshold` - account names with specified (or above) similarity ratio are used for joining Anchor and Salesforce account data. Number between 0 and 100; by default, 75.

Aforementioned spreadsheet column names are mandatory. Though spreadsheets are allowed to additionally contain arbitrary columns - they will be simply ignored during data reconciliation.

<h2>Output</h2>

* `--result-file` - path to result Excel workbook. The file will have 2 spreadsheets for accounts and contacts reconciliation.

<h2>Example</h2>

```bash
run.py --anchor-file anchor_usage_data.xlsx --northstar-file northstar_users.xlsx ----salesforce-file "X360sync - Anchor Partner Contacts.xlsx" --account-name-match-ratio-threshold 85 --result-file output.xlsx 
```   
 
<h1>Script algorithm (briefly)</h1>

1. Read Anchor workbook.
1. Read Northstar workbook.
1. Select Northstar entries where user role != 'Regular User' and not empty.
1. Join Anchor and Northstar data by license key.
1. Read Salesforce workbook.
1. Join Salesforce and Anchor/Northstar accounts by Salesforce ID.
1. Join Salesforce and Anchor/Northstar accounts by license key.
1. Join Salesforce and Anchor/Northstar accounts by name fuzzy match ratio.
1. Join Salesforce and Anchor/Northstar contacts by e-mail.
1. Export joined accounts and contacts to Excel workbook.

<h1>Installation</h1>

The script requires Python >= 3.7

Install all the required dependencies system-wide with 
```shell script
cd /path/to/script/dir
pip install -r ./requirements.txt
``` 

Or you may want to use virtual environment (venv) to keep your current Python installation intact.
```shell script
cd /path/to/script/dir
# create venv
python -m venv ./venv
# activate venv
source ./venv/bin/activate
# install dependencies
pip install -r requirements.txt
```