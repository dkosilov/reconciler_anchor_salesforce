import argparse

from libs.data_model import AnchorNorthstarDataframe, SalesForceDataframe, \
    AnchorSalesforceAccountsDataframe, AnchorSalesforceContactsDataframe
from libs.utils import save_dataframes_to_excel

parser = argparse.ArgumentParser(description='Reconcile accounts and contacts between Anchor and Salesforce')
parser.add_argument('-a', '--anchor-file', help='Path to Anchor Excel workbook', required=True)
parser.add_argument('-n', '--northstar-file', help='Path to Northstar Excel workbook', required=True)
parser.add_argument('-s', '--salesforce-file', help='Path to Salesforce Excel workbook', required=True)
parser.add_argument('-t', '--account-name-match-ratio-threshold', type=int,
                    help='Account names with specified (or above) similarity ratio will be used for joining Anchor and '
                         'Salesforce account data. Number between 0 and 100.', default=75)
parser.add_argument('-r', '--result-file',
                    help='Path to result Excel workbook. The file will have 2 spreadsheets for accounts and '
                         'contacts reconciliation', required=True)

args = parser.parse_args()

anchor_ns = AnchorNorthstarDataframe(args.anchor_file, args.northstar_file)
salesforce = SalesForceDataframe(args.salesforce_file)

anchor_sf_accounts = AnchorSalesforceAccountsDataframe(anchor_ns, salesforce, args.account_name_match_ratio_threshold)
anchor_sf_contacts = AnchorSalesforceContactsDataframe(anchor_ns, salesforce)

save_dataframes_to_excel(args.result_file, {'Accounts': anchor_sf_accounts.df, 'Contacts': anchor_sf_contacts.df},
                         wrap_text=False)

