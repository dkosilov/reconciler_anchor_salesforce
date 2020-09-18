from abc import abstractproperty
from datetime import datetime
from numpy import nan
from pandas import concat, isnull, notnull, read_excel, DataFrame, MultiIndex, Series
from typing import Union, List, Dict
from fuzzywuzzy import fuzz, process

from libs.utils import save_dataframes_to_excel


class DataframeColumn:
    def __init__(self, name, order=0, src_name=None):
        self.name = name
        self.src_name = src_name
        self.order = order


class BaseDataframe:
    def __init__(self, src_filepath):
        src_cols = [c.src_name for c in self._get_columns().values() if c.src_name is not None]
        dest_cols = {c.src_name: c.name for c in self._get_columns().values() if c.src_name is not None}
        self.df = read_excel(src_filepath, usecols=src_cols).rename(dest_cols, axis='columns'). \
            drop_duplicates(ignore_index=True)
        self.orderize_columns()

    def orderize_columns(self):
        sorted_cols = sorted(self._get_columns().values(), key=lambda c: c.order)
        col_names = [c.name for c in sorted_cols]
        self.df = self.df[col_names]

    @classmethod
    def _get_columns(cls):
        return {name: col for name, col in cls.__dict__.items() if isinstance(col, DataframeColumn)}

    def save_to_excel(self, filepath):
        save_dataframes_to_excel(filepath, sheets_dataframes={'Result': self.df}, wrap_text=False)

    @staticmethod
    def log(msg):
        dt_now = datetime.now()
        print(f'{str(dt_now)}: {msg}')


class NorthStarDataframe(BaseDataframe):
    license_key = DataframeColumn('License Key', src_name='license key')
    # company_id = DataframeColumn('Company ID', 'company id')
    user_role = DataframeColumn('User Role', src_name='user role')

    USER_ROLE_REGULAR_USER = 'Regular User'

    def __init__(self, src_filepath):
        self.log('Reading Northstar data...')
        super().__init__(src_filepath)
        self.df = self.df[self.df[self.user_role.name].notnull() &
                          (self.df[self.user_role.name] != self.USER_ROLE_REGULAR_USER)]


class AnchorDataframe(BaseDataframe):
    salesforce_id = DataframeColumn('Salesforce ID', src_name='Salesforce ID')
    company_name = DataframeColumn('Company Name', src_name='Company')
    contact_name = DataframeColumn('Contact Name', src_name='Name')
    contact_email = DataframeColumn('Contact Email', src_name='Email')
    license_key = DataframeColumn('License Key', src_name='License Key')
    status = DataframeColumn('Status', src_name='Status')

    def __init__(self, src_filepath):
        self.log('Reading Anchor data...')
        super().__init__(src_filepath)


class SalesForceDataframe(BaseDataframe):
    salesforce_id = DataframeColumn('Salesforce ID', src_name='Account 18 digit Id')
    company_name = DataframeColumn('Company Name', src_name='Account Name')
    country = DataframeColumn('Billing Country', src_name='Billing Country')
    brand_id = DataframeColumn('Brand ID', src_name='Brand ID')
    products = DataframeColumn('Products', src_name='Current Products')
    contact_first_name = DataframeColumn('Contact First Name', src_name='First Name')
    contact_last_name = DataframeColumn('Contact Last Name', src_name='Last Name')
    contact_email = DataframeColumn('Contact Email', src_name='Email')
    license_key = DataframeColumn('License Key', src_name='TPS License Information')

    PRODUCT_ANCHOR = 'Anchor'
    PRODUCT_X360SYNC = 'x360Sync'

    def __init__(self, src_filepath):
        self.log('Reading Salesforce data...')
        super().__init__(src_filepath)
        # self.df[self.products.name] = self.df[self.products.name].apply(
        #     lambda x: x.split('; ') if isinstance(x, str) else [x]
        # )


class AnchorNorthstarDataframe(BaseDataframe):
    salesforce_id = DataframeColumn(AnchorDataframe.salesforce_id.name)
    company_name = DataframeColumn(AnchorDataframe.company_name.name)
    contact_name = DataframeColumn(AnchorDataframe.contact_name.name)
    contact_email = DataframeColumn(AnchorDataframe.contact_email.name)
    license_key = DataframeColumn(AnchorDataframe.license_key.name)
    status = DataframeColumn(AnchorDataframe.status.name)
    user_role = DataframeColumn(NorthStarDataframe.user_role.name)

    def __init__(self, src_anchor_filepath, src_northstar_filepath):
        anchor = AnchorDataframe(src_anchor_filepath)
        northstar = NorthStarDataframe(src_northstar_filepath)

        self.log('Joining Anchor/Northstar data by license key...')
        self.df = anchor.df.merge(right=northstar.df, how="inner", left_on=anchor.license_key.name,
                                  right_on=northstar.license_key.name, suffixes=(None, '_ns'))
        self.df.drop_duplicates(inplace=True)
        self.orderize_columns()


class AnchorSalesforceMixin:
    @classmethod
    def rebuild_dataframe(cls, dataframe: DataFrame, columns: Dict[str, DataframeColumn], top_level_name: str,
                          columns_key_prefix: str) -> DataFrame:
        """
        Rebuild dataframe by restricting it to columns with specified key prefix and
        adding one level to column structure

        :param dataframe: dataframe to rebuild
        :param columns: dictionary containing dataframe column and its key
        :param top_level_name: column level name to add to the top
        :param columns_key_prefix: columns key prefix
        :return:
        """
        df = dataframe[[v.name[1] for k, v in columns.items() if k.startswith(columns_key_prefix)]]
        df = df.drop_duplicates()
        df.columns = MultiIndex.from_product(iterables=([top_level_name], df.columns))
        return df


class AnchorSalesforceAccountsDataframe(BaseDataframe, AnchorSalesforceMixin):
    top_anchor = 'Anchor'
    anchor_salesforce_id = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.salesforce_id.name), order=0)
    anchor_company_name = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.company_name.name), order=10)
    anchor_license_key = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.license_key.name), order=30)
    anchor_status = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.status.name), order=40)
    anchor_user_role = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.user_role.name), order=50)

    top_salesforce = 'Salesforce'
    sf_salesforce_id = DataframeColumn(name=(top_salesforce, SalesForceDataframe.salesforce_id.name), order=60)
    sf_company_name = DataframeColumn(name=(top_salesforce, SalesForceDataframe.company_name.name), order=70)
    sf_products = DataframeColumn(name=(top_salesforce, SalesForceDataframe.products.name), order=80)
    sf_license_key = DataframeColumn(name=(top_salesforce, SalesForceDataframe.license_key.name), order=90)

    top_match = 'Matches'
    match_sf_id = DataframeColumn(name=(top_match, 'Salesforce ID'), order=90)
    match_license_key = DataframeColumn(name=(top_match, 'License Key'), order=100)
    match_fuzzy_ratio = DataframeColumn(name=(top_match, 'Fuzzy ratio'), order=110)
    # match_fuzzy_ratio_1st_chars = DataframeColumn(name=(top_match, 'Fuzzy ratio/n(1st 10 chars)'), order=120)

    def __init__(self, anchor_ns: AnchorNorthstarDataframe, salesforce: SalesForceDataframe,
                 name_fuzzy_match_ratio_threshold: int = 75):
        """
        Join accounts in Anchor and Salesforce by salesforce id, license key and name fuzzy matching

        :param anchor_ns: Anchor/Northstar dataframe object
        :param salesforce: Salesforce dataframe object
        :param name_fuzzy_match_ratio_threshold: account names with specified (or above) similarity ratio will be used
            for joining Anchor and Salesforce account data. Number between 0 and 100; by default, 75.
        """
        self.name_fuzzy_match_ratio_threshold = name_fuzzy_match_ratio_threshold
        df = self.rebuild_dataframe(dataframe=anchor_ns.df, columns=self._get_columns(), top_level_name=self.top_anchor,
                                    columns_key_prefix='anchor_')

        df[self.match_sf_id.name] = nan
        df[self.match_license_key.name] = nan
        df[self.match_fuzzy_ratio.name] = nan
        # df[self.match_fuzzy_ratio_1st_chars.name] = nan

        df_sf = self.rebuild_dataframe(dataframe=salesforce.df, columns=self._get_columns(),
                                       top_level_name=self.top_salesforce, columns_key_prefix='sf_')

        self.log('Joining Anchor/Salesforce accounts by Salesforce ID...')
        df = df.merge(right=df_sf, how="left", left_on=[self.anchor_salesforce_id.name],
                      right_on=[self.sf_salesforce_id.name], suffixes=(None, '_sf'))
        sf_id_nulls = df[self.sf_salesforce_id.name].isnull()
        self.df = df[~sf_id_nulls]
        df = df[sf_id_nulls]

        self.log('Joining Anchor/Salesforce accounts by license key...')
        df = df[[self.top_anchor]]
        df = df.merge(right=df_sf, how="left", left_on=[self.anchor_license_key.name],
                      right_on=[self.sf_license_key.name])
        license_key_nulls = df[self.sf_license_key.name].isnull()
        self.df = concat([self.df, df[~license_key_nulls]], ignore_index=True)
        df = df[license_key_nulls]

        self.log('Joining Anchor/Salesforce accounts by name fuzzy match...')
        df = df[[self.top_anchor]]
        df = self._merge_by_fuzzy_match(df, df_sf,
                                        self.anchor_company_name.name, self.sf_company_name.name)
        self.df = concat([self.df, df], ignore_index=True)

        self.log('Finalizing result Anchor/Salesforce accounts...')
        self.df[self.match_sf_id.name] = \
            self.df[self.anchor_salesforce_id.name] == self.df[self.sf_salesforce_id.name]
        self.df[self.match_license_key.name] = \
            self.df[self.anchor_license_key.name] == self.df[self.sf_license_key.name]
        self.df[self.match_fuzzy_ratio.name] = self.df.apply(
            lambda x: fuzz.ratio(str(x[self.anchor_company_name.name]), str(x[self.sf_company_name.name]))
            if notnull(x[self.anchor_company_name.name]) and notnull(x[self.sf_company_name.name]) and
            isnull(x[self.match_fuzzy_ratio.name]) else x[self.match_fuzzy_ratio.name],
            axis="columns"
        )
        self.orderize_columns()

    def _merge_by_fuzzy_match(self, left_df, right_df, left_on, right_on):
        tmp_col_match = ('tmp', 'fuzzy match')
        df = DataFrame()
        for _, left_row in left_df.iterrows():
            row = left_row.copy(deep=True)
            matches = process.extract(str(row[left_on]), right_df[right_on].to_list(), scorer=fuzz.ratio, limit=10)
            row[tmp_col_match] = list(filter(lambda x: x[1] >= self.name_fuzzy_match_ratio_threshold, matches))
            df_exploded = DataFrame([row]).explode(tmp_col_match)
            df_exploded[[tmp_col_match, self.match_fuzzy_ratio.name]] = df_exploded[tmp_col_match].apply(
                lambda x: Series(x) if isinstance(x, tuple) else Series([nan, nan])
            )
            df = concat([df, df_exploded])
        df = df.merge(right_df, how='left', left_on=[tmp_col_match], right_on=[right_on])
        df.drop(columns=[tmp_col_match], inplace=True)
        return df


class AnchorSalesforceContactsDataframe(BaseDataframe, AnchorSalesforceMixin):
    top_anchor = 'Anchor'
    anchor_salesforce_id = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.salesforce_id.name), order=0)
    anchor_company_name = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.company_name.name), order=10)
    anchor_contact_name = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.contact_name.name), order=20)
    anchor_contact_email = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.contact_email.name), order=30)
    anchor_status = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.status.name), order=40)
    anchor_user_role = DataframeColumn(name=(top_anchor, AnchorNorthstarDataframe.user_role.name), order=50)

    top_salesforce = 'Salesforce'
    sf_salesforce_id = DataframeColumn(name=(top_salesforce, SalesForceDataframe.salesforce_id.name), order=60)
    sf_company_name = DataframeColumn(name=(top_salesforce, SalesForceDataframe.company_name.name), order=70)
    sf_contact_first_name = DataframeColumn(name=(top_salesforce, SalesForceDataframe.contact_first_name.name),
                                            order=80)
    sf_contact_last_name = DataframeColumn(name=(top_salesforce, SalesForceDataframe.contact_last_name.name),
                                           order=90)
    sf_contact_email = DataframeColumn(name=(top_salesforce, SalesForceDataframe.contact_email.name),
                                       order=100)

    def __init__(self, anchor_ns: AnchorNorthstarDataframe, salesforce: SalesForceDataframe):
        """
        Join contacts in Anchor and Salesforce by e-mail

        :param anchor_ns: Anchor/Northstar dataframe object
        :param salesforce: Salesforce dataframe object
        """
        df = self.rebuild_dataframe(dataframe=anchor_ns.df, columns=self._get_columns(), top_level_name=self.top_anchor,
                                    columns_key_prefix='anchor_')

        df_sf = self.rebuild_dataframe(dataframe=salesforce.df, columns=self._get_columns(),
                                       top_level_name=self.top_salesforce, columns_key_prefix='sf_')

        self.log('Joining Anchor/Salesforce contacts by e-mail...')
        self.df = df.merge(right=df_sf, how="left", left_on=[self.anchor_contact_email.name],
                           right_on=[self.sf_contact_email.name], suffixes=(None, '_sf'))
        self.orderize_columns()
