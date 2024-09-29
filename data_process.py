import pandas as pd
import numpy as np
from tqdm import tqdm
from utils import *

def process_data():
    print('Processing data...')
    app_dir = get_app_dir()

    # Path to the external files
    open_windows_fp = os.path.join(app_dir, 'data', 'Availability Open Windows - By Territory and Right (Copy) 1.xlsx')
    check_file(open_windows_fp)
    contract_summary_fp = os.path.join(app_dir, 'data', 'Contract Summary.xlsx')
    check_file(contract_summary_fp)
    titles_fp = os.path.join(app_dir, 'data', 'Project List.xlsx')
    check_file(titles_fp)
    rights_fp = os.path.join(app_dir, 'data', 'rights.csv')
    check_file(rights_fp)
    countries_fp = os.path.join(app_dir, 'data', 'countries.csv')
    check_file(countries_fp)
    title_metadata_fp = os.path.join(app_dir, 'data', 'Project Data ID.xlsx')
    check_file(title_metadata_fp)
    ratings_fp = os.path.join(app_dir, 'data', 'Ratings & Titles.xlsx')
    check_file(ratings_fp)

    # Load data
    rights = pd.read_csv(rights_fp, encoding = 'unicode_escape')

    right_groups_map = {}

    unique_rights = rights['name'].unique()

    for right in unique_rights:
        if right in [
                'AVOD',
                'Ad-VOD',
    #            'Free-VOD',
        ]:
            right_groups_map[right] = 'AVOD'
        elif right in [
                'Airline',
    #            'Airline (U)',
                'Ancillary',
                'Buses/Transportation (Long-Haul)',
                'Hotel/Motel',
                'NVOD',
                'Ship',
        ]:
            right_groups_map[right] = 'Ancillary'
        elif right in [
                'Advertising / Promotions',
                'Clip',
                'Live Performance',
                'Merchandising',
                'Other Ancillary',
                'Novelization / Print Pub.',
                'Remake',
                'Sequel / Prequel',
                'Soundtrack / Music Pub.',
                'Souvenir / Alternate Markets',
                'Video Game / Interactive',
        ]:
            right_groups_map[right] = 'Other Ancillary'
        elif right in [
                'Basic Pay TV (Local)',
                'Basic Pay TV (Local) CC',
                'Basic Pay TV (Local) Sat.,Cable',
                'Basic Pay TV (Local) Catch-Up/FVOD/SVOD-AS',
                'Basic Pay TV (Local) Terrestrial',
                'Basic Pay TV (Local) TV Everywhere',
        ]:
            right_groups_map[right] = 'Basic Pay TV (Local)'
        elif right in [
                'Basic Pay TV (Pan Regional) Sat.,Cable',
                'Basic Pay TV (Pan Regional) Catch-Up/FVOD/SVOD-AS',
                'Basic Pay TV (Pan Regional) Terrestrial',
                'Basic Pay TV (Pan Regional) TV Everywhere',
                'Basic Pay TV (Pan Regional)',
                'Basic Pay TV (Pan Regional) CC',
        ]:
            right_groups_map[right] = 'Basic Pay TV (Pan Regional)'
        elif right in [
                'Cinematic',
                'Non-Theatrical',
                'Public Video',
                'Theatrical',
        ]:
            right_groups_map[right] = 'Cinematic'
        elif right in [
                'ClosedNet',
                'INT-Download',
                'INT-Stream',
                'IPTV-Webcast-Simulcast',
                'Internet',
        ]:
            right_groups_map[right] = 'Internet'
        elif right in [
                'Demand',
                'EST',
                'Electronic Rental (Download-To-Rent)',
                'Electronic Sell-Thru (Dowload-To-Own)',
                'PPV-Non-Residential',
                'PPV-Residential',
                'Pay-Per-View',
                'TVOD',
    #            'TVOD (U)',
        ]:
            right_groups_map[right] = 'TVOD'
        elif right in [
                'Free TV',
                'Free TV-CAB',
                'Free TV-Catch-Up',
                'Free TV-Closed Circuit',
                'Free TV-SAT',
                'Free TV-TER',
        ]:
            right_groups_map[right] = 'Free TV'
        elif right in [
                'Home Video',
                'Video-Commercial',
                'Video-Rental',
                'Video-Sell-Thru',
        ]:
            right_groups_map[right] = 'Home Video'
        elif right in [
                'MOB-Download',
                'MOB-Ringtone',
                'MOB-Stream',
                'MOB-Wallpaper',
                'Wireless',
        ]:
            right_groups_map[right] = 'Wireless'
        elif right in [
                'Premium Pay TV (Local)',
                'Premium Pay TV (Local) CC',
                'Premium Pay TV (Local) Sat., Cable',
                'Premium Pay TV (Local) Catch-Up/FVOD/SVOD-AS',
                'Premium Pay TV (Local) Terrestrial',
                'Premium Pay TV (Local) TV Everywhere',
        ]:
            right_groups_map[right] = 'Premium Pay TV (Local)'
        elif right in [
                'Premium Pay TV (Pan Regional) Sat.,Cable',
                'Premium Pay TV (Pan Regional) Catch-Up/FVOD/SVOD-AS',
                'Premium Pay TV (Pan Regional) Terrestrial',
                'Premium Pay TV (Pan Regional) CC',
                'Premium Pay TV (Pan Regional) TV Everywhere',
                'Premium Pay TV (Pan Regional)',
        ]:
            right_groups_map[right] = 'Premium Pay TV (Pan Regional)'
        elif right in [
                'SVOD',
    #            'SVOD (U)',
        ]:
            right_groups_map[right] = 'SVOD'
        else:
            right_groups_map[right] = right

    rights = pd.DataFrame(
        list(right_groups_map.items()),
        columns=['name', 'group']
)
    rights['right'] = np.arange(1, rights.shape[0] + 1)

    countries = pd.read_csv(countries_fp, encoding = 'unicode_escape')
    countries = countries.drop_duplicates(subset='name', keep='first')
    countries['country'] = np.arange(1, countries.shape[0] + 1)
    countries['market_region'] = countries['market_region'].fillna(countries['geo_region'])

    contracts_df = pd.read_excel(contract_summary_fp, header=0)
    contracts_df.columns = [
        colname.replace(" ","_").replace("/Conditions", "").replace("#","Id").lower() for colname in contracts_df.columns
    ]
    contract_summary_cols = [
    #    'support_code', 
        'contract_id', 
        'contract_type', 
        'licensor',
        'distributor', 
    #    'projects', 
    #    'project_codes', 
    #    'territories', 
    #    'regions',
    #    'zone', 
    #    'languages', 
        'status', 
        'deal_status', 
        'creation_date',
        'deal_type', 
    #    'isestimatedstartdate', 
    #    'start_date', 
    #    'end_date',
    #    'isestimatedenddate', 
    #    'start_date_note', 
    #    'end_date_note',
        'fully_executed', 
    #    'nod', 
    #    'outside_delivery_date', 
    #    'deal_memo_date',
    #    'mg_usd', 
        'mg', 
        'cur', 
    #   'exchange_rate', 
    #    'payment_data', 
    #    'invoice_data',
        'additional_terms', 
    #    'total_invoiced', 
    #    'total_paid', 
    #    'total_balance',
    #    'total_invoiced_usd', 
    #    'total_paid_usd', 
    #    'right_group_summary',
    #    'right_summary', 
    #  'notes'
    ]

    contracts = contracts_df[contract_summary_cols].copy().sort_values(by='creation_date')
    contracts['contract'] = np.arange(1, contracts.shape[0]+1)
    contracts.rename(columns={'contract_id': 'contract_code'}, inplace=True)

    titles_df = pd.read_excel(titles_fp)
    titles_df.dropna(axis=1, how='all', inplace=True)
    titles_cols = [
        'Title', 
        'AKA 1', 
        'AKA 2', 
        'Adj. Running Time', 
    #    'Associate Producer',
    #    'Budget', 
    #    'Business Unit', 
    #    'Cast Crew - Summary Tab', 
    #    'Cast Member',
        'Copyright Holder', 
    #  'Copyright Year', 
        'Country of Origin',
        'Dialogue Language', 
    #    'Director', 
    #    'Exploitation', 
    #    'External Comments',
        'Genre', 
        'IMDB Code', 
    #   'Internal Comments', 
        'Logline',
    #   'Motion Picture Association of America', 
        'Number of Episodes',
        'Number of Seasons', 
        'Original Format', 
        'Original Language', 
    #    'Producer',
    #    'Production Company', 
        'Project Code', 
        'Project Group', 
        'Project Type',
        'Rating', 
    #   'Release Date', 
        'Running Time', 
    #  'Sales Agency', 
        'Season',
        'Short Synopsis', 
        'Status', 
        'Subtitle Language', 
        'Synopsis',
        'Title Code', 
        'Unique Id', 
    #   'Web Synopsis', 
        'Website', 
    #    'Writer',
        'Year Completed'
    ]
    talent_df = tidy_split(titles_df[['Cast Member', 'Unique Id']], 'Cast Member')
    talent_df['Role'] = 'Cast'
    talent_df = talent_df.rename({'Cast Member': 'Talent Full Name'}, axis=1)
    directors = tidy_split(titles_df[['Director', 'Unique Id']], 'Director')
    directors['Role'] = 'Director'
    directors = directors.rename({'Director': 'Talent Full Name'}, axis=1)
    producers = tidy_split(titles_df[['Producer', 'Unique Id']], 'Producer')
    producers['Role'] = 'Producer'
    producers = producers.rename({'Producer': 'Talent Full Name'}, axis=1)
    writers = tidy_split(titles_df[['Writer', 'Unique Id']], 'Writer')
    writers['Role'] = 'Writer'
    writers = writers.rename({'Writer': 'Talent Full Name'}, axis=1)
    talent_df = pd.concat([talent_df, directors, producers, writers])
    talent_df = talent_df.reset_index(drop=True)
    titles_df.columns = [
        colname.replace(" ","_").replace(".","").lower() for colname in titles_df.columns
    ]
    titles_cols = [
        colname.replace(" ","_").replace(".","").lower() for colname in titles_cols
    ]
    titles = titles_df[titles_cols].copy()
    titles.rename(columns={'title':'name'}, inplace=True)
    titles.rename(columns={'unique_id':'title'}, inplace=True)

    title_metadata = pd.read_excel(title_metadata_fp)
    title_metadata.columns = [
        colname.replace(" ","_").replace(".","").lower() for colname in title_metadata.columns
    ]
    title_metadata.drop('title', axis=1, inplace=True)
    title_metadata['title'] = title_metadata['unique_identifier'].astype(pd.Int64Dtype())
    title_metadata = title_metadata.set_index('title')
    title_metadata.drop(['unique_identifier'], axis=1, inplace=True)

    titles = titles.set_index('title').join(title_metadata, how='left')

    ratings = pd.read_excel(ratings_fp, index_col='Unique Identifier')
    ratings.columns = [
        colname.strip().replace(" ","_").replace(".","").lower() for colname in ratings.columns
    ]
    ratings.index.name = 'title'
    imdb = ratings['imdb']
    ratings.drop(['title', 'imdb'], axis=1, inplace=True)
    ratings.columns = ['rating_'+col.strip() for col in ratings.columns]
    ratings['imdb'] = imdb

    titles = titles.join(ratings, how='left')

    xls = pd.ExcelFile(open_windows_fp)
    open_windows_sn = xls.sheet_names
    open_windows_sn.remove('All Rights')
    open_windows_sn.remove('Filter Values')
    open_windows_sn = [sn for sn in open_windows_sn if not sn.endswith(' (U)')]

    df_list = []  # use list to store individual sheet dataframes

    for sn in tqdm(open_windows_sn):
        sheet_df = pd.read_excel(xls, sheet_name=sn)
        sheet_df = sheet_df.loc[~sheet_df['Contract Code'].isna()]
        sheet_df = sheet_df.loc[sheet_df['Territory'].isin(countries['name']), :]
        sheet_df = sheet_df.loc[sheet_df['Right'].isin(rights['name']), :]
        sheet_df['Start Date'] = sheet_df['Start Date'].apply(clean_date)
        sheet_df['End Date'] = sheet_df['End Date'].apply(
            lambda x: clean_date(x, start_date=False))
        sheet_df.columns = [
            colname.replace(" ", "_").lower() for colname in sheet_df.columns
        ]
        sheet_df.loc[(sheet_df['contract_code'].map(len) > 6), 'contract_code'] = sheet_df.loc[(sheet_df['contract_code'].map(len) > 6), 'contract_code'].apply(lambda x: x.split()[0][:-1])
        sheet_df['window_id_name'] = sheet_df['contract_code'] + \
            sheet_df['territory'] + sheet_df['right'] + \
            sheet_df['unique_id'].astype(str)
        sheet_df['start_confirmed'] = sheet_df['start_e/a'].map({
            'E': False,
            'A': True
        })
        sheet_df['end_confirmed'] = sheet_df['end_e/a'].map({
            'E': False,
            'A': True
        })
        sheet_df.drop(['start_e/a', 'end_e/a'], axis=1, inplace=True)
        sheet_df['title'] = sheet_df['unique_id'].astype(int)
        sheet_df.drop('unique_id', axis=1, inplace=True)
        df_list.append(sheet_df)

    # concatenating all the dataframes at once
    open_windows = pd.concat(df_list, axis=0, ignore_index=True)
    open_windows['window'] = np.arange(1, open_windows.shape[0] + 1)
    
    # Merge dataframes
    open_windows = open_windows.merge(
        contracts[['contract_code', 'contract']], 
        on='contract_code', 
        how='left'
    )
    open_windows = open_windows.merge(
        countries[['country', 'name', 'market_region']],
        left_on='territory',
        right_on='name',
        how='left'
    )
    open_windows['country_name'] = open_windows['territory']
    open_windows.drop(['territory', 'name'], axis=1, inplace=True)

    open_windows = open_windows.merge(
        rights[['right', 'name', 'group']],
        left_on='right',
        right_on='name',
        how='left'
    )
    open_windows['right_name'] = open_windows['name']
    open_windows.drop(['name', 'right_x'], axis=1, inplace=True)
    open_windows.rename(columns={'right_y': 'right'}, inplace=True)

    open_windows = open_windows.merge(
        titles.reset_index()[['title', 'name']],
        how='left',
        on='title'
    )
    open_windows['title_name'] = open_windows['name']
    open_windows.drop('name', axis=1, inplace=True)

    # Create roles table
    talent = talent_df.copy()
    talent.columns = [
        colname.replace(" ","_").lower() for colname in talent.columns
    ]
    people = pd.DataFrame(talent['talent_full_name'].unique(), columns=['name'])
    people['person'] = np.arange(1, people.shape[0] + 1)

    roles = talent.merge(
        people,
        left_on='talent_full_name',
        right_on='name',
        how='left'
    )
    roles.drop(['talent_full_name','name'], axis=1, inplace=True)
    roles.rename(columns={'unique_id': 'title'}, inplace=True)
    roles['row'] = np.arange(1, roles.shape[0] + 1)

    # Save tables to pickle
    print('Saving data to disk...')

    # Create an 'data/tables' directory if it does not exist
    if not os.path.exists(os.path.join(app_dir, 'data', 'tables')):
        os.makedirs(os.path.join(app_dir, 'data', 'tables'))

    open_windows.to_pickle(os.path.join(app_dir, 'data', 'tables', 'windows.pkl'))
    contracts.to_pickle(os.path.join(app_dir, 'data', 'tables', 'contracts.pkl'))
    titles.to_pickle(os.path.join(app_dir, 'data', 'tables', 'titles.pkl'))
    people.to_pickle(os.path.join(app_dir, 'data', 'tables', 'people.pkl'))
    roles.to_pickle(os.path.join(app_dir, 'data', 'tables', 'roles.pkl'))