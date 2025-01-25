import pandas as pd
import numpy as np
import datetime
from collections import Counter



# change to current working month and year accordingly
WORKING_MONTH = "JULY"
WORKING_YEAR = 2024
# change to last day of the current working month
FIRST_DAY = 16
LAST_DAY = 31


chain_conv_dict = {
    "eleven": "7ELEVEN",
    "alfamart": "ALFAMART",
    "capital": "GAISANO CAPITAL",
    "citimart": "CITIMART",
    "csi": "CSI",
    "dsg": "DSG SONS",
    "easymart": "ROBINSONS EASYMART",
    "ever": "EVER",
    "grand": "GAISANO GRAND",
    "harddiscount": "HARD DISCOUNT",
    "kcc": "KCC",
    "lawson": "LAWSON",
    "mdc": "MDC",
    "mercury": "MDC",
    "metro": "METRO GAISANO",
    "ministop": "MINISTOP",
    "ncccmin": "NCCC",
    "ncccsm": "NCCC",
    "ncccpal": "NCCC-PAL",
    "ncccswl": "NCCC-PAL",
    "puregold": "PUREGOLD",
    "puremart": "PUREMART",
    "robinsonssuper": "ROBINSONS SUPERMARKET",
    "robinsonseasy": "ROBINSONS EASYMART",
    "smh": "SMH",
    "southstardrug": "SSDI",
    "ssdi": "SSDI",
    "stephen": "GAISANO STEPHEN",
    "super8": "SUPER8",
    "s&r": "S&R",
    "svi": "SVI",
    "threesixty": "THREE SIXTY PHARMACY",
    "umret": "UM RETAIL",
    "ultramega": "UM RETAIL",
    "umws": "UM WHOLESALE",
    "unclejohns": "UNCLE JOHN'S / MINISTOP",
    "waltermart": "WALTERMART"
}

act_type_dict1 = {
    'bundling': ['Bundling In-Store', 'BUNDLING'],
    'discount': ['Discount/Price Rollback', 'DISCOUNT'],
    'deployment': ['Other Activations', 'OTHER ACTIVATIONS'],
    'loyalty': ['Loyalty Program', 'LOYALTY'],
    'promo': ['Promo Packs', 'PROMO'],
    'redemption': ['Redemption w Premium Items', 'REDEMPTION'],
    'thematic': ['Thematic', 'THEMATIC'],
    'tactical': ['Tactical Display', 'TACTICAL'],
    'trade deal': ['Trade Deal', 'TRD DEAL'],
    'deals': ['Trade/Case Deals', 'TRD CASE DEALS'],
    'contest': ['Merchandising Contest', 'CONTEST'],
    'generic': ['Merchandising Generic', 'GENERIC'],
    'paid': ['Merchandising Paid', 'PAID'],
    'new products': ['New Products', 'NEW PRODUCTS'],
    'product launch': ['New Product Launch', 'NEW PRODUCTS'],
    'product renovation': ['New Product Renovation', 'NEW PRODUCTS'],
    'other activation': ['Other Activations', 'OTHER ACTIVATIONS'],
    'other activations': ['Other Activations', 'OTHER ACTIVATIONS'],
    'placement': ['Placement', 'PLACEMENT'],
    'price off': ['Price Off', 'PRICE OFF'],
    'renovation': ['Renovation (existing product)', 'RENOVATION'],
    'sleeving': ['Promo Packs', 'PROMO']
}

month_conv_dict = {
    "jan": "JANUARY",
    "feb": "FEBRUARY",
    "mar": "MARCH",
    "apr": "APRIL",
    "may": "MAY",
    "jun": "JUNE",
    "jul": "JULY",
    "aug": "AUGUST",
    "sep": "SEPTEMBER",
    "oct": "OCTOBER",
    "nov": "NOVEMBER",
    "dec": "DECEMBER"
}

form_id_count_dict = {
    0: '',
    1: 'A',
    2: 'B',
    3: 'C',
    4: 'D',
    5: 'E',
    6: 'F',
    7: 'G',
    8: 'H',
    9: 'I',
    10: 'J',
    11: 'K',
    12: 'L',
    13: 'M',
    14: 'N',
    15: 'O',
    16: 'P',
    17: 'Q',
    18: 'R',
    19: 'S',
    20: 'T',
    21: 'U',
    22: 'V',
    23: 'W',
    24: 'X',
    25: 'Y',
    26: 'Z'
}

# check if chain column in template is named CHAIN or ACCOUNT, and change col name here accordingly
# check if activity column in template is named MARS_NAME or MAINTAINED_ACTIVITY_NAME, and change col name here accordingly
data_cols = ['ACCOUNT', 'CATEGORY', 'ACTIVITY_TYPE', 'MARS_NAME', 'START_DATE', 'END_DATE', 'START_MONTH', 'END_MONTH']

main_df = pd.read_excel(f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CS_RawFiles/B2/CUST SPEC JULY 2024 MARS UPLOADING - BATCH 2.xlsb", sheet_name="PER CHAIN", index_col=None, header=1, usecols=data_cols, dtype={"ACCOUNT": str, "START_MONTH": object, "END_MONTH": object}, keep_default_na=False)

# -------------------------- Parsing Account Names -------------------------- #
# TO-DO #1: convert accounts to lowercase
main_df['ACCOUNT'] = [account.lower() for account in main_df['ACCOUNT']]

# TO-DO #2: remove whitespaces from accounts
main_df['ACCOUNT'] = [account.replace(' ', '') for account in main_df['ACCOUNT']]

# TO-DO #3: remove '-', '/', and ' from accounts
main_df['ACCOUNT'] = [account.replace('-', '') for account in main_df['ACCOUNT']]
main_df['ACCOUNT'] = [account.replace('/', '') for account in main_df['ACCOUNT']]
main_df['ACCOUNT'] = [account.replace("'", '') for account in main_df['ACCOUNT']]


main_df['ACCOUNT_adj'] = [
    ''.join(list(set(chain_conv_dict[key] if key in account
    else "SHOPWISE/ THE MARKETPLACE" if account == "shopwisemarketplace" or account == "shopwisethemarketplace"
    else "THE MARKETPLACE" if account == "themarketplace"
    else "SHOPWISE" if account == "shopwise"
    else "NCCC" if account == "nccc"
    else ''
    for key in chain_conv_dict))) for account in main_df['ACCOUNT']
]

main_df['ACCOUNT_REMARKS'] = [''.join(list(set("maintain" if key in account else '' for key in chain_conv_dict))) for account in main_df['ACCOUNT']]
main_df['ACCOUNT_REMARKS'] = ['drop' if value == '' else 'maintain' for value in main_df['ACCOUNT_REMARKS']]

# -------------------------- Parsing Date Formats -------------------------- #
if str(main_df['START_DATE'].dtype) == 'int64' and str(main_df['END_DATE'].dtype) == 'int64':
    main_df['START_DATE'] = pd.to_timedelta(abs(main_df['START_DATE']), unit='d') + datetime.datetime(1899, 12, 30)
    main_df['START_DATE'] = pd.to_datetime(main_df['START_DATE'], format="%B/%d/%Y", errors="raise")
    main_df['START_DATE'].dt.strftime("%B-%d-%Y")

    main_df['END_DATE'] = pd.to_timedelta(abs(main_df['END_DATE']), unit='d') + datetime.datetime(1899, 12, 30)
    main_df['END_DATE'] = pd.to_datetime(main_df['END_DATE'], format="%B/%d/%Y", errors="raise")
    main_df['END_DATE'].dt.strftime("%B-%d-%Y")
else:
    main_df['START_DATE'] = pd.to_datetime(main_df['START_DATE'], format="%B/%d/%Y", errors='raise')
    main_df['START_DATE'].dt.strftime("%B-%d-%Y")
    main_df['END_DATE'] = pd.to_datetime(main_df['END_DATE'], format="%B/%d/%Y", errors='raise')
    main_df['END_DATE'].dt.strftime("%B-%d-%Y")

main_df['START_DAY'] = main_df['START_DATE'].dt.day
main_df['START_YEAR'] = [WORKING_YEAR for index in main_df.index.values]
main_df['END_DAY'] = main_df['END_DATE'].dt.day
main_df['END_YEAR'] = [WORKING_YEAR for index in main_df.index.values]

# ------------------- Adding cols for grouping reference ------------------- #
# TO-DO 1: replacing ':' with '-' in MARS_NAME column
activity_list = [element.replace(":", "-") if ":" in element else element for element in main_df['MARS_NAME'].to_list()]
main_df['MARS_NAME'] = activity_list

# TO-DO 2: adding WIN tag to MAINSTREAM POWDERED MILKS and PREMIUM POWDERED MILKS categories
main_df['CATEGORY'] = ["WIN MAINSTREAM POWDERED MILKS" if main_df['CATEGORY'].iloc[index] == "MAINSTREAM POWDERED MILKS"
 else "WIN PREMIUM POWDERED MILKS" if main_df['CATEGORY'].iloc[index] == "PREMIUM POWDERED MILKS"
 else main_df['CATEGORY'].iloc[index]
 for index in main_df.index.values]

# TO-DO 3: create ACTIVITY_TYPE_ADJ column for FORM_NAME/ID
main_df['ACTIVITY_TYPE_lower'] = [act_type.lower() for act_type in main_df['ACTIVITY_TYPE']]
main_df['ACTIVITY_TYPE_drop'] = [''.join(list(set("maintain" if key in act_type else "drop" for key in act_type_dict1))) for act_type in main_df['ACTIVITY_TYPE_lower']]
main_df['ACTIVITY_TYPE'] = [''.join(list(set(act_type_dict1[key][0] if key in act_type else '' for key in act_type_dict1))) for act_type in main_df['ACTIVITY_TYPE_lower']]

main_df['FORM_ID_ACTIVITY_TYPE'] = [''.join(list(set(act_type_dict1[key][1] if key in act_type else '' for key in act_type_dict1))) for act_type in main_df['ACTIVITY_TYPE_lower']]

# TO-DO 4: adding 'ACTIVITY' column
main_df['ACTIVITY'] = [f"{main_df['ACTIVITY_TYPE'].iloc[index]}: {main_df['MARS_NAME'].iloc[index]}" for index in main_df.index.values]

# TO-DO 5: adding 'DURATION' column as reference for 'GROUPING_REF'
# TO-DO 5.1: convert 'START_MONTH' and 'END_MONTH' to values in month_conv_dict
main_df['START_MONTH'] = [''.join(list(set(month_conv_dict[key] if key in month.lower() else '' for key in month_conv_dict))) for month in main_df['START_MONTH']]

main_df['END_MONTH'] = [''.join(list(set(month_conv_dict[key] if key in month.lower() else '' for key in month_conv_dict))) for month in main_df['END_MONTH']]

# TO-DO 6: adding 'DUPLICATE_REF' and 'DUPLICATE_drop' columns and filtering out last instances of duplicates from df
main_df['DUPLICATE_REF'] = [f"{main_df['ACCOUNT_adj'].iloc[index]} - {main_df['CATEGORY'].iloc[index]} - {main_df['ACTIVITY'].iloc[index]}" for index in main_df.index.values]
main_df['DUPLICATE_drop'] = main_df.duplicated(keep='last', subset=['DUPLICATE_REF'])

# TO-DO 7: filter out duplicate activities, drop #N/A customer codes, empty cells, and save to new main_df_filtered
main_df_filtered = main_df.dropna()
main_df_filtered = main_df_filtered[main_df_filtered.DUPLICATE_drop != True]
main_df_filtered = main_df_filtered[main_df_filtered.ACTIVITY_TYPE_drop != 'drop']
main_df_filtered.reset_index(drop=True, inplace=True)


'''
TO-DO 8: Add FORM_ID_DURATION column whose underlying values fit CS batch 1 timings, i.e durations outside of May should be converted to May 1 to 31, and those that fall between May 1 to 31 should be maintained as such
Note 8:
  Case 1 - Both START_MONTH and END_MONTH == WORKING_MONTH, in which case just follow START_DAY / END_DAY
  Case 2 - START_MONTH != WORKING_MONTH and END_MONTH == WORKING_MONTH, in which case START_DAY is hard-coded to '1', and END_DAY is maintained as is (e.g MAY 1 TO 31; MAY 1 TO 16)
  Case 3 - START_MONTH == WORKING_MONTH and END_MONTH != WORKING_MONTH, in which case START_DAY is maintained as is, and END_DAY is assigned with value from LAST_DAY constant (e.g MAY 5 TO 31; MAY 16 TO 31)
  Case 4 - Both START_MONTH and END_MONTH != WORKING_MONTH, in which case START_DAY is set to '1' TO {LAST_DAY}'
'''

main_df_filtered['FORM_ID_DURATION'] = [
    f"{WORKING_MONTH} {int(main_df_filtered['START_DAY'].iloc[index])} TO {int(main_df_filtered['END_DAY'].iloc[index])} {WORKING_YEAR}" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] == WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) >= FIRST_DAY and int(main_df_filtered['END_DAY'].iloc[index]) <= LAST_DAY
    else f"{main_df_filtered['START_MONTH'].iloc[index]} {main_df_filtered['START_DAY'].iloc[index]} TO {main_df_filtered['END_DAY'].iloc[index]} {WORKING_YEAR}" if
    main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[
        index] == WORKING_MONTH and main_df_filtered['END_DAY'].iloc[index] < FIRST_DAY
    else f"{WORKING_MONTH} 16 TO {int(main_df_filtered['END_DAY'].iloc[index])} {WORKING_YEAR}" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] == WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) <= FIRST_DAY and int(main_df_filtered['END_DAY'].iloc[index]) <= LAST_DAY

    else f"{WORKING_MONTH} {FIRST_DAY} TO {LAST_DAY} {WORKING_YEAR}" if main_df_filtered['START_MONTH'].iloc[index] != WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] != WORKING_MONTH

    else f"{WORKING_MONTH} {int(main_df_filtered['START_DAY'].iloc[index])} TO {LAST_DAY} {WORKING_YEAR}" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] != WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) >= FIRST_DAY
    else f"{WORKING_MONTH} {FIRST_DAY} TO {LAST_DAY} {WORKING_YEAR}" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] != WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) <= FIRST_DAY

    else f"{WORKING_MONTH} {FIRST_DAY} TO {int(main_df_filtered['END_DAY'].iloc[index])} {WORKING_YEAR}" if main_df_filtered['START_MONTH'].iloc[index] != WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] == WORKING_MONTH and int(main_df_filtered['END_DAY'].iloc[index]) <= LAST_DAY

    else "drop"
    for index in main_df_filtered.index.values
]

main_df_filtered['DURATION_REMARKS'] = [
    f"maintain" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] == WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) >= FIRST_DAY and int(main_df_filtered['END_DAY'].iloc[index]) <= LAST_DAY
    else f"drop - activity duration outside of CS Batch 2 timing(i.e 16 to 30/31)" if
    main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[
        index] == WORKING_MONTH and main_df_filtered['END_DAY'].iloc[index] < FIRST_DAY
    else f"maintain" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] == WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) <= FIRST_DAY and int(main_df_filtered['END_DAY'].iloc[index]) <= LAST_DAY

    else f"maintain" if main_df_filtered['START_MONTH'].iloc[index] != WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] != WORKING_MONTH

    else f"maintain" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] != WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) >= FIRST_DAY
    else f"maintain" if main_df_filtered['START_MONTH'].iloc[index] == WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] != WORKING_MONTH and int(main_df_filtered['START_DAY'].iloc[index]) <= FIRST_DAY

    else f"maintain" if main_df_filtered['START_MONTH'].iloc[index] != WORKING_MONTH and main_df_filtered['END_MONTH'].iloc[index] == WORKING_MONTH and int(main_df_filtered['END_DAY'].iloc[index]) <= LAST_DAY

    else "drop"
    for index in main_df_filtered.index.values
]


# TO-DO 9: add 'GROUPING_REF' column that uses 'FORM_ID_DURATION' to create unique primary keys
main_df_filtered['GROUPING_REF'] = [f"{main_df_filtered['ACCOUNT_adj'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['FORM_ID_DURATION'].iloc[index]}" for index in main_df_filtered.index.values]

# TO-DO 10: add 'COUNT' column that uses 'GROUPING_REF' for unique count values
main_df_filtered['COUNT'] = [[row for row in main_df_filtered['GROUPING_REF']].count(value) for value in main_df_filtered['GROUPING_REF']]


# -------------------------- Creating Groupings -------------------------- #

# TO-DO 1: group activities using unique values in 'GROUPING_REF' column as primary key; returns series of grouped activities as lists with unique values in 'GROUPING_REF' as index
grouped_activities_ser = main_df_filtered.groupby('GROUPING_REF')['ACTIVITY'].apply(lambda x: list(x))
# create dict from grouped_activities_ser using its index as keys, and the grouped activities in list form as values
grouped_activities_dict = {value: grouped_activities_ser[value] for value in grouped_activities_ser.index}
# create list grouped_activities_list and populating it with values from grouped_activities_dict using values from 'GROUPING_REF' as keys
grouped_activities_list = [grouped_activities_dict[value] for value in main_df_filtered['GROUPING_REF']]
# converting elements in grouped_activities_list in list form to strings with // separator for each activity and saving values in new 'GROUPED_ACTIVITIES' column
main_df_filtered['GROUPED_ACTIVITIES'] = [" // ".join(element) for element in grouped_activities_list]


# TO-DO 2: same thing as creating groupings for activities, but this one for activity types
grouped_activity_types_ser = main_df_filtered.groupby('GROUPING_REF')['FORM_ID_ACTIVITY_TYPE'].apply(lambda x: list(x))
grouped_activity_types_dict = {value: grouped_activity_types_ser[value] for value in grouped_activity_types_ser.index}
grouped_activity_types_list = [grouped_activity_types_dict[value] for value in main_df_filtered['GROUPING_REF']]
main_df_filtered['GROUPED_ACTIVITY_TYPES'] = [" // ".join(element) for element in grouped_activity_types_list]


# TO-DO 3: Add FORM_ID_ACTIVITY_TYPE column from grouped_activity_types_list; include each unique activity type
form_id_act_type = []
for value in grouped_activity_types_list:
    act_type_list = []
    for element in value:
        if element not in act_type_list:
            act_type_list.append(element)
        else:
            pass
    act_type_string = " / ".join(act_type_list)
    form_id_act_type.append(act_type_string)

main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'] = form_id_act_type

# ------------------- Creating Count IDs ------------------- #

# TO-DO 1: create form ID count reference to serve as primary key
main_df_filtered['FORM_ID_COUNT_REF'] = [f"{main_df_filtered['ACCOUNT_adj'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['FORM_ID_DURATION'].iloc[index]} - {main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'].iloc[index]} - {main_df_filtered['COUNT'].iloc[index]}" for index in main_df_filtered.index.values]

# TO-DO 2: create series of grouped activities based on primary key from ['FORM_ID_COUNT_REF']
FORM_ID_COUNT_GROUPED_ser = main_df_filtered.groupby('FORM_ID_COUNT_REF')['GROUPED_ACTIVITIES'].apply(lambda x: list(x))
# TO-DO 3: create a dict using values from FORM_ID_COUNT_GROUPED_ser's index as keys (these index values are values under the [FORM_ID_COUNT_REF'] column), and the keys' values from the same series. Structure is as follows: {'FOOD - JUNE 1 TO 30 - DISCOUNT / LOYALTY - 2': 'grouped DISCOUNT / LOYALTY activity'}
FORM_ID_COUNT_GROUPED_dict = {value: FORM_ID_COUNT_GROUPED_ser[value] for value in FORM_ID_COUNT_GROUPED_ser.index}
# TO-DO 4: create column ['FORM_ID_COUNT_GROUPED'] containing list of grouped activities belonging to the same value primary key/s under ['FORM_ID_COUNT_REF']
main_df_filtered['FORM_ID_COUNT_GROUPED'] = [list(set(FORM_ID_COUNT_GROUPED_dict[value])) for value in main_df_filtered['FORM_ID_COUNT_REF']]

# TO-DO 5: create column ['FORM_ID_COUNT_dict'] from values under ['FORM_ID_COUNT_GROUPED'] col. Each value is a list whose elements are grouped activities based on ['FORM_ID_COUNT_REF']. Those elements are then used as keys for dicts, after which they are used to query under ['GROUPED_ACTIVITIES'] to find the matching value under separate column ['COUNT']. They value for the key takes the form of <activity count><letter indentifier based on key's index in list>, e.g 2A, 2B, etc.
# structure for each column element is as follows: [{grouped_activity1: '1'}, {grouped_activity1A: '1A'}, {grouped_activity1B: '1B'}...{grouped_activity1N: '1N'}]
main_df_filtered['FORM_ID_COUNT_dict'] = [[{key: f"{main_df_filtered.loc[main_df_filtered['GROUPED_ACTIVITIES'] == key, 'COUNT'].iloc[0]}{form_id_count_dict[value.index(key)]}"} for key in value] for value in main_df_filtered['FORM_ID_COUNT_GROUPED']]

# TO-DO 6: loop through values under ['GROUPED_ACTIVITIES'] col and check them against values under ['FORM_ID_COUNT_dict'], going through each element in the latter and checking if the key matches the former, and then retrieving the key's value if it does
# populating form_id_count_list with unique form ID count (e.g 1A, 2B, etc) based on whether each list value under column 'FORM_ID_COUNT_dict' has a length of 1 or greater
# if length is equal to 1, list is appended with the value based on the current iteration of grouped_activity as the key
# if length is greater than 1, the list containing dicts of grouped activities:unique ID count pairs are iterated through, accessing their keys to check if the current iteration of grouped_activity matches the key of the current iteration of element
form_id_count_list = []
for index in main_df_filtered.index.values:
    grouped_activity = main_df_filtered['GROUPED_ACTIVITIES'].iloc[index]
    grouped_activity_list = main_df_filtered['FORM_ID_COUNT_dict'].iloc[index]
    if len(grouped_activity_list) == 1:
        form_id_count_list.append(grouped_activity_list[0].get(grouped_activity))
    else:
        for element in grouped_activity_list:
            keys_list = element.keys()
            if grouped_activity in keys_list:
                form_id_count_list.append(element.get(grouped_activity))

# TO-DO 7: create final column ['FORM_ID_COUNT'] from the populated form_id_count_list
main_df_filtered['FORM_ID_COUNT'] = form_id_count_list


# ------------------- Creating Form IDs and Form Names ------------------- #
main_df_filtered['FORM_ID'] = [f"CS - MID - {main_df_filtered['FORM_ID_DURATION'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['ACCOUNT_adj'].iloc[index]} - {main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'].iloc[index]} {main_df_filtered['FORM_ID_COUNT'].iloc[index]}" for index in main_df_filtered.index.values]

main_df_filtered['FORM_NAME'] = [f"CS - {WORKING_MONTH} - {main_df_filtered['CATEGORY'].iloc[index]}" for index in main_df_filtered.index.values]

# -------------------------- Writing to Excel -------------------------- #

final_columns = ['ACCOUNT_adj', 'ACCOUNT_REMARKS', 'CATEGORY', 'GROUPED_ACTIVITY_TYPES', 'GROUPED_ACTIVITIES', 'ACTIVITY', 'DURATION_REMARKS', 'FORM_ID', 'FORM_NAME']
main_df_filtered = main_df_filtered.filter(items=final_columns)

with (pd.ExcelWriter(
        f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CS_OutputFiles/B2/CS_{WORKING_MONTH}{WORKING_YEAR}_EXECUTION_B2_NCM_CHAINS_v1.1_test(2 - updated form timings).xlsx",
        engine="xlsxwriter", datetime_format='mmmm/dd/yy')
as writer): main_df_filtered.to_excel(writer, sheet_name="PER CHAIN")
