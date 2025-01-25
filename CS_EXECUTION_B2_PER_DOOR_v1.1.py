import pandas as pd
import numpy as np
import datetime

'''
Things to check before reading in CP-CS files:
1. Pre-format column headers in data template by replacing ' ' with an underscore '_'
2. Check if CATEGORY column in data template is named CATEGORY_FOR_MARS_FILL_UP, in which case change to just 'CATEGORY'
3. Check if MARS Activity column is named MARS_NAME. Change to MARS_NAME otherwise
4. Check data template if START_DATE and END_DATE columns are completely formatted, i.e not missing day, month, or year
5. Change WORKING_MONTH to current working month
6. Change LAST_DAY to last day of the current working month
7. Check that the data template's file name is copied exactly when reading it in
8. Check that the csv file containing maintained customer codes in AFS database is up-to-date, and its file name copied exactly when reading it in
9. Change output file names for csv and excel files towards the end of the code accordingly

'''

# change to current working month and year accordingly
WORKING_MONTH = "JULY"
WORKING_YEAR = 2024
# change to last day of the current working month
# no need to change FIRST_DAY
FIRST_DAY = 16
LAST_DAY = 31

data_cols = ["ACCOUNT", "CUSTOMER_CODE", "CATEGORY", "ACTIVITY_TYPE", "MARS_NAME", "START_DATE", "END_DATE", "START_MONTH", "END_MONTH"]


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

######################## Change excel file name/path here ########################
main_df = pd.read_excel(f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CS_RawFiles/B2/CUST SPEC JULY 2024 MARS UPLOADING - BATCH 2.xlsb", sheet_name="PER DOOR", index_col=None, header=1, usecols=data_cols, keep_default_na=False, dtype={"CUSTOMER_CODE": str, "START_MONTH": str, "END_MONTH": str})

######################### Change csv file name/path here #########################
afs_customers_df = pd.read_csv(f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/customers.csv", dtype={'AFS DOORS': str})
afs_customers_list = afs_customers_df['AFS DOORS'].to_list()
print(main_df.START_DATE.dtype)
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


# TO-DO 2: adding 'AMI-' to Alfamart customer codes or replacing 'ALF' with 'AMI-'
main_df['CUSTOMER_CODE_ADJ'] = [
    f"AMI-{main_df['CUSTOMER_CODE'].iloc[index]}" if main_df['ACCOUNT'].iloc[index] == "Alfamart"
    else f"AMI-{main_df['CUSTOMER_CODE'].iloc[index][3:]}" if str(main_df['CUSTOMER_CODE'].iloc[index][0:3]) == "ALF" and main_df['ACCOUNT'].iloc[index] == "ALFAMART"
    else main_df['CUSTOMER_CODE'].iloc[index]
    for index in main_df.index.values
]

# TO-DO 3: adding WIN tag to MAINSTREAM POWDERED MILKS and PREMIUM POWDERED MILKS categories
main_df['CATEGORY'] = ["WIN MAINSTREAM POWDERED MILKS" if main_df['CATEGORY'].iloc[index] == "MAINSTREAM POWDERED MILKS"
 else "WIN PREMIUM POWDERED MILKS" if main_df['CATEGORY'].iloc[index] == "PREMIUM POWDERED MILKS"
 else main_df['CATEGORY'].iloc[index]
 for index in main_df.index.values]

# TO-DO 4: check to see which customer codes are in the AFS customer database
# Note: 'CUSTOMER_CODE' column needs to be converted to str dtype and customers.csv file needs to be converted to list after being read in in order to look up customer codes in AFS customer database
main_df['AFS_CHECK'] = [
    main_df['CUSTOMER_CODE_ADJ'].iloc[index] if main_df['CUSTOMER_CODE_ADJ'].iloc[index] in afs_customers_list
    else "drop" if main_df['CUSTOMER_CODE_ADJ'].iloc[index] == "No Customer Code"
    else "drop"
    for index in main_df.index.values
]


# TO-DO 5: create ACTIVITY_TYPE_ADJ column for FORM_NAME/ID
main_df['ACTIVITY_TYPE_lower'] = [act_type.lower() for act_type in main_df['ACTIVITY_TYPE']]
main_df['ACTIVITY_TYPE_drop'] = [''.join(list(set("maintain" if key in act_type else "drop" for key in act_type_dict1))) for act_type in main_df['ACTIVITY_TYPE_lower']]
main_df['ACTIVITY_TYPE'] = [''.join(list(set(act_type_dict1[key][0] if key in act_type else '' for key in act_type_dict1))) for act_type in main_df['ACTIVITY_TYPE_lower']]

main_df['FORM_ID_ACTIVITY_TYPE'] = [''.join(list(set(act_type_dict1[key][1] if key in act_type else '' for key in act_type_dict1))) for act_type in main_df['ACTIVITY_TYPE_lower']]


# TO-DO 6: adding 'ACTIVITY' column
main_df['ACTIVITY'] = [f"{main_df['ACTIVITY_TYPE'].iloc[index]}: {main_df['MARS_NAME'].iloc[index]}" for index in main_df.index.values]

#  TO-DO 7: adding 'DUPLICATE_REF' and 'DUPLICATE_DROP' columns and filtering out last instances of duplicates from entire main_df
main_df['DUPLICATE_REF'] = [f"{main_df['CUSTOMER_CODE_ADJ'].iloc[index]} - {main_df['ACTIVITY'].iloc[index]}" for index in main_df.index.values]
main_df['DUPLICATE_drop'] = main_df.duplicated(keep='last', subset=['DUPLICATE_REF'])


# TO-DO 8: adding 'DURATION' column as reference for 'GROUPING_REF'
# TO-DO 8.1: convert 'START_MONTH' and 'END_MONTH' to values in month_conv_dict
main_df['START_MONTH'] = [''.join(list(set(month_conv_dict[key] if key in month.lower() else '' for key in month_conv_dict))) for month in main_df['START_MONTH']]

main_df['END_MONTH'] = [''.join(list(set(month_conv_dict[key] if key in month.lower() else '' for key in month_conv_dict))) for month in main_df['END_MONTH']]


# TO-DO 9: filter out duplicate activities, drop #N/A customer codes, empty cells, and save to new main_df_filtered
main_df_filtered = main_df[main_df.DUPLICATE_drop != True]
main_df_filtered = main_df_filtered[main_df_filtered.AFS_CHECK != "drop"]
main_df_filtered = main_df_filtered[main_df_filtered.ACTIVITY_TYPE_drop != "drop"]
main_df_filtered = main_df_filtered.dropna()
main_df_filtered.reset_index(drop=True, inplace=True)


'''
TO-DO 10: Add FORM_ID_DURATION column whose underlying values fit CS batch 1 timings, i.e durations outside of May should be converted to May 1 to 31, and those that fall between May 1 to 31 should be maintained as such
Note 10:
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

# TO-DO 11: 'GROUPING_REF' column that uses 'FORM_ID_DURATION' to create unique primary keys
main_df_filtered['GROUPING_REF'] = [f"{main_df_filtered['CUSTOMER_CODE_ADJ'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['FORM_ID_DURATION'].iloc[index]}" for index in main_df_filtered.index.values]

# TO-DO 12: 'COUNT' column that uses 'GROUPING_REF' for unique count values
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
# print(main_df_filtered['GROUPED_ACTIVITIES'])


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
'''
Some forms will contain the same category, duration, combination of activity types, and activity count. To account for this, the individual and unique activities are used to distinguish between forms, and are labeled as such: "Bundling 1" and "Bundling 1A", or "Discount / Loyalty 2" and "Discount / Loyalty 2A" -- where the form ID "Discount / Loyalty 2A" denotes a different combination of Discount and Loyalty activties from "Discount / Loyalty 2". 

    Ex:
        CATEGORY      ACTIVITY TYPE     ACTIVITY                 COUNT   START DATE  END DATE        FORM ID WITHOUT UNIQUE COUNT                                               FORM ID WITH UNIQUE COUNT
        MILO          Discount          Discount activity 1      2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2
        MILO          Loyalty           Loyalty activity 1       2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2
        MILO          Discount          Discount activity 1A     2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2A
        MILO          Loyalty           Loyalty activity 1A      2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / LOYALTY 2A

To achieve this, a reference is created to serve as the primary key for grouping the df according to 1. category (['CATEGORY']), 2. form duration (['FORM_ID_DURATION']), 3. combination of activity types (['GROUPED_ACTIVITY_TYPES']), and 4. activity count (['COUNT']). This is then used to collect and group values under 'GROUPED_ACTIVITIES' column that correspond or match each the value under 'FORM_ID_COUNT_REF' column into a series where the index are values of the latter and the column contains values of the former. From this, a dictionary is created where the keys are the individual values of the index (i.e values in the form of ['CATEGORY'] - ['FORM_ID_DURATION'] - ['GROUPED_ACTIVITY_TYPES'] - ['COUNT']), and its values are the grouped activities.

A new column ['FORM_ID_COUNT_GROUPED'] is then created from ['FORM_ID_COUNT_REF'] and is populated with lists containing grouped activities that share a common category, duration, activity types, and activity count. Some of these lists contain just 1 grouped activity, while others contain 3 to 5 -- all belonging to the same '[FORM_ID_COUNT_REF'] value. 

Using ['FORM_ID_COUNT_GROUPED'] column, a new column ['FORM_ID_COUNT_dict'] is created, and is populated with lists containing dictionaries, where keys of each dict is taken from the individual elements within list values under ['FORM_ID_COUNT_GROUPED'] (i.e unique combinations of grouped activities). Those keys are then used to filter and lookup under column ['GROUPED_ACTIVITIES'] in order to query its matching values under the column ['COUNT']. The unique ID count is then formed from form_id_count_dict, where the index of the grouped activity element corresponds to its unique count letter. (e.g ['first DISCOUNT / LOYALTY grouped activity', ' second DISCOUNT / LOYALTY grouped activity'] where the first element with index 0 corresponds to '', therefore having a form id count of 'DISCOUNT / LOYALTY 2', and the second element having an index of 1 corresponding to 'A' in the dict, in which case the resulting form id count will be 'DISCOUNT / LOYALTY 2A'. This will form the value corresponding to its grouped activities keys.  
'''

# TO-DO 1: create form ID count reference to serve as primary key
main_df_filtered['FORM_ID_COUNT_REF'] = [f"{main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['FORM_ID_DURATION'].iloc[index]} - {main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'].iloc[index]} - {main_df_filtered['COUNT'].iloc[index]}" for index in main_df_filtered.index.values]

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
main_df_filtered['FORM_ID'] = [f"CS - MID - {main_df_filtered['FORM_ID_DURATION'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]} - NCM PER DOOR - {main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'].iloc[index]} {main_df_filtered['FORM_ID_COUNT'].iloc[index]}" for index in main_df_filtered.index.values]

main_df_filtered['FORM_NAME'] = [f"CS - {WORKING_MONTH} - {main_df_filtered['CATEGORY'].iloc[index]}" for index in main_df_filtered.index.values]

# add ['ACTIVITY_ID'] col used as reference to list out all unique activities belonging to a form in an ungrouped, one-activity-per-cell manner
main_df_filtered['ACTIVITY_ID'] = [f"{main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['ACTIVITY'].iloc[index]} - {main_df_filtered['FORM_ID'].iloc[index]}" for index in main_df_filtered.index.values]

# --------- Creating seprate DF for Customer Code/Group Name csv --------- #
'''
Creates the separate .csv file to be uploaded to the AFS data loading tool. The .csv file collects the customer codes and are assigned to the groupings they're assigned to. The groupings need to be created beforehand in the AFS console before uploading the data from the .csv file
'''
afs_grouping_df = main_df_filtered.filter(items=['CUSTOMER_CODE_ADJ', 'FORM_ID'])

afs_grouping_df['DUPLICATE_REF'] = [f"{afs_grouping_df['CUSTOMER_CODE_ADJ'].iloc[index]} - {afs_grouping_df['FORM_ID'].iloc[index]}" for index in afs_grouping_df.index.values]
afs_grouping_df['DUPLICATE_DROP'] = afs_grouping_df.duplicated(keep='last', subset=['DUPLICATE_REF'])

afs_grouping_filtered_df = afs_grouping_df[afs_grouping_df.DUPLICATE_DROP != True]
afs_grouping_filtered_df = afs_grouping_filtered_df.filter(items=['FORM_ID', 'CUSTOMER_CODE_ADJ'])

afs_grouping_filtered_df['GroupType'] = [1 for index in afs_grouping_filtered_df.index.values]
afs_grouping_filtered_df['Delete'] = [0 for index in afs_grouping_filtered_df.index.values]

afs_grouping_filtered_df = afs_grouping_filtered_df.reindex(columns=['FORM_ID', 'GroupType', 'CUSTOMER_CODE_ADJ', 'Delete'])
afs_grouping_filtered_df = afs_grouping_filtered_df.rename(columns={'FORM_ID': 'Group_ID', 'CUSTOMER_CODE_ADJ': 'Reference_ID'})
afs_grouping_filtered_df.to_csv(f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CS_OutputFiles/B2/CS_{WORKING_MONTH}{WORKING_YEAR}_EXECUTION_B2_PER_DOOR_GROUPINGS_v1.1_test(1).csv", index=False)

# -------------------------- Writing to Excel -------------------------- #

final_columns1 = ['CUSTOMER_CODE_ADJ', 'CATEGORY', 'COUNT', 'GROUPED_ACTIVITY_TYPES', 'GROUPED_ACTIVITIES', 'FORM_ID', 'FORM_NAME']
main_df_filtered1 = main_df_filtered.filter(items=final_columns1)
main_df_filtered1.reset_index(drop=True, inplace=True)

final_columns2 = ['CATEGORY', 'COUNT', 'GROUPED_ACTIVITY_TYPES', 'GROUPED_ACTIVITIES', 'ACTIVITY_ID', 'ACTIVITY', 'FORM_ID', 'FORM_NAME']
main_df_filtered2 = main_df_filtered.filter(items=final_columns2)
main_df_filtered2 = main_df_filtered2.drop_duplicates(subset=['ACTIVITY_ID'])
main_df_filtered2.reset_index(drop=True, inplace=True)

names = ['PER DOOR (w customer codes)', 'PER DOOR (for form creation)']
dataframes = [main_df_filtered1, main_df_filtered2]

writer = pd.ExcelWriter(f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CS_OutputFiles/B2/CS_{WORKING_MONTH}{WORKING_YEAR}_EXECUTION_B2_NCM_DOORS_v1.1_test(1).xlsx", engine="xlsxwriter", datetime_format='mmmm/dd/yy')

for i, df in enumerate(dataframes):
    df.to_excel(writer, sheet_name=names[i])

writer.close()