import pandas as pd
import numpy as np
import datetime
'''
Things to check before reading in CP NFO Excel Template:
1. Check that column names for both CP NFO templates match those in data_cols exactly (especially the 'PRODUCT NAME' column; check that there are no spaces before and after the column name in Excel
2. Change WORKING_MONTH constant
3. Copy file names for both CP NFO templates and paste into ################### Change excel file names/paths here ################### section
4. Run programnag
'''

# change to current working month and year accordingly
WORKING_MONTH = "AUGUST"
WORKING_YEAR = 2024
# change to last day of the current working month
LAST_DAY = 31


month_conv_dict = {
    1: "JANUARY",
    2: "FEBRUARY",
    3: "MARCH",
    4: "APRIL",
    5: "MAY",
    6: "JUNE",
    7: "JULY",
    8: "AUGUST",
    9: "SEPTEMBER",
    10: "OCTOBER",
    11: "NOVEMBER",
    12: "DECEMBER",
}

category_conv_dict = {
    "KOKO KRUNCH": "BREAKFAST CEREALS",
    "KOKOKRUNCH": "KOKO KRUNCH",
    "KITKAT": "CHOCOLATES",
    "COFFEE-MATE": "COFFEE AND CREAMER",
    "NESCAFE": "COFFEE AND CREAMER",
    "NESCAFE RTD": "COFFEE AND CREAMER",
    "STARBUCKS": "COFFEE AND CREAMER",
    "MAGGI": "FOOD",
    "NESTLE": "FOOD",
    "CHUCKIE": "FRESH MILK / READY TO DRINK - CHUCKIE",
    "BOOST": "HEALTH SCIENCE",
    "NUTREN": "HEALTH SCIENCE",
    "NESTOGEN": "INFANT / GROWING UP MILKS & SNACKS",
    "NAN": "INFANT / GROWING UP MILKS & SNACKS",
    "NIDO": "INFANT / GROWING UP MILKS & SNACKS",
    "MILO": "MILO (CHOCO POWDERED BEV)",
    "BEAR BRAND": "POWDERED MILKS",
    "BONNA": "WIN MAINSTREAM POWDERED MILKS",
    "BONAMIL": "WIN MAINSTREAM POWDERED MILKS",
    "BONAKID": "WIN MAINSTREAM POWDERED MILKS",
    "S-26": "WIN PREMIUM POWDERED MILKS",
    "PROMIL": "WIN PREMIUM POWDERED MILKS"
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


act_type_master_list = []
for value in act_type_dict1.values():
    if value[1] in act_type_master_list:
        pass
    else:
        act_type_master_list.append(value[1])


channel_drop_list = ['01', '02', '03', '04', '05', '06', '07', '09', '10', '11', '12', '13', '14', '15', '17', '19', '20', '21', '22', '23', '24', '25', '26', '27']

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

data_cols = ["BRAND NAME", "PRODUCT NAME", "Target Channel", "Target Region", "Activity Type", "Start Date", "End Date", "MARS"]

######################## Change excel file names/paths here ########################
main_df = pd.read_excel(
    f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CP_RawFiles/NFO/BEES x MARS x MyBuddy Template for Aug2024.xlsx",
    sheet_name="Proposed_Channel", index_col=None, header=3, usecols=data_cols, keep_default_na=False,
    dtype={"Target Channel": str, "Target Region": str})

ref_df = pd.read_excel(
    f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CP_RawFiles/NFO/BEES x MARS x MyBuddy Template for Jul2024.xlsx",
    sheet_name="Proposed_Channel", index_col=None, header=3, usecols=data_cols, keep_default_na=False,
    dtype={"Target Channel": str, "Target Region": str})
# -------------------------- Parsing Date Formats, Pre-filters, and Column Names -------------------------- #
# TO-DO 1: rename cols for both DFs
main_df = main_df.rename(
    columns={"BRAND NAME": "BRAND_NAME", "PRODUCT NAME": "PRODUCT_NAME", "Target Channel": "TARGET_CHANNEL",
             "Target Region": "TARGET_REGION", "Activity Type": "ACTIVITY_TYPE", "Start Date": "START_DATE",
             "End Date": "END_DATE"})
ref_df = ref_df.rename(
    columns={"BRAND NAME": "BRAND_NAME", "PRODUCT NAME": "PRODUCT_NAME", "Target Channel": "TARGET_CHANNEL",
             "Target Region": "TARGET_REGION", "Activity Type": "ACTIVITY_TYPE", "Start Date": "START_DATE",
             "End Date": "END_DATE"})

# TO-DO 2: parse dates
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

if str(ref_df['START_DATE'].dtype) == 'int64' and str(ref_df['END_DATE'].dtype) == 'int64':
    ref_df['START_DATE'] = pd.to_timedelta(abs(ref_df['START_DATE']), unit='d') + datetime.datetime(1899, 12, 30)
    ref_df['START_DATE'] = pd.to_datetime(ref_df['START_DATE'], format="%B/%d/%Y", errors="raise")
    ref_df['START_DATE'].dt.strftime("%B-%d-%Y")

    ref_df['END_DATE'] = pd.to_timedelta(abs(ref_df['END_DATE']), unit='d') + datetime.datetime(1899, 12, 30)
    ref_df['END_DATE'] = pd.to_datetime(ref_df['END_DATE'], format="%B/%d/%Y", errors="raise")
    ref_df['END_DATE'].dt.strftime("%B-%d-%Y")
else:
    ref_df['START_DATE'] = pd.to_datetime(ref_df['START_DATE'], format="%B/%d/%Y", errors='raise')
    ref_df['START_DATE'].dt.strftime("%B-%d-%Y")
    ref_df['END_DATE'] = pd.to_datetime(ref_df['END_DATE'], format="%B/%d/%Y", errors='raise')
    ref_df['END_DATE'].dt.strftime("%B-%d-%Y")

main_df['START_DATE'] = pd.to_datetime(main_df['START_DATE'], format="%B/%d/%Y", errors='coerce')
main_df['START_DATE'].dt.strftime("%B-%d-%Y")

main_df['END_DATE'] = pd.to_datetime(main_df['END_DATE'], format="%B/%d/%Y", errors='coerce')
main_df['END_DATE'].dt.strftime("%B-%d-%Y")

main_df['START_DAY'] = main_df['START_DATE'].dt.day
main_df['START_MONTH'] = [month_conv_dict[month] if month in month_conv_dict else month for month in main_df['START_DATE'].dt.month]
main_df['START_YEAR'] = main_df['START_DATE'].dt.year

main_df['END_DAY'] = main_df['END_DATE'].dt.day
main_df['END_MONTH'] = [month_conv_dict[month] if month in month_conv_dict else month for month in main_df['END_DATE'].dt.month]
main_df['END_YEAR'] = main_df['END_DATE'].dt.year

ref_df['START_DATE'] = pd.to_datetime(ref_df['START_DATE'], format="%B/%d/%Y", errors='coerce')
ref_df['START_DATE'].dt.strftime("%B-%d-%Y")

ref_df['END_DATE'] = pd.to_datetime(ref_df['END_DATE'], format="%B/%d/%Y", errors='coerce')
ref_df['END_DATE'].dt.strftime("%B-%d-%Y")

ref_df['START_DAY'] = ref_df['START_DATE'].dt.day
ref_df['START_MONTH'] = [month_conv_dict[month] if month in month_conv_dict else month for month in ref_df['START_DATE'].dt.month]
ref_df['START_YEAR'] = ref_df['START_DATE'].dt.year

ref_df['END_DAY'] = ref_df['END_DATE'].dt.day
ref_df['END_MONTH'] = [month_conv_dict[month] if month in month_conv_dict else month for month in ref_df['END_DATE'].dt.month]
ref_df['END_YEAR'] = ref_df['END_DATE'].dt.year

# TO-DO 3: pre-filter DF with 'Y' under ['MARS'] col
main_df = main_df.loc[main_df['MARS'] == 'Y']
main_df.reset_index(drop=True, inplace=True)

ref_df = ref_df.loc[ref_df['MARS'] == 'Y']
ref_df.reset_index(drop=True, inplace=True)

# ------------------- Adding cols for grouping reference ------------------- #
# TO-DO 1: convert and standardize activity types; add col for dropping activity types not included in act_type_dict1

main_df['ACTIVITY_TYPE_lower'] = [act_type.lower() for act_type in main_df['ACTIVITY_TYPE']]
ref_df['ACTIVITY_TYPE_lower'] = [act_type.lower() for act_type in ref_df['ACTIVITY_TYPE']]

main_df['ACTIVITY_TYPE'] = [
    ''.join(list(set(act_type_dict1[key][0] if key in act_type else '' for key in act_type_dict1))) for act_type in
    main_df['ACTIVITY_TYPE_lower']
]
ref_df['ACTIVITY_TYPE'] = [
    ''.join(list(set(act_type_dict1[key][0] if key in act_type else '' for key in act_type_dict1))) for act_type in
    ref_df['ACTIVITY_TYPE_lower']
]

main_df['FORM_ID_ACTIVITY_TYPE'] = [
    ''.join(list(set(act_type_dict1[key][1] if key in act_type else '' for key in act_type_dict1))) for act_type in
    main_df['ACTIVITY_TYPE_lower']
]
ref_df['FORM_ID_ACTIVITY_TYPE'] = [
    ''.join(list(set(act_type_dict1[key][1] if key in act_type else '' for key in act_type_dict1))) for act_type in
    ref_df['ACTIVITY_TYPE_lower']
]

main_df['ACTIVITY_TYPE_DROP'] = ["drop" if activity_type not in act_type_master_list else activity_type for
                                 activity_type in main_df['FORM_ID_ACTIVITY_TYPE']]

# TO-DO 2: replace ':' with '-' in PRODUCT_NAME and create ['ACTIVITY'] col
main_activity_list = [element.replace(":", "-") if ":" in element else element for element in
                 main_df['PRODUCT_NAME'].to_list()]
main_df['PRODUCT_NAME'] = main_activity_list
main_df['ACTIVITY'] = [f"{main_df['ACTIVITY_TYPE'].iloc[index]}: {main_df['PRODUCT_NAME'].iloc[index]}" for index in
                       main_df.index.values]

ref_activity_list = [element.replace(":", "-") if ":" in element else element for element in
                 ref_df['PRODUCT_NAME'].to_list()]
ref_df['PRODUCT_NAME'] = ref_activity_list
ref_df['ACTIVITY'] = [f"{ref_df['ACTIVITY_TYPE'].iloc[index]}: {ref_df['PRODUCT_NAME'].iloc[index]}" for index in
                      ref_df.index.values]

# TO-DO 3: convert Brand Names to MARS Category Names
main_df['CATEGORY'] = [category_conv_dict[brand] if brand in category_conv_dict else brand for brand in
                       main_df['BRAND_NAME']]
ref_df['CATEGORY'] = [category_conv_dict[brand] if brand in category_conv_dict else brand for brand in
                      ref_df['BRAND_NAME']]

'''
TO-DO 4: Add FORM_ID_DURATION and FORM_NAME_DURATION colS
Note 4:
  Case 1 - START_MONTH != END_MONTH, in which case both months are maintained
  Case 2 - START_MONTH == END_MONTH, in which case the duration follows the format 'JUNE 1 TO 30 2024'
'''
main_df['FORM_ID_DURATION'] = [
    f"{main_df['START_MONTH'].iloc[index]} {int(main_df['START_DAY'].iloc[index])} TO {main_df['END_MONTH'].iloc[index]} {int(main_df['END_DAY'].iloc[index])} {int(main_df['END_YEAR'].iloc[index])}" if
    main_df['START_MONTH'].iloc[index] != main_df['END_MONTH'].iloc[index]
    else f"{main_df['START_MONTH'].iloc[index]} {int(main_df['START_DAY'].iloc[index])} TO {int(main_df['END_DAY'].iloc[index])} {int(main_df['END_YEAR'].iloc[index])}"
    for index in main_df.index.values
]

main_df['FORM_NAME_DURATION'] = [
    f"{main_df['START_MONTH'].iloc[index]} TO {main_df['END_MONTH'].iloc[index]}" if main_df['START_MONTH'].iloc[index] != main_df['END_MONTH'].iloc[index]
    else f"{main_df['START_MONTH'].iloc[index]}"
    for index in main_df.index.values
]

ref_df['FORM_ID_DURATION'] = [
    f"{ref_df['START_MONTH'].iloc[index]} {int(ref_df['START_DAY'].iloc[index])} TO {ref_df['END_MONTH'].iloc[index]} {int(ref_df['END_DAY'].iloc[index])} {int(ref_df['END_YEAR'].iloc[index])}" if
    ref_df['START_MONTH'].iloc[index] != ref_df['END_MONTH'].iloc[index]
    else f"{ref_df['START_MONTH'].iloc[index]} {int(ref_df['START_DAY'].iloc[index])} TO {int(ref_df['END_DAY'].iloc[index])} {int(ref_df['END_YEAR'].iloc[index])}"
    for index in ref_df.index.values
]

# TO-DO 5: add ['CHANNEL-REGION'] cols
main_df['CHANNEL_REGION'] = [f"{main_df['TARGET_CHANNEL'].iloc[index]}-{main_df['TARGET_REGION'].iloc[index]}" for index
                             in main_df.index.values]
ref_df['CHANNEL_REGION'] = [f"{ref_df['TARGET_CHANNEL'].iloc[index]}-{ref_df['TARGET_REGION'].iloc[index]}" for index in
                            ref_df.index.values]

# TO-DO 6: add ['ACTIVITY_CHECK'] cols; main_df['ACTIVITY_CHECK'] will be checked against ref_df['ACTIVITY_CHECK']
main_df['ACTIVITY_CHECK'] = [
    f"{main_df['CATEGORY'].iloc[index]} - {main_df['CHANNEL_REGION'].iloc[index]} - {main_df['ACTIVITY'].iloc[index]} - {main_df['FORM_ID_DURATION'].iloc[index]}"
    for index in main_df.index.values]
ref_df['ACTIVITY_CHECK'] = [
    f"{ref_df['CATEGORY'].iloc[index]} - {ref_df['CHANNEL_REGION'].iloc[index]} - {ref_df['ACTIVITY'].iloc[index]} - {ref_df['FORM_ID_DURATION'].iloc[index]}"
    for index in ref_df.index.values]

# TO-D0 7: check if there are any repeating activities in current month that have already been maintained from previous month
ref_activity_list = ref_df['ACTIVITY_CHECK'].to_list()
main_df['ACTIVITY_DROP'] = ["drop" if value in ref_activity_list else value for value in main_df['ACTIVITY_CHECK']]

# TO-DO 8: add cols for dropping channels that are not in the AFS console
main_df['CHANNEL_DROP_REF'] = [main_df['TARGET_CHANNEL'].iloc[index][0:2] for index in main_df.index.values]
main_df['CHANNEL_DROP'] = [
    "drop" if main_df['CHANNEL_DROP_REF'].iloc[index] not in channel_drop_list else main_df['CHANNEL_DROP_REF'].iloc[index]
    for index in main_df.index.values]

# TO-DO 9: filter main_df and drop rows that correspond to repeating activities, as well as rows whose activity types are not in act_type_dict1
main_df_filtered = main_df.loc[main_df['ACTIVITY_DROP'] != 'drop']
main_df_filtered = main_df_filtered.loc[main_df_filtered['ACTIVITY_TYPE_DROP'] != 'drop']
main_df_filtered = main_df_filtered.loc[main_df_filtered['CHANNEL_DROP'] != 'drop']
main_df_filtered = main_df_filtered.dropna()
main_df_filtered.reset_index(drop=True, inplace=True)



# TO-DO 10: create ['GROUPING_REF'] col
main_df_filtered['GROUPING_REF'] = [
    f"{main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['CHANNEL_REGION'].iloc[index]} - {main_df_filtered['FORM_ID_DURATION'].iloc[index]}"
    for index in main_df_filtered.index.values]

# TO-DO 11: create ['COUNT'] col
main_df_filtered['COUNT'] = [[row for row in main_df_filtered['GROUPING_REF']].count(value) for value in
                             main_df_filtered['GROUPING_REF']]


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

main_df_filtered['CHANNEL_REGION_ACTIVITIES'] = [
    [f"{main_df_filtered['CHANNEL_REGION'].iloc[index]}, {main_df_filtered['GROUPED_ACTIVITIES'].iloc[index]}"] for
    index in main_df_filtered.index.values]
# ------------------- Creating Count IDs ------------------- #
'''
Some NFO channel-region forms will contain the same category, duration, combination of activity types, and activity count, but have different channel-region combinations, 
    Ex:
        CATEGORY    CHANNEL-REGION  ACTIVITY            COUNT   START DATE  END DATE        FORM ID WITHOUT UNIQUE COUNT                                                FORM ID WITH UNIQUE COUNT
        MILO        01,02-00,01     Discount Activity   2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2
        MILO        01,02-00,01     Bundling Activity   2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2
        MILO        02,03-01,02     Discount Activity   2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2A
        MILO        02,03-01,02     Bundling Activity   2       06/01/2024  06/30/2024      CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2     CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2A

where the form CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2 is different from CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 2A because they are tagged under two different region-channel combinations.

If, for example, the four activities above were tagged under the same channel-region combination, and assuming all other aspects remain the same, all four activities would instead be under the same form with a form ID of CP - JUNE 1 TO 30 2024 - MILO - NFO PER CHANNEL - DISCOUNT / BUNDLING 4
'''

# TO-DO 1: create form ID count reference to serve as primary key
main_df_filtered['FORM_ID_COUNT_REF'] = [
    f"{main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['FORM_ID_DURATION'].iloc[index]} - {main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'].iloc[index]} - {main_df_filtered['COUNT'].iloc[index]}"
    for index in main_df_filtered.index.values]

# TO-DO 2: create series of grouped activities based on primary key from ['FORM_ID_COUNT_REF']
FORM_ID_COUNT_GROUPED_ser = main_df_filtered.groupby('FORM_ID_COUNT_REF')['CHANNEL_REGION'].apply(lambda x: list(x))

# TO-DO 3: create a dict using values from FORM_ID_COUNT_GROUPED_ser's index as keys (these index values are values under the [FORM_ID_COUNT_REF'] column), and the keys' values from the same series. Structure is as follows: {'FOOD - JUNE 1 TO 30 - DISCOUNT / LOYALTY - 2': 'grouped DISCOUNT / LOYALTY activity'}
FORM_ID_COUNT_GROUPED_dict = {value: FORM_ID_COUNT_GROUPED_ser[value] for value in FORM_ID_COUNT_GROUPED_ser.index}

# TO-DO 4: create column ['FORM_ID_COUNT_GROUPED'] containing list of grouped channel-region combinations belonging to the same value primary key/s under ['FORM_ID_COUNT_REF']
main_df_filtered['FORM_ID_COUNT_GROUPED'] = [list(set(FORM_ID_COUNT_GROUPED_dict[value])) for value in
                                             main_df_filtered['FORM_ID_COUNT_REF']]

# TO-DO 5: create column ['FORM_ID_COUNT_dict'] from values under ['FORM_ID_COUNT_GROUPED'] col. Each value is a list whose elements are grouped activities based on ['FORM_ID_COUNT_REF']. Those elements are then used as keys for dicts, after which they are used to query under ['GROUPED_ACTIVITIES'] to find the matching value under separate column ['COUNT']. They value for the key takes the form of <activity count><letter indentifier based on key's index in list>, e.g 2A, 2B, etc.
# structure for each column element is as follows: [{channel_region1: '1'}, {channel_region1A: '1A'}, {channel_region1B: '1B'}...{channel_region1N: '1N'}]
main_df_filtered['FORM_ID_COUNT_dict'] = [[{
                                               key: f"{main_df_filtered.loc[main_df_filtered['CHANNEL_REGION'] == key, 'COUNT'].iloc[0]}{form_id_count_dict[value.index(key)]}"}
                                           for key in value] for value in main_df_filtered['FORM_ID_COUNT_GROUPED']]

# TO-DO 6: loop through values under ['GROUPED_ACTIVITIES'] col and check them against values under ['FORM_ID_COUNT_dict'], going through each element in the latter and checking if the key matches the former, and then retrieving the key's value if it does
# populating form_id_count_list with unique form ID count (e.g 1A, 2B, etc) based on whether each list value under column 'FORM_ID_COUNT_dict' has a length of 1 or greater
# if length is equal to 1, list is appended with the value based on the current iteration of grouped_activity as the key
# if length is greater than 1, the list containing dicts of grouped activities:unique ID count pairs are iterated through, accessing their keys to check if the current iteration of grouped_activity matches the key of the current iteration of element
form_id_count_list = []
for index in main_df_filtered.index.values:
    grouped_activity = main_df_filtered['CHANNEL_REGION'].iloc[index]
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

main_df_filtered['TEST'] = [
    f"{main_df_filtered['CATEGORY'].iloc[index]} - {main_df_filtered['FORM_ID_ACTIVITY_TYPE'].iloc[index]} {main_df_filtered['FORM_ID_COUNT'].iloc[index]}"
    for index in main_df_filtered.index.values]

# ------------------- Creating Form IDs and Form Names ------------------- #
main_df_filtered['FORM_ID'] = [
    f"CP - {main_df_filtered['FORM_ID_DURATION'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]} - NFO PER CHANNEL - {main_df_filtered['FORM_ID_ACTIVITY_TYPEgrouped'].iloc[index]} {main_df_filtered['FORM_ID_COUNT'].iloc[index]}"
    for index in main_df_filtered.index.values]

main_df_filtered['FORM_NAME'] = [
    f"CP - {main_df_filtered['FORM_NAME_DURATION'].iloc[index]} - {main_df_filtered['CATEGORY'].iloc[index]}" for index
    in main_df_filtered.index.values]

# -------------------------- Writing to Excel -------------------------- #
final_columns = ['CHANNEL_REGION', 'TARGET_CHANNEL', 'TARGET_REGION', 'CATEGORY', 'COUNT', 'GROUPED_ACTIVITY_TYPES', 'GROUPED_ACTIVITIES', 'FORM_ID',
                 'FORM_NAME']
main_df_filtered = main_df_filtered.filter(items=final_columns)

with (pd.ExcelWriter(
        f"{WORKING_MONTH}{WORKING_YEAR}_CPCS/CP_OutputFiles/NFO/CP_{WORKING_MONTH}{WORKING_YEAR}_NFO_CHANNEL_REGION_update_test.xlsx",
        engine="xlsxwriter", datetime_format='mmmm/dd/yy')
as writer): main_df_filtered.to_excel(writer, sheet_name="PER CHANNEL-REGION")