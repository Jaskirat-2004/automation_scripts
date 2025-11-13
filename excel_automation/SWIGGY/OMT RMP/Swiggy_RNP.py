import pandas as pd

import duckdb

con = duckdb.connect()

RNP_dump1 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump1.xlsx",
               header  = True,
               sheet = "SFDC Ticket",
               all_varchar=true)
""").df()

RNP_dump2 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump2.xlsx",
               header  = True,
               sheet = "SFDC Ticket",
               all_varchar=true)
""").df()

RNP_dump3 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump3.xlsx",
               header  = True,
               sheet = "SFDC Ticket",
               all_varchar=true)
""").df()

RNP_dump4 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump4.xlsx",
               header  = True,
               sheet = "SFDC Ticket",
               all_varchar=true)
""").df()


print("Combining dataframes...")
df = pd.concat([RNP_dump1, RNP_dump2 ,RNP_dump3,RNP_dump4], ignore_index=True)
print(f"Combined data: {len(df)} rows")

# Remove '+05:30', convert to datetime, and extract only the date
df['AGENT_ASSIGNED_TIME'] = (
    df['AGENT_ASSIGNED_TIME']
    .str.replace(r'\+05:30', '', regex=True)     # remove timezone part
    .pipe(pd.to_datetime, errors='coerce')       # convert to datetime
    .dt.date                                      # keep only date (YYYY-MM-DD)
)

# ✅ Step A: Handle missing 'FTR' column safely
df['FTR'] = df.get('FTR', 0)

# ✅ Step B: Fill missing values in group_cols to avoid losing rows 
group_cols = ['TYPE', 'AGENT_ASSIGNED_TIME', 'AGENT_EMAIL', 'CAMPAIGN', 'LANGUAGE']
#group_cols = ['TYPE', 'Date', 'AGENT_EMAIL', 'CAMPAIGN', 'LANGUAGE']
df[group_cols] = df[group_cols].fillna('Unknown')

# ✅ Step C: Create FTR_Final column
df['FTR_Final'] = df['FTNR'].fillna(df['FTR'])
df['FTR_Final'] = df['FTR_Final'].apply(lambda x: 1 if x == 1 else x)

# ✅ Step D: Create DNR Flag
df['DNR_Flag'] = df['L2_DISPOSITION'].isin(['DP Call Drop', 'DX Call Drop', 'Dx Call Drop'])

# ✅ Step E: Create SD Flag
sd_values = [
    'DP Call Drop',
    'Swiggy Captain - Selected a Different Actual Node',
    'Issue Type - Cusomer is not reachable',
    'Issue Type - Customer is not reachable',
    'Service Denial',
    'Dx Call Drop',
    'DX Call Drop'
]
df['SD_Flag'] = df['L2_DISPOSITION'].isin(sd_values)

# ✅ Step 1: Calculate DNR Count properly
dnr_count = (
    df[df['DNR_Flag']]
    .groupby(group_cols)
    .size()
    .reset_index(name='DNR_Count')
)

# ✅ Step 1.1: Calculate SD Count
sd_count = (
    df[df['SD_Flag']]
    .groupby(group_cols)
    .size()
    .reset_index(name='SD_Count')
)

# ✅ Step 2: Create DNR summary with Yes/No
df1 = (
    df.groupby(group_cols)['DNR_Flag']
      .any()
      .reset_index()
)
df1['DNR'] = df1['DNR_Flag'].apply(lambda x: 'Yes' if x else 'No')
df1 = df1.drop(columns=['DNR_Flag'])

# ✅ Step 2.1: Create SD summary with Yes/No
sd_summary = (
    df.groupby(group_cols)['SD_Flag']
      .any()
      .reset_index()
)
sd_summary['SD'] = sd_summary['SD_Flag'].apply(lambda x: 'Yes' if x else 'No')
sd_summary = sd_summary.drop(columns=['SD_Flag'])

# ✅ Step 3: Merge DNR and SD counts
df1 = df1.merge(dnr_count, on=group_cols, how='left')
df1 = df1.merge(sd_summary, on=group_cols, how='left')
df1 = df1.merge(sd_count, on=group_cols, how='left')

df1['DNR_Count'] = df1['DNR_Count'].fillna(0).astype(int)
df1['SD_Count'] = df1['SD_Count'].fillna(0).astype(int)

# ✅ Step 4: Calculate Like and Dislike counts
like_count = (
    df[df['DE_SAT_SCORE'] == 1]
    .groupby(group_cols)
    .size()
    .reset_index(name='Like')
)

dislike_count = (
    df[df['DE_SAT_SCORE'] == 0]
    .groupby(group_cols)
    .size()
    .reset_index(name='Dislike')
)

# ✅ Step 5: Merge Like/Dislike counts
df1 = df1.merge(like_count, on=group_cols, how='left')
df1 = df1.merge(dislike_count, on=group_cols, how='left')

# ✅ Step 6: Fill NaN with 0
df1[['Like', 'Dislike']] = df1[['Like', 'Dislike']].fillna(0).astype(int)

# ✅ Step 7: Add total call count per agent per day
call_count = (
    df.groupby(group_cols)
    .size()
    .reset_index(name='Call_Count')
)
df1 = df1.merge(call_count, on=group_cols, how='left')

# ✅ Step 8: Keep FTR, FTNR, and FTR_Final for reference (first occurrence)
extra_cols = df[group_cols + ['FTR', 'FTNR', 'FTR_Final']].drop_duplicates(subset=group_cols)
df1 = df1.merge(extra_cols, on=group_cols, how='left')

# ✅ Step 9: Export final result
df1.to_csv(r"C:\Users\8242K\Desktop\WFM\SWIGGY\OMT RMP\dumps\rnp.csv", index=False)


##########################################################################################################################
# Read the Excel file

aht_dump1 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump1.xlsx",
               header  = True,
               sheet = "SFDC Calldetails",
               all_varchar=true)
""").df()

aht_dump2 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump2.xlsx",
               header  = True,
               sheet = "SFDC Calldetails",
               all_varchar=true)
""").df()

aht_dump3 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump3.xlsx",
               header  = True,
               sheet = "SFDC Calldetails",
               all_varchar=true)
""").df()

aht_dump4 = con.execute("""
SELECT *
FROM read_xlsx("C:/Users/8242K/Desktop/WFM/SWIGGY/OMT RMP/dumps/dump4.xlsx",
               header  = True,
               sheet = "SFDC Calldetails",
               all_varchar=true)
""").df()


print("Combining dataframes...")
aht_df = pd.concat([aht_dump1, aht_dump2 ,aht_dump3,aht_dump4], ignore_index=True)
print(f"Combined data: {len(aht_df)} rows")


# Define grouping columns
req_cols = ['TYPE', 'DATE', 'AGENT_EMAIL']

# Filter out Conference calls
aht_df = aht_df[aht_df['CALL_DIRECTION'] != 'Conference']

# Fill missing values
aht_df[req_cols] = aht_df[req_cols].fillna('Unknown')

# ✅ Group and aggregate instead of transform
final_df = (
    aht_df
    .groupby(req_cols, as_index=False)
    .agg({'AHT': 'sum', 'UNIQUE_CASE_COUNT': 'sum'})
    .rename(columns={'AHT': 'Total_AHT'})
)

# Save to CSV
final_df.to_csv(r"C:\Users\8242K\Desktop\WFM\SWIGGY\OMT RMP\dumps\aht.csv", index=False)




