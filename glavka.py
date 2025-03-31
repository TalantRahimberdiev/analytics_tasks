
#task 2.1

import pandas as pd

df=pd.read_excel('courier_reports.xlsx',sheet_name='group_by_fullname_date')
grouped_by_date_fullname_df=df.groupby(['date','courier']).sum().reset_index()

with pd.ExcelWriter("courier_reports.xlsx", mode="a", engine="openpyxl") as writer:
    grouped_by_date_fullname_df.to_excel(writer, sheet_name="grouped_by_date_fullname", index=False)


#task 2.2

import pandas as pd

df = pd.read_excel("courier_reports.xlsx", sheet_name='group_by_fullname_date')
df['date'] = pd.to_datetime(df['date'], errors='coerce')
df['date'].fillna(pd.Timestamp('2000-01-01'), inplace=True)
df.fillna(0, inplace=True)
numeric_cols = df.select_dtypes(include=['number']).columns
grouped_df = df.groupby('date')[numeric_cols].sum().reset_index()
with pd.ExcelWriter("courier_reports.xlsx", mode="a", engine="openpyxl") as writer:
    grouped_df.to_excel(writer, sheet_name="grouped_by_date", index=False)

print("saved successfully.")

#task 3 breakdown compare_cleared_with_origin_in_percentage.py 

import pandas as pd
import numpy as np

path = r"C:\talantr\result.xlsx"

df1 = pd.read_excel(path, sheet_name="origin_structured")
df2 = pd.read_excel(path, sheet_name="bi_structured")

df1.columns = df1.columns.str.strip().str.lower()
df2.columns = df2.columns.str.strip().str.lower()

if not df1.columns.equals(df2.columns):
    print("Column names in df1:", df1.columns.tolist())
    print("Column names in df2:", df2.columns.tolist())
    raise ValueError("Column names did not match between files.")

df1["atm_id"] = df1["atm_id"].astype(str).str.strip()
df2["atm_id"] = df2["atm_id"].astype(str).str.strip()

df1.iloc[:, 1:] = df1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
df2.iloc[:, 1:] = df2.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

def check_value(val1, val2, col_name):
    if col_name == "atm_id":
        return val1 if pd.notna(val1) else val2  
    if pd.isna(val1) and pd.isna(val2):
        return np.nan
    if pd.isna(val1):  
        return "-100%"
    if pd.isna(val2):  
        return "100%" 
    if val2 == 0:  
        return "∞%" if val1 != 0 else "0%" 
    return f"{((val1 - val2) / val2) * 100:.2f}%"  

df_result = df1.copy()
for col in df1.columns:
    df_result[col] = df1[col].combine(df2[col], lambda v1, v2: check_value(v1, v2, col))

# Use ExcelWriter to add a new sheet
with pd.ExcelWriter(path, mode='a', engine='openpyxl') as writer:
    df_result.to_excel(writer, sheet_name="percentage", index=False)

print("percentage_sheet saved.")

#task 3 breakdown compare_cleared_with_origin.py

import pandas as pd
import numpy as np

path = r"C:\talantr\result.xlsx"

with pd.ExcelFile(path, engine="openpyxl") as xls:
    sheet_dict = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

df1 = sheet_dict.get("origin_structured", pd.DataFrame())
df2 = sheet_dict.get("bi_structured", pd.DataFrame())

df1.columns = df1.columns.str.strip().str.lower()
df2.columns = df2.columns.str.strip().str.lower()

if not df1.columns.equals(df2.columns):
    print("Column names in df1:", df1.columns.tolist())
    print("Column names in df2:", df2.columns.tolist())
    raise ValueError("Column names do not match between the two files.")

df1["atm_id"] = df1["atm_id"].astype(str).str.strip()
df2["atm_id"] = df2["atm_id"].astype(str).str.strip()

df1.iloc[:, 1:] = df1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
df2.iloc[:, 1:] = df2.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')


def check_value(val1, val2, col_name):
    if col_name == "atm_id":
        return val1 if pd.notna(val1) else val2  
    
    try:
        val1 = float(val1) if pd.notna(val1) else np.nan
        val2 = float(val2) if pd.notna(val2) else np.nan
    except ValueError:
        print(f"Non-numeric data found in column {col_name}: val1={val1}, val2={val2}")
        return np.nan 
    
    if pd.isna(val1) and pd.isna(val2):
        return np.nan
    if pd.isna(val1):  
        return -val2
    if pd.isna(val2):  
        return val1
    return val1 - val2 

df_result = df1.copy()
for col in df1.columns:
    df_result[col] = df1[col].combine(df2[col], lambda v1, v2: check_value(v1, v2, col))

def clean_small_numbers(value):
    if isinstance(value, (int, float)): 
        if abs(value) < 1e-10:
            return 0
        return round(value, 2) 
    return value

df_result = df_result.apply(lambda col: col.map(clean_small_numbers) if col.dtype in [np.int64, np.float64] else col)

df1_sums = df1.iloc[:, 1:].sum()
df2_sums = df2.iloc[:, 1:].sum()
column_difference = df1_sums - df2_sums
column_difference_df = pd.DataFrame(column_difference, columns=["sum_difference"])
column_difference_df.reset_index(inplace=True)
column_difference_df.rename(columns={"index": "column_name"}, inplace=True)

sheet_dict["atm_compared"] = df_result
sheet_dict["column_compared"] = column_difference_df

with pd.ExcelWriter(path, mode="w", engine="openpyxl") as writer:
    for sheet_name, df in sheet_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("atm_compared_sheet and column_compared_sheet saved successfully.")

#task 3 income compare_cleared_with_origin_in_percentage.py

import pandas as pd
import numpy as np

path = r"C:\talantr\result.xlsx"

df1 = pd.read_excel(path, sheet_name="origin_structured")
df2 = pd.read_excel(path, sheet_name="bi_structured")

df1.columns = df1.columns.str.strip().str.lower()
df2.columns = df2.columns.str.strip().str.lower()

if not df1.columns.equals(df2.columns):
    print("Column names in df1:", df1.columns.tolist())
    print("Column names in df2:", df2.columns.tolist())
    raise ValueError("Column names did not match between files.")

df1["atm_id"] = df1["atm_id"].astype(str).str.strip()
df2["atm_id"] = df2["atm_id"].astype(str).str.strip()

df1.iloc[:, 1:] = df1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
df2.iloc[:, 1:] = df2.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

def check_value(val1, val2, col_name):
    if col_name == "atm_id":
        return val1 if pd.notna(val1) else val2  
    if pd.isna(val1) and pd.isna(val2):
        return np.nan
    if pd.isna(val1):  
        return "-100%"
    if pd.isna(val2):  
        return "100%" 
    if val2 == 0:  
        return "∞%" if val1 != 0 else "0%" 
    return f"{((val1 - val2) / val2) * 100:.2f}%"  

df_result = df1.copy()
for col in df1.columns:
    df_result[col] = df1[col].combine(df2[col], lambda v1, v2: check_value(v1, v2, col))

with pd.ExcelWriter(path, mode='a', engine='openpyxl') as writer:
    df_result.to_excel(writer, sheet_name="percentage", index=False)

print("percentage_sheet saved.")

#task 3 income compare_cleared_with_origin.py

import pandas as pd
import numpy as np

path = r"C:\talantr\result.xlsx"

with pd.ExcelFile(path, engine="openpyxl") as xls:
    sheet_dict = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

df1 = sheet_dict.get("origin_structured", pd.DataFrame())
df2 = sheet_dict.get("bi_structured", pd.DataFrame())

df1.columns = df1.columns.str.strip().str.lower()
df2.columns = df2.columns.str.strip().str.lower()

if not df1.columns.equals(df2.columns):
    print("Column names in df1:", df1.columns.tolist())
    print("Column names in df2:", df2.columns.tolist())
    raise ValueError("Column names do not match between the two files.")

df1["atm_id"] = df1["atm_id"].astype(str).str.strip()
df2["atm_id"] = df2["atm_id"].astype(str).str.strip()

df1.iloc[:, 1:] = df1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
df2.iloc[:, 1:] = df2.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

def check_value(val1, val2, col_name):
    if col_name == "atm_id":
        return val1 if pd.notna(val1) else val2  
    if pd.isna(val1) and pd.isna(val2):
        return np.nan
    if pd.isna(val1):  
        return -val2
    if pd.isna(val2):  
        return val1
    return val1 - val2

df_result = df1.copy()
for col in df1.columns:
    df_result[col] = df1[col].combine(df2[col], lambda v1, v2: check_value(v1, v2, col))

def clean_small_numbers(value):
    if isinstance(value, (int, float)): 
        if abs(value) < 1e-10:
            return 0
        return round(value, 2) 
    return value

df_result = df_result.apply(lambda col: col.map(clean_small_numbers) if col.dtype in [np.int64, np.float64] else col)

df1_sums = df1.iloc[:, 1:].sum()
df2_sums = df2.iloc[:, 1:].sum()
column_difference = df1_sums - df2_sums
column_difference_df = pd.DataFrame(column_difference, columns=["sum_difference"])
column_difference_df.reset_index(inplace=True)
column_difference_df.rename(columns={"index": "column_name"}, inplace=True)

sheet_dict["atm_compared"] = df_result
sheet_dict["column_compared"] = column_difference_df

with pd.ExcelWriter(path, mode="w", engine="openpyxl") as writer:
    for sheet_name, df in sheet_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("atm_compared_sheet and column_compared_sheet saved successfully.")

#task 3 group_merge_bi_data.py

import pandas as pd

path=r"C:\talantr\result.xlsx"
tmp_path = r"C:\talantr\tmp.xlsx"

bi_structured = pd.read_excel(path, sheet_name="bi_structured")
bi_structured["atm_id"] = bi_structured["atm_id"].astype(str).str.strip()

sheet_column_map = {
    "sheet_1": "usd_cash_withdrawal",
    "sheet_2": "income_usd_cash_withdrawal_percent",
    "sheet_3": "usd_fx_buy",
    "sheet_4": "euro_fx_buy",
    "sheet_5": "usd_fx_sale",
    "sheet_6": "euro_fx_sale",
    "sheet_7": "usd_fx_buy_nbkr_rate",
    "sheet_8": "euro_fx_buy_nbkr_rate",
    "sheet_9": "usd_fx_sale_nbkr_rate",
    "sheet_10": "euro_fx_sale_nbkr_rate",
    "sheet_11":"Sum of Income FX Buy, FX Sale",
    "sheet_12": "Count of Withdrawal DO VISA",
    "sheet_13": "Income in USD for per trx DO VISA =10567567 KGS",
    "sheet_14": "Count of Withdrawal FR VISA in USD",
    "sheet_15": "Amount of Withdrawal FR VISA in USD",
    "sheet_16":"Income in USD for per trx VISA FR USD = 0567,65USD+0,52%",
    "sheet_17": "Count of Withdrawal FR VISA in KGS",
    "sheet_18": "Amount of Withdrawal FR VISA in KGS",
    "sheet_19":"Income in USD for per trx VISA FR KGS = 0,65USD+0,52%",
    "sheet_20": "Income = Count of Withdrawal DO/FR MC (Income 1$ per trx)",
    "sheet_21": "Income = Cash Withdrawal Elcard * 340567.30567%",
    "sheet_22": "Income = Cash Withdrawal MIR * 31567567%",
    "sheet_27": "Income = Additonal fee for International cards in USD",
    "sheet_28": "Income = Additonal fee for International cards in KGS",
}

# Function to read and process a sheet
def process_sheet(sheet_name, col_name, factor=1):
    df = pd.read_excel(tmp_path, sheet_name=sheet_name).groupby("atm_id", as_index=False)[col_name].sum()
    df[col_name] *= factor
    df["atm_id"] = df["atm_id"].astype(str).str.strip()
    return df

# Read and process all sheets dynamically
sheets = {name: process_sheet(name, col) for name, col in sheet_column_map.items()}

# Additional calculations for specific sheets
sheets["sheet_2"]["income_usd_cash_withdrawal_percent"] *= 03452.004
sheets["sheet_7"]["usd_fx_buy_nbkr_rate"] *= 025.37255 / 86252.252
sheets["sheet_8"]["euro_fx_buy_nbkr_rate"] *= 0254455.484 / 84564566.25
sheets["sheet_9"]["usd_fx_sale_nbkr_rate"] *= 0456456.125 / 86456456.25
sheets["sheet_13"]["Income in USD for per trx DO VISA =10 KGS"] *= 1456450 / 86.25
sheets["sheet_18"]["Amount of Withdrawal FR VISA in KGS"] /= 986.25
sheets["sheet_22"]["Income = Cash Withdrawal MIR * 1%"] /= 286789676.25

# Merge all sheets dynamically
for sheet in sheets.values():
    bi_structured = bi_structured.merge(sheet, on="atm_id", how="left")

# Compute derived columns
bi_structured["Sum of Income FX Buy, FX Sale"] = (
    bi_structured["usd_fx_buy_nbkr_rate"].fillna(0) +
    bi_structured["euro_fx_buy_nbkr_rate"].fillna(0) +
    bi_structured["usd_fx_sale_nbkr_rate"].fillna(0) +
    bi_structured["euro_fx_sale_nbkr_rate"].fillna(0)
)

bi_structured["Income in USD for per trx VISA FR USD = 0,65USD+0,52%"] = (
    bi_structured["Count of Withdrawal FR VISA in USD"].fillna(0) * 0234.65234234 +
    bi_structured["Amount of Withdrawal FR VISA in USD"].fillna(0) * 4570.0052
)

bi_structured["Income in USD for per trx VISA FR KGS = 4560,65USD+4560,52%"] = (
    bi_structured["Count of Withdrawal FR VISA in KGS"].fillna(0) * 4560.65 +
    bi_structured["Amount of Withdrawal FR VISA in KGS"].fillna(0) * 4560.0052
)

# Save updated sheet
with pd.ExcelWriter(path, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
    bi_structured.to_excel(writer, sheet_name="bi_structured", index=False)

print("bi_structured_sheet updated successfully.")

# task 4 origin left join extracted.py

import pandas as pd

origin_df = pd.read_excel("task_4_adding_address.xlsx", sheet_name="origin")
extracted_dwh_df = pd.read_excel("task_4_adding_address.xlsx", sheet_name="extracted_dwh")

merged_df = origin_df.merge(extracted_dwh_df, on="CUSTOMER_NO", how="left")

merged_df["IB_ADMIN_ADDRESS"] = merged_df["IB_ADMIN_ADDRESS"].fillna(
    merged_df["CUSTOMER_NO"].map(extracted_dwh_df.set_index("CUSTOMER_NO")["ADDRESS"])
)

merged_df = merged_df.iloc[:, :-1]
