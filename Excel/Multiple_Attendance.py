"""
This script checks attendance consistency across multiple Excel registration files.

It:
- Reads a list of expected usernames from a master Excel file
- Scans a folder containing multiple registration Excel files
- Extracts and normalizes usernames from each file
- Counts how many times each user appears across all files
- Compares the actual count with the expected count (number of files)
- Identifies users who have attended fewer times than expected
- Outputs the result to a new Excel file

Useful for tracking attendance across multiple sessions and identifying missing participation.
"""

from pathlib import Path
import pandas as pd

# Update these paths
user_file = Path("data/users.xlsx")
registration_folder = Path("data/registrations")
output_file = Path("output/users_with_missing_attendance.xlsx")

def to_username(series: pd.Series) -> pd.Series:
    normalized = series.astype("string").str.strip().str.lower()
    normalized = normalized.str.split("@", n=1).str[0]
    return normalized


# Find all .xlsx files in the folder
xlsx_files = [
    file_path
    for file_path in xlsx_registry_path.iterdir()
    if file_path.is_file() and file_path.suffix.lower() == ".xlsx"
]

file_count = len(xlsx_files)

if file_count == 0:
    raise ValueError("No .xlsx files found in the registry folder.")

# Read usernames from Info_3.xlsx, column 1
df_users = pd.read_excel(
    xlsx_users_path,
    header=None,
    usecols=[0],
    engine="openpyxl"
)

usernames = to_username(df_users.iloc[:, 0]).dropna()
usernames = usernames[usernames.ne("")]
usernames = usernames.drop_duplicates()

# Read emails/usernames from all registry files
registry_list = []

for file_path in xlsx_files:
    df_temp = pd.read_excel(
        file_path,
        header=None,
        usecols=[1],
        engine="openpyxl"
    )
    registry_list.append(df_temp.iloc[:, 0])

all_registry_values = pd.concat(registry_series_list, ignore_index=True)
registered_usernames = to_username(all_registry_values).dropna()
registered_usernames = registered_usernames[registered_usernames.ne("")]

# Count occurrences in registry files
username_counts = registered_usernames.value_counts()

# Compare against expected count = number of xlsx files
result_df = pd.DataFrame({"username": usernames})
result_df["registered_count"] = result_df["username"].map(username_counts).fillna(0).astype(int)
result_df["expected_count"] = file_count
result_df["missing_count"] = result_df["expected_count"] - result_df["registered_count"]

users_with_too_few = result_df[
    result_df["registered_count"] < result_df["expected_count"]
].copy()

# Save result
users_with_too_few.to_excel(output_path, index=False)

print(f"Number of .xlsx files: {file_count}")
print(f"Found {len(users_with_too_few)} usernames with too few occurrences.")
print(f"Saved to: {output_path}")
