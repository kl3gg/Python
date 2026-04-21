"""
This script compares two Excel files (.xlsx) to identify usernames that are missing
from an email registration list.

It:
- Reads a list of usernames and a list of email addresses
- Normalizes both by removing whitespace, converting to lowercase, and stripping domains
- Compares the two lists to find usernames not present in the email list
- Outputs the missing usernames to a new Excel file

Useful for attendance tracking or registration validation
"""
import pandas as pd
from pathlib import Path


# File paths 
xlsx_users_path = Path("data/users.xlsx")  #List over usernames
xlsx_emails_path = Path("data/emails.xlsx") #Attendance list with usernames
output_path = Path("output/missing_usernames.xlsx")

def to_username(series: pd.Series) -> pd.Series:
    s = series.astype("string").str.strip().str.lower()
    s = s.str.split("@", n=1).str[0].str.strip()
    return s

# Load data (no header, specific columns only)
df_users = pd.read_excel(xlsx_users_path, header=None, usecols=[0], engine="openpyxl")
df_emails = pd.read_excel(xlsx_emails_path, header=None, usecols=[1], engine="openpyxl")

# Normalize usernames
users = to_username(df_users.iloc[:, 0]).dropna()
emails = to_username(df_emails.iloc[:, 0]).dropna()

# Remove empty strings
users = users[users.ne("")]
emails = emails[emails.ne("")]

# Find usernames that are not present in the email list (unique only)
missing_unique = users[~users.isin(emails)].drop_duplicates()

# Output results
print(f"Found {len(missing_unique)} usernames missing from the email list.")
missing_unique.to_frame("missing_usernames").to_excel(output_path, index=False)

print(f"Saved to: {output_path}")
