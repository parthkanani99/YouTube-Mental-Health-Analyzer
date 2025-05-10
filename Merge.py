import pandas as pd

# 📥 Paths to your two Excel files
file1 = r"C:\Users\parth\OneDrive\Documents\IDS 506 Mental health\.venv\New_Final_2.xlsx"
file2 = r"C:\Users\parth\OneDrive\Documents\IDS 506 Mental health\.venv\YT_Extracted Data.xlsx"

# 📥 Sheet names (if different from default "Sheet1")
sheet1 = "Sheet1"
sheet2 = "YouTubeData"

# 📥 Common column name to merge on (case-sensitive!)
common_column = "YouTube URL"

# 📄 Load the two Excel files
df1 = pd.read_excel(file1, sheet_name=sheet1)
df2 = pd.read_excel(file2, sheet_name=sheet2)

# 🔗 Merge the two DataFrames on the common column
merged_df = pd.merge(df1, df2, on=common_column, how="left") 
# 🔥 'how="inner"' keeps only matching rows
# You can use 'left', 'right', 'outer' instead depending on your need.

# 💾 Save the merged file
output_file = r"C:\Users\parth\OneDrive\Documents\IDS 506 Mental health\.venv\Ultimate.xlsx"
merged_df.to_excel(output_file, index=False)

print(f"✅ Merged file saved at: {output_file}")
