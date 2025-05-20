import pandas as pd

# Load Excel file
file_path = "attendance.xlsx"
df = pd.read_excel(file_path)

# Get total students from Excel
total_students = df.shape[0]
print(f"Total students in sheet: {total_students}")

# How many students to display
num_students = int(input("Please enter the Number of Students: "))
print("\n.................................................")

# Trim DataFrame to selected number of students
df = df.head(num_students)

# Extract attendance columns (assumes 4th column onward are attendance)
attendance_data = df.iloc[:, 3:]

# Calculate attendance percentage
def calculate_percentage(row):
    total_classes = (row == 'P').sum() + (row == 'A').sum()
    if total_classes == 0:
        return 0
    present_count = (row == 'P').sum()
    return round((present_count / total_classes) * 100, 2)

df['Attendance'] = attendance_data.apply(calculate_percentage, axis=1)

# Assign marks based on percentage
def assign_marks(pct):
    if pct >= 70:
        return 5
    elif pct >= 60:
        return 4
    elif pct >= 45:
        return 3
    elif pct >= 30:
        return 2
    elif pct <= 30:
        return 1
    else:
        return 0

df['Marks'] = df['Attendance'].apply(assign_marks)

# Print Results
print("\nCalculated Attendance Percentage:")
print("No.  Name                                 ID                   Percentage  Marks")
for idx, row in df.iterrows():
    print(f"{idx + 1:<4} {row['Student\'s Name']:<35} {row['Student\'s ID']:<20} {row['Attendance']:<10}% {row['Marks']}")

# Summary counts
count_70 = df[df['Attendance'] >= 70].shape[0]
count_60 = df[(df['Attendance'] >= 60) & (df['Attendance'] < 70)].shape[0]
count_45 = df[(df['Attendance'] >= 45) & (df['Attendance'] < 60)].shape[0]
count_30_44 = df[(df['Attendance'] >= 30) & (df['Attendance'] < 45)].shape[0]
count_30 = df[df['Attendance'] <= 30].shape[0]

print("\n.....................................................")
print("\nAttendance Percentage (Student Count):")
print("No. Percentage     Count")
print(f"1. >= 70%         {count_70}")
print(f"2. 60% - 69%      {count_60}")
print(f"3. 45% - 59%      {count_45}")
print(f"4. 30% - 44%      {count_30_44}")
print(f"5. <= 30%         {count_30}")
