# %% [markdown]
# ## Analyzing the data extracts for key info
# 
# After running this notebook, we'll have key data points ready to plug into our coverage of the new test scores.
# 
# *Here's what we want:*
#  - Statewide proficiency in math, reading, science and social studies
#  - Equity insights into statewide proficiency in 8th-grade math and 3rd-grade reading
#  - Statewide 8th-grade PSAT scores and proficiency
#  - Statewide SAT scores (overall and broken down by subject)
#  - Most of the above for Detroit schools
#  - Top districts for 3rd-grade reading, 8th-grade math, PSAT and SAT scores in the state

# %%
import pandas as pd

# %% [markdown]
# **Statewide proficiency, grades 3-8**

# %%
#load data
data_3_8 = pd.read_csv("outputs/grades_3_8_proficiency_statewide.csv")

# %%
#filter to just M-STEP results (and PSAT for 8th grade)
mstep_3_8 = data_3_8[data_3_8["TestType"].isin(["M-STEP", "PSAT"])]
mstep_3_8.head()

# %%
#filter to relevant columns
mstep_3_8_small = mstep_3_8[["SchoolYear", "GradeContentTested", "Subject", "PercentMet", "PercentDidNotMeet"]]
mstep_3_8_small.head()

# %%
#readable outputs
grades = [3, 4, 5, 6, 7, 8]
subjects = ["ELA", "Mathematics", "Science", "Social Studies"]

for grade in grades:
    for subj in subjects:
        try:
            proficiency = mstep_3_8_small[(mstep_3_8_small["GradeContentTested"] == grade) & (mstep_3_8_small["Subject"] == subj)]["PercentMet"].values[0]
            print("----------------")
            print(f"{proficiency}% of students in grade {grade} met proficiency in {subj}.")
        except:
            continue

# %% [markdown]
# **Equity in ELA3 and Math8**

# %%
#load data
equity_data = pd.read_csv('outputs/grade_subject_report_categories_statewide.csv')

#filter to M-STEP and PSAT
equity_data = equity_data[equity_data['TestType'].isin(['M-STEP', 'PSAT'])]
equity_data.head()

# %%
#what are the groups we can break it out by?
equity_data["ReportCategory"].unique()

# %%
#split into categories
race = equity_data[equity_data['ReportCategory'].isin([
    "American Indian or Alaska Native",
    "Asian",
    "Black or African American",
    "Hispanic of Any Race",
    "Native Hawaiian or Other Pacific Islander",
    "Two or More Races",
    "White"
])]

econ = equity_data[equity_data['ReportCategory'].isin([
    "Economically Disadvantaged",
    "Not Economically Disadvantaged"
])]

gender = equity_data[equity_data['ReportCategory'].isin([
    "Female",
    "Male"
])]

language = equity_data[equity_data['ReportCategory'].isin([
    "English Learners",
    "Not English Learners"
])]

disability = equity_data[equity_data['ReportCategory'].isin([
    "Students With Disabilities",
    "Students Without Disabilities"
])]

# %%
#race
for group in race['ReportCategory'].unique():
    for grade in [3, 8]:
        for subj in ["ELA", "Mathematics"]:
            proficiency = race[
                (race["GradeContentTested"] == grade) &
                (race["Subject"] == subj) &
                (race["ReportCategory"] == group)
                ]["PercentMet"].values
            
            if proficiency.size > 0:
                print("----------------")
                print(f"{proficiency}% of students in grade {grade} who are {group} met proficiency in {subj}.")

# %%
#gender
for group in gender['ReportCategory'].unique():
    for grade in [3, 8]:
        for subj in ["ELA", "Mathematics"]:
            proficiency = gender[
                (gender["GradeContentTested"] == grade) &
                (gender["Subject"] == subj) &
                (gender["ReportCategory"] == group)
                ]["PercentMet"].values
            
            if proficiency.size > 0:
                print("----------------")
                print(f"{proficiency}% of students in grade {grade} who are {group} met proficiency in {subj}.")

# %%
#language
for group in language['ReportCategory'].unique():
    for grade in [3, 8]:
        for subj in ["ELA", "Mathematics"]:
            proficiency = language[
                (language["GradeContentTested"] == grade) &
                (language["Subject"] == subj) &
                (language["ReportCategory"] == group)
                ]["PercentMet"].values
            
            if proficiency.size > 0:
                print("----------------")
                print(f"{proficiency}% of students in grade {grade} who are {group} met proficiency in {subj}.")

# %%
#economic
for group in econ['ReportCategory'].unique():
    for grade in [3, 8]:
        for subj in ["ELA", "Mathematics"]:
            proficiency = econ[
                (econ["GradeContentTested"] == grade) &
                (econ["Subject"] == subj) &
                (econ["ReportCategory"] == group)
                ]["PercentMet"].values
            
            if proficiency.size > 0:
                print("----------------")
                print(f"{proficiency}% of students in grade {grade} who are {group} met proficiency in {subj}.")

# %%
#disability
for group in disability['ReportCategory'].unique():
    for grade in [3, 8]:
        for subj in ["ELA", "Mathematics"]:
            proficiency = disability[
                (disability["GradeContentTested"] == grade) &
                (disability["Subject"] == subj) &
                (disability["ReportCategory"] == group)
                ]["PercentMet"].values
            
            if proficiency.size > 0:
                print("----------------")
                print(f"{proficiency}% of students in grade {grade} who are {group} met proficiency in {subj}.")

# %% [markdown]
# **Statewide SAT scores**

# %%
#load data
sat = pd.read_csv('outputs/sat_statewide.csv')
sat.head()

# %%
sat.columns

# %%
#filter to "all students"
sat_all_students = sat[sat['Subgroup'] == 'All Students']
sat_all_students

# %%
#readable output
num_assessed = sat_all_students['MathNumAssessed'].values[0]
math_avg = sat_all_students["MathScoreAverage"].values[0]
EBRW_avg = sat_all_students["EBRWScoreAverage"].values[0]
math_ready = sat_all_students["MathPercentReady"].values[0]
EBRW_ready = sat_all_students["EBRWPercentReady"].values[0]

print("----------------")
print(f"Out of {num_assessed} total students who took the SAT statewide, the average Math score was {math_avg} and the average EBRW score was {EBRW_avg}.")
print(f"The percentage of students who were ready for college-level Math was {math_ready}% \n and the percentage of students who were ready for college-level EBRW was {EBRW_ready}%.")

# %% [markdown]
# **Summary of test scores for Detroit schools**

# %%
#load data
detroit_equity = pd.read_csv('outputs/grade_subject_report_categories_dpscd.csv')
detroit_3_8 = pd.read_csv("outputs/grades_3_8_proficiency_dpscd.csv")
detroit_psat = pd.read_csv("outputs/psat8_dpscd.csv")
detroit_sat = pd.read_csv("outputs/sat_dpscd.csv")

# %%
#3rd-8th grade proficiency
detroit_mstep = detroit_3_8[detroit_3_8["TestType"].isin(["M-STEP", "PSAT"])]
detroit_mstep_small = detroit_mstep[["SchoolYear", "GradeContentTested", "Subject", "PercentMet", "PercentDidNotMeet"]]
detroit_mstep_small

# %%
#detroit PSAT
detroit_psat[["DistrictName", "GradeContentTested", "Subject", "PercentMet", "PercentDidNotMeet"]]

# %%
#detroit SAT
detroit_sat[detroit_sat['Subgroup'] == 'All Students'][["DistrictName", "MathNumAssessed", "MathScoreAverage", "EBRWScoreAverage", "MathPercentReady", "EBRWPercentReady"]]

# %% [markdown]
# **Top districts for key metrics**

# %%
#load data
sat_all_districts = pd.read_csv('outputs/sat_all_districts.csv')
proficiency_all_districts = pd.read_csv('outputs/grades_3_8_proficiency_all_districts.csv')

# %%
#Top 10 districts for SAT score
sat_all_districts = sat_all_districts[sat_all_districts["AllSubjectScoreAverage"] != "*"]
sat_all_districts["AllSubjectScoreAverage"] = sat_all_districts["AllSubjectScoreAverage"].astype(float)
sat_all_districts.sort_values(by="AllSubjectScoreAverage", ascending=False)[["DistrictName", "AllSubjectScoreAverage"]].head(10)

# %%
#Top 10 districts for SAT (math)
sat_all_districts.sort_values(by="MathScoreAverage", ascending=False)[["DistrictName", "MathScoreAverage"]].head(10)

# %%
#Top 10 districts for SAT (EBRW)
sat_all_districts.sort_values(by="EBRWScoreAverage", ascending=False)[["DistrictName", "EBRWScoreAverage"]].head(10)

# %%
#Top 10 districts for ELA3 proficiency
proficiency_all_districts = proficiency_all_districts[proficiency_all_districts["PercentMet"] != "*"]
ela3_all_districts = proficiency_all_districts[(proficiency_all_districts["GradeContentTested"] == 3) & (proficiency_all_districts["Subject"] == "ELA")]
ela3_all_districts[['DistrictName', 'GradeContentTested', 'Subject', 'PercentMet']].sort_values(by="PercentMet", ascending=False).head(10)

# %%
#Top 10 districts for Math8 proficiency
proficiency_all_districts = proficiency_all_districts[proficiency_all_districts["PercentMet"] != "*"]
math8_all_districts = proficiency_all_districts[(proficiency_all_districts["GradeContentTested"] == 8) & (proficiency_all_districts["Subject"] == "Mathematics")]
math8_all_districts[['DistrictName', 'GradeContentTested', 'Subject', 'PercentMet']].sort_values(by="PercentMet", ascending=False).head(10)


