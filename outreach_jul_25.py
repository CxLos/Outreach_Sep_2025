# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
from datetime import datetime
import folium
import os
import sys
from collections import Counter
# -------------------------------
import requests
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# --------------------------------
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)

# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
# data_path = 'data/MarCom_Responses.xlsx'
# file_path = os.path.join(script_dir, data_path)
# data = pd.read_excel(file_path)
# df = data.copy()

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1xZ-OulU-SOfd6jraH2fEvvVdbSXIUOg-RA3PKZHP_GQ/edit?gid=0#gid=0"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
worksheet = sheet.worksheet("July")
data = pd.DataFrame(worksheet.get_all_records())
df = data.copy()

# Strip whitespace from columns and cell values
df.columns = df.columns.str.strip()
df = df.apply(
        lambda col: col.str.strip() if col.dtype == "object" or pd.api.types.is_string_dtype(col) else col
    )

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# Get the reporting month:
current_month = datetime(2025, 7, 1).strftime("%B")
report_year = datetime(2025, 7, 1).strftime("%Y")
report = 'Outreach'
# -------------------------------------------------
# print(df)
# print(df[["Date of Activity", "Total travel time (minutes):"]])
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns.tolist())
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())
# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

columns =  [
    'Client',
    'Project', 
    'Task', 
    'Kiosk', 
    'User', 
    'Group',
    'Tags', 
    'Description',
    'Collaborated Entity', 
    '# of People Engaged', 
    'Duration (h)',
    'Total Travel Time',

    'Email', 
    'Billable', 
    'Start Date',
    'Start Time', 
    'End Date', 
    'End Time', 
    'Duration (h)',
    'Duration (decimal)', 
    'Billable Rate (USD)', 
    'Billable Amount (USD)',
]

df = df[df['Project'] == 'Community Outreach Activity']
# print(df.head())

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

#  Please provide public information:    137
# Please explain event-oriented:        13

# ============================== Data Preprocessing ========================== #

# Check for duplicate columns
# duplicate_columns = df.columns[df.columns.duplicated()].tolist()
# print(f"Duplicate columns found: {duplicate_columns}")
# if duplicate_columns:
#     print(f"Duplicate columns found: {duplicate_columns}")

# Rename columns
df.rename(
    columns={
        "Client": "Client",
        "Project": "Project",
        # --------------------
        "Duration (h)": "Duration",
        "Total Travel Time": "Travel",
        "# of People Engaged": "Engaged",
        
        "Group": "Group",
        "Task": "Task",
        "Tags": "Tags",
        "User": "User",
        "Collaborated Entity": "Collab",
        # --------------------------------
        "Kiosk": "Kiosk",
        "Description": "Description",
        # "": "",
    }, 
inplace=True)

# print(df.dtypes)

# -------------------------- Care Events --------------------------- #

admin_events = len(df)
# print("Total Marcom events:", marcom_events)

# ---------------------------- Care Hours ---------------------------- #

# print("Duration Unique Before:", df['Duration'].unique().tolist())

# Convert duration strings (HH:MM:SS) to timedelta
df['Duration'] = pd.to_timedelta(df['Duration'], errors='coerce')
df['Duration'] = df['Duration'].dt.total_seconds() / 3600
df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')

# print("Duration Unique After:", df['Duration'].unique().tolist())

admin_hours = df['Duration'].sum()
admin_hours=round(admin_hours)
# print('Total Activity Duration:', admin_hours, 'hours')

# ------------------------- Care Time ------------------------------ #

# print("Travel Unique before:", df['Travel'].unique().tolist())

df['Travel'] = df['Travel'].fillna('0')

travel_unique =[
    '', 'None', 
    '1-30 Minutes', 
    '31-60 Minutes', 
    '61-90 Minutes', 
    '0-30 Minutes'
]

df['Travel'] = (
    df['Travel']
        .astype(str)
        .str.strip()
            .replace({
                '0-30 Minutes': '30',
                '1-30 Minutes': '30',
                '31-60 Minutes': '60',
                '61-90 Minutes': '90',
                '': '0',
                'None': '0',
            })
)

df['Travel'] = pd.to_numeric(df['Travel'], errors='coerce')

# print("Travel Unique After:", df['Travel'].unique().tolist())

df_travel = df['Travel'].sum()/60
df_travel = round(df_travel)
# print('Total Travel Time:', df_travel)

# ------------------------- Total Engaged ------------------------------ #

# print("Engaged Unique before:", df['Engaged'].unique().tolist())

df['Engaged'] = df['Engaged'].fillna('0')

engaged_unique = [
    'Between 1 and 10', 
    'None', 
    '', 
    'Between 11 and 20', 
    'Between 20 and 30'
]

df['Engaged'] = (
    df['Engaged']
        .astype(str)
        .str.strip()
            .replace({
                '': '0',
                'None': '0',
                'Between 1 and 10': '10',
                'Between 11 and 20': '20',
                'Between 20 and 30': '30',
                # '': '',
            })
)
# print("Engaged Unique after:", df['Engaged'].unique().tolist())

df['Engaged'] = pd.to_numeric(df['Engaged'], errors='coerce')

# print("Travel Unique After:", df['Travel'].unique().tolist())

df_engaged = df['Engaged'].sum()
# print('Total Engaged:', df_engaged)

# --------------------------- Care Group -------------------------- #

# print("Group Unique before:", df['Group'].unique().tolist())

df['Group'] = (
    df['Group']
        .astype(str)
            .str.strip()
            .replace({
                "" : "N/A",
            })
    )

group_categories = [
    'Coordination & Navigation', 
    'Information Technology', 
    'Outreach & Engagement',
    'Permanent Supportive Housing',
    'Administration',
    'Communications',
    'Marketing',
]

group_normalized = {cat.lower().strip(): cat for cat in group_categories}

# Counter to count matches
counter = Counter()

for entry in df['Group']:
    
    # Split and clean each category
    items = [i.strip().lower() for i in entry.split(",")]
    for item in items:
        if item in group_normalized:
            counter[group_normalized[item]] += 1

# for category, count in counter.items():
#     print(f"Support Counts: \n {category}: {count}")

df_group = pd.DataFrame(counter.items(), columns=['Group', 'Count']).sort_values(by='Count', ascending=False)

# df_group = df.groupby('Group').size().reset_index(name='Count')
# print('Admin Groups: \n', df_group)

group_bar=px.bar(
    df_group,
    x='Group',
    y='Count',
    color='Group',
    text='Count',
).update_layout(
    height=800, 
    width=1500,
    title=dict(
        text=f'{current_month} {report} Groups',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        # tickangle=-15,
        tickfont=dict(size=18), 
        title=dict(
            # text=None,
            text="Group",
            font=dict(size=20), 
        ),
        showticklabels=False  
        # showticklabels=True  
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  
        x=1.05,  
        y=1,  
        xanchor="left",  
        yanchor="top",  
        # visible=False
        visible=True
    ),
    hovermode='closest', 
    bargap=0.08,  
    bargroupgap=0,  
).update_traces(
    textposition='auto',
    hovertemplate='<b>%{label}</b>: %{y}<extra></extra>'
)

group_pie=px.pie(
    df_group,
    names="Group",
    values='Count' 
).update_layout(
    height=800,
    width=950,
    title=f'{current_month} Ratio of {report} Groups',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=150,
    textposition='auto',
    insidetextorientation='horizontal', 
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# --------------------------- Care Task -------------------------- #

# print("Task Unique before:", df['Task'].unique().tolist())

df['Task'] = (
    df['Task']
        .astype(str)
            .str.strip()
            .replace({
                "Newsletter - writing, editing, proofing" : "Newsletter",
                "" : "N/A",
                # "" : "",
            })
    )

# print("Task Unique after:", df['Task'].unique().tolist())

task_categories = [
'Communication & Correspondence',
'HR Support', 
'Research & Planning', 
'Key Event', 
'Data Archiving',
'General Maintenance',
'Record Keeping & Documentation',
'Desk Help Support', 
'Workforce Development', 
'Academic',
'Content, Line Editing, or Proofing', 
'Financial & Budgetary Management',
'Device Management', 
'Training',
'Social Media/YouTube', 
'Compliance & Policy Enforcement', 
'Health Education or Awareness', 
'Clinical Provider', 
'Website or Intranet Updates', 
'Field Outreach', 
'Tabling',
'Advocacy Partner',
'Board Support', 
'', 
'Selfcare Healing', 
'SDoH Provider', 
'Travel',
'Team Meeting',
'Newsletter - writing, editing, proofing', 
'Office Management'
]

# task_normalized = {cat.lower().strip(): cat for cat in task_categories}
# counter = Counter()

# for entry in df['Task']:
    
#     items = [i.strip().lower() for i in entry.split(",")]
#     for item in items:
#         if item in task_normalized:
#             counter[task_normalized[item]] += 1

# # for category, count in counter.items():
# #     print(f"Support Counts: \n {category}: {count}")

# df_group = pd.DataFrame(counter.items(), columns=['Group', 'Count']).sort_values(by='Count', ascending=False)

df_task = df.groupby('Task').size().reset_index(name='Count')
df_task = df_task.sort_values(by='Count', ascending=False)
# print('Admin Tasks: \n', df_group)

task_bar=px.bar(
    df_task,
    x='Task',
    y='Count',
    color='Task',
    text='Count',
).update_layout(
    height=1000, 
    width=1500,
    title=dict(
        text=f'{current_month} {report} Tasks',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        # tickangle=-15,
        tickfont=dict(size=18), 
        title=dict(
            # text=None,
            text="Task",
            font=dict(size=20), 
        ),
        showticklabels=False  
        # showticklabels=True  
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  
        x=1.05,  
        y=1,  
        xanchor="left",  
        yanchor="top",  
        # visible=False
        visible=True
    ),
    hovermode='closest', 
    bargap=0.08,  
    bargroupgap=0,  
).update_traces(
    textposition='auto',
    hovertemplate='<b>%{label}</b>: %{y}<extra></extra>'
)

task_pie=px.pie(
    df_task,
    names="Task",
    values='Count' 
).update_layout(
    height=1000,
    width=950,
    title=f'{current_month} Ratio of {report} Tasks',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=170,
    textposition='auto',
    insidetextorientation='horizontal', 
    texttemplate='%{value}<br>(%{percent:.2%})',
    # textinfo='none',  # Hides all labels
    # textposition='none',  # Ensures nothing is placed inside or outside the pie
    # texttemplate=None,  # Optional, but reinforces no custom text
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# --------------------------- Care Tags -------------------------- #

# print("Tags Unique before:", df['Tags'].unique().tolist())

df['Tags'] = (
    df['Tags']
        .astype(str)
            .str.strip()
            .replace({
                "" : "N/A",
                # "" : "",
            })
    )

tag_unique = [
"N/A",
'Tabling Event', 'AmeriCorps Duties, Handout', 'Presentation', 'Training', 'Movement Is Medicine', 'Timesheet / Impact Reporting', 'Handout', 'Care Network, Handout, Tabling Event', 'Tabling Event, Handout', 'Data Archiving, Documentation', 'Meeting, Movement Is Medicine', 'Meeting', 'Materials Review', '', 'Impromptu Discussion', 'AmeriCorps Duties', 'Documentation, Event Planning, Handout, Movement Is Medicine, Presentation, Research, writing, and editing, Sustainability Binder', 'PSH Work', 'Movement Is Medicine, Tabling Event', 'Newsletter / Announcements', 'HealthyCuts', 'Graphic and/or Creatives Design', 'Event Planning, Meeting, Movement Is Medicine, Phone Call', 'Phone Call', 'Event Planning, Meeting, Movement Is Medicine', 'Continuation, PSH Work', 'Social Media and/or Youtube', 'Board Support', 'Movement Is Medicine, Phone Call, Meeting', 'Documentation', 'Care Network', 'Event Planning', 'Movement Is Medicine, Data Archiving, Email, Documentation', 'Social Media and/or Youtube, Videography, Graphic and/or Creatives Design', 'Movement Is Medicine, Meeting', 'Graphic and/or Creatives Design, Social Media and/or Youtube, Videography', 'Documentation, HealthyCuts, Recent Change', 'Meeting, Event Planning, Phone Call', 'Polls/Surveys, Tabling Event', 'Meeting, Movement Is Medicine, Phone Call', 'Brand Messaging Strategy, Newsletter / Announcements'
]

# Flatten and clean the list
tag_categories = sorted(set(
    item.strip()
    for entry in tag_unique if entry  # Ensure we only process non-empty entries
    for item in entry.split(',')      # Split by comma if there are multiple entries
))

# Normalize the categories for matching
tag_normalized = {cat.lower(): cat for cat in tag_categories}
counter = Counter()

# Count occurrences of each category, regardless of combinations
for entry in df['Tags']:
    items = [i.strip().lower() for i in entry.split(",")]
    for item in items:
        if item in tag_normalized:
            counter[tag_normalized[item]] += 1

# for category, count in counter.items():
#     print(f"Support Counts: \n {category}: {count}")

df_tag = pd.DataFrame(
                counter.items(), 
                columns=['Tags', 'Count']
            ).sort_values(
                by='Count', 
                ascending=False)

# df_tag = df.groupby('Tags').size().reset_index(name='Count')
# print('Admin Tags:: \n', df_group)

tag_bar=px.bar(
    df_tag,
    x='Tags',
    y='Count',
    color='Tags',
    text='Count',
).update_layout(
    height=1100, 
    width=1500,
    title=dict(
        text=f'{current_month} {report} Tags',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        # tickangle=-15,
        tickfont=dict(size=18), 
        title=dict(
            # text=None,
            text="Tags",
            font=dict(size=20), 
        ),
        showticklabels=False  
        # showticklabels=True  
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  
        x=1.05,  
        y=1,  
        xanchor="left",  
        yanchor="top",  
        # visible=False
        visible=True
    ),
    hovermode='closest', 
    bargap=0.08,  
    bargroupgap=0,  
).update_traces(
    textposition='auto',
    hovertemplate='<b>%{label}</b>: %{y}<extra></extra>'
)

tag_pie=px.pie(
    df_tag,
    names="Tags",
    values='Count' 
).update_layout(
    height=1200,
    width=950,
    title=f'{current_month} Ratio of {report} Tags',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=160,
    textposition='auto',
    insidetextorientation='horizontal', 
    texttemplate='%{percent:.2%}',
    # texttemplate='%{value}<br>(%{percent:.2%})',
    # textinfo='none',  # Hides all labels
    # textposition='none',  # Ensures nothing is placed inside or outside the pie
    # texttemplate=None,  # Optional, but reinforces no custom text
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# --------------------------- Collaborated Entity -------------------------- #

# print("Collab Unique before:", df['Collab'].unique().tolist())
# print("Collab Value Counts before: \n", df['Collab'].value_counts())

# Replace blank strings with "N/A" and strip whitespace
df['Collab'] = (
    df['Collab']
    .astype(str)
    .str.strip()
    .replace({
        "": "N/A"
    })
)

# Create a list of normalized categories from your reference list
collab_unique = [
    'University of Texas at Austin',
    'CommunityCare',
    'Other', 
    '', 
    'N/A', 
    "Black Men's Health Clinic",
    "Black Men's Health Clinic, Other", 
    'Sustainable Food Center', 
    'Integral Care', 
    'GudLife, Integral Care', 
    'Downtown Austin Community Court', 
    'Austin-Travis County ECHO'
]

# Clean and split all entries, then flatten into a single list
all_collabs = []
for entry in df['Collab']:
    items = [i.strip() for i in str(entry).split(",") if i.strip() and i.strip().lower() != "n/a"]
    all_collabs.extend(items)

# Count occurrences
counter = Counter(all_collabs)

# for category, count in counter.items():
#     print(f"Support Counts: \n {category}: {count}")

df_collab = pd.DataFrame(counter.items(), columns=['Collab', 'Count']).sort_values(by='Count', ascending=False)

# print(df_collab)

# df_collab = df.groupby('Collab').size().reset_index(name='Count')
# print('Admin Collabs: \n', df_collab)

collab_bar=px.bar(
    df_collab,
    x='Collab',
    y='Count',
    color='Collab',
    text='Count',
).update_layout(
    height=850, 
    width=1500,
    title=dict(
        text=f'{current_month} {report} Collaborated Entities',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        # tickangle=-15,
        tickfont=dict(size=18), 
        title=dict(
            # text=None,
            text="Collaborations",
            font=dict(size=20), 
        ),
        showticklabels=False  
        # showticklabels=True  
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  
        x=1.05,  
        y=1,  
        xanchor="left",  
        yanchor="top",  
        # visible=False
        visible=True
    ),
    hovermode='closest', 
    bargap=0.08,  
    bargroupgap=0,  
).update_traces(
    textposition='auto',
    hovertemplate='<b>%{label}</b>: %{y}<extra></extra>'
)

collab_pie=px.pie(
    df_collab,
    names="Collab",
    values='Count' 
).update_layout(
    height=850,
    width=1500,
    title=f'{current_month} Ratio of Collaborated Entities',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=100,
    textposition='auto',
    insidetextorientation='horizontal', 
    # texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)


# --------------------------- Care Users -------------------------- #

# print("User Unique before:", df['User'].unique().tolist())

user_unique = [
'larrywallace.jr', 'Coby Albrecht', 'kiounis williams', 'Areebah Mubin', 'jaqueline.oviedo', 'Jordan Calbert', 'Sashricaa Manoj Kumar', 'Eric Roberts', 'pamela.parker', 'Angelita Delagarza', 'lavonne.williams', 'kimberly.holiday', 'Azaniah Israel', 'arianna.williams', 'antonio.montgomery', 'Michael Lambert', 'steve kemgang', 'tramisha.pete', 'toyacraney', 'felicia.chandler', 'Dominique Holman'
]

df['User'] = (
    df['User']
        .astype(str)
            .str.strip()
            .replace({
                "steve kemgang" : "Steve Kemgang",
                "toyacraney" : "Toya Craney",
                "felicia.chandler" : "Felicia Chandler",
                "tramisha.pete" : "Tramisha Pete",
                "jaqueline.oviedo" : "Jaqueline Oviedo",
                "larrywallace.jr" : "Larry Wallace Jr.",
                "kiounis williams" : "Kiounis Williams",
                "pamela.parker" : "Pamela Parker",
                "lavonne.williams" : "Lavonne Williams",
                "kimberly.holiday" : "Kimberly Holiday",
                "antonio.montgomery" : "Antonio Montgomery",
                "arianna.williams" : "Arianna Williams",
                "carlos.bautista" : "Carlos Bautista",
                "christi.freeman" : "Christi Freeman",
                "" : "",
            })
    )

# print("User Unique after:", df['User'].unique().tolist())

# user_categories = [

# ]

# user_normalized = {cat.lower().strip(): cat for cat in user_categories}
# counter = Counter()

# for entry in df['User']:
    
#     # Split and clean each category
#     items = [i.strip().lower() for i in entry.split(",")]
#     for item in items:
#         if item in user_normalized:
#             counter[user_normalized[item]] += 1

# for category, count in counter.items():
#     print(f"Support Counts: \n {category}: {count}")

# df_user = pd.DataFrame(counter.items(), columns=['User', 'Count']).sort_values(by='Count', ascending=False)

df_user = df.groupby('User').size().reset_index(name='Count')
df_user = df_user.sort_values(by='Count', ascending=False)
# print('Admin Groups: \n', df_group)

user_bar=px.bar(
    df_user,
    x='User',
    y='Count',
    color='User',
    text='Count',
).update_layout(
    height=850, 
    width=1500,
    title=dict(
        text=f'{current_month} User Submissions',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        # tickangle=-15,
        tickfont=dict(size=18), 
        title=dict(
            # text=None,
            text="User",
            font=dict(size=20), 
        ),
        showticklabels=False  
        # showticklabels=True  
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  
        x=1.05,  
        y=1,  
        xanchor="left",  
        yanchor="top",  
        # visible=False
        visible=True
    ),
    hovermode='closest', 
    bargap=0.08,  
    bargroupgap=0,  
).update_traces(
    textposition='auto',
    hovertemplate='<b>%{label}</b>: %{y}<extra></extra>'
)

user_pie=px.pie(
    df_user,
    names="User",
    values='Count' 
).update_layout(
    height=900,
    width=1500,
    title=f'{current_month} Ratio of User Submissions',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=180,
    textposition='auto',
    insidetextorientation='horizontal', 
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# ========================== DataFrame Table ========================== #

# MarCom Table
admin_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df[col] for col in df.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server

app.layout = html.Div(
    children=[ 
    html.Div(
        className='divv', 
        children=[ 
        html.H1(
            f'Community {report} Report', 
            className='title'),
        html.H1(
            f'{current_month} {report_year}', 
            className='title2'),
    html.Div(
        className='btn-box', 
        children=[
        html.A(
            'Repo',
            href=f'https://github.com/CxLos/{report}_{current_month}_{report_year}',
            className='btn'),
        ]),
    ]),    

# Data Table
# html.Div(
#     className='row0',
#     children=[
#         html.Div(
#             className='table',
#             children=[
#                 html.H1(
#                     className='table-title',
#                     children='Data Table'
#                 )
#             ]
#         ),
#         html.Div(
#             className='table2', 
#             children=[
#                 dcc.Graph(
#                     className='data',
#                     figure=admin_table
#                 )
#             ]
#         )
#     ]
# ),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_month} {report} Events']
            ),
            html.Div(
                className='circle',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high2',
                            children=[admin_events]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high3',
                children=[f'{current_month} People Engaged']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[df_engaged]
                    ),
                        ]
                    )
                ],
            ),
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_month} {report} Hours']
            ),
            html.Div(
                className='circle',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high6',
                            children=[admin_hours]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high3',
                children=[f'{current_month} {report} Travel Hours']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high8',
                            children=[df_travel]
                    ),
                        ]
                    )
                ],
            ),
            ]
        ),
    ]
),

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=group_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=group_pie
                )
            ]
        ),
    ]
),   
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=task_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=task_pie
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=tag_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=tag_pie
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=collab_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=collab_pie
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=user_bar
                )
            ]
        ),
    ]
),   

html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph33',
            children=[
                dcc.Graph(
                    figure=user_pie
                )
            ]
        ),
    ]
),   
])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# =================================== Updated Database ================================= #

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# updated_path = f'data/Admin_{current_month}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)

# with pd.ExcelWriter(data_path, engine='xlsxwriter') as writer:
#     df.to_excel(
#             writer, 
#             sheet_name=f'MarCom {current_month} {report_year}', 
#             startrow=1, 
#             index=False
#         )

#     # Create the workbook to access the sheet and make formatting changes:
#     workbook = writer.book
#     sheet1 = writer.sheets['MarCom April 2025']
    
#     # Define the header format
#     header_format = workbook.add_format({
#         'bold': True, 
#         'font_size': 13, 
#         'align': 'center', 
#         'valign': 'vcenter',
#         'border': 1, 
#         'font_color': 'black', 
#         'bg_color': '#B7B7B7',
#     })
    
#     # Set column A (Name) to be left-aligned, and B-E to be right-aligned
#     left_align_format = workbook.add_format({
#         'align': 'left',  # Left-align for column A
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })

#     right_align_format = workbook.add_format({
#         'align': 'right',  # Right-align for columns B-E
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })
    
#     # Create border around the entire table
#     border_format = workbook.add_format({
#         'border': 1,  # Add border to all sides
#         'border_color': 'black',  # Set border color to black
#         'align': 'center',  # Center-align text
#         'valign': 'vcenter',  # Vertically center text
#         'font_size': 12,  # Set font size
#         'font_color': 'black',  # Set font color to black
#         'bg_color': '#FFFFFF'  # Set background color to white
#     })

#     # Merge and format the first row (A1:E1) for each sheet
#     sheet1.merge_range('A1:Q1', f'MarCom Report {current_month} {report_year}', header_format)

#     # Set column alignment and width
#     # sheet1.set_column('A:A', 20, left_align_format)  

#     print(f"MarCom Excel file saved to {data_path}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create admin-jun-25
# heroku git:remote -a admin-jun-25
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx