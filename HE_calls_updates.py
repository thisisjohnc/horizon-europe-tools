"""
Script Name: HE_calls_updates.py
Description: This script retrieves call information from the EU Funding and Tenders Portal for the Horizon Europe program and organises them into a spreadsheet 
that can be filtered or sorted by cluster or date. It can optionally produces visual calendars of the calls and deadlines
Author: John Creech
Date: May 2023

Requirements:
- Python 3.6+
- requests
- pandas
- openpyxl
- tqdm

Usage:
- Run the script without any arguments to fetch updated calls and save to an Excel sheet
- Use the '-n' or '--newonly' flag to save outputs only if there are changes from the compare file
- Use the '-l' or '--local' option to run on local data without downloading new calls (mainly for testing)
- Use the '-c' or '--calendars' option to produce a visual calendar of call opening and closing dates as a pdf
- Optionally specify a file to compare to using the 'file' argument, otherwise it defaults to most recent one in the script folder if no file is specified

Example:
    python HE_calls_updates.py -nc HE_calls_2024-04-16.xlsx
    This will download updates, compare with the file HE_calls_2024-04-16.xlsx, and, if new calls are found, save an updated spreadsheet and calendars
"""

import requests
import datetime
import json
import re
import pandas as pd
import os
import glob
import sys
import openpyxl
import warnings
warnings.filterwarnings('ignore')
import subprocess
from tqdm import tqdm
import argparse
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

def parse_arguments():
    parser = argparse.ArgumentParser(description="Get updated Horizon Europe calls from the Funding and Tenders Portal and put into a spreadsheet.")
    parser.add_argument("-n", "--newonly", help="save output only when new calls are found", action="store_true") 
    parser.add_argument("-l", "--local", help="skip downloading new calls and use previously saved local data", action="store_true")
    parser.add_argument("-c", "--calendars", help="save visual calendar of call dates", action="store_true")
    parser.add_argument("file", nargs='?', help="optionally specify file to compare to")
    return parser.parse_args()

def download_json_with_progress(url):
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))
    progress_bar = tqdm(total=total_size, unit='B', unit_scale=True)
    data = b""
    for chunk in response.iter_content(chunk_size=1024):
        data += chunk
        progress_bar.update(len(chunk))
    progress_bar.close()
    # Decode the byte string and parse as JSON
    return json.loads(data.decode('utf-8')) 

def process_data(grantsdata):
    # # Unpack data
    df = pd.json_normalize(grantsdata['fundingData']['GrantTenderObj'])

    print("Processing data...")
    ## filter to just pillar 2
    df['divAbbrev']=pd.json_normalize(pd.json_normalize(df['programmeDivision'])[0])['abbreviation']
    df=df[df['divAbbrev'].str.contains("HORIZON.2", na=False)]
    # Some calls are in the database twice, with slightly different descriptions. Only want one
    df.drop_duplicates(subset='ccm2Id',inplace=True)

    ## # Format the date columns with the datetime format so they output nicely (from ms after unix epoch)
    df['openDate'] = pd.to_datetime(df['plannedOpeningDateLong'],unit='ms').dt.strftime('%Y-%m-%d')
    df['pubDate'] = pd.to_datetime(df['publicationDateLong'],unit='ms').dt.strftime('%Y-%m-%d')
    df['closeDate']=pd.to_datetime([df['deadlineDatesLong'].iloc[i][0] for i in range(len(df))],unit='ms')
    df['closeDate'] = df['closeDate'].dt.strftime('%Y-%m-%d')
    # get second stage closing dates
    df['s2Date'] = pd.NaT  # Initialize with NaT (Not a Time) values
    df['s2Date'] = df.apply(lambda row: pd.to_datetime(row['deadlineDatesLong'][1], unit='ms') if len(row['deadlineDatesLong']) > 1 else pd.NaT, axis=1)
    df['s2Date'] = df['s2Date'].dt.strftime('%Y-%m-%d')


    # Flatten some JSON columns
    for i in range(0,2):
        df=df.join(pd.json_normalize(pd.json_normalize(df['programmeDivision'])[i]).add_prefix('programmeDivision.').add_suffix('.'+str(i)).set_index(df.index))
    a=df.explode('topicActions')
    df=df.join(pd.json_normalize(a.topicActions).add_prefix('topicActions.').set_index(df.index))

    # Extract destination names
    for i in range(len(df)):
        # if(len(df['programmeDivision.abbreviation.0'].iloc[i])==13): 
        if(len(df['programmeDivision.abbreviation.0'].iloc[i])>len(df['programmeDivision.abbreviation.1'].iloc[i])): ## whichever has longer ID is the right one
            df['destination'].iloc[i]=df['programmeDivision.description.0'].iloc[i] 
        else:
            df['destination'].iloc[i]=df['programmeDivision.description.1'].iloc[i]

    # Extract info from call identifiers (e.g., HORIZON-CL5-2023-D2-01)
    df['clusterCode'] = [re.split(r"-[0-9]{4}",x)[0] for x in df['callIdentifier']]
    df['callYear'] = [re.search(r"[0-9]{4}",x)[0] for x in df['callIdentifier']]
    # df['destCode'] = [re.split(r"-[0-9]{4,7}-",x)[1] for x in df['callIdentifier']] ## indexes out of range when applied to all
    df['destCode']=df['callYear'] #make a new column
    for x in range(len(df)):
        df['destCode'].iloc[x]=df['callIdentifier'].iloc[x][len(df['clusterCode'].iloc[x])+len(df['callYear'].iloc[x])+2:]

    # Use a dictionary to get cluster names from cluster codes
    clusterDict = {
        'HORIZON-HLTH': '1. Health',
        'HORIZON-CL2': '2. Culture, Creativity and Inclusive Society',
        'HORIZON-CL3': '3. Civil Security for Society',
        'HORIZON-CL4': '4. Digital, Industry and Space',
        'HORIZON-CL5': '5. Climate, Energy and Mobility',
        'HORIZON-CL6': '6. Food, Bioeconomy, Natural Resources, Agriculture and Environment',
        'HORIZON-MISS': 'Missions',
        'HORIZON-ER-JU':'Europe\'s Rail Joint Undertaking',
        'HORIZON-JU-ER':'Europe\'s Rail Joint Undertaking',
        'HORIZON-EUROHPC-JU':'European High Performance Computing Joint Undertaking',
        'HORIZON-EUSPA':'European Union Agency for the Space Programme',
        'HORIZON-JTI-CLEANH2':"Clean Hydrogen Joint Technology Initiatives",
        'HORIZON-JU-CBE':"Circular Bio-based Europe Joint Undertaking",
        'HORIZON-JU-Clean-Aviation':"Clean Aviation Joint Undertaking",
        'HORIZON-JU-GH-EDCTP3':"Global Health - European and Developing Countries Clinical Trials Partnership Joint Undertaking",
        'HORIZON-JU-IHI':"Innovative Health Initiative Joint Undertaking",
        'HORIZON-JU-SNS':"Smart Networks and Services Joint Undertaking",
        'HORIZON-KDT-JU':"Key Digital Technologies Joint Undertaking",
        'HORIZON-SESAR':"Single European Sky ATM (Air Traffic Management) Research Joint Undertaking",
        'HORIZON-JU-Chips':"Chips Joint Undertaking"}

    df['clusterName'] = df['clusterCode'].map(clusterDict)
    # set callYear to integer type (not text)
    df['callYear']=df['callYear'].astype(int)

    # Rename columns
    df = df.rename(columns={'callIdentifier': 'callId', 'identifier': 'topicId','title':'topicTitle', 'status.abbreviation': 'status','sumbissionProcedure.abbreviation':"process",'topicActions.abbreviation':'actionType'})

    # Select the columns we want to keep
    cols = ['ccm2Id','callYear','clusterCode','clusterName','destination','destCode','pubDate','openDate','closeDate','s2Date','status','process','actionType','callId','callTitle','topicId','topicTitle']
    df = df[cols]

    # Create hyperlinks for topic IDs
    base_url = "https://ec.europa.eu/info/funding-tenders/opportunities/portal/screen/opportunities/topic-details/"
    df['topicId'] = df['topicId'].apply(lambda x: f'=HYPERLINK("{base_url}{x.lower()}", "{x}")')

    return df

def write_to_excel(df, file_path):
    # Export the DataFrame to Excel
    ## output using openpyxl engine so we can format the excel file
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        # Set autofilter on all columns
        num_rows = len(df)
        num_cols = len(df.columns)
        worksheet.autofilter(0, 0, num_rows, num_cols - 1)
        # Set column widths
        worksheet.set_column('C:C', width=11)
        worksheet.set_column('D:F', width=18.5)
        worksheet.set_column('G:J', width=10)
        worksheet.set_column('J:J', width=11)
        worksheet.set_column('L:L', width=11)
        worksheet.set_column('M:M', width=13)
        worksheet.set_column('N:N', width=42)
        worksheet.set_column('O:O', width=27)
        worksheet.set_column('Q:Q', width=60)
        # Define a format for hyperlinks (blue and underlined) and apply to link column
        hyperlink_format = workbook.add_format({'font_color': 'blue', 'underline': True})
        worksheet.set_column(15,15,40, hyperlink_format)
        # Save the Excel file
        writer.close()
    print("Saved",file_path)

def get_last_file():
    path=os.getcwd()
    excel_files = glob.glob(os.path.join(path, '*.xlsx'))
    excel_files = [file for file in excel_files if 'new' not in os.path.basename(file) and not os.path.basename(file).startswith("~$") and os.path.basename(file).startswith("HE_calls")]
    if not excel_files:
        return False
    excel_files.sort(key=os.path.getmtime)
    last_file = os.path.basename(excel_files[-1])
    return last_file
    
def compare_calls(df, compare_file):
    df_prev = pd.read_excel(compare_file, engine='openpyxl')
    new_calls = df[~df['ccm2Id'].isin(df_prev['ccm2Id'])]
    return new_calls

def prep_calendar(df, call_year):
    df = df.drop_duplicates(subset=['callId'])

    # Convert openDate and closeDate to datetime objects
    df['openDate'] = pd.to_datetime(df['openDate'])
    df['closeDate'] = pd.to_datetime(df['closeDate'])
    df['s2Date'] = pd.to_datetime(df['s2Date'])

    # Filter by callYear
    df = df[(df['callYear'] == call_year) | (df['closeDate'].dt.year == call_year)] # now getting all calls closing in the given year

    # Reverse the order of rows
    df = df.iloc[::-1]

    # Use a dictionary to get cluster names from cluster codes
    clusterDict = {
        'HORIZON-HLTH': '1. Health',
        'HORIZON-CL2': '2. Culture, Creativity and Inclusive Society',
        'HORIZON-CL3': '3. Civil Security for Society',
        'HORIZON-CL4': '4. Digital, Industry and Space',
        'HORIZON-CL5': '5. Climate, Energy and Mobility',
        'HORIZON-CL6': '6. Food, Bioeconomy, Natural Resources, Agriculture and Environment',
        'HORIZON-MISS': 'Missions',
        'HORIZON-ER-JU':'5. Climate, Energy and Mobility',
        'HORIZON-JU-ER':'5. Climate, Energy and Mobility',
        'HORIZON-EUROHPC-JU':'4. Digital, Industry and Space',
        'HORIZON-EUSPA':'4. Digital, Industry and Space',
        'HORIZON-JTI-CLEANH2':"5. Climate, Energy and Mobility",
        'HORIZON-JU-CBE':"6. Food, Bioeconomy, Natural Resources, Agriculture and Environment",
        'HORIZON-JU-Clean-Aviation':"5. Climate, Energy and Mobility",
        'HORIZON-JU-GH-EDCTP3':"1. Health",
        'HORIZON-JU-IHI':"1. Health",
        'HORIZON-JU-SNS':"4. Digital, Industry and Space",
        'HORIZON-KDT-JU':"4. Digital, Industry and Space",
        'HORIZON-SESAR':"5. Climate, Energy and Mobility",
        'HORIZON-JU-Chips':"4. Digital, Industry and Space"}

    df['clusterName2'] = df['clusterCode'].map(clusterDict)

    # Sort the DataFrame
    df = df.sort_values(by=['clusterName2', 'callId'])
    df = df.iloc[::-1]

    return df

def save_call_calendar(df, call_year):
    # Define colormap based on cluster
    cluster_colors = {
        '1. Health': '#ADD8E6',  # lightskyblue
        'Global Health - European and Developing Countries Clinical Trials Partnership Joint Undertaking': '#87CEFA',  # Sky Bluelightblue
        'Innovative Health Initiative Joint Undertaking': '#87CEEB',  # Baby Blue
        '2. Culture, Creativity and Inclusive Society': '#F4A460',  # sandybrown
        '3. Civil Security for Society': '#F08080',  # lightcoral
        '4. Digital, Industry and Space': '#F6851F',  # orange
        'European High Performance Computing Joint Undertaking': "#FAA41A",
        'European Union Agency for the Space Programme': "#F8BF1C",
        'Smart Networks and Services Joint Undertaking': "#DFC423",
        'Key Digital Technologies Joint Undertaking': "#F7E542",
        'Chips Joint Undertaking': "#F9EC00",
        '5. Climate, Energy and Mobility': '#D8BFD8', 
        'Clean Hydrogen Joint Technology Initiatives': '#C591C1',
        'Clean Aviation Joint Undertaking': '#B972AF',
        'Europe\'s Rail Joint Undertaking': '#A081B5',
        'Single European Sky ATM (Air Traffic Management) Research Joint Undertaking': '#C94E9C', 
        '6. Food, Bioeconomy, Natural Resources, Agriculture and Environment': '#9ACD32',  # yellowgreen
        'Circular Bio-based Europe Joint Undertaking': '#7FFF00',  # Lime Green
        'Missions': '#FFC0CB'  # lightpink
    }

    # Plot the Gantt chart with larger figure size and smaller font size
    fig, ax = plt.subplots(figsize=(12, 15))  # Larger figure size
    plt.rc('xtick', labelsize=10)  # Adjust font size for x-axis labels
    plt.rc('ytick', labelsize=10)  # Adjust font size for y-axis labels
    # Define colors for alternating shading
    shading_colors = ['whitesmoke', 'white']
    # Get min and max dates from data
    min_date = df['openDate'].min()
    max_date = df['closeDate'].max()
    # Plot alternating shading for months
    month_starts = pd.date_range(start=min_date, end=max_date, freq='MS')
    for i, month_start in enumerate(month_starts):
        if i % 2 == 0:
            shading_color = shading_colors[0]
        else:
            shading_color = shading_colors[1]
        month_end = month_start + pd.DateOffset(months=1)
        ax.axvspan(month_start, month_end, facecolor=shading_color, alpha=0.5, zorder=-1)
    # Plot the bars
    bar_height = 0.8  # Adjust the height of the bars
    bar_spacing = 0.2  # Adjust the spacing between bars
    # Add data to plot
    for i, row in df.iterrows():
        # color = cluster_colors.get(row['clusterName2'], 'skyblue')  # Default color for unknown clusters
        color = cluster_colors.get(row['clusterName'], 'skyblue')  # Default color for unknown clusters
        ax.barh(y=row['callId'], width=row['closeDate'] - row['openDate'], left=row['openDate'], height=bar_height, color=color, zorder=0)
        ax.text(row['openDate'] + (row['closeDate'] - row['openDate']) / 2, row['callId'], str(row['callId']), ha='center', va='center', color='black', fontsize=8)
        # Add marker for second stage closing date if it exists
        if not pd.isnull(row['s2Date']):
            ax.plot(row['s2Date'], row['callId'], 'o', color='black', markersize=5)
    # Set y-axis ticks and labels
    y_labels = df['clusterCode']
    # Set y-axis ticks and labels
    ax.set_yticks(range(len(y_labels)))
    ax.set_yticklabels(y_labels)
    # Format the dates on the x-axis
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%Y'))
    # Add grid lines
    ax.grid(axis='x',which='both', linestyle='--', linewidth=0.5, zorder=3, alpha=0.3)
    # Rotate the x-axis labels
    plt.xticks(rotation=90)
    # Set the title
    plt.title('Horizon Europe calls for callYear ' + str(call_year), weight="bold")
    # # Add datestamp
    date_stamp = 'Generated: ' + datetime.datetime.now().strftime("%Y-%m-%d")
    # Create an annotation object with the date and desired position
    annotation = ax.annotate(
        date_stamp,
        xy=(1.02, -0.08),  # Place the text at the right edge, slightly above the center
        xycoords="axes fraction",
        horizontalalignment="right",
        verticalalignment="center",
    )
    # Add vertical line for year dividers
    for year in range(min_date.year, max_date.year + 1):
        january_first = datetime.datetime(year, 1, 1) # Create a datetime object for January 1st of that year
        if min_date <= january_first <= max_date: # Check if January 1st falls within the range of dates in the plot
            # Add a vertical line at January 1st
            ax.axvline(x=january_first, color='black', linestyle='--', linewidth=0.5, alpha=0.5, zorder=-4)
    # Set the annotation text properties
    annotation.set_fontsize(8)
    annotation.set_color("gray")
    # Show the plot
    plt.tight_layout()
    savename="HE_calendar-" + str(call_year) + "_calls-" + datetime.datetime.now().strftime("%Y%m%d")+ ".pdf"
    plt.savefig(savename)
    return savename



def main():
    args = parse_arguments()
    filename="grantsTenders.json"
    if args.local:
        print("Local mode: trying using stored grantsTenders.json file")
        if os.path.isfile("grantsTenders.json"):
            with open("grantsTenders.json", 'r') as f:
                grantsdata = json.load(f)
        else:
            print(">grantsTenders.json not found; aborting.")
            sys.exit(0)
    else:
        print("Getting updates from European Commission server...")
        url = 'https://ec.europa.eu/info/funding-tenders/opportunities/data/referenceData/grantsTenders.json'
        grantsdata = download_json_with_progress(url)
        # Save a copy of grantsTenders.json to disk
        with open("grantsTenders.json", 'w') as f:
            # Dump the JSON data to the file with indentation for readability (optional)
            json.dump(grantsdata, f, indent=4)

    df = process_data(grantsdata)
    
    if args.file: # check for compare file as argument
        arg_file = args.file
        if arg_file.endswith('.xlsx') and os.path.isfile(arg_file):
            # print("Compare file:", arg_file)
            compare_file = arg_file
        else:
            print("Invalid file specified; skipping comparison")
            compare_file = False
    else: # if no argument given, look for most recent file
        if get_last_file():
            compare_file = get_last_file()
        else: # if no argument or recent files, just save the new file and exit
            print("No previous HE_calls files to compare to")
            compare_file = False

    out_file="HE_calls_"+str(datetime.date.today())+".xlsx"

    new_data = False 
    if compare_file:
        # if there is a compare file, look for new calls and only write files if there are some
        new_calls = compare_calls(df, compare_file)

        if(len(new_calls)>0):
            print("Found {} changes compared to {}".format(str(len(new_calls)), compare_file))
            new_data = True
        else:
            print("No new calls in latest file.")

    if new_data or not args.newonly:
        write_to_excel(df, out_file)

        # option to also save calendars
        if args.calendars:
            for year in df['callYear'].unique():
                cal_data = prep_calendar(df, year)
                cal_file = save_call_calendar(cal_data, year)
                print("Saved calendar from {} for {} as {}".format(out_file, year, cal_file))

    print("Done.")


if __name__ == "__main__":
    main()
    
