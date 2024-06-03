"""
Script Name: HE_CORDIS_updates.py
Description: This script processes data from the CORDIS database, checks for updates, and produces Excel summary sheets for given countries.
Author: John Creech
Date: May 2023

Requirements:
- Python 3.6+
- requests
- pandas
- openpyxl
- pycountry
- tqdm
- argparse

Usage:
python script.py [options] [country ...]

Options:
-n, --newonly  Save outputs only if new data were found in CORDIS.
-l, --local    Skip CORDIS check and use previously saved local data.
-f, --force    Force download of data from CORDIS even if not newer than local copy.
country        List of two-letter country codes or sets of countries for summary (default: NZ).

Examples:
    python HE_CORDIS_updates.py -n NZ pacific

Available predefined country sets:
    pacific
    eu_members
    associated_countries
    nordics
"""

import requests
import datetime
import pandas as pd
import os
import sys
import openpyxl
import warnings
warnings.filterwarnings('ignore')
import pycountry
import zipfile
import xml.etree.ElementTree as ET
from tqdm import tqdm
import argparse
# pycountry fix to add Kosovo
pycountry.countries.add_entry(alpha_2="XK", alpha_3="XXK", name="Kosovo", numeric="926")

def parse_arguments():
    parser = argparse.ArgumentParser(description="Check for updated data in CORDIS and produce summary sheets",
        epilog="preset country groupings: pacific, eu_members, associated_countries")
    parser.add_argument("-n", "--newonly", help="process data only if we got new data", action="store_true")
    parser.add_argument("-l", "--local", help="skip CORDIS check use local data", action="store_true")
    parser.add_argument("-f", "--force", help="force download of data from CORDIS even if not newer", action="store_true")
    parser.add_argument("country", nargs='*', default=["NZ"], help="country (or set of countries) for summary (default: NZ)")
    return parser.parse_args()

def checkCordisDate(rss_url):
    ### Check the publication date of data in CORDIS.
    response = requests.get(rss_url)
    rss_data = response.text
    root = ET.fromstring(rss_data)
    for item in root.findall(".//item"):
        pub_date_element = item.find("pubDate")
        pub_date_str = pub_date_element.text
        pub_date = datetime.datetime.strptime(pub_date_str, "%a, %d %b %Y %H:%M:%S %z")
    return(pub_date)

def checkLocalDataDate(path):
    # Grab previous data date from cordis_date.txt
    if os.path.exists(path+"/cordis_date.txt"):
        with open(path+"/cordis_date.txt", "r") as file:
            fp_date = datetime.datetime.strptime(file.read().strip(), "%Y-%m-%d %H:%M:%S %z")
            return fp_date
    else:
        return "none"

def updateCordisData(url,path):
    filename=path+".zip"
    # Fetch and unzip cordis data
    download_with_progress(url, filename)
    extract_without_paths(filename, path)
    return True


def download_with_progress(url, filename):
    # Download a file from a URL with progress bar.
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))
    progress_bar = tqdm(total=total_size, unit='B', unit_scale=True) ## standalone script version
    # progress_bar = tqdm_notebook(total=total_size, unit='B', unit_scale=True) ## jupyter version
    with open(filename, 'wb') as file:
        for chunk in response.iter_content(chunk_size=1024):
            file.write(chunk)
            progress_bar.update(len(chunk))
    progress_bar.close()

def extract_without_paths(zipfile_path, extract_path):
    # Extract contents of a zip file, junking paths
    try:
        if not os.path.exists(zipfile_path):
            raise FileNotFoundError(f"The zip file '{zipfile_path}' does not exist.")
        with zipfile.ZipFile(zipfile_path, "r") as zip_ref:
            for file_info in zip_ref.infolist():
                if file_info.is_dir():
                    continue
                else:
                    file_info.filename = os.path.basename(file_info.filename)
                    zip_ref.extract(file_info, extract_path)
        print(f"Extracted {zipfile_path} to {extract_path}")
    except Exception as e:
        print(f"Error occurred: {e}")

def get_country_name(country_code):
    if country_code == 'UK': # the ISO 3166-1 alpha-2 code for the United Kingdom is GB but the data uses UK
        country_code = 'GB'

    try:
        country_name = pycountry.countries.get(alpha_2=country_code.upper()).name
        return country_name
    except AttributeError:
        return ""

def processCordisData():
    print("Processing CORDIS data (this can take some time, please be patient)")
    ## Read in data files and combine
    # Org data
    df_fp7 = pd.read_excel("cordis-fp7projects-xlsx/organization.xlsx")#, index_col="projectID")
    df_H2020 = pd.read_excel("cordis-h2020projects-xlsx/organization.xlsx")#, index_col="projectID")
    df_HEU = pd.read_excel("cordis-HORIZONprojects-xlsx/organization.xlsx")#, index_col="projectID")
    # Project data
    df_fp7_project = pd.read_excel("cordis-fp7projects-xlsx/project.xlsx")#, index_col='id')
    df_H2020_project = pd.read_excel("cordis-h2020projects-xlsx/project.xlsx")#, index_col='id')
    df_HEU_project = pd.read_excel("cordis-HORIZONprojects-xlsx/project.xlsx")#, index_col='id')
    # Add a column indicating the framework program
    df_fp7['frameworkProgramme'] = 'FP7'
    df_H2020['frameworkProgramme'] = 'H2020'
    df_HEU['frameworkProgramme'] = 'HORIZON'
    # Combine the orgs DataFrames
    df_orgs = pd.concat([df_fp7, df_H2020, df_HEU])
    # Fix some number formats
    df_orgs['ecContribution'] = pd.to_numeric(df_orgs['ecContribution'], errors='coerce')  # Convert 'ecContribution' to numeric
    df_orgs['netEcContribution'] = pd.to_numeric(df_orgs['netEcContribution'], errors='coerce')  # Convert 'ecContribution' to numeric
    df_orgs['totalCost'] = pd.to_numeric(df_orgs['totalCost'], errors='coerce')  # Convert 'ecContribution' to numeric
    # Combine the projects data into a new dataframe
    df_projects = pd.concat([df_fp7_project, df_H2020_project, df_HEU_project])
    df_projects = df_projects.rename(columns={'id': 'projectID'}) # called id in the projects file, but projectID elsewhere, so renaming

    # pickle them so we can reuse them if no new data
    df_projects.to_pickle("cordis_projects.pkl")
    df_orgs.to_pickle("cordis_orgs.pkl")
    return df_projects, df_orgs

## Function to produce summary spreadsheets
def do_cordis_summary(df_projects, df_orgs, country, filename=None): # country/countries as list, e.g., ['NZ'] or ['NZ' ,'AU']
    print("Preparing summary for country list: {}".format(country))
    # Produce new dataframe with projects with partners from selected country/countries, and get titles and other info from df_projects
    df_country = df_orgs[df_orgs['country'].isin(country)]
    df_country = pd.merge(df_country, df_projects[['projectID', 'title', 'fundingScheme', 'subCall', 'ecSignatureDate', 'startDate', 'endDate']], on='projectID', how='left')
    df_country = df_country.rename(columns={'organisationID': 'PIC', 'fundingScheme':'Type of action', 'name':'Organisation'})
    df_country['countryName'] = df_country['country'].map(get_country_name)
    df_country['PIC'] = df_country['PIC'].fillna(-1).astype(int)

    # pull out the columns I want
    df_country=df_country[['frameworkProgramme', 'projectID', 'projectAcronym', 'title', 'ecSignatureDate', 'startDate', 'endDate',
    'country', 'countryName', 'PIC', 'Organisation', 'shortName', 'activityType', 'SME', 'Type of action', 'subCall', 'order', 
    'role', 'ecContribution', 'netEcContribution', 'totalCost']]
    
    ### Produce the summary by organisation
    # Get a list of Framework Programmes each org has participated in
    grouped_orgs = df_country.groupby('PIC')['frameworkProgramme'].agg(lambda x: ', '.join(set(x.dropna()))).reset_index()
    # Merge the aggregated data with the original dataframe
    df_country_orgs = df_country.drop_duplicates('PIC').merge(grouped_orgs, on='PIC', how='left')
    
    # Get the total funding for each org
    grouped_sum_ec = df_country.groupby('PIC')['ecContribution'].sum().reset_index()
    # Merge the aggregated sum with the merged dataframe
    df_country_orgs = df_country_orgs.merge(grouped_sum_ec, on='PIC', how='left')
    
    # Count the projects for each org
    project_count = df_country.groupby('PIC')['projectID'].nunique().reset_index()
    project_count.columns = ['PIC', 'projectCount']
    # Merge the aggregated data with the original dataframe and project count
    df_country_orgs = df_country_orgs.merge(project_count, on='PIC', how='left')
    
    # tidy up
    df_country_orgs = df_country_orgs.rename(columns={'ecContribution_y':'Total EU funding (€)', 'frameworkProgramme_y':'Framework programmes', 'name':'Organisation'})
    df_country_orgs['PIC'] = df_country_orgs['PIC'].astype(int)
    
    # pull columns to keep
    df_country_orgs = df_country_orgs[['PIC', 'Organisation', 'shortName', 'country', 'countryName', 'activityType', 'SME', 'projectCount', 'Framework programmes', 'Total EU funding (€)']]

    # ## Get the summary stats (count of funded projects, and sum of ecContribution) for orgs across each Framework Programme
    # # # Group by 'PIC' and 'frameworkProgramme' and aggregate the count of projects,
    grouped_orgs_fp = df_country.groupby(['PIC', 'frameworkProgramme']).agg(
        num_projects=('projectID', 'nunique'),
        num_funded_projects=('ecContribution', lambda x: (x > 0).sum()),
        total_ecContribution=('ecContribution', 'sum')
        ).reset_index()
    
    # Pivot the table to have 'frameworkProgramme' as columns
    pivot_table = grouped_orgs_fp.pivot_table(
        index='PIC',
        columns='frameworkProgramme',
        values=['num_projects', 'num_funded_projects', 'total_ecContribution'],
        fill_value=0
        )
    
    # Flatten MultiIndex columns
    pivot_table.columns = ['_'.join(col).strip() for col in pivot_table.columns.values]
    # Reset index to make 'PIC' a regular column
    pivot_table.reset_index(inplace=True)
    
    # Merge the pivot_table with the summary DataFrame on 'PIC'
    df_country_orgs = df_country_orgs.merge(pivot_table, on='PIC', how='left')
    df_country_orgs=df_country_orgs[['PIC', 'Organisation', 'shortName', 'country', 'countryName', 'activityType', 'SME',
    'projectCount', 'Framework programmes', 'Total EU funding (€)',
    'num_projects_FP7', 'num_funded_projects_FP7', 'total_ecContribution_FP7',
    'num_projects_H2020', 'num_funded_projects_H2020', 'total_ecContribution_H2020',
    'num_projects_HORIZON', 'num_funded_projects_HORIZON', 'total_ecContribution_HORIZON']]
    
    # rename some columns to make easier to read
    df_country_orgs = df_country_orgs.rename(columns={'num_projects_FP7':'FP7 projects', 'num_funded_projects_FP7':'FP7 funded projects', 'total_ecContribution_FP7':'FP7 total funding (€)',
        'num_projects_H2020':'H2020 projects', 'num_funded_projects_H2020':'H2020 funded projects', 'total_ecContribution_H2020':'H2020 total funding (€)',
        'num_projects_HORIZON':'HEU projects', 'num_funded_projects_HORIZON':'HEU funded projects', 'total_ecContribution_HORIZON':'HEU total funding (€)'})

    ### Produce summary of projects
    # # Get searched country involvement
    df_country_byproject = df_country.groupby('projectAcronym')['country'].unique()
    # # Merge the aggregated data with the original dataframe
    df_country_projects = df_country.drop_duplicates('projectAcronym').merge(df_country_byproject, on='projectAcronym', how='left')
    df_country_projects = df_country_projects.rename(columns={'country_y': 'Matched countries'})
    
    # Get all countries involvement
    df_country_byproject = df_orgs.groupby('projectAcronym')['country'].unique()
    # # Merge the aggregated data with the original dataframe
    df_country_projects = df_country_projects.drop_duplicates('projectAcronym').merge(df_country_byproject, on='projectAcronym', how='left')
    df_country_projects = df_country_projects.rename(columns={'country': 'All countries'})

    # return df_country_projects
    # # Get sum of ecContribution for each project (for selected countries) and merge back in
    grouped_countries_x = df_country.groupby(['projectID']).agg(
        country_ecContribution=('ecContribution', lambda x: x[x > 0].sum())  # Sum only non-zero values of ecContribution
        ).reset_index()
    df_country_projects = df_country_projects.merge(grouped_countries_x, on='projectID', how='left')
    # # Get sum of ecContribution for each project (for all countries) and merge back in
    grouped_countries_x = df_orgs.groupby(['projectID']).agg(
        total_ecContribution=('ecContribution', lambda x: x[x > 0].sum())  # Sum only non-zero values of ecContribution
        ).reset_index()
    df_country_projects = df_country_projects.merge(grouped_countries_x, on='projectID', how='left')
    # select columns for output
    df_country_projects=df_country_projects[['frameworkProgramme', 'projectID', 'projectAcronym', 'title',
    'ecSignatureDate', 'startDate', 'endDate', 'Type of action', 'subCall',
    'country_ecContribution', 'total_ecContribution', 'Matched countries', 'All countries']]
    
    ### Produce a summary by country
    # Group by country and aggregate the unique values of project acronym
    df_bycountry = df_country.groupby('country')['projectAcronym'].unique()
    
    # Merge the aggregated data with the original dataframe
    df_countries = df_country.drop_duplicates('country').merge(df_bycountry, on='country', how='left')
    df_countries = df_countries.rename(columns={'projectAcronym_y': 'Project acronyms'})
    
    # Group by 'country' and 'frameworkProgramme' and aggregate the count of projects,
    # count of funded projects, and sum of ecContribution
    grouped_countries_fp = df_country.groupby(['country', 'frameworkProgramme']).agg(
        num_projects=('projectID', 'nunique'),
        num_funded_projects=('ecContribution', lambda x: (x > 0).sum()),
        total_ecContribution=('ecContribution', 'sum')
        # num_funded_projects=('projectID', lambda x: (x[x['ecContribution'] > 0]).nunique()),  # Count projects with ecContribution > 0
        # total_ecContribution=('ecContribution', lambda x: x[x > 0].sum())  # Sum only non-zero values of ecContribution
        ).reset_index()
    
    # Pivot the table to have 'frameworkProgramme' as columns
    pivot_table_countries = grouped_countries_fp.pivot_table(
        index='country',
        columns='frameworkProgramme',
        values=['num_projects', 'num_funded_projects', 'total_ecContribution'],
        fill_value=0
        )
    
    # # Flatten MultiIndex columns
    pivot_table_countries.columns = ['_'.join(col).strip() for col in pivot_table_countries.columns.values]
    
    # Reset index to make 'country' a regular column
    pivot_table_countries.reset_index(inplace=True)
    
    # Merge the pivot table with the original summary DataFrame on 'country'
    df_countries_summary = df_countries.merge(pivot_table_countries, on='country', how='left')
    
    # Print the resulting DataFrame
    df_countries_summary=df_countries_summary[['country','countryName', 'Project acronyms',
    'num_projects_FP7', 'num_funded_projects_FP7', 'total_ecContribution_FP7',
    'num_projects_H2020', 'num_funded_projects_H2020', 'total_ecContribution_H2020',
    'num_projects_HORIZON', 'num_funded_projects_HORIZON', 'total_ecContribution_HORIZON']]
    
    df_countries_summary = df_countries_summary.rename(columns={'num_projects_FP7':'FP7 projects', 'num_funded_projects_FP7':'FP7 funded projects', 'total_ecContribution_FP7':'FP7 funding (€)',
        'num_projects_H2020':'H2020 projects', 'num_funded_projects_H2020':'H2020 funded projects', 'total_ecContribution_H2020':'H2020 funding (€)',
        'num_projects_HORIZON':'HEU projects', 'num_funded_projects_HORIZON':'HEU funded projects', 'total_ecContribution_HORIZON':'HEU funding (€)'})
    
    ## Add country stats from across all FPs
    # count of funded projects, and sum of ecContribution
    grouped_countries_overall = df_country.groupby(['country']).agg(
        total_projects=('projectID', 'nunique'),
        total_funded_projects=('ecContribution', lambda x: (x > 0).sum()),
        total_funding=('ecContribution', 'sum')
        ).reset_index()
    df_countries_summary = df_countries_summary.merge(grouped_countries_overall, on='country', how='left')
    df_countries_summary = df_countries_summary.rename(columns={'total_projects':'Total projects', 'total_funded_projects':'Total funded projects', 'total_funding':'Total funding (€)'})
    
    ### Generate URLs to include links to org profiles and projects in output spreadsheets
    # Org profiles on FTP
    base_url = "https://ec.europa.eu/info/funding-tenders/opportunities/portal/screen/how-to-participate/org-details/"
    df_country['PIC'] = df_country['PIC'].apply(lambda x: f'=HYPERLINK("{base_url}{x}", "{x}")')
    # project details on CORDIS
    base_url = "https://cordis.europa.eu/project/id/"
    df_country['projectID'] = df_country['projectID'].apply(lambda x: f'=HYPERLINK("{base_url}{x}", "{x}")')
    df_country_projects['projectID'] = df_country_projects['projectID'].apply(lambda x: f'=HYPERLINK("{base_url}{x}", "{x}")')
    
    ### save the tables to excel sheets
    file_path='Country_summary.xlsx'
    if filename:
        file_path=filename
        
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            ## save the total participation sheet
            df_country.to_excel(writer, index=False, sheet_name='FP_participation')
            # Access the workbook
            workbook = writer.book
            worksheet = writer.sheets['FP_participation']
            # Set autofilter on all columns
            num_rows = len(df_country)
            num_cols = len(df_country.columns)
            worksheet.autofilter(0, 0, num_rows, num_cols - 1)
            
            # Define a format for hyperlinks (blue and underlined) and apply to link column
            hyperlink_format = workbook.add_format({'font_color': 'blue', 'underline': True})
            # Define a cell format for left-aligned text
            left_aligned_format = workbook.add_format({'align': 'left', 'bold':True})
            header_format = workbook.add_format({'align': 'left', 'bold':True,'text_wrap': True})
            
            # Set column widths
            # worksheet.set_column('A:A', width=10)
            worksheet.set_column('C:C', width=12)
            worksheet.set_column('D:D', width=30)
            worksheet.set_column('E:G', width=10)
            worksheet.set_column('K:K', width=35)
            worksheet.set_column('L:L', width=10)
            worksheet.set_column('M:N', width=6)
            worksheet.set_column('O:P', width=10)
            worksheet.set_column('Q:Q', width=6)
            worksheet.set_column('R:R', width=10)
            worksheet.set_column(1,1,11, hyperlink_format) # projectID
            worksheet.set_column(9,9,10, hyperlink_format) # PIC  
            worksheet.set_column('S:U', None, workbook.add_format({'num_format': '#,##0'}))
            
            # Apply the left-aligned format to the header row
            for col_num, value in enumerate(df_country.columns.values):
                worksheet.write(0, col_num, value, header_format)
                
            ## ORG SUMMARY
            df_country_orgs.to_excel(writer, index=False, sheet_name='Orgs_summary')
            # Access the workbook
            workbook = writer.book
            worksheet = writer.sheets['Orgs_summary']
            # Define a format for hyperlinks (blue and underlined) and apply to link column
            hyperlink_format = workbook.add_format({'font_color': 'blue', 'underline': True})
            # Define a cell format for left-aligned text
            left_aligned_format = workbook.add_format({'align': 'left', 'bold':True})
            cell_format = workbook.add_format({'text_wrap': True})
            header_format = workbook.add_format({'align': 'left', 'bold':True,'text_wrap': True})
            
            # Set autofilter on all columns
            num_rows = len(df_country_orgs)
            num_cols = len(df_country_orgs.columns)
            worksheet.autofilter(0, 0, num_rows, num_cols - 1)
            # Set column widths
            worksheet.set_column('B:B', width=55)
            worksheet.set_column('C:C', width=15)
            worksheet.set_column('D:H', width=6)
            # worksheet.set_column('H:H', width=14)
            worksheet.set_column('I:I', width=11)
            worksheet.set_column(0,0,10, hyperlink_format)
            # Apply the left-aligned format to the header row
            for col_num, value in enumerate(df_country_orgs.columns.values):
                worksheet.write(0, col_num, value, header_format)
            # Set the number format for specific columns
            worksheet.set_column('J:J', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('M:M', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('P:P', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('S:S', 12, workbook.add_format({'num_format': '#,##0'}))

            ## save the project summary
            df_country_projects.to_excel(writer, index=False, sheet_name='FP_projects')
            # Access the workbook
            workbook = writer.book
            worksheet = writer.sheets['FP_projects']
            # Set autofilter on all columns
            num_rows = len(df_country_projects)
            num_cols = len(df_country_projects.columns)
            worksheet.autofilter(0, 0, num_rows, num_cols - 1)
            # Set column widths
            worksheet.set_column('C:C', width=12)
            worksheet.set_column('D:D', width=25)
            worksheet.set_column('E:M', width=11)
            # Define a format for hyperlinks (blue and underlined) and apply to link column
            worksheet.set_column(1,1,10, hyperlink_format)
            worksheet.set_column('J:J', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('K:K', 12, workbook.add_format({'num_format': '#,##0'}))
            # Apply the left-aligned format to the header row
            for col_num, value in enumerate(df_country_projects.columns.values):
                worksheet.write(0, col_num, value, header_format)

            ## save the country summary
            df_countries_summary.to_excel(writer, index=False, sheet_name='Countries_summary')
            # Access the workbook
            workbook = writer.book
            worksheet = writer.sheets['Countries_summary']
            # Set autofilter on all columns
            num_rows = len(df_countries_summary)
            num_cols = len(df_countries_summary.columns)
            worksheet.autofilter(0, 0, num_rows, num_cols - 1)
            # Set column widths
            worksheet.set_column('A:A', width=7)
            worksheet.set_column('B:B', width=14)
            worksheet.set_column('C:C', width=54)
            worksheet.set_column('D:E', width=7)
            worksheet.set_column('G:H', width=7)
            worksheet.set_column('J:K', width=7)
            worksheet.set_column('M:N', width=7)
            worksheet.set_column('F:F', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('I:I', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('L:L', 12, workbook.add_format({'num_format': '#,##0'}))
            worksheet.set_column('O:O', 12, workbook.add_format({'num_format': '#,##0'}))
            # Apply the left-aligned format to the header row
            for col_num, value in enumerate(df_countries_summary.columns.values):
                worksheet.write(0, col_num, value, header_format)
                
            writer.close()

    print(f"Saved {file_path}")

    
def main():
    args = parse_arguments()

    ## Setup things
    country_sets = {
    "pacific": ["FJ", "KI", "MH", "FM", "NR", "PW", "PG", "WS", "SB", "TO", "TV", "VU"],
    "eu_members": ['AT', 'BE', 'BG', 'CY', 'CZ', 'DE', 'DK', 'EE', 'ES', 'FI', 'FR', 'GR', 'HR', 'HU', 'IE', 'IT', 'LT', 'LU', 'LV', 'MT', 'NL', 'PL', 'PT', 'RO', 'SE', 'SI', 'SK'],
    "associated_countries": ['AL', 'AM', 'BA', 'FO', 'GE', 'IS', 'IL', 'XK', 'MD', 'ME', 'NZ', 'MK', 'NO', 'RS', 'TN', 'TR', 'UA', 'UK'],
    "nordics": ["NO", "SE", "DK", "FI", "IS"]
    }

    fp_data = {
      "FP7": {
        "rss_url": "https://data.europa.eu/api/hub/search/en/feeds/datasets/cordisfp7projects.rss",
        "download_url": "https://cordis.europa.eu/data/cordis-fp7projects-xlsx.zip",
        "path": "cordis-fp7projects-xlsx",
        "cordis_date":"",
        "stored_date":""
      },
      "H2020": {
        "rss_url": "https://data.europa.eu/api/hub/search/en/feeds/datasets/cordish2020projects.rss",
        "download_url": "https://cordis.europa.eu/data/cordis-h2020projects-xlsx.zip",
        "path": "cordis-h2020projects-xlsx",
        "cordis_date":"",
        "stored_date":""
      },
      "HORIZON": {
        "rss_url": "https://data.europa.eu/api/hub/search/en/feeds/datasets/cordis-eu-research-projects-under-horizon-europe-2021-2027.rss",
        "download_url": "https://cordis.europa.eu/data/cordis-HORIZONprojects-xlsx.zip",
        "path": "cordis-HORIZONprojects-xlsx",
        "cordis_date":"",
        "stored_date":""
      }
    }

    newdata = False ## flag to set if we download new data
    ## Data checking
    if args.local:
        print("Skipped checking for new CORDIS data")
    else:
        print("Checking for CORDIS data...")
        for fp in fp_data:
            fp_data[fp]["cordis_date"] = checkCordisDate(fp_data[fp]["rss_url"]) # check date of data in CORDIS
            fp_data[fp]["stored_date"] = checkLocalDataDate(fp_data[fp]["path"]) # check date of stored data
            print(" ",fp,"   \tCORDIS:",fp_data[fp]["cordis_date"],"\tLOCAL:",fp_data[fp]["stored_date"])

        for fp in fp_data:
            if fp_data[fp]["stored_date"] == "none" or fp_data[fp]["cordis_date"] > fp_data[fp]["stored_date"] or args.force:
                print("Fetching online data for",fp)
                if(updateCordisData(fp_data[fp]["download_url"],fp_data[fp]["path"])):
                    newdata = True
                    with open(fp_data[fp]["path"]+"/cordis_date.txt", "w") as file:
                        file.write(fp_data[fp]["cordis_date"].strftime("%Y-%m-%d %H:%M:%S %z"))
    
    
    ## Load local data (for speed) if no new data are available
    if not newdata: 
        if args.newonly: # If -n option used, quit if no new data
            print("No updates available in CORDIS \nDone.")
            sys.exit(0)
        else:
            print("No updates available in CORDIS \nProceeding with local data")
            try:
                df_projects = pd.read_pickle("cordis_projects.pkl")
                df_orgs = pd.read_pickle("cordis_orgs.pkl")
                print("Loaded local data from pkl")
            except FileNotFoundError:
                print("No pkl files; processing from scratch")
                df_projects, df_orgs = processCordisData()
    else:        
        df_projects, df_orgs = processCordisData()

    ## Validate country parameters; note that it defaults to NZ if no arguments are given
    selected_countries = []
    if args.country:
        for item in args.country: # validate arguments as two-character country codes
            if item in country_sets:
                selected_countries += country_sets[item]
            elif item == 'GB': # the ISO 3166-1 alpha-2 code for the United Kingdom is GB, but the data uses UK
                selected_countries.append('UK')
            elif not get_country_name(item):
                print("Invalid country code:", item)
                sys.exit(0)
            else:
                selected_countries.append(item)

    ## Set save file name
    cordis_date = checkLocalDataDate(fp_data["HORIZON"]["path"]).strftime("[%Y%m%d_cordis_data]")
    # out_name = datetime.datetime.now().strftime("%Y%m%d_")+'CORDIS_summary-'+'_'.join(args.country)+"-"+cordis_date+'.xlsx'
    out_name = 'HE_CORDIS_summary-'+'_'.join(args.country)+"-"+cordis_date+'_'+datetime.datetime.now().strftime("%Y%m%d")+'.xlsx'
    ## Produce the summary
    do_cordis_summary(df_projects, df_orgs, selected_countries, out_name)
   

if __name__ == "__main__":
    main()

