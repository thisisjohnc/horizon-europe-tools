# Horizon Europe tools [![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fthisisjohnc%2Fhorizon-europe-tools&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://hits.seeyoufarm.com)

This repo contains Python scripts for producing overviews of funding calls and participation in Horizon Europe and previous EU Framework Programmes for Research and Innovation. I wrote them to support my work on Horizon Europe in New Zealand, but they may be useful for anyone interested in monitoring for new calls or understanding participation from different countries, such as NCPs, funding agencies, or research managers. They use data from public databases of the European Commission (the [EU Funding and Tenders Portal](https://ec.europa.eu/info/funding-tenders/opportunities/portal/screen/home) and [CORDIS](https://cordis.europa.eu/projects)). 

*Note that these tools and files are the work of the author and do not represent an official output of the New Zealand Government*

#### Don't want to run the code yourself?

Check back here periodically for updated output files in the Outputs folder.

#### Questions or suggestions?

You can try emailing [john@isospike.org](mailto:john@isospike.org)

## Horizon Europe calls spreadsheet generator

This script fetches updated Horizon Europe calls from the Funding and Tenders Portal and puts them into a filterable/sortable spreadsheet.

### Requirements

- Python 3.6+
- `requests`
- `pandas`
- `openpyxl`
- `tqdm`

Install the required Python packages using pip:

```
pip install requests pandas openpyxl tqdm
```

### Usage

You can run the script with the following command:

```
python HE_calls_updates.py [options] [compare_file]
```

This will download new calls and save to and Excel sheet with the current date (e.g., HE_calls_2024-05-19.xlsx) and compare with the optionally specified compare_file or the most recent previous output (if available).

#### Options

- `-n`, `--newonly`: Save outputs only if there are changes from the compare file
- `-l`, `--local`: Testing mode that will runs using previously saved local data
- `-c`, `--calendars`: Option to produce a visual calendar of call opening and closing dates
- `[compare_file]`: Optionally specify a file to compare to.

#### Examples
Download new calls and compare with a specific previous output, and save a new Excel output and calendars if there are new calls:

```
python HE_calls_updates.py -nc HE_calls_2024-04-16.xlsx
```


## CORDIS Data Processing

This repository contains a Python script designed to process data from the CORDIS (Community Research and Development Information Service) database. The script fetches, extracts, and processes project and organization data, and generates Excel reports summarizing the data for specified countries.

### Requirements

- Python 3.6+
- `requests`
- `pandas`
- `openpyxl`
- `pycountry`
- `tqdm`
- `argparse`

Install the required Python packages using pip:

```
pip install requests pandas openpyxl pycountry tqdm argparse
```

### Usage

Run the script with the following command:

```
python HE_CORDIS_updates.py [options] [country ...]
```

#### Options

- `-n`, `--new`: Process data and save summaries only if new data were found in CORDIS.
- `-f`, `--force`: Force download of data from CORDIS even if not newer.
- `-s`, `--skip`: Skip CORDIS check and use local data.
- `country`: List of two-character country codes (or predefined sets of countries) for summary (default: NZ).

Predefined country sets:
- `pacific`
- `eu_members`
- `associated_countries`
- `nordics`

#### Examples

Generate summaries for New Zealand (NZ), Canada (CA), and pacific island nations (pacific):

```
python HE_CORDIS_updates.py -n NZ CA pacific
```

## Contributing

Contributions are welcome. Please submit a pull request or open an issue to discuss any changes.

## License

[![CC BY-SA 4.0][cc-by-sa-shield]][cc-by-sa]

This work is licensed under a
[Creative Commons Attribution-ShareAlike 4.0 International License][cc-by-sa].

[![CC BY-SA 4.0][cc-by-sa-image]][cc-by-sa]

[cc-by-sa]: http://creativecommons.org/licenses/by-sa/4.0/
[cc-by-sa-image]: https://licensebuttons.net/l/by-sa/4.0/88x31.png
[cc-by-sa-shield]: https://img.shields.io/badge/License-CC%20BY--SA%204.0-lightgrey.svg
