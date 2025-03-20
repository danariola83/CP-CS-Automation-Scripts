# Overview
A set of scripts written in Python for automation of CP-CS (Cycle Plan-Customer Specific) data processing.

Using Pandas and NumPy, the scripts ingest xlsx/b and csv files, process these files for merchandise activity forms, and outputs a file/s that's ready for use on the AFS / MARS retail execution platform.

# Introduction
This is a personal project built while self-studying Python during my previous job as a data analyst at a multinational FMCG company. Around 70% of my time during this tenure was spent manually cleaning, validating, and processing large datasets for the company's merchandise operations, and so I figured I'd try building a program that addresses this issue.

Because this was built during my time as a beginner with Python and programming in general, the code is quite messy--no modularity, overly long comments, no concept of the DRY principle, and just an overall lack of adherence to best practices. Documentation was also poorly written. I only knew that, after several months of testing on actual data, it worked as intended and was eventually deployed for the position's future use after my resignation.

The program shortens what is usually 2-3 days worth of manual work, requiring use of multiple tools (i.e Excel, Power BI, DAX, and/or Kutools), to just a few clicks. Depending on the volume of data, a script usually parses through data and outputs a file in less than a minute. 

For being my first ever project implemented at a work environment, despite how poorly written the codebase may have been, it still alleviates ~60-70% of the weekly workload for the position, and for that I am incredibly proud of it.

# Architecture
![System Architecture](CP-CS-Automation-Scripts.drawio.png)


## Pre-requisites

- An IDE
- Anaconda/Miniconda

## Dependencies

et-xmlfile==1.1.0
numpy==1.26.4
openpyxl==3.1.2
pandas==2.2.2
python-dateutil==2.9.0.post0
pytz==2024.1
pyxlsb==1.0.10
six==1.16.0
tzdata==2024.1
xlrd==2.0.1
XlsxWriter==3.2.0

# Usage
Apart from packages and dependencies, directory structure is perhaps the most critical aspect in running a script.
Using these scripts assumes you are thoroughly familiar with CP-CS procedures.

## Method 1: No GitHub
Assuming you don't have a GitHub account and are not familiar with cloning a repository, you will need to setup your folder tree as follows:
```
├── CP-CS-Automation-Scripts.drawio.png
├── CP-CS-Automation-Scripts.drawio.png:Zone.Identifier
├── CPCS_Files
│   ├── AUGUST2024_CPCS                 <-- folders containing CP-CS files using the naming convention '<WORKING_MONTH><WORKING_YEAR>_CPCS'
│   │   ├── CP_OutputFiles              <-- folder where output files are generated upon running a script
│   │   │   ├── NCM
│   │   │   └── NFO
│   │   ├── CP_RawFiles                 <-- folder where the user should place raw files for the month of August 2024 into the appropriate sub-directories
│   │   │   ├── NCM
│   │   │   └── NFO
│   │   ├── CS_OutputFiles
│   │   │   ├── B1
│   │   │   └── B2
│   │   └── CS_RawFiles
│   │       ├── B1
│   │       └── B2
│   └── JULY2024_CPCS
│       ├── CP_OutputFiles
│       │   ├── NCM
│       │   └── NFO
│       ├── CP_RawFiles
│       │   ├── NCM
│       │   └── NFO
│       ├── CS_OutputFiles
│       │   ├── B1
│       │   └── B2
│       └── CS_RawFiles
│           ├── B1
│           └── B2
├── CPCS_Scripts
│   ├── CP_NFO_PER_CHANNEL_REGION_v1.1.py
│   ├── CS_EXECUTION_B1_PER_CHAIN_v1.1.py
│   ├── CS_EXECUTION_B1_PER_DOOR_v1.1.py
│   ├── CS_EXECUTION_B2_PER_CHAIN_v1.1.py
│   ├── CS_EXECUTION_B2_PER_DOOR_v1.1.py
│   ├── CS_preEXECUTION_B1_PER_CHAIN_v1.1.py
│   ├── CS_preEXECUTION_B1_PER_DOOR_v1.1.py
│   ├── CS_preEXECUTION_B2_PER_CHAIN_v1.1.py
│   └── CS_preEXECUTION_B2_PER_DOOR_v1.1.py
├── README.md
└── requirements.txt
```

## Method 2: Cloning the repository from GitHub

1. Open a terminal and run the command `git clone https://github.com/danariola83/CP-CS-Automation-Scripts.git`
2. `cd` into the repo's root and run the command `pip install -r requirements.txt`
3. Copy raw excel and csv files into the appropriate folders
4. Run the pertinent script/s


# Planned Changes

- place dicts and lists in a separate `references.py` file
- create `functions.py` for tasks that are common across all scripts:
    - reading in excel and csv files
    - parsing dates
    - chain and category adjustments
    - creating fields for groupings, form names, and form IDs
    - creating and processing separate dataframe for Customer Code/Group Name csv (Per Door scripts)
    - writing DFs to excel
        - consider different `functions.py` that contain specific functions for Per Region-Channel, Per Chain, and Per Door operations
- fuzzy string matching to replace convoluted list comprehensions that standardize chain and category names
- transfer all comments regarding CP-CS operations and instructions here on the readme file