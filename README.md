# Goolge Analytics Reporting API
## Notice
It is possible that GitHub fails to display Jupyter Notebooks. Should such circumstances arise, please refer to ***Part 4. Steps*** listed below for code samples.

## Part 1. Objective
Automatically update Google Analytics reporting data to the existing Excel workbook by R and Python.

## Part 2. Data
Extract custom data from the Google Analytics API.

## Part 3. Outline
### 3.1. Data Extraction 
Retrieve Google Analytics reporting data by R which provides simpler authentication without having to create another project in the Developer Console. 
- Tool: ```googleAnalyticsR``` ```magrittr``` ```dplyr``` ```xlsx``` 

### 3.2. Data Updates
Update the existing Excel workbook with extracted Google Analytics reporting data by Python.
- Tools: ```xlwings```

### 3.3. Facilitate Reporting Process
Integrate the process of data extraction and data updates into a batch file that can be executed as the scheduled task on a daily basis.
- Tool: ```rp2``` ```xlwings```
> Notice:  
To successfully execute R script from Python, the following steps should be taken:

## Part 4. Steps
