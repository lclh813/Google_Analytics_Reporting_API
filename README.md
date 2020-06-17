# Goolge Analytics Reporting API
## Python Project

## Jupyter Notebook
It is possible that GitHub fails to display Jupyter Notebooks. Should such circumstances arise, please refer to ***Part 4. Steps*** listed below for code samples.

## Part 1. Objective
Automatically update Google Analytics reporting data to the existing Excel workbook by R and Python.
<br>
<div align=center><img src="https://github.com/lclh813/Google_Analytics_Reporting_API/blob/master/0_Intro.png"/></div>
<br>

## Part 2. Data
Extract custom data from the Google Analytics API.

## Part 3. Outline
### 3.1. Data Extraction 
Retrieve Google Analytics reporting data by R which provides simpler authentication without having to create another project in the Developer Console. 
- Tool: ```googleAnalyticsR``` ```magrittr``` ```dplyr``` ```xlsx``` 

### 3.2. Data Updates
Update the existing Excel workbook with extracted Google Analytics reporting data by Python.
- Tools: ```rpy2``` ```xlwings```

### 3.3. Facilitate Reporting Process
Integrate the process of data extraction and data updates into a batch file that can be executed as the scheduled task on a daily basis.

### Notice: 
To execute R script from Python successfully, the following steps should be taken:  
- Install ***rpy2*** package with a ***whl*** file.  
- Set environmental variables which can affect an R session, incliuding:  
```R_HOME```: Locate top-level directory of R.  
```R_PATH```: Locate ***R.dll*** file.  
```R_USER```: Locate ***rpy2*** package.  
```R_LIBS_USER```: Locate ***R Library***.
- If R library and Python library are not in the same directory, it is necessary to specify where R library is when importing R packages in Python.

## Part 4. Steps
> [***Complete Code***](https://nbviewer.jupyter.org/github/lclh813/Google_Analytics_Reporting_API/blob/master/4_CompleteCode.ipynb)  

### Step 1. Preparation
[1. Preparation](https://nbviewer.jupyter.org/github/lclh813/Google_Analytics_Reporting_API/blob/master/1_Preparation.ipynb)  

### Step 2. Data Extraction
[2. Data Extraction](https://nbviewer.jupyter.org/github/lclh813/Google_Analytics_Reporting_API/blob/master/2_DataExtraction.ipynb)  
[2.2.1. Extract Data of a Given Day](https://github.com/lclh813/Google_Analytics_Reporting_API/blob/master/2_2_1_Oneday.R)  
[2.2.2. Extract Data of a Given Period of Time](https://github.com/lclh813/Google_Analytics_Reporting_API/blob/master/2_2_2_Period.R)  

### Step 3. Data Updates
[3. Data Updates](https://nbviewer.jupyter.org/github/lclh813/Google_Analytics_Reporting_API/blob/master/3_DataUpdates.ipynb)  

## Part 5. Reference
- [How to Install rpy2 in Windows](https://www.cnblogs.com/Xeonilian/p/windows_rpy2_install.html) 
- [How to Install R Packages in Python](https://stackoverflow.com/questions/46140624/unable-to-install-r-package-in-python-jupyter-notebook)
- [R Manual](https://stat.ethz.ch/R-manual/)
