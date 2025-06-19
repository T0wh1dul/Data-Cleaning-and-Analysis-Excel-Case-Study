# Data-Cleaning-and-Analysis-Excel-Case-Study
## Project Overview

This repository showcases cleaning and analyzing fleet equipment inventory data using Microsoft Excel, demonstrating key skills in data preparation, data manipulation, and basic data summarization with pivot tables.

## Problem Statement

I was tasked with two primary objectives:

1.  **Data Cleaning:** Import raw, comma-separated value (CSV) inventory data for the fleet of vehicles, convert it to an Excel workbook, and perform extensive data cleaning to ensure data quality and usability.
2.  **Data Analysis:** Utilize the cleaned data to perform initial analyses using Excel's powerful features, specifically focusing on summarizing fleet information by department and equipment class using pivot tables.

## Dataset

The dataset used in this project is a modified subset of the "Fleet Equipment Inventory" data from Montgomery County, MD, publicly available under a Public Domain license.

* **Part 1 (Cleaning):** `Montgomery_Fleet_PART_1_START.CSV`
* **Part 2 (Analysis):** `Montgomery_Fleet_PART_2_START.XLSX`

## Tools and Technologies

* **Microsoft Excel for the Web**: Used for all data cleaning, formatting, and analysis tasks.

## Key Skills Demonstrated

This project highlights proficiency in the following data analysis skills:

### Data Cleaning Techniques:

* **File Format Conversion**: Converting CSV to XLSX.
* **Column Formatting**: Adjusting column widths for readability.
* **Data Filtering**: Identifying and removing empty rows.
* **Duplicate Management**: Detecting and removing duplicate records.
* **Data Validation**: Checking for and correcting spelling errors.
* **Text Manipulation**: Removing extraneous whitespace (double spaces).
* **Data Transformation**: Consolidating split data (department names) using Flash Fill and removing redundant columns.

### Data Analysis Techniques:

* **Data Formatting**: Structuring data as an Excel Table.
* **Basic Aggregations**: Calculating SUM, AVERAGE, MIN, MAX, and COUNT using AutoSum.
* **Pivot Table Creation**: Summarizing data by department and equipment count.
* **Pivot Table Sorting**: Ordering pivot table results (descending by equipment count).
* **Multi-dimensional Analysis**: Creating pivot tables with multiple row fields (`Department` and `Equipment Class`) to show hierarchical data.
* **Data Drill-down**: Collapsing and expanding fields within pivot tables for targeted analysis (e.g., focusing on 'Transportation' or 'CUV' categories).

## Project Structure and Deliverables

The project is divided into two main parts, corresponding to the initial data cleaning and subsequent analysis phases:

### Part 1: Data Cleaning

* **Input File:** `Montgomery_Fleet_PART_1_START.CSV`
* **Output File:** `Montgomery_Fleet_PART_1_END.XLSX`
* **Tasks Performed:**
    * Converted CSV to XLSX.
    * Adjusted column widths.
    * Removed empty rows.
    * Removed duplicate records.
    * Corrected spelling mistakes.
    * Removed double spaces.
    * Consolidated `Department` names using Flash Fill.

### Part 2: Data Analysis with Pivot Tables

* **Input File:** `Montgomery_Fleet_PART_2_START.XLSX` (This is the cleaned data from Part 1, or a similar pre-cleaned version provided for Part 2)
* **Output File:** `Montgomery_Fleet_PART_2_END.XLSX`
* **Tasks Performed:**
    * Formatted data as an Excel Table.
    * Calculated SUM, AVERAGE, MIN, MAX, COUNT for 'C' column using AutoSum.
    * Created three identical pivot tables showing `Sum of Equipment Count` by `Department`.
    * Sorted all pivot tables by `Sum of Equipment Count` in descending order.
    * **Pivot Table 2 Analysis:** Added `Equipment Class` below `Department` and collapsed all fields except `Transportation`.
    * **Pivot Table 3 Analysis:** Added `Equipment Class` above `Department` and collapsed all fields except `CUV`.

## Author

**Towhidul Islam**
https://www.linkedin.com/in/towhidul-islam01/
---
