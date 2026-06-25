# Café Sales — Data Cleaning & Exploration

## Overview

The goal behind this project was to build an interactive Microsoft Excel dashboard to allow users to identify trends in coffee bean sales and compare these sales across roast types and countries.

**Data Source:** [Mo Chen YouTube](https://github.com/mochen862/excel-project-coffee-sales)
**Tools:** Excel 
**Skills practiced:** Data Cleaning/Preparation and Dashboard Building 

---

## Objectives

- This project involves end to end data analysis in Microsoft Excel, from data cleaning and transformation to creating meaningful visualizations in the form of a dynamic, interactive dashboard.
- The dataset contains three worksheets — `orders`, `customers` and `products`, with `orders` being related to `customers` and `products` with the Customer ID and Product ID as the "foreign keys". I will refer to the worksheets as tables for the rest of this documentation.

---

## What I Did

### 1. Data Preparation Using Advanced Formulas and Functions
- Removed duplicate rows and blank rows, trimmed whitespace, and added `kg` to the `Size` column for a better understanding of what the data means.
- Used `XLOOKUP` to look up the customer’s full name from the `customers` table to the `orders` table 
- Used `IF` and `XLOOKUP` to look up the email address from the `customers` table to the `orders` table 
- Used `XLOOKUP` to look up the country of the customer from the `customers` table to the `orders` table     
- Used one dynamic `INDEX` `MATCH` formula to look up the coffee type, roast type, size and unit price values from the `products` table to the `orders` table. I used the correct cell locking (“$”) for an absolute reference so that the formula in cell H1 can be automatically filled to the right and to the bottom of the table
- Calculated the sales as the product of `unit_price * quantity`  
- Used multiple `IF` functions to get the full coffee type and roast type names    
- Used `XLOOKUP` to look up the loyalty card status from the `customers` table to the `orders` table

### 2. Data Visualizations And Dashboard Build Using `Pivot Tables`, `Pivot Charts`, and `Timeline Slicers`
- Tracked total coffee bean sales by type over time using a Line Chart
- Tracked sales by country using a Bar Chart
- Identified the top 5 customers based on sales using a Bar Chart
- Created customized timeline and slicers (roast type, size, loyalty card) to give the user control on the exact dashboard view they want to see. I made sure to connect the slicers to all the Pivot Tables/Charts to ensure that a user is able to filter everything when they click on them

---

## Files in This Repo

| File | Description |
|------|-------------|
| `coffeeOrdersDataProject.xlsx` | New Dynamic Dashboard in Excel File |
| `coffeeOrdersData.xlsx` | Original Data in Excel File |


---

## What I Learned

- How to work with multiple worksheets in an Excel workbook
- How to build an interactive dashboard
- How to create pivot charts, multiple slicers, and timeline filters
- Practice with advanced formulas and functions

---

*Part of my data science portfolio — [View Portfolio](#)*
