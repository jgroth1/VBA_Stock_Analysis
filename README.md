# VBA_Stock_Analysis

## Goal of Project:

The Goal of this project is to create a VBA script that provides easy analysis of yearly stock performance.  The worksheets are organized by year and contain the starting, final, highest, and lowest value of each stock for each day of the year.  The VBA script will give the yearly increase/decrease in stock value for each stock, the percentage change over the year, and the total volume for the year.  In addition the stock will give the greatest increase, decrease, and total volume.   

## Assumptions:

1. The data has been sorted such that the stocks are grouped together by ticker value
2. The data is sorted by date such that the dates are in order with the first date for a given stock is the begining of the year and the last value for a give stock is the last date of the year traded.
3. Non-zero stock value is not assumed
4. The data in the spread sheet represents the complete data with no missing values

## Steps:

### 1. extract ticker values:

The Unique ticker values are extracted and inserted into a column.

### 2. extract initial and final values for stock price:

The initial and final values for the stock price. As non-zero values are not assumed, the first non-zero value is taken to avoid division by zero.

### 3. Calculate values

The values are calculated for yearly change in stock values, yearly percentage change, and yealy total volume. Values are displayed in sheet next to corresponding ticker value.

yearly change = final value - initial (non-zero) value 

percentage change = yearly change / initial (non-zero) value

yearly volume = Sum(daily volume)

### 4. Determine Max and Min values:

The Greatest % increase, greatest % decrease, and greatest yearly volume are determined from the values calculated in the last step and displayed in the chart.
