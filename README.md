# Quarterly Stock Analysis Tool

## Overview

This VBA script is designed to calculate and summarize quarterly results for stock data. It processes raw daily stock data and generates a concise quarterly report, providing key metrics for each stock ticker.

## Features

1. **Data Processing**: Sorts the raw data by ticker and date for accurate calculations.
2. **Quarterly Metrics**: Calculates the following metrics for each stock ticker on a quarterly basis:
   - Quarterly Change (in price)
   - Percentage Change
   - Total Volume
   

## Output

The script generates a new worksheet with the following columns:
1. Ticker (Stock)
2. Quarter (in the format "YYYY Q#")
3. Quarterly Change
4. Percentage Change
5. Total Volume

## Usage

To use this script:
1. Ensure your raw stock data is in a sheet named "Q1".
2. Run the `CalculateQuarterlResults` macro.
3. Review the results in the newly created "Quarter 1 Results" sheet.

## Notes

- The script assumes that the input data is sorted by date within each ticker.
- It's designed to handle multiple quarters of data


