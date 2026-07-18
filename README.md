# Ecommerce Data Cleaning Project in Excel

## Project Overview

This project focuses on cleaning and preparing an ecommerce dataset using Microsoft Excel. The dataset contained 1,200 order records and was checked for missing values, duplicate records, incorrect dates, invalid numeric values, and calculation errors.

## Project Objectives

- Identify and handle missing values
- Check and remove duplicate records
- Validate unique Order IDs
- Standardize and validate date formats
- Validate numeric columns
- Verify Total Price calculations
- Prepare a clean dataset for analysis

## Dataset Columns

The dataset includes:

- Order ID
- Date
- Customer ID
- Product
- Quantity
- Unit Price
- Shipping Address
- Payment Method
- Order Status
- Tracking Number
- Items in Cart
- Coupon Code
- Referral Source
- Total Price

## Data Cleaning Performed

### Missing Values

309 blank values were found in the `CouponCode` column. These blanks were replaced with:

`No Coupon`

### Duplicate Records

Excel's Remove Duplicates feature was used to check the complete dataset.

Result:

- 0 duplicate rows found
- 0 duplicate Order IDs found

### Date Validation

The Date column was standardized to:

`DD/MM/YYYY`

An Excel validation formula was used to confirm that all dates were recognized as valid Excel dates.

Result:

- 0 incorrectly formatted dates found

### Numeric Validation

Helper columns and Excel formulas were used to validate:

- Quantity
- Unit Price
- Items in Cart
- Total Price

Result:

- 0 invalid numeric values found

### Price Calculation Validation

Total Price was checked using:

`Total Price = Quantity × Unit Price`

Result:

- 0 calculation errors found

## Tools and Excel Features Used

- Microsoft Excel
- Filters
- Remove Duplicates
- COUNTIF
- IF
- ISNUMBER
- AND
- INT
- ABS
- Helper columns
- Cleaning Log

## Final Results

- Total records checked: 1,200
- Missing Coupon Codes handled: 309
- Duplicate rows: 0
- Duplicate Order IDs: 0
- Invalid dates: 0
- Invalid numeric values: 0
- Price calculation errors: 0

## Project Files

- Cleaned Excel workbook
- Cleaning Log
- Dataset validation helper columns
- Project screenshots

## Key Learning

This project helped me understand that reliable analysis begins with clean and accurate data. I learned how to detect missing values, validate data formats, identify duplicate records, and use Excel formulas to reduce human error.
