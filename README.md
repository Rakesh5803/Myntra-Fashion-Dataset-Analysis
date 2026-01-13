DATA LINK 🔗
EXCEL Unclean file
🔗https://1drv.ms/x/c/7181b8536c05b374/IQBqUwGOPHYZSIJ2LxsVeRn9Af_NEpVzL6Aqu4asHlT79gc?e=7oV6Wv
🔗https://1drv.ms/x/c/7181b8536c05b374/IQAlAMCfOrR1QIWXuaBY6cEUAU-5VWECGR5FG-mIvyRde-4?e=06rNrV clean file 
🔗
📊 Data Cleaning & Analysis Project (Excel)
📌 Project Overview
This project focuses on data cleaning, preparation, analysis, and lookup operations using an e-commerce fashion dataset (Myntra Fashion Clothing data).
The main goal is to improve data quality, analyze product pricing and discounts, and retrieve specific product details using Excel functions.

🧩 Dataset Description
The dataset contains information related to fashion products, including:
Product ID
Brand
Category
Original Price
Discount Price
Discount Offer
Size Options
Ratings
The data required cleaning before performing meaningful analysis.

🛠 Tools & Techniques Used
Microsoft Excel
Excel Functions:
oAVERAGE
oCOUNTIF
oIF
oVLOOKUP
oXLOOKUP
oINDEX
oMATCH
 Data Cleaning Techniques:
oRemoving duplicates
oHandling missing values
oStandardizing text formats

🔹 A. Data Cleaning and Preparation
1. Removing Duplicate Records
 Checked the dataset for duplicate rows.
 Removed duplicate entries to ensure data accuracy.
2. Standardizing DiscountOffer Column
 Converted all discount values into a uniform percentage format (e.g., 50% OFF).
 This helped in easier comparison and analysis.
3. Handling Missing Discount Values
 Identified rows where both DiscountPrice and DiscountOffer were missing.
 Filled the DiscountPrice with the average discount price of the respective product category.
4. Handling Missing Size Options
 Replaced all NULL values in the SizeOption column with:
 Copy codeNotAvailable

🔹 B. Data Analysis
1. Average Original Price (Rating > 4)
 Calculated the overall average original price for products having ratings greater than 4.
2. Products with High Discount
 Counted products where Discount Offer > 50% OFF.
3. Size Availability Analysis
 Counted the number of products available in Size "M".
4. Discount Category Labeling
 Created a new column Discount Category:
oHigh Discount → Discount > 50% OFF
oLow Discount → Discount ≤ 50% OFF
Formula Used:
Copy code=IF(DiscountOffer>50,"High Discount","Low Discount")

🔹 C. Data Retrieval and Lookup
1. Product Details using VLOOKUP / XLOOKUP
 Retrieved Brand, Price, and Rating for:
 Copy codeProduct ID:11226634
2. Discount Price using INDEX & MATCH
 Found the DiscountPrice for:
 Copy codeProduct ID:6744434
Formula Example:
Copy code=INDEX(DiscountPriceRange, MATCH(ProductID, ProductIDRange, 0))
3. Nested XLOOKUP
 Used nested XLOOKUP to retrieve any column’s information using Product ID dynamically.

✅ Key Findings
 Products with higher ratings often have higher average original prices.
 A significant number of products offer discounts above 50%, attracting customers.
 Size "M" is one of the most commonly available sizes.
 Data consistency greatly improved after cleaning.

🎯 What We Accomplished
 Cleaned and standardized raw data
 Improved data accuracy and reliability
 Performed meaningful pricing and discount analysis
 Implemented powerful Excel lookup techniques



📝 Conclusion
This project demonstrates how Excel can be effectively used for data cleaning, analysis, and lookup operations.
By applying proper data preparation techniques and Excel formulas, valuable insights can be extracted from raw datasets.
