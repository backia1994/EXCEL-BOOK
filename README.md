Product Dataset Analysis Using Microsoft Excel

Introduction :

This assignment demonstrates the application of Microsoft Excel
functions to analyze a product dataset. The project focuses on
performing descriptive statistics, conditional analysis, and text
manipulation to generate meaningful insights from raw data.
The objective is to showcase fundamental Excel skills commonly used in
data analytics, including data summarization, categorization, and
information extraction.
Dataset Overview
The dataset contains information related to products, including:
•	Product ID
•	Product Category
•	Product Price
•	Additional product-related attributes
The data was analysed using built-in Excel formulas to support business
decision-making and reporting.

Data Analysis Performed

Total Product Price :

The  ”SUM” function was used to calculate the total value of all
products in the dataset.
**Formula**

=SUM(B2:B100)

Purpose - Measures the overall value of inventory/products. -
Provides a quick financial summary.

Product Count :

The **COUNT** function was used to determine the total number of
products.
**Formula**

=COUNT(B2:B100)

Purpose - Identifies the total number of valid product records. -
Supports inventory and dataset validation.

Average Product Price :

The **AVERAGE** function was used to calculate the mean product price.
**Formula**

=AVERAGE(B2:B100)

Purpose - Helps understand the typical pricing level across
products. - Useful for pricing analysis.

Minimum Product Price :

The **MIN** function was used to identify the lowest-priced product.

**Formula**

=MIN(B2:B100)

Purpose - Detects the least expensive product. - Supports price
range analysis.

Maximum Product Price :

The **MAX** function was used to identify the highest-priced product.

**Formula**

=MAX(B2:B100)

Purpose - Detects premium-priced products. - Supports market
positioning analysis.

Conditional Analysis :

Product Price Classification

The **IF** function was used to classify products based on price.

Business Rule
Price Condition   Category
----------------- ----------------
Price \>= 500     High Price
Price \< 500      Standard Price

**Formula**

=IF(B2>=500,"High Price","Standard Price")

Purpose - Segments products into pricing categories. - Simplifies
business reporting.



Total Electronics Sales Value :

The **SUMIF** function was used to calculate the total price of products
belonging to the Electronics category.
**Formula**

=SUMIF(C2:C100,"Electronics",B2:B100)

Purpose - Measures category-specific revenue. - Supports category
performance analysis.

Count of Low-Priced Products :

The **COUNTIF** function was used to count products priced below 100.

**Formula**

=COUNTIF(B2:B100,"<100")

Purpose - Identifies budget-friendly products. - Supports pricing
strategy analysis.

Product ID Analysis Using Text Functions :

To extract meaningful information from Product IDs, Excel text functions
were used. 

Date Extraction :

The **LEFT** function was used to extract the first two characters from
the Product ID.

**Formula**

=LEFT(A2,2)

**Output Column** - Date

Country Code Extraction :

The **RIGHT** function was used to extract the last two characters from
the Product ID.

**Formula**

=RIGHT(A2,2)

**Output Column** - Country Code


Month Extraction :

The **MID** function was used to extract characters from the middle of
the Product ID.

**Formula**
=MID(A2,3,2)

**Output Column** - Month

Key Skills Demonstrated :

o	Microsoft Excel
o	Data Analysis
o	Data Cleaning
o	Descriptive Statistics
o	Conditional Logic
o	Data Categorization
o	Spreadsheet Reporting
o	Business Data Interpretation

Conclusion :

This assignment demonstrates the effective use of Excel functions for
data analysis and reporting. By applying statistical, logical,
conditional, and text functions, meaningful insights were extracted from
the product dataset. The project highlights essential analytical skills
required for entry-level Data Analyst and Data Analytics Internship
roles.
The completed analysis provides a structured summary of product pricing,
category performance, and product identification details, showcasing
practical Excel proficiency in a business analytics environment.

