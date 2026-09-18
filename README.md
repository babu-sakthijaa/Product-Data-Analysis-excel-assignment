# 📊 Excel Data Exploration – Product Dataset

## 📌 Project Overview

This project is part of my journey toward becoming a **Data Analyst**. The objective of this assignment is to build foundational skills in **Microsoft Excel** by performing basic data exploration, calculations, conditional analysis, logical categorization, and text manipulation.

The project uses a product dataset containing information such as:

* Product ID
* Product Name
* Brand Name
* Price
* Quantity
* Category

Using Excel formulas and functions, I analyzed the dataset and created additional columns to extract meaningful information from the Product IDs.

---

## 🎯 Project Objectives

The main objectives of this project are to:

* Explore and summarize product data.
* Calculate basic statistical measures.
* Apply logical conditions using the `IF` function.
* Perform conditional calculations using `SUMIF` and `COUNTIF`.
* Extract information from text using `LEFT`, `RIGHT`, and `MID`.
* Develop practical Excel skills used in data analysis.
* Present the analysis as part of my Data Analyst portfolio.

---

## 📁 Dataset

The dataset contains product-level information with the following attributes:

| Column       | Description                                                     |
| ------------ | --------------------------------------------------------------- |
| Product ID   | Unique identifier containing date/month and country information |
| Product Name | Name of the product                                             |
| Brand Name   | Brand associated with the product                               |
| Price ($)    | Price of the product in US dollars                              |
| Quantity     | Quantity associated with the product                            |
| Category     | Product category                                                |

### Example Product ID

`28-JAN-US`

The Product ID was also used to demonstrate Excel text extraction functions.

---

# 🔎 Analysis Performed

## 1. Basic Data Exploration

I used Excel functions to calculate the following:

### Total Price

The total price of all products was calculated using:

```excel
=SUM(D2:D35)
```

This provides the sum of all product prices in the dataset.

### Number of Products

The number of products was calculated using:

```excel
=COUNTA(C2:C35)
```

This counts the populated product records.

### Average Price

The average product price was calculated using:

```excel
=AVERAGE(D2:D35)
```

This provides the average price across all products.

---

## 2. Minimum and Maximum Price

To understand the price range of the products, I calculated the minimum and maximum prices.

### Minimum Price

```excel
=MIN(D2:D35)
```

### Maximum Price

```excel
=MAX(D2:D35)
```

These calculations help identify the lowest- and highest-priced products in the dataset.

---

# 3. 🏷️ Price Categorization Using IFS

A new column called **Price Range** was created to categorize products based on their price.

The business rule was:

* Price ≥ $500 → **High Price**
* Price < $500 → **Standard Price**

The Excel formula used was:

```excel
=IFS(D2>=500,"HIGH PRICE","STANDARD PRICE")
```

This demonstrates the use of conditional logic in Excel.

### Example

|  Price | Price Range    |
| -----: | -------------- |
| $1,000 | HIGH PRICE     |
|    $80 | STANDARD PRICE |
|   $130 | STANDARD PRICE |
|   $900 | HIGH PRICE     |

---

# 4. 🧮 Conditional Analysis

## SUMIF – Electronics Products

I used the `SUMIF` function to calculate the total price of products belonging to the **Electronics** category.

Formula:

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

Where:

* `F2:F35` → Category range
* `"Electronics"` → Condition
* `D2:D35` → Price range

This demonstrates how Excel can aggregate values based on a specific condition.

---

## COUNTIF – Products Below $100

I used `COUNTIF` to determine how many products have a price below $100.

Formula:

```excel
=COUNTIF(D2:D35,"<100")
```

This identifies the number of products that fall below the specified price threshold.

---

# 5. 🔤 Text Extraction Using LEFT, RIGHT and MID

The Product ID contains multiple pieces of information.

For example:

```text
28-JAN-US
```

I used Excel text functions to extract individual parts of the Product ID.

## Day

A new **Day** column was created using the `LEFT` function.

```excel
=LEFT(A2,2)
```

For:

```text
28-JAN-US
```

The result is:

```text
28
```

---

## Country Code

A new **Country Code** column was created using the `RIGHT` function.

```excel
=RIGHT(A2,2)
```

For:

```text
28-JAN-US
```

The result is:

```text
US
```

---

## Month

A new **Month** column was created using the `MID` function.

```excel
=MID(A2,4,3)
```

For:

```text
28-JAN-US
```

The result is:

```text
JAN
```

---

# 📋 Excel Functions Used

| Function  | Purpose                               |
| --------- | ------------------------------------- |
| `SUM`     | Calculate total price                 |
| `COUNTA`  | Count product records                 |
| `AVERAGE` | Calculate average price               |
| `MIN`     | Find minimum price                    |
| `MAX`     | Find maximum price                    |
| `IFS`       | Categorize products by price          |
| `SUMIF`   | Calculate price total for Electronics |
| `COUNTIF` | Count products below $100             |
| `LEFT`    | Extract the Day from Product ID       |
| `RIGHT`   | Extract the Country Code              |
| `MID`     | Extract the Month                     |

---

# 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Formulas & Functions
* Data Exploration
* Conditional Analysis
* Text Manipulation

---

# 📂 Project Structure

```text
Excel-Data-Exploration/
│
├── README.md
│
└── Excel Assignment 1 - Data Exploration.xlsx
```

---

# 💡 Key Learnings

Through this assignment, I practiced several fundamental data analysis concepts:

* Working with structured datasets in Excel.
* Performing basic descriptive analysis.
* Applying logical conditions to classify data.
* Performing conditional aggregation.
* Counting records based on conditions.
* Extracting meaningful information from text fields.
* Using Excel formulas to automate repetitive calculations.
* Organizing an analysis project for portfolio presentation.

---

# 🚀 Future Improvements

As I continue developing my Data Analyst skills, I plan to extend this project by:

* Creating Excel PivotTables.
* Building interactive dashboards.
* Adding charts and visualizations.
* Performing category-level analysis.
* Analyzing quantity and revenue-related metrics.
* Applying data cleaning techniques.
* Exploring the same dataset using **SQL** and **Python/Pandas**.
* Comparing insights across different analytical tools.

---

# 👨‍💻 About Me

I am currently building my skills toward a career as a **Data Analyst** and using practical projects to develop my portfolio.

This repository represents one of my initial projects focused on **Excel-based data analysis**.

I am continuously learning and expanding my skills in:

* 📊 Microsoft Excel
* 🗄️ SQL
* 🐍 Python
* 📈 Data Visualization
* 📉 Statistics
* 🔍 Data Analysis

More projects will be added as I continue my learning journey.

---

⭐ **This project is part of my Data Analyst portfolio and learning journey.**
