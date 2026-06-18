# Excel for Business Analytics — Complete Project Repository 

> A hands-on project repository demonstrating real-world application of Microsoft Excel for Business Analytics — covering formula engineering, Power Query, DAX, data modeling, and dashboard design.

---

## 📌 Table of Contents

- [About This Repository](#about-this-repository)
- [Skills & Topics Covered](#skills--topics-covered)
  - [1. Logical Functions](#1-logical-functions)
  - [2. Text Cleaning Functions](#2-text-cleaning-functions)
  - [3. Lookup & Relationship Functions](#3-lookup--relationship-functions)
  - [4. Data Validation & Error Detection](#4-data-validation--error-detection)
  - [5. Date & Time Functions](#5-date--time-functions)
  - [6. Aggregation & Statistical Functions](#6-aggregation--statistical-functions)
  - [7. Conditional Aggregation Functions](#7-conditional-aggregation-functions)
  - [8. Dynamic Array Functions](#8-dynamic-array-functions)
  - [9. Advanced Formula Engineering](#9-advanced-formula-engineering)
  - [10. Power Query](#10-power-query)
  - [11. Data Modeling](#11-data-modeling)
  - [12. DAX Functions](#12-dax-functions)
  - [13. Dashboard Engineering](#13-dashboard-engineering)
  - [14. Statistical Analysis in Excel](#14-statistical-analysis-in-excel)
  - [15. Automation in Excel](#15-automation-in-excel)
- [Learning Roadmap](#learning-roadmap)
- [Professional Competencies](#professional-competencies)

---

## About This Repository

This repository contains hands-on Excel projects built while mastering business analytics through Microsoft Excel. Each project applies a specific skill set — from core formula engineering to Power Query ETL, DAX-powered data models, and executive-grade interactive dashboards.

All projects are Excel-native, reflecting a practical, career-focused approach to Data Analytics and Business Intelligence.

---

## Skills & Topics Covered

---

### 1. Logical Functions

Mastered conditional logic to control data flow and automate decision-making directly within Excel worksheets.

| Function | Purpose | Example |
|----------|---------|---------|
| `IF()` | Conditional logic | `=IF(A2>100,"High","Low")` |
| `IFS()` | Multiple conditions without nesting | `=IFS(A2>90,"A",A2>80,"B")` |
| `AND()` | All conditions must be true | `=AND(A2>0,B2<100)` |
| `OR()` | At least one condition must be true | `=OR(A2="Yes",B2="Approved")` |
| `NOT()` | Reverse a logical result | `=NOT(A2="Closed")` |
| `IFERROR()` | Catch and replace any error | `=IFERROR(A2/B2,0)` |
| `IFNA()` | Handle only `#N/A` errors | `=IFNA(VLOOKUP(A2,D:E,2,FALSE),"Missing")` |

---

### 2. Text Cleaning Functions

Developed proficiency in standardizing and transforming raw, inconsistent text data into clean, analysis-ready formats.

| Function | Purpose | Example |
|----------|---------|---------|
| `TRIM()` | Remove extra spaces | `=TRIM(A2)` |
| `CLEAN()` | Remove non-printable characters | `=CLEAN(A2)` |
| `UPPER()` | Convert text to uppercase | `=UPPER(A2)` |
| `LOWER()` | Convert text to lowercase | `=LOWER(A2)` |
| `PROPER()` | Capitalize each word | `=PROPER(A2)` |
| `LEFT()` | Extract characters from the left | `=LEFT(A2,3)` |
| `RIGHT()` | Extract characters from the right | `=RIGHT(A2,4)` |
| `MID()` | Extract text from a specific position | `=MID(A2,2,5)` |
| `LEN()` | Count total characters | `=LEN(A2)` |
| `FIND()` | Case-sensitive position search | `=FIND("@",A2)` |
| `SEARCH()` | Case-insensitive position search | `=SEARCH("apple",A2)` |
| `SUBSTITUTE()` | Replace specific text patterns | `=SUBSTITUTE(A2,"-","")` |
| `REPLACE()` | Replace text by character position | `=REPLACE(A2,1,3,"XYZ")` |
| `TEXTBEFORE()` | Extract text before a delimiter | `=TEXTBEFORE(A2,"-")` |
| `TEXTAFTER()` | Extract text after a delimiter | `=TEXTAFTER(A2,"-")` |
| `TEXTSPLIT()` | Split text into an array | `=TEXTSPLIT(A2,",")` |
| `CONCAT()` | Combine text values | `=CONCAT(A2,B2)` |
| `TEXTJOIN()` | Join text with a delimiter | `=TEXTJOIN(",",TRUE,A2:A10)` |

---

### 3. Lookup & Relationship Functions

Built expertise in retrieving and cross-referencing data across large datasets and multiple Excel tables.

| Function | Purpose | Example |
|----------|---------|---------|
| `XLOOKUP()` | Modern, flexible lookup with exact/approximate match | `=XLOOKUP(A2,D:D,E:E)` |
| `VLOOKUP()` | Classic vertical lookup by column index | `=VLOOKUP(A2,D:E,2,FALSE)` |
| `INDEX()` | Return a value from a range by position | `=INDEX(B:B,5)` |
| `MATCH()` | Find the position of a value in a range | `=MATCH(A2,D:D,0)` |
| `INDEX + MATCH` | Flexible two-dimensional lookup | `=INDEX(E:E,MATCH(A2,D:D,0))` |
| `XMATCH()` | Enhanced match with richer options | `=XMATCH(A2,D:D)` |

---

### 4. Data Validation & Error Detection

Applied IS-functions to enforce data integrity, detect anomalies, and build robust, error-tolerant spreadsheet models.

| Function | Purpose | Example |
|----------|---------|---------|
| `ISBLANK()` | Detect empty cells | `=ISBLANK(A2)` |
| `ISNUMBER()` | Check if a value is numeric | `=ISNUMBER(A2)` |
| `ISTEXT()` | Check if a value is text | `=ISTEXT(A2)` |
| `ISERROR()` | Detect any error type | `=ISERROR(A2)` |
| `ISNA()` | Detect `#N/A` errors specifically | `=ISNA(A2)` |
| `EXACT()` | Case-sensitive string comparison | `=EXACT(A2,B2)` |

---

### 5. Date & Time Functions

Gained command over temporal data — calculating durations, building calendars, and extracting time-based insights for reporting.

| Function | Purpose | Example |
|----------|---------|---------|
| `TODAY()` | Return today's date dynamically | `=TODAY()` |
| `NOW()` | Return current date and time | `=NOW()` |
| `DATE()` | Construct a date from year, month, day | `=DATE(2026,5,24)` |
| `YEAR()` | Extract the year from a date | `=YEAR(A2)` |
| `MONTH()` | Extract the month from a date | `=MONTH(A2)` |
| `DAY()` | Extract the day from a date | `=DAY(A2)` |
| `DATEDIF()` | Calculate elapsed time between two dates | `=DATEDIF(A2,B2,"D")` |
| `NETWORKDAYS()` | Count working days between two dates | `=NETWORKDAYS(A2,B2)` |

---

### 6. Aggregation & Statistical Functions

Applied core aggregation and descriptive statistics to summarize and profile business datasets efficiently.

| Function | Purpose | Example |
|----------|---------|---------|
| `SUM()` | Total a range of values | `=SUM(A2:A10)` |
| `AVERAGE()` | Calculate the mean | `=AVERAGE(A2:A10)` |
| `MIN()` | Return the minimum value | `=MIN(A2:A10)` |
| `MAX()` | Return the maximum value | `=MAX(A2:A10)` |
| `COUNT()` | Count numeric cells | `=COUNT(A2:A10)` |
| `COUNTA()` | Count all non-empty cells | `=COUNTA(A2:A10)` |
| `MEDIAN()` | Find the middle value | `=MEDIAN(A2:A10)` |
| `MODE()` | Find the most frequent value | `=MODE(A2:A10)` |
| `STDEV.P()` | Population standard deviation | `=STDEV.P(A2:A10)` |
| `PERCENTILE()` | Return a value at a given percentile | `=PERCENTILE(A2:A10,0.9)` |

---

### 7. Conditional Aggregation Functions

Built multi-condition summaries and KPI scorecards to drive segmented business performance reporting.

| Function | Purpose | Example |
|----------|---------|---------|
| `SUMIF()` | Sum values based on one condition | `=SUMIF(A:A,"East",B:B)` |
| `SUMIFS()` | Sum values based on multiple conditions | `=SUMIFS(C:C,A:A,"East",B:B,"Laptop")` |
| `COUNTIF()` | Count cells matching one condition | `=COUNTIF(A:A,"Yes")` |
| `COUNTIFS()` | Count cells matching multiple conditions | `=COUNTIFS(A:A,"East",B:B,">100")` |
| `AVERAGEIF()` | Average values based on one condition | `=AVERAGEIF(A:A,"East",B:B)` |
| `AVERAGEIFS()` | Average values based on multiple conditions | `=AVERAGEIFS(C:C,A:A,"East")` |

---

### 8. Dynamic Array Functions

Leveraged Excel's modern dynamic array engine to build self-updating, spill-based analytical outputs without manual range management.

| Function | Purpose | Example |
|----------|---------|---------|
| `FILTER()` | Return rows that meet a condition | `=FILTER(A2:C100,B2:B100="East")` |
| `SORT()` | Sort a range dynamically | `=SORT(A2:C20,2,1)` |
| `SORTBY()` | Sort a range using another column as key | `=SORTBY(A2:C20,B2:B20)` |
| `UNIQUE()` | Extract distinct values from a range | `=UNIQUE(A2:A100)` |
| `SEQUENCE()` | Generate a sequence of numbers | `=SEQUENCE(10)` |
| `TAKE()` | Extract the first N rows from a range | `=TAKE(A2:C20,5)` |
| `DROP()` | Remove the first N rows from a range | `=DROP(A2:C20,2)` |

---

### 9. Advanced Formula Engineering

Engineered sophisticated, performance-optimized formula architectures using Excel's most powerful modern capabilities.

| Function | Purpose | Example |
|----------|---------|---------|
| `LET()` | Define named variables inside a formula | `=LET(x,A2*10,x+5)` |
| `LAMBDA()` | Create reusable custom functions | `=LAMBDA(x,x*2)(5)` |
| `MAP()` | Apply a function across each element of an array | `=MAP(A2:A10,LAMBDA(x,x*2))` |
| `BYROW()` | Perform row-wise calculations across a range | `=BYROW(A2:C10,LAMBDA(r,SUM(r)))` |
| `BYCOL()` | Perform column-wise calculations across a range | `=BYCOL(A2:C10,LAMBDA(c,AVERAGE(c)))` |
| `SCAN()` | Return running/cumulative calculations | `=SCAN(0,A2:A10,LAMBDA(a,b,a+b))` |
| `REDUCE()` | Collapse an array into a single aggregated value | `=REDUCE(0,A2:A10,LAMBDA(a,b,a+b))` |
| `MAKEARRAY()` | Generate a dynamic array from row/column logic | `=MAKEARRAY(3,3,LAMBDA(r,c,r*c))` |
| `OFFSET()` | Create a dynamic reference offset from a starting cell | `=SUM(OFFSET(A1,1,1,3,1))` |
| `INDIRECT()` | Build dynamic cell or range references from text | `=INDIRECT("A1")` |
| `CHOOSECOLS()` | Select specific columns from a range | `=CHOOSECOLS(A1:D10,1,3)` |
| `CHOOSEROWS()` | Select specific rows from a range | `=CHOOSEROWS(A1:D10,1,3)` |
| `WRAPROWS()` | Reshape a range into rows of a fixed width | `=WRAPROWS(A1:A10,2)` |
| `WRAPCOLS()` | Reshape a range into columns of a fixed height | `=WRAPCOLS(A1:A10,2)` |

---

### 10. Power Query

Mastered Excel's built-in ETL (Extract, Transform, Load) engine to automate data preparation and cleaning pipelines at scale — eliminating manual data wrangling.

| Concept | Description |
|---------|-------------|
| **Remove Duplicates** | Identify and eliminate repeated records from datasets |
| **Merge Queries** | Combine tables horizontally using join logic (similar to SQL JOINs) |
| **Append Queries** | Stack multiple tables vertically into a single unified dataset |
| **Split Columns** | Break a single column into multiple columns by delimiter or position |
| **Unpivot Columns** | Transform wide-format data into long-format for analysis |
| **Pivot Columns** | Reshape long-format data into a cross-tabulated summary |
| **Change Data Types** | Enforce correct data types (text, number, date, boolean) across columns |
| **Fill Down / Fill Up** | Propagate values into blank cells in a column |
| **Group By** | Aggregate rows by category with SUM, COUNT, AVERAGE, and more |
| **Conditional Columns** | Add calculated columns based on IF/ELSE business logic |
| **Data Profiling** | Analyze column quality, distribution, and error rates |
| **Refresh Automation** | Configure queries to refresh on demand or on a schedule |

---

### 11. Data Modeling

Designed professional relational data models in Excel's Power Pivot to support scalable, multi-table analytical solutions.

| Concept | Description |
|---------|-------------|
| **Relationships** | Define connections between tables using matching key columns |
| **Star Schema** | Central fact table surrounded by dimension tables — optimal for analytics |
| **Snowflake Schema** | Normalized extension of the star schema with linked dimension tables |
| **Fact Tables** | Tables containing measurable business events (sales, transactions) |
| **Dimension Tables** | Descriptive reference tables (products, customers, dates) |
| **Primary Keys** | Unique identifiers for each row in a dimension table |
| **Foreign Keys** | Columns in a fact table that reference a dimension table's primary key |
| **Calendar Tables** | Dedicated date dimension tables enabling time intelligence calculations |
| **Measures vs. Calculated Columns** | Measures aggregate dynamically in context; calculated columns compute row-by-row at load time |

---

### 12. DAX Functions

Wrote DAX (Data Analysis Expressions) to power dynamic, context-aware calculations in Excel's Power Pivot environment.

| Function | Purpose |
|----------|---------|
| `CALCULATE()` | Override or extend the current filter context |
| `SUMX()` | Iterate over table rows and sum an expression |
| `RELATED()` | Retrieve a value from a related dimension table |
| `RELATEDTABLE()` | Return all related rows from a linked table |
| `FILTER()` | Return a filtered version of a table |
| `ALL()` | Remove all filters from a table or column |
| `DIVIDE()` | Perform safe division with a fallback for divide-by-zero |
| `TOTALYTD()` | Calculate year-to-date totals using time intelligence |
| `RANKX()` | Rank items dynamically within a filtered context |
| `COUNTROWS()` | Count the number of rows in a table or filtered result |

---

### 13. Dashboard Engineering

Designed executive-grade, interactive dashboards in Excel that communicate business performance clearly and support data-driven decision-making.

| Concept | Description |
|---------|-------------|
| **KPI Design** | Define, calculate, and visually represent key performance indicators |
| **Executive Dashboards** | High-level summary views tailored for leadership and stakeholders |
| **Interactive Dashboards** | Dynamic reports using Slicers, Timelines, and drop-down controls |
| **Business Storytelling** | Structure data narratives to guide decisions, not just display numbers |
| **Visual Hierarchy** | Organize chart layouts to draw attention to the most critical insights |
| **UX Principles for Reports** | Apply layout, color theory, and whitespace for maximum clarity |

---

### 14. Statistical Analysis in Excel

Applied statistical techniques natively in Excel to uncover trends, validate assumptions, and build forecasting outputs for business use.

| Technique | Description |
|-----------|-------------|
| **Correlation** | Measure the strength of relationships between two variables |
| **Regression** | Model and predict outcomes based on historical data |
| **Hypothesis Testing** | Validate business assumptions using statistical significance |
| **Forecasting** | Project future values from historical trends using Excel's built-in tools |
| **Trend Analysis** | Identify directional patterns over time in time-series data |
| **Z-Scores** | Detect outliers and anomalies by measuring standard deviations from the mean |
| **Moving Averages** | Smooth time-series volatility to reveal underlying trends |
| **Variance Analysis** | Compare actual performance against budgets, targets, or benchmarks |

---

### 15. Automation in Excel

Built automated, low-maintenance reporting systems within Excel to eliminate repetitive manual work and enable scalable workflows.

| Tool / Concept | Description |
|----------------|-------------|
| **VBA (Visual Basic for Applications)** | Write macros to automate complex, multi-step Excel tasks programmatically |
| **Office Scripts** | Cloud-compatible automation scripts for Excel Online and Microsoft 365 |
| **Power Automate** | Connect Excel to other Microsoft 365 apps for no-code workflow automation |
| **Macro Recording** | Capture and replay repetitive Excel actions without writing code |
| **Dynamic Refresh Workflows** | Configure data connections to update automatically on a schedule or trigger |

---

## Learning Roadmap

Skills in this repository were developed and applied across a structured, phase-based progression:

```
Phase 1 — Core Excel Functions
          IF · XLOOKUP · SUMIFS · COUNTIFS · TRIM · CLEAN · IFERROR

Phase 2 — Dynamic Arrays & PivotTables
          FILTER · SORT · UNIQUE · SEQUENCE · Pivot Table design

Phase 3 — Power Query & Dashboard Design
          ETL automation · Interactive dashboards · KPI reporting

Phase 4 — Data Modeling & DAX
          Star schema · Power Pivot · CALCULATE · Time intelligence

Phase 5 — Advanced BI & Excel Automation
          VBA · Office Scripts · Power Automate · Statistical analysis
```

---

## Professional Competencies

Projects in this repository demonstrate the following professional-level Excel capabilities:

- ✅ Cleaning and transforming messy, real-world datasets efficiently using Excel-native tools
- ✅ Building automated, scalable reporting workflows with Power Query and VBA
- ✅ Designing normalized, relational data models in Power Pivot
- ✅ Creating interactive, executive-ready dashboards with slicers and dynamic visuals
- ✅ Writing advanced Excel formulas and production-grade DAX measures
- ✅ Applying statistical analysis techniques directly within Excel
- ✅ Delivering clear, actionable business insights from raw data
- ✅ Interpreting and communicating business KPIs and performance metrics

---

<div align="center">

*Built entirely in Microsoft Excel — formula by formula, model by model.*

</div>
