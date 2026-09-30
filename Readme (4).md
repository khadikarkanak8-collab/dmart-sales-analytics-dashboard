# EX2: Revenue Metrics Calculation

---

You are working as a Junior Data Analyst in the Business Intelligence Team at DMart.
After completing initial data cleaning in Sprint 1, your manager Rajesh has assigned a new sprint task.

```
Senior management wants a quick summary of:
- Total Revenue
- Total Profit
- Profit Margin
- Discount Impact
- Quantity Sold
- Category-wise Performance

```

---

# Sprint Objective

Build business KPIs using Excel formulas and prepare the dataset for dashboard reporting in future sprints.

```
Business Requirements (BRD)
Requirement ID	Business Requirement
- BR-1	Calculate Revenue for each order
- BR-2	Calculate Profit earned
- BR-3	Calculate Discount Amount
- BR-4	Calculate Profit Margin %
- BR-5	Calculate Total Sales Summary
- BR-6	Identify Top Performing Categories
- BR-7	Prepare KPI summary for dashboard
```
---

```
Functional Requirements (FRD)
Function	Description
- Revenue Calculation	Sales Price × Quantity
- Profit Calculation	Revenue − Cost
- Discount Calculation	Original Price − Discounted Price
- Margin %	Profit ÷ Revenue
- Aggregation	SUMIFS / Pivot Table
- Lookup Validation	VLOOKUP / XLOOKUP
```

# Functions and Formulas to be explore:
---

```
| Sr No | Function / Formula  | Syntax                                         | Business Use Case                   | Example                                   |
| ----- | ------------------- | ---------------------------------------------- | ----------------------------------- | ----------------------------------------- |
| 1     | SUM                 | `=SUM(range)`                                  | Calculate total revenue/profit      | `=SUM(J2:J100)`                           |
| 2     | AVERAGE             | `=AVERAGE(range)`                              | Find average sales                  | `=AVERAGE(H2:H100)`                       |
| 3     | MAX                 | `=MAX(range)`                                  | Highest revenue/product sale        | `=MAX(H2:H100)`                           |
| 4     | MIN                 | `=MIN(range)`                                  | Lowest sales value                  | `=MIN(H2:H100)`                           |
| 5     | COUNT               | `=COUNT(range)`                                | Count numeric records               | `=COUNT(A2:A100)`                         |
| 6     | COUNTA              | `=COUNTA(range)`                               | Count non-empty cells               | `=COUNTA(B2:B100)`                        |
| 7     | COUNTIF             | `=COUNTIF(range,criteria)`                     | Count category occurrences          | `=COUNTIF(C:C,"Grocery")`                 |
| 8     | SUMIF               | `=SUMIF(range,criteria,sum_range)`             | Category-wise revenue               | `=SUMIF(C:C,"Snacks",H:H)`                |
| 9     | SUMIFS              | `=SUMIFS(sum_range,criteria_range1,criteria1)` | Multi-condition revenue analysis    | `=SUMIFS(H:H,C:C,"Beverages",F:F,"West")` |
| 10    | IF                  | `=IF(condition,true,false)`                    | Profit/Loss identification          | `=IF(K2>0,"Profit","Loss")`               |
| 11    | IFS                 | `=IFS(condition1,result1,condition2,result2)`  | Multi-level grading/profit check    | `=IFS(K2>5000,"High",K2>1000,"Medium")`   |
| 12    | AND                 | `=AND(condition1,condition2)`                  | Multiple condition validation       | `=AND(K2>0,H2>1000)`                      |
| 13    | OR                  | `=OR(condition1,condition2)`                   | Flexible condition checking         | `=OR(K2<0,H2<500)`                        |
| 14    | ROUND               | `=ROUND(number,digits)`                        | Round financial values              | `=ROUND(K2,2)`                            |
| 15    | ROUNDUP             | `=ROUNDUP(number,digits)`                      | Revenue rounding up                 | `=ROUNDUP(H2,0)`                          |
| 16    | ROUNDDOWN           | `=ROUNDDOWN(number,digits)`                    | Remove decimals                     | `=ROUNDDOWN(H2,0)`                        |
| 17    | ABS                 | `=ABS(number)`                                 | Convert negative values to positive | `=ABS(K2)`                                |
| 18    | CONCAT              | `=CONCAT(text1,text2)`                         | Create product IDs                  | `=CONCAT(A2,"-",B2)`                      |
| 19    | TEXT                | `=TEXT(value,format)`                          | Format dates/revenue                | `=TEXT(A2,"MMM")`                         |
| 20    | LEFT                | `=LEFT(text,n)`                                | Extract first characters            | `=LEFT(B2,3)`                             |
| 21    | RIGHT               | `=RIGHT(text,n)`                               | Extract last characters             | `=RIGHT(B2,2)`                            |
| 22    | MID                 | `=MID(text,start,n)`                           | Extract middle text                 | `=MID(B2,2,4)`                            |
| 23    | LEN                 | `=LEN(text)`                                   | Product name length                 | `=LEN(B2)`                                |
| 24    | TRIM                | `=TRIM(text)`                                  | Remove extra spaces                 | `=TRIM(B2)`                               |
| 25    | UPPER               | `=UPPER(text)`                                 | Standardize uppercase text          | `=UPPER(B2)`                              |
| 26    | LOWER               | `=LOWER(text)`                                 | Convert to lowercase                | `=LOWER(B2)`                              |
| 27    | PROPER              | `=PROPER(text)`                                | Proper case formatting              | `=PROPER(B2)`                             |
| 28    | TODAY               | `=TODAY()`                                     | Current date                        | `=TODAY()`                                |
| 29    | NOW                 | `=NOW()`                                       | Current date & time                 | `=NOW()`                                  |
| 30    | YEAR                | `=YEAR(date)`                                  | Extract sales year                  | `=YEAR(A2)`                               |
| 31    | MONTH               | `=MONTH(date)`                                 | Monthly sales analysis              | `=MONTH(A2)`                              |
| 32    | DAY                 | `=DAY(date)`                                   | Daily sales analysis                | `=DAY(A2)`                                |
| 33    | VLOOKUP             | `=VLOOKUP(value,table,col,FALSE)`              | Fetch product/category data         | `=VLOOKUP(A2,$P$2:$R$20,2,FALSE)`         |
| 34    | XLOOKUP             | `=XLOOKUP(lookup,array,return)`                | Advanced lookup                     | `=XLOOKUP(A2,P:P,Q:Q)`                    |
| 35    | INDEX               | `=INDEX(range,row,col)`                        | Retrieve value dynamically          | `=INDEX(B2:D10,2,2)`                      |
| 36    | MATCH               | `=MATCH(value,range,0)`                        | Find position of value              | `=MATCH("Rice",B:B,0)`                    |
| 37    | INDEX + MATCH       | Combined formula                               | Dynamic lookup replacement          | `=INDEX(C:C,MATCH(A2,B:B,0))`             |
| 38    | IFERROR             | `=IFERROR(formula,value)`                      | Handle lookup errors                | `=IFERROR(VLOOKUP(...),"Not Found")`      |
| 39    | UNIQUE              | `=UNIQUE(range)`                               | Unique categories/products          | `=UNIQUE(C2:C100)`                        |
| 40    | FILTER              | `=FILTER(range,condition)`                     | Filter profitable products          | `=FILTER(A2:H100,K2:K100>0)`              |
| 41    | SORT                | `=SORT(range,column)`                          | Sort revenue data                   | `=SORT(A2:H100,5,-1)`                     |
| 42    | SUBTOTAL            | `=SUBTOTAL(function_num,range)`                | Filter-aware totals                 | `=SUBTOTAL(9,H2:H100)`                    |
| 43    | LARGE               | `=LARGE(range,k)`                              | Top revenue values                  | `=LARGE(H:H,1)`                           |
| 44    | SMALL               | `=SMALL(range,k)`                              | Lowest sales values                 | `=SMALL(H:H,1)`                           |
| 45    | PERCENTAGE Formula  | `=(part/total)*100`                            | Profit margin %                     | `=(K2/H2)*100`                            |
| 46    | Revenue Formula     | `=Qty*Price`                                   | Gross revenue                       | `=D2*E2`                                  |
| 47    | Discount Formula    | `=Revenue*Discount%`                           | Discount amount                     | `=H2*G2`                                  |
| 48    | Net Revenue Formula | `=Revenue-Discount`                            | Final sales amount                  | `=H2-I2`                                  |
| 49    | Total Cost Formula  | `=Qty*Cost`                                    | Product cost                        | `=D2*F2`                                  |
| 50    | Profit Formula      | `=Net Revenue-Total Cost`                      | Profit calculation                  | `=J2-K2`                                  |
```
---
