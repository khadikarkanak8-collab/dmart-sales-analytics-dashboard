# EX3: Formula Auditing & Named Ranges

---

You are working as a Junior Data Analyst in the Business Intelligence Team at DMart.
After calculating revenue metrics in EX2, your manager Rajesh observed that:

---

Management now wants a cleaner and auditable Excel model where products are categorized as:

- High Performers

- Medium Performers

- Low Performers

---

Revenue Categorization Logic:

if Revenue > 100000 then High Performer

if Revenue > 50000 then Medium Performer

else  Low Performer

---

# Functions and Formulas to be Explored:

```
| Sr No | Function / Feature | Syntax                             | Business Use Case           | Example                                               |
| ----- | ------------------ | ---------------------------------- | --------------------------- | ----------------------------------------------------- |
| 1     | IF                 | `=IF(condition,true,false)`        | Revenue categorization      | `=IF(D2>100000,"High","Low")`                         |
| 2     | Nested IF          | Multiple IF conditions             | Multi-level categorization  | `=IF(D2>100000,"High",IF(D2>50000,"Medium","Low"))`   |
| 3     | IFS                | `=IFS(condition1,result1)`         | Cleaner multiple conditions | `=IFS(D2>100000,"High",D2>50000,"Medium",TRUE,"Low")` |
| 4     | AND                | `=AND(condition1,condition2)`      | Multi-condition checks      | `=AND(D2>50000,E2>10000)`                             |
| 5     | OR                 | `=OR(condition1,condition2)`       | Flexible validations        | `=OR(D2<0,E2<0)`                                      |
| 6     | NOT                | `=NOT(condition)`                  | Reverse logic               | `=NOT(D2>50000)`                                      |
| 7     | IFERROR            | `=IFERROR(formula,value)`          | Handle formula errors       | `=IFERROR(A2/B2,0)`                                   |
| 8     | ISERROR            | `=ISERROR(value)`                  | Detect formula issues       | `=ISERROR(F2)`                                        |
| 9     | ISBLANK            | `=ISBLANK(cell)`                   | Check missing data          | `=ISBLANK(D2)`                                        |
| 10    | FORMULATEXT        | `=FORMULATEXT(cell)`               | Display formula used        | `=FORMULATEXT(F2)`                                    |
| 11    | CELL               | `=CELL(info_type,cell)`            | Retrieve cell properties    | `=CELL("address",D2)`                                 |
| 12    | ERROR.TYPE         | `=ERROR.TYPE(error)`               | Identify error category     | `=ERROR.TYPE(F2)`                                     |
| 13    | N                  | `=N(value)`                        | Convert values to numbers   | `=N(D2)`                                              |
| 14    | EXACT              | `=EXACT(text1,text2)`              | Exact text validation       | `=EXACT(A2,B2)`                                       |
| 15    | LEN                | `=LEN(text)`                       | Validate product names      | `=LEN(B2)`                                            |
| 16    | TRIM               | `=TRIM(text)`                      | Remove extra spaces         | `=TRIM(B2)`                                           |
| 17    | UPPER              | `=UPPER(text)`                     | Standardize text            | `=UPPER(B2)`                                          |
| 18    | LOWER              | `=LOWER(text)`                     | Convert lowercase           | `=LOWER(B2)`                                          |
| 19    | PROPER             | `=PROPER(text)`                    | Proper naming format        | `=PROPER(B2)`                                         |
| 20    | CONCAT             | `=CONCAT(text1,text2)`             | Create audit IDs            | `=CONCAT(A2,"-",C2)`                                  |
| 21    | TEXT               | `=TEXT(value,format)`              | Revenue formatting          | `=TEXT(D2,"₹#,##0")`                                  |
| 22    | ROUND              | `=ROUND(number,digits)`            | Financial precision         | `=ROUND(D2,2)`                                        |
| 23    | ABS                | `=ABS(number)`                     | Positive value conversion   | `=ABS(E2)`                                            |
| 24    | SUM                | `=SUM(range)`                      | Total revenue               | `=SUM(D2:D100)`                                       |
| 25    | AVERAGE            | `=AVERAGE(range)`                  | Avg revenue analysis        | `=AVERAGE(D2:D100)`                                   |
| 26    | MAX                | `=MAX(range)`                      | Highest performer           | `=MAX(D2:D100)`                                       |
| 27    | MIN                | `=MIN(range)`                      | Lowest performer            | `=MIN(D2:D100)`                                       |
| 28    | COUNTIF            | `=COUNTIF(range,criteria)`         | Count High performers       | `=COUNTIF(G:G,"High")`                                |
| 29    | SUMIF              | `=SUMIF(range,criteria,sum_range)` | Revenue by category         | `=SUMIF(G:G,"High",D:D)`                              |
| 30    | SUBTOTAL           | `=SUBTOTAL(function_num,range)`    | Filter-aware calculations   | `=SUBTOTAL(9,D2:D100)`                                |

```

---

# Name Range to Create:

```
| Named Range     | Refers To | Purpose                |
| --------------- | --------- | ---------------------- |
| Revenue         | `D2:D100` | Revenue calculations   |
| Profit          | `E2:E100` | Profit analysis        |
| ProductName     | `B2:B100` | Product lookup         |
| HighThreshold   | `100000`  | High performer logic   |
| MediumThreshold | `50000`   | Medium performer logic |

```
