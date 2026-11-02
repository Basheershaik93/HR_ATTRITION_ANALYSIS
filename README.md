# HR Attrition Analysis in Excel

**Who is leaving, why are they leaving, what does it cost, and which retention actions are worth the money?**

An end-to-end Excel project on the IBM HR Analytics dataset (1,470 employees). It uses conditional formatting to spot leavers and risk, a rule-based risk-scoring model to find who might leave next, and what-if analysis (scenarios, sensitivity tables, Goal Seek) to estimate the ROI of retention programmes.

![Dashboard](dashboard.jpeg)

---

## Table of Contents
1. [Why I built this](#1-why-i-built-this)
2. [Business questions](#2-business-questions)
3. [Dataset](#3-dataset)
4. [Headline findings](#4-headline-findings)
5. [Workbook structure](#5-workbook-structure)
6. [How I built it, step by step](#6-how-i-built-it-step-by-step)
7. [Conditional formatting used](#7-conditional-formatting-used)
8. [What-if analysis in detail](#8-what-if-analysis-in-detail)
9. [Recommendations](#9-recommendations)
10. [Assumptions and limitations](#10-assumptions-and-limitations)
11. [How to use this workbook](#11-how-to-use-this-workbook)
12. [Skills demonstrated](#12-skills-demonstrated)
13. [Future improvements](#13-future-improvements)
14. [Tools](#14-tools)
15. [Contact](#15-contact)

---

## 1. Why I built this
I am building my portfolio for data analyst roles. Employee attrition is a problem every company has, it is easy to explain to a non-technical audience, and it lets me practise the skills analysts use every day in Excel: cleaning and structuring data, summarising it with formulas, visualising it, and turning the numbers into a business recommendation with a cost attached.

## 2. Business questions
1. How many employees have left, and what is the attrition rate?
2. **Who** is leaving? (department, job role, age, tenure, income, overtime)
3. **Why** are they leaving? Which factors separate leavers from stayers the most?
4. What has attrition **cost** the company so far?
5. Which **current** employees look most at risk of leaving?
6. If the company ran retention programmes, how many leavers could it prevent, and **is it worth the money**?

## 3. Dataset
| Item | Detail |
|---|---|
| Name | IBM HR Analytics Employee Attrition & Performance |
| Source | Kaggle (dataset by pavansubhasht), file `WA_Fn-UseC_-HR-Employee-Attrition.csv` |
| Size | 1,470 rows x 35 columns, one row per employee |
| Nature | A fictional dataset created by IBM data scientists for analytics practice. It is not real IBM employee data. |
| Target | `Attrition` (Yes = left, No = still employed) |

Columns I used in the model: EmployeeNumber, Age, Gender, Department, JobRole, MonthlyIncome, YearsAtCompany, YearsSinceLastPromotion, DistanceFromHome, OverTime, JobSatisfaction, WorkLifeBalance, Attrition.

## 4. Headline findings
Results from the workbook (all numbers are formula-driven and update if the data changes):

| Metric | Result |
|---|---|
| Employees | **1,470** |
| Employees who left | **237** |
| Attrition rate | **16.1%** |
| Replacement cost of attrition so far (assumed 50% of annual salary) | **about 6.8 million** (same currency as the income column) |
| Current employees rated High risk | **23** |

**Why people leave**
- **Overtime is the clearest driver.** Employees working overtime leave at **30.5%**, versus **10.4%** for those who do not.
- **The first two years are the danger zone.** Employees with 0 to 2 years at the company leave at **29.8%** (company average 16.1%). Attrition falls steadily with tenure: 13.8% (3 to 5 years), 12.3% (6 to 10 years), 8.1% (over 10 years).
- **Job role matters a lot.** Sales Representatives leave at **39.8%**, while the lowest-attrition role (among roles with 50+ employees) sits at 2.5%.
- **Department:** Sales 20.6%, Human Resources 19.0%, Research & Development 13.8%.
- **Satisfaction:** employees rating their job satisfaction 1 or 2 leave at 19.7% versus 13.9% for those rating it 3 or 4. By individual rating: 22.8% (1), 16.4% (2), 16.5% (3), 11.9% (4).

**What could be done**
- In the Base scenario (assumptions in section 8), the retention programmes could prevent roughly **52 leavers a year**, giving a **net benefit of about 407 thousand** and bringing projected attrition from 16.1% down to about **12.6%**.

> Note: these findings describe a fictional dataset. They show the method, not facts about any real company.

## 5. Workbook structure
| Sheet | Purpose |
|---|---|
| `Dashboard` | KPI cards, auto-updating insight sentences, 4 charts, live Top-10 at-risk employees |
| `IBM_Raw` | The raw Kaggle data pasted exactly as downloaded (35 columns). A header check cell confirms the column order matches. |
| `Data` | Pulls the 13 columns the model needs from `IBM_Raw`, then adds calculated helper columns (bands, risk score, risk level) |
| `Analysis` | Attrition breakdowns across 11 factors, leavers-vs-stayers profile, driver ranking |
| `WhatIf` | Every assumption, the risk model weights, scenarios, ROI, projection, two sensitivity tables |

Colour convention: **blue text on yellow = an input you can change**, green = link to another sheet, black = formula.

## 6. How I built it, step by step

### Step 1: Define the questions
Before opening Excel I wrote down the six business questions in section 2. Every table and chart in the workbook answers one of them.

### Step 2: Get the data
I downloaded the CSV from Kaggle and opened it in Excel to understand it: 1,470 rows, 35 columns, no missing values, Attrition stored as Yes/No text.

### Step 3: Load the raw data (`IBM_Raw`)
I pasted the full CSV, headers included, into `IBM_Raw!A1` and did not edit it. Keeping raw data untouched means I can always trace a number back to the source. A check formula in `IBM_Raw!AK1` returns "Headers OK" when the column order matches the Kaggle file.

### Step 4: Build the working table (`Data`)
Columns A to M link to `IBM_Raw` so the model always reads from the raw sheet:
```excel
=IF(IBM_Raw!$J2="","",IBM_Raw!A2)
```
The `IF` guard keeps the sheet clean if fewer rows are pasted. The model supports up to 1,500 employees.

### Step 5: Create helper columns
Bands turn continuous numbers into groups I can compare.
```excel
Annual salary   =F2*12
Age band        =IF(B2<30,"Under 30",IF(B2<40,"30 to 39",IF(B2<50,"40 to 49","50 and above")))
Tenure band     =IF(G2<=2,"0 to 2 yrs",IF(G2<=5,"3 to 5 yrs",IF(G2<=10,"6 to 10 yrs","Over 10 yrs")))
Income band     =IF(F2<WhatIf!$C$13,"Low",IF(F2<WhatIf!$C$14,"Medium","High"))
Promotion gap   =IF(H2<=1,"0 to 1 yrs",IF(H2<=3,"2 to 3 yrs","4 or more yrs"))
Distance band   =IF(I2<=10,"0 to 10 km",IF(I2<=20,"11 to 20 km","21 or more km"))
```
The income cut-offs (3,000 and 8,500 monthly income) live on the `WhatIf` sheet so I can change them without editing formulas. They roughly mark the bottom and top quarter of this dataset.

### Step 6: Build the risk score for current employees
I wanted a simple, explainable way to flag current employees who look like past leavers. Each risk factor earns points, and the weights are inputs on `WhatIf`:

| Risk factor | Points |
|---|---|
| Works overtime | 2 |
| Job satisfaction 2 or below | 1 |
| Work-life balance 2 or below | 1 |
| No promotion for 4+ years | 1 |
| 2 years or less at the company | 1 |
| Low income band | 1 |

```excel
Risk score  =WhatIf!$C$5*(J2="Yes")+WhatIf!$C$6*(K2<=2)+WhatIf!$C$7*(L2<=2)
            +WhatIf!$C$8*(H2>=4)+WhatIf!$C$9*(G2<=2)+WhatIf!$C$10*(Q2="Low")
Risk level  =IF(M2="Yes","Left",IF(T2>=WhatIf!$C$11,"High",IF(T2>=WhatIf!$C$12,"Medium","Low")))
```
A score of 5 or more is High, 3 or more is Medium. I also added a ranking key so the dashboard can list the Top 10 without ties:
```excel
Rank key    =IF(M2="No",T2+ROW()/100000,0)
```

### Step 7: Build the analysis tables (`Analysis`)
For each of 11 factors (department, job role, overtime, job satisfaction, work-life balance, age band, tenure band, income band, promotion gap, distance, gender) I built the same table:
```excel
Headcount   =COUNTIFS(Data!$D$2:$D$1501,B29)
Leavers     =COUNTIFS(Data!$D$2:$D$1501,B29,Data!$M$2:$M$1501,"Yes")
Attrition % =IFERROR(D29/C29,0)
vs average  =E29-$C$7
Signal      =IF(E29>$C$7*1.25,"▲ Hot spot",IF(E29<$C$7*0.75,"▼ Healthy","● Near avg"))
```
A group is flagged a hot spot when its attrition is 25% higher than the company rate.

### Step 8: Profile leavers versus stayers
A second table compares averages for the two groups (age, income, tenure, years since promotion, distance, satisfaction, work-life balance, % on overtime) with `AVERAGEIFS`, plus a percentage difference. It answers "what do leavers look like?" in one glance.
```excel
=AVERAGEIFS(Data!$F$2:$F$1501,Data!$M$2:$M$1501,"Yes")
```

### Step 9: Rank the drivers
To answer "which factor matters most?" I measured the **spread**: the highest group attrition rate minus the lowest, for each factor. A factor with a large spread separates leavers from stayers more.
```excel
Highest rate  =MAXIFS(E29:E31,C29:C31,">="&$K$17)
Lowest rate   =MINIFS(E29:E31,C29:C31,">="&$K$17)
Spread        =K5-L5
Rank          =RANK(M5,$M$5:$M$15)
```
`$K$17` is a **minimum group size (50 employees)**. Without it, tiny groups produce extreme rates that distort the ranking, which I saw happen in an early version.

### Step 10: Add conditional formatting
I used it to make leavers and hot spots visible without reading numbers. Full list in section 7.

### Step 11: Build the what-if analysis
Covered in section 8.

### Step 12: Build the dashboard
- **KPI cards** link to `Analysis` cells.
- **Insight sentences** are formulas, so they update with the data, for example:
```excel
="Overtime: staff working overtime leave at "&TEXT(Analysis!E47,"0.0%")&" vs "&TEXT(Analysis!E48,"0.0%")&" for those who do not."
```
- **Four charts**: driver strength, attrition by department, by tenure band, and by job satisfaction.
- **Top-10 at-risk list**, built with `LARGE`, `MATCH` and `INDEX`:
```excel
Data row   =MATCH(LARGE(Data!$V$2:$V$1501,B48),Data!$V$2:$V$1501,0)
Emp ID     =INDEX(Data!$A$2:$A$1501,$K48)
```

### Step 13: Validate
I checked the totals against what is known about this dataset: **1,470 employees, 237 leavers, 16.1% attrition**. I also spot-checked a few `COUNTIFS` results by filtering the `Data` sheet manually and counting.

## 7. Conditional formatting used
| Where | Rule | Purpose |
|---|---|---|
| `Data` rows | Formula rule `=$M2="Yes"` shades the whole row light red | See at a glance who left |
| `Data` Risk Level | Cell-value rules: High = red, Medium = amber, Low = green, Left = grey | Risk triage |
| `Data` Risk Score | 3-colour scale (green to red) | Spot the highest scores |
| `Data` Satisfaction / Work-life balance | Traffic-light icon set | Low ratings stand out |
| `Data` OverTime | Bold orange for "Yes" | Highlight overtime |
| `Data` MonthlyIncome | Data bars | Compare income visually |
| `Analysis` Attrition % | 3-colour scale on each table | Hot spots in red |
| `Analysis` Leavers | Data bars | Size of the problem |
| `Analysis` Signal | Formula rules on ▲ / ▼ | Flag hot spots and healthy groups |
| `Analysis` Leavers-vs-stayers % difference | Diverging colour scale (blue to red) | Where leavers differ most |
| `Analysis` Driver rank | Rank 3 or better highlighted | Top drivers |
| `WhatIf` ROI / Net benefit | Data bars, red when negative, green when positive | Which programmes pay off |
| `WhatIf` Sensitivity tables | Colour scales | Best and worst cases |
| `Dashboard` Top 10 | Data bars, risk-level fills, icon set | Priority list |

## 8. What-if analysis in detail
All assumptions are inputs on the `WhatIf` sheet (blue text on yellow), so anyone can change them and see the result immediately.

**Cost assumptions**
- Replacement cost = **50% of annual salary** (an assumption; commonly quoted ranges are much wider)
- Overlap adjustment = **75%** (one employee can belong to several target groups, so combined savings are scaled down)
- Target attrition rate = **12%**

**Four retention programmes and the attrition reduction I assumed for each**

| Programme | Target group | Conservative | Base | Optimistic | Cost per targeted employee |
|---|---|---|---|---|---|
| Cap overtime / shift support | Works overtime | 10% | 20% | 35% | 1,500 |
| Engagement and manager coaching | Satisfaction 2 or below | 10% | 20% | 30% | 800 |
| Promotion / career-path review | 4+ years without promotion | 10% | 15% | 25% | 1,200 |
| Structured onboarding and mentoring | 2 years or less at company | 10% | 15% | 25% | 600 |

**Scenario selector** (`WhatIf!C18`): enter 1, 2 or 3 and every result updates.
```excel
Selected reduction  =CHOOSE($C$18,C24,D24,E24)
```

**ROI per programme**
```excel
Targeted employees   =COUNTIFS(Data!$J$2:$J$1501,"Yes",Data!$M$2:$M$1501,"No")
Baseline leavers     =COUNTIFS(Data!$J$2:$J$1501,"Yes",Data!$M$2:$M$1501,"Yes")
Leavers prevented    =D31*F24
Replacement saved    =E31*F31*$C$17        (F31 = average annual salary of leavers in the group)
Programme cost       =C31*G24
Net benefit          =G31-H31
ROI                  =IFERROR(I31/H31,0)
Total leavers saved  =SUM(E31:E34)*C19     (after overlap adjustment)
```

**Sensitivity tables.** Two formula-driven grids (they behave like Excel Data Tables):
1. Net annual savings of the overtime programme for different attrition reductions and replacement-cost percentages:
```excel
=$D$31*$B52*$F$31*C$51-$H$31
```
2. Projected attrition rate for different combinations of overtime reduction and low-satisfaction reduction.

**Goal Seek.** To find the overtime attrition reduction needed to hit the 12% target: *Data > What-If Analysis > Goal Seek*, set `WhatIf!C43` to `0.12` by changing `WhatIf!D24` (with scenario 2 selected).

## 9. Recommendations
Based on the analysis (and subject to the limitations below):
1. **Tackle overtime first.** It is the clearest and most controllable driver (30.5% vs 10.4%). Review workloads, staffing and shift patterns.
2. **Invest in the first two years.** Structured onboarding, a buddy or mentor, and a check-in at 3, 6 and 12 months. New joiners leave at almost double the company rate.
3. **Look closely at Sales Representatives.** Their attrition is far above every other role. Investigate pay structure, targets, and manager quality before assuming the cause.
4. **Use the at-risk list for conversations, not decisions.** Managers can use the Top-10 list to prioritise stay interviews, but a score is not a verdict on any individual.
5. **Re-run the what-if with the company's real costs.** The replacement cost and programme costs are my assumptions; the ranking of programmes could change with real numbers.

## 10. Assumptions and limitations
- **The data is fictional.** It was created for practice, so the findings illustrate a method.
- **Correlation is not causation.** Showing that overtime workers leave more does not prove overtime causes people to leave.
- **The risk score is a rule-based model, not a statistical one.** I chose the weights using judgement and the patterns in the analysis. I have not tested its accuracy on unseen data.
- **Cost and reduction assumptions are illustrative.** The 50% replacement cost, the percentage reductions, the programme costs and the 75% overlap factor are my assumptions, clearly labelled and editable.
- **Small groups are noisy.** That is why the driver ranking only includes groups with 50 or more employees.
- **Single snapshot.** The dataset has no dates, so I cannot analyse trends over time.
- **Currency.** The dataset does not state a currency, so I left monetary values unitless.

## 11. How to use this workbook
1. Download `HR_Attrition_Analysis.xlsx` from this repo and open it in Excel (desktop recommended; it uses `MAXIFS` and `MINIFS`, which need Excel 2019 or Microsoft 365).
2. Go to the `Dashboard` sheet for the summary.
3. Go to `WhatIf`, change any blue input cell (for example the replacement cost or the scenario number), and watch the dashboard update.
4. To use your own data, paste it into `IBM_Raw` starting at cell A1 using the same 35 columns in the same order. Check that `IBM_Raw!AK1` says "Headers OK".

## 12. Skills demonstrated
- Data cleaning and structuring in Excel (raw sheet kept separate from the working sheet)
- Formulas: `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `INDEX/MATCH`, `LARGE`, `RANK`, `CHOOSE`, `IFERROR`, `TEXT`, nested `IF`
- Conditional formatting: formula rules, colour scales, data bars, icon sets
- What-if analysis: scenario selector, sensitivity tables, Goal Seek
- Dashboard design with formula-driven insights
- Translating data into business recommendations and costing them
- Documenting assumptions and limitations honestly

## 13. Future improvements
- Rebuild the dashboard in **Power BI** or Tableau
- Query the same data with **SQL** to cross-check the Excel results
- Use **Python (logistic regression)** to test which factors predict attrition statistically, and compare it with my rule-based risk score
- Add a pivot-table version of the analysis for comparison
- Test the risk score on a hold-out sample

## 14. Tools
- Microsoft Excel
- Dataset: IBM HR Analytics Employee Attrition & Performance (Kaggle)
- I used an AI assistant (Claude by Anthropic) to help generate the workbook structure and formulas. I reviewed the formulas, checked the totals against the dataset, and wrote up the assumptions and conclusions.

## 15. Contact
**SHAIK BASHEER**
LinkedIn: [https://www.linkedin.com/in/shaik-basheer-b69478241/]
Email: [shaik.basheer.me@gmail.com]

If you have feedback or suggestions on this project, I would be glad to hear them.
