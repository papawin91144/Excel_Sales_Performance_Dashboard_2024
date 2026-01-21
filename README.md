# Sales Performance Dashboard (Microsoft Excel 2024)

### Interactive Excel dashboard using PivotTables, PivotCharts, Slicers, and VBA macros to track sales performance vs. target across regions.

---

## Introduction
This project is my first end-to-end dashboard build in **Microsoft Excel 2024**.  
The goal was to turn a flat sales dataset into an **interactive performance dashboard** that quickly answers the most common business questions: *Who is selling the most? Who is underperforming? How close are we to targets?* — with simple filtering by **region/city** and clean executive-level visuals.

The final file is an **.xlsm Excel dashboard** with PivotTables, charts, slicers, and macros used for dashboard navigation.

---

## Background & Motivation
When I started learning Excel analytics, I noticed a gap between:
- knowing features (PivotTables, slicers, charts), and  
- building something that feels like a real “business dashboard”.

So I built this project to practice the full workflow:
1. structuring raw data,
2. creating calculated metrics,
3. building PivotTables and PivotCharts,
4. connecting slicers across visuals,
5. designing a clean dashboard layout,
6. adding **VBA macros** to connect dashboard views like an app-like experience.

---

## Questions This Project Answers
This dashboard is designed to answer questions like:

- **Top performers:** Who are the highest selling sales executives?
- **Low performers:** Who are the lowest selling sales executives?
- **Target achievement:** Who has the highest *Target Hit %*?
- **Target gap:** Who is most *Away From Target %*?
- **Regional view:** How do these results change when filtering by region/city?
- **Quick comparisons:** What’s the performance spread across executives at a glance?

---

## Tools & Technologies Used
- **Microsoft Excel 2024**
- **PivotTables** (Top/Bottom analysis, sorting, filters)
- **PivotCharts** (bar and line visuals)
- **Slicers** (interactive filtering across pivots)
- **VBA Macros (.xlsm)** for dashboard navigation/connection
- **GitHub** for versioning and portfolio presentation

---

## The Analysis
### Dataset Structure
The raw dataset contains:
- `Emp Code`, `Sales Executive`, `Region`
- Daily sales: `Day1` to `Day5`
- `Total Sales` (calculated)
- `Target` (constant per executive)
- `Target Hit %` (calculated)
- `Away From Target %` (calculated)

### Key Calculations
- **Total Sales** = Sum of Day1…Day5  
- **Target Hit %** = Total Sales / Target  
- **Away From Target %** = 100% − Target Hit %

### Dashboard Components
The dashboard is built using 4 PivotTables (and connected charts) that summarize:
- **Highest Selling Executives** (Top 5 by Total Sales)
- **Lowest Selling Executives** (Bottom 5 by Total Sales)
- **Target Hit % Wise** (Top performers by target achievement)
- **Away From Target % Wise** (largest gap to target)

Slicers allow filtering by region/city so the same dashboard layout instantly updates all visuals.

---

## What I Learned
- How to design PivotTables for “Top/Bottom N” reporting (and keep them readable)
- How to connect slicers to multiple PivotTables for a true dashboard experience
- How to create clean chart layouts that match the story (bars for ranking, line for trend/contrast)
- How to structure raw data so PivotTables stay stable and easy to refresh
- Basics of using **VBA macros** to make Excel dashboards feel more interactive and navigable
- Dashboard design habits: spacing, alignment, consistency, and “glanceable” KPIs

---

## Conclusions (Key Insights)
From the dashboard view (example filter), the analysis highlights:
- A clear separation between top sellers and low sellers by **Total Sales**
- Target achievement varies significantly across executives (roughly mid-40% to mid-70% in the sample)
- The “Away From Target %” view makes underperformance obvious and actionable
- Filtering by region/city quickly reveals how rankings and gaps change across locations

---

## Final Thoughts
This project was built as a practical learning milestone: not just using Excel features, but building a complete, decision-friendly dashboard.

Next improvements I want to add:
- KPI cards (Total Sales, Avg Hit %, Best/Worst executive)
- A trend view by day (Day1–Day5) by region
- Automated refresh + timestamp of last refresh
- Cleaner macro buttons + documentation of VBA modules

---

## How to Use
1. Download the `.xlsm` file from this repo.
2. Open in **Microsoft Excel (desktop)**.
3. Click **Enable Editing** and **Enable Content** (macros) if prompted.
4. Use slicers/buttons to filter by region/city and navigate dashboard views.
5. If new data is added to `Raw_Data`, refresh PivotTables to update visuals.

> Note: This is a learning/portfolio project dataset and dashboard.


---

## Dashboard Preview
![Dashboard Preview](/Dashboard_Preview.PNG)

