# 📞 PwC Call Center Analytics Dashboard

![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![DAX](https://img.shields.io/badge/DAX-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query-217346?style=for-the-badge&logo=microsoft&logoColor=white)
![Data](https://img.shields.io/badge/Data-Kaggle-20BEFF?style=for-the-badge&logo=kaggle&logoColor=white)

> A comprehensive **Microsoft Excel** analytics solution for call center performance data — built entirely in Excel using **DAX measures** and **Pivot Tables**. Covers agent performance, customer experience, operational efficiency, and resolution trends across 5,000 calls.

---

## 📌 Table of Contents

- [Project Overview](#-project-overview)
- [Dashboard Pages](#️-dashboard-pages)
- [Data Preparation](#-data-preparation)
- [DAX Measures](#-dax-measures)
- [Key Insights](#-key-insights)
- [Data Source](#-data-source)
- [How to Use](#-how-to-use)
- [Repository Structure](#-repository-structure)
- [Author](#-author)

---

## 📊 Project Overview

This project analyzes call center data from **PwC (via Kaggle)**, fully built inside **Microsoft Excel** — no Power BI, no Python. The entire pipeline from raw data to interactive dashboards lives in a single `.xlsx` file.

| Layer | Tool Used |
|-------|-----------|
| 🔄 Data Preparation | Excel — Basic transformations |
| 📐 Calculations | DAX Measures (Power Pivot) |
| 📊 Visualization | Pivot Charts, Pivot Tables & Slicers |

**Four dashboard pages cover:**

| Page | Focus |
|------|-------|
| 📋 Executive Overview | High-level KPIs, call volume trends, resolution by topic |
| 👤 Agent Performance | Per-agent metrics: calls handled, resolution rate, satisfaction, speed |
| 😊 Customer Experience | Satisfaction distribution, topic analysis, daily and monthly trends |
| ⚙️ Operational Insights | Hourly patterns, topic volumes, handle time efficiency |

**Key Numbers at a Glance:**

| Metric | Value |
|--------|-------|
| Total Calls | 5,000 |
| Total Agents | 8 |
| Answer Rate | 81.1% |
| Resolution Rate | 72.9% |
| AVG Satisfaction | 3.4 / 5 |
| AVG Speed to Answer | 67.5 Seconds |
| AVG Handle Time | 0:03:45 |
| Resolved Calls | 3,646 (73%) |
| Unanswered Calls | 946 (19%) |

---

## 🖥️ Dashboard Pages

### 1. 📋 Executive Overview

> High-level snapshot of call center health across all agents, topics, and time periods.


<img width="1887" height="858" alt="Excutive Pverview" src="https://github.com/user-attachments/assets/9fd8f4df-787e-4678-9590-bb5cd8b0767f" />



**Highlights:**

- Total calls received: **5,000** across January, February, and March
- Overall answer rate: **81.1%** — **946 calls went unanswered**
- Resolution rate: **72.9%** — **3,646 calls resolved out of 5,000**
- Call volume is **declining month-over-month**: 1,772 (Jan) → 1,616 (Feb) → 1,612 (Mar)
- **Admin Support** has the highest resolution rate at **74.1%**
- **Technical Support** has the lowest resolution rate at **72.2%**
- Peak call hours: **11AM (590)** and **17PM (566)**
- Call volume drops sharply at **18:00** — only **4 calls** in the last hour
- **Martha** leads agent satisfaction at **3.47/5**, **Joe** is lowest at **3.33/5**

---

### 2. 👤 Agent Performance

> Individual agent metrics covering workload, resolution quality, customer satisfaction, and response efficiency.


<img width="1887" height="858" alt="Excutive Pverview" src="https://github.com/user-attachments/assets/400c2bd9-9a7b-4b88-b90f-96737a230ad8" />



**Highlights:**

- **Jim** handles the most calls: **666** — highest workload of all 8 agents
- **Stewart** handles the fewest: **582** — lowest workload
- **Dan** has the best resolution rate: **74.4%**
- **Diane** has the lowest resolution rate: **71.4%**
- **Martha** scores highest in customer satisfaction: **3.47/5**
- **Joe** scores lowest in customer satisfaction: **3.33/5**
- **Diane** answers calls fastest: **0:03:39** average handle time
- **Dan** has the longest handle time: **0:03:51**
- **Becky** answers calls with lowest speed to answer: **65.3 seconds**
- **Joe** takes longest to answer: **71.0 seconds**
- Positive correlation between calls handled and satisfaction — busier agents tend to score slightly higher

---

### 3. 😊 Customer Experience

> Explores customer satisfaction patterns across topics, days of the week, and monthly trends.


<img width="1877" height="850" alt="Customer Experience" src="https://github.com/user-attachments/assets/107f01a6-9463-4cf2-94a2-d1958bc208dd" />


**Highlights:**

- Customer ratings concentrate at **3 (1,218)** and **4 (1,180)** — moderate satisfaction overall
- **946 calls have no rating** — a significant data quality opportunity
- **Admin Support** delivers the highest satisfaction: **3.43/5**
- **Contract related** topics score lowest: **3.38/5**
- **Admin Support** also has the best resolution rate: **74.1%**
- **Technical Support** resolution rate is lowest: **72.2%**
- **Monday** is the busiest day: **770 calls**
- **Tuesday** is the quietest: **675 calls**
- **Saturday** sees a spike: **768 calls** — second busiest day
- Negative correlation between waiting time and satisfaction — longer wait = lower rating
- Resolution rate improved in **March (73%)** after dipping in **February (72%)**

---

### 4. ⚙️ Operational Insights

> Operational efficiency analysis covering call volume patterns by hour, day, topic, and handle time trends.


<img width="1887" height="837" alt="Operational Insights" src="https://github.com/user-attachments/assets/e4c41105-7286-476a-a5f8-f1e74d4a96dc" />


**Highlights:**

- **Streaming** is the most common topic: **1,022 calls**
- **Admin Support** and **Contract related** are the least frequent: **976 calls** each
- All topics have nearly **equal handle times** (~0:03:45 to 0:03:48) — no major efficiency gap
- **Payment related** calls are handled fastest: **0:03:36**
- Call volume peaks at **11AM (590)** and **17PM (583)**
- **Answer rate peaks at 18:00: 85.7%** — best performance in the last active hour
- **Answer rate is lowest at 9AM: 79.7%** — morning starts slower
- **February** had the longest average handle time: **0:03:48**
- **March** dropped significantly: **0:03:40** — efficiency improved
- **Monday and Saturday** are the two peak days — workload is not evenly distributed

---

## 🧹 Data Preparation

All data preparation was performed directly in **Excel** before building the dashboards.

> ⚠️ **Note:** This project does not use a dimensional data model or Snowflake/Star Schema. Data was imported into Excel as a flat table and lightly transformed.

| # | Action Taken | Purpose |
|---|-------------|---------|
| 1 | Loaded raw CSV into Excel | Starting point for all analysis |
| 2 | Converted `AvgTalkDuration` from seconds to `mm:ss` time format | Readable handle time display |
| 3 | Extracted `Hour` from call timestamp | Enable hourly analysis |
| 4 | Extracted `Day Name` from call date | Enable day-of-week analysis |
| 5 | Standardized `Resolved` and `Answered` columns to Yes/No | Consistent filtering |
| 6 | Created helper columns for DAX measures | Enable calculated metrics |

**No missing values removed. No records dropped.**
Dataset is clean and complete at **5,000 rows**.

---

## 📐 DAX Measures

All measures are built using **Power Pivot DAX** inside Excel.

```dax
-- ─────────────────────────────────────────
-- VOLUME METRICS
-- ─────────────────────────────────────────

Total Calls =
COUNTROWS(Sheet1)

Total Answered Calls =
CALCULATE(
    COUNTROWS(Sheet1),
    Sheet1[Answered (Y/N)] = "Y"
)

Answered Calls % =
DIVIDE(
    [Total Answered Calls],
    [Total Calls],
    0
)

Total Resolved Calls =
CALCULATE(
    COUNTROWS(Sheet1),
    Sheet1[Resolved] = "Y"
)

-- ─────────────────────────────────────────
-- RESOLUTION & SATISFACTION
-- ─────────────────────────────────────────

Resolution Rate =
DIVIDE(
    [Total Resolved Calls],
    [Total Calls],
    0
)

Avg Satisfaction =
AVERAGE(Sheet1[Satisfaction rating])

-- ─────────────────────────────────────────
-- EFFICIENCY METRICS
-- ─────────────────────────────────────────

Average Handle Time (AHT) =
AVERAGE(Sheet1[AvgTalkDuration])

Average Speed of Answer (seconds) =
AVERAGE(Sheet1[Speed of answer in seconds])

Agent number =
DISTINCTCOUNT(Sheet1[Agent])
```

---

## 💡 Key Insights

### 🔴 Problems Identified

1. **946 unanswered calls (18.9%)** — nearly 1 in 5 calls is being missed
2. **946 calls with no customer rating** — significant gap in feedback data
3. **Resolution rate of 72.9%** — over 1,350 calls end without a resolution
4. **Joe scores lowest in both satisfaction (3.33) and speed (71 seconds)** — needs coaching
5. **Tuesday has consistently low volume AND lower resolution** — possible staffing mismatch
6. **Morning answer rates start at 79.7%** — team isn't fully warmed up at 9AM

### 🟢 Opportunities

1. **Dan has the best resolution rate (74.4%)** — document his approach as a best practice
2. **Martha leads satisfaction (3.47)** — pair her with underperforming agents for coaching
3. **85.7% answer rate at 18:00** — team performs best in the last hour; understand why
4. **Payment calls handle fastest (0:03:36)** — study this topic's workflow for efficiency gains
5. **Admin Support has best satisfaction AND resolution** — identify what makes it work

### 📋 Recommendations

- **Implement a coaching program** pairing top performers (Dan, Martha) with lower performers (Joe, Diane)
- **Add mandatory rating collection** at end of each call — reduce the 946 no-rating gap
- **Review Monday and Saturday staffing** — two peak days need better resource allocation
- **Set a 9AM team warm-up protocol** — structured opening to improve morning answer rates
- **Investigate Streaming topic** — highest call volume (1,022) with average resolution — room to improve
- **Set individual improvement targets** for Joe: reduce speed to answer below 68 seconds, improve resolution rate above 73%

---

## 📁 Repository Structure

```
PwC-Call-Center-Analytics-Dashboard-Excel/
│
├── 📊 PwC_Call_Center_Dashboard.xlsx        ← Main Excel file (Dashboards + DAX)
├── 📋 01_Call_Center_Dataset.csv            ← Source data from Kaggle (5,000 records)
│
├── screenshots/
│   ├── Excutive_Pverview.png
│   ├── Agent_Performance.png
│   ├── Customer_Experience.png
│   ├── Operational_Insights.png
│   └── DAX.png
│
└── README.md
```

---

## 📦 Data Source

- **Platform:** Kaggle — PwC Switzerland Power BI Job Simulation
- **File Format:** CSV
- **Total Records:** 5,000 calls
- **Period Covered:** January – March (Q1)
- **Agents:** 8 agents (Becky, Dan, Diane, Greg, Jim, Joe, Martha, Stewart)
- **Topics:** Streaming, Technical Support, Payment related, Admin Support, Contract related

---

## 🚀 How to Use

1. **Download** `PwC_Call_Center_Dashboard.xlsx` from this repository
2. **Open in Microsoft Excel** (Excel 2016 or later — requires Power Pivot)
3. **Enable Power Pivot** if prompted:
   `File → Options → Add-ins → COM Add-ins → ✅ Microsoft Power Pivot for Excel`
4. Navigate between the **4 dashboard tabs** using the sidebar buttons
5. Use the **Agent slicer** on the left panel to filter by individual agent

> ⚠️ Requires **Power Pivot** — available in Excel 2016+, Microsoft 365 (Windows). Some features may not work on Excel for Mac.

---

## 👤 Author

**Mohamed Sayed Taha**  
Data Analyst | Excel | Power BI | SQL | Python | Power Query | DAX

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/mohamed-sayed71)
[![Email](https://img.shields.io/badge/Email-D14836?style=for-the-badge&logo=gmail&logoColor=white)](mailto:mohamedsaidtaha1@gmail.com)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/Mohamed-Sayed-Taha)

---

> ⭐ If you found this project useful, please consider giving it a star!
