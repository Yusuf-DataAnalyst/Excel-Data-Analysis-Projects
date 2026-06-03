# 📊 Advanced Excel Data Analysis Portfolio

This repository serves as a comprehensive collection of my data analysis projects, where I leverage Microsoft Excel's advanced features to derive actionable insights. My focus is not just on data visualization, but on the entire data pipeline: from raw data extraction and cleaning to building professional, automated, and interactive business intelligence dashboards.

---

## 🚀 Featured Projects

### 1️⃣ World Population Analysis Dashboard 🌍
* **Project Overview:** A sophisticated demographic analytics project. I transformed raw population datasets into a dynamic decision-support dashboard to explore growth patterns across continents and track population shifts over time.
* **Analytical Workflow:**
  * **Data Enrichment:** Utilized **XLOOKUP** to perform complex cross-referencing and **IF-Logic** to categorize nations into "Large" or "Small" based on population density.
  * **Aggregated Insights:** Used **SUMIF & COUNTIF** to dynamically calculate regional contributions to total global figures.
  * **Statistical Modeling:** Leveraged the **MAX** function to identify demographic outliers and highest-growth clusters, providing a strategic view of global demographic trends.
* **Impact:** This dashboard enables users to quickly compare continental growth rates and visualize demographic distributions through interactive charts and pie visualizers.
* **Project Preview:**
![Population Dashboard](population.png)
![Population Data](population2.png)

---

### 2️⃣ Zomato vs Swiggy Sales Performance Analysis 🛒
* **Project Overview:** A deep-dive comparative business study. By processing transactional data from two industry giants, I uncovered critical performance metrics and customer purchasing behaviors that drive competitive advantage in the food-delivery sector.
* **Analytical Workflow:**
  * **Business Intelligence:** Built custom **Pivot Tables** to segment sales by region and platform.
  * **Trend Identification:** Applied **Conditional Formatting** to visually flag "Hotspots" (top-performing sales regions), which helps in optimizing marketing and delivery resource allocation.
* **Key Findings:** Successfully identified which platform leads in specific demographic segments, providing a blueprint for potential business scaling.
* **Project Preview:**
![Zwiggy vs Zomato Preview](Zwiggy_vs_Zomato.png)

---

### 3️⃣ Netflix Data Cleaning & Quality Assurance 🎬
* **Project Overview:** A hands-on project dedicated to **Data Integrity**. Raw datasets are rarely "analysis-ready"; I applied professional cleaning protocols to transform unstructured data into a high-quality format suitable for reporting.
* **Analytical Workflow:**
  * **Consistency Protocol:** Automated the identification of null values in critical fields like `director` and `cast`, replacing them with "Not Given" to ensure complete data coverage.
  * **Redundancy Management:** Implemented systematic deduplication strategies using `show_id` to ensure no overlapping data points skew final results.
* **Impact:** Established a clean, robust, and reliable dataset that guarantees high-precision results for any future analysis or reporting.
* **Project Preview:**
![Netflix Preview](netflix.png)

---

### 4️⃣ Medical Analytics & Healthcare Dashboard 🏥
* **Project Overview:** A specialized project focused on clinical data management and healthcare performance metrics. I structured complex, raw medical tracking records to build an insightful, interactive dashboard that visualizes patient data and clinical performance trends.
* **Analytical Workflow:**
  * **Data Structuring:** Cleaned and organized patient records to ensure consistency in tracking clinical metrics and patient outcomes.
  * **Dashboard Architecture:** Designed an intuitive, user-friendly interface that aggregates key healthcare indicators, allowing for quick assessment of clinical performance.
  * **Automated Trend Tracking:** Implemented automated calculations and visual elements to monitor health metric distribution and operational efficiency over time.
* **Impact:** This dashboard transforms fragmented medical records into a clear strategic overview, assisting in the identification of trends in patient care and clinical performance.
* **Project Preview:**
![Medical Dashboard Preview 1](medical.png.png)
![Medical Dashboard Preview 2](medical2.png.png)

---
---

### 5️⃣ Global Freelancers Data Cleaning & Analysis 🌍
* **📌 Project Overview:** This project demonstrates the end-to-end process of transforming a raw, unformatted single-column CSV dataset of global freelancers into a clean, structured, and analyzed Excel dashboard.
* **🛠️ Tools Used:** Microsoft Excel (Text-to-Columns, Advanced Formulas, Pivot Tables, Pivot Charts).
* **🧼 Key Data Cleaning Steps:**
  * **Data Parsing:** Split the raw comma-delimited text into 13 structured columns.
  * **Text Standardization:** Uniformed inconsistent gender values (e.g., `female`, `FEMALF`, `m`) into `Male` and `Female`.
  * **Currency Cleaning:** Removed text prefixes (`$`, `USD`) from hourly rates to enable mathematical calculations.
  * **Handling Missing Values:** Imputed null values in `rating`, `age`, and `client_satisfaction` using statistical averages.
  * **Boolean Uniformity:** Converted mixed indicators (`yes`, `1`, `0`) into standard `TRUE` / `FALSE` values.
* **📊 Analysis & Insights:**
  * **Feature Engineering:** Grouped freelancers into `Junior`, `Mid-Level`, and `Senior` based on years of experience.
  * **Pivot Tables:** Analyzed the average hourly rate per skill and freelancer distribution by country.
  * **Data Visualization:** Created Pivot Charts to visually capture market trends and high-paying tech domains.
* **📸 Project Previews:**
![Global Freelancers Preview 1](Global_Freelance.png)
![Global Freelancers Preview 2](Global_Freelance2.png)



------- 



6️⃣   Financial Sales Performance Dashboard
An interactive Excel Data Analysis project focused on cleaning, transforming, and analyzing global sales data to extract actionable business insights.
## 🚀 Project Overview
This project takes a raw, unformatted financial dataset and transforms it into a dynamic, production-ready analytics dashboard. The primary goal was to handle typical real-world data quality issues and build an interactive reporting system for executive decision-making. 

## 🛠️ Key Analytical Steps
### 1. Data Cleaning & Engineering
* **Text-to-Columns Transformation:** Handled raw CSV parsing issues by restructuring data from a single string column into a proper relational table layout.
* **Data Type Rectification:** Resolved strict formatting anomalies (e.g., `#VALUE!` errors) by stripping hidden string characters and converting metrics into calculated float/integer fields.
* **Feature Engineering:** Developed advanced logical structures using nested conditional logic:
  $$\text{Order Size} = f(\text{Units Sold})$$
  Categorized as **Small**, **Medium**, or **Large** based on distribution bounds to enhance operational filtering.
* **Financial Modeling:** Added core custom metrics including **Profit Margin %** calculations:
  $$\text{Profit Margin} = \frac{\text{Profit}}{\text{Sales}}$$

### 2. Aggregation & Pivot Tables
* Restructured over 700 operational rows into multidimensional summaries using Excel Pivot Tables.
* Switched aggregation types from simple counts to **Summatons (SUM)** to track multi-million dollar performance across geographical and sectoral dimensions.

### 3. Dynamic Visualization & UX
* Developed a high-impact visual interface optimized for clarity.
* **Geographical Matrix:** Utilized a distribution Pie Chart to break down global profitability.
* **Product-Segment Matrix:** Implemented a clustered Column Chart to track product performance across distinct market categories.
* **Interactive Slicers:** Embedded synchronized timeline filters (`Month Name` & `Year`) for seamless, on-the-fly interactive filtering.
* 
## 📊 Key Executive Insights
* **The Growth Engine:** The **Government** sector represents the highest volume driver for the organization, showing exponential traction when paired with the **Paseo** product line.
* **Geographical Balance:** Profitability is highly synchronized across international borders, with **France** and **Germany** securing a marginal lead in absolute net returns.
* **Product Optimization:** While **Paseo** dominates the charts as a top-tier seller, **Carretera** exhibits prolonged stagnation across multiple segments, indicating a critical need for marketing re-evaluation or pricing adjustments.
## 🧰 Tech Stack Used
* **Microsoft Excel:** Advanced Formulas (`IF`, `VALUE`, `SUBSTITUTE`), Pivot Tables, Pivot Charts, Slicers, Data Cleaning Tools.
* * **📸 Project Previews:**
![financial preview](financial1.png)
![financial preview](financial2.png)





## 🛠️ Skills & Technologies
* **Advanced Excel Functions:** XLOOKUP, SUMIF, COUNTIF, MAX, IF-Logic, Nested Functions.
* **Business Intelligence:** Pivot Table Architecture, Dynamic Dashboarding, Trend Analysis.
* **Data Methodology:** Structured Data Cleaning, Quality Assurance, and Business Metric Visualization.
