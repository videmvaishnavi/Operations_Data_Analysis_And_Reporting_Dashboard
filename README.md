# 📊 Operations Data Analysis & Reporting Dashboard

## 📌 Project Overview

The **Operations Data Analysis & Reporting Dashboard** is an end-to-end data analytics portfolio project developed using **Microsoft Excel, MySQL, SQL, and Microsoft Power BI**.

The project analyzes operational task data to understand task completion, pending workload, turnaround time (TAT), delayed tasks, and departmental performance. It demonstrates how raw data can be cleaned, validated, analyzed, and transformed into interactive dashboards that support data-driven decision-making.

**Project Type:** Personal Data Analytics Project  
**Domain:** Operations Analytics  
**Tools:** Excel, MySQL, SQL, Power BI, DAX  
**Dataset Size:** 220 unique task records and 12 employee records

---

## 🎯 Project Objectives

- Clean and validate operational task data.
- Identify and remove duplicate records.
- Standardize inconsistent text values.
- Validate assigned and completed dates.
- Calculate actual turnaround time and identify delayed tasks.
- Analyze completed and pending tasks across departments.
- Evaluate on-time performance and average turnaround time.
- Use SQL to perform structured data analysis.
- Develop interactive Power BI dashboards.
- Present key performance indicators (KPIs) and business insights.

---

## 🛠️ Tools and Technologies

| Tool | Purpose |
|---|---|
| Microsoft Excel | Data cleaning, validation, formulas, PivotTables, and dashboard creation |
| MySQL | Database creation, storage, and data retrieval |
| SQL | Filtering, aggregation, joins, KPI calculations, and operational analysis |
| Power BI Desktop | Interactive dashboards and data visualization |
| DAX | Calculating measures and performance indicators |
| GitHub | Project documentation, version control, and portfolio presentation |

---

## 📂 Repository Structure

```text
OPERATIONS-DATA-ANALYSIS/
│
├── Employee_master.csv
├── Employee_master.xlsx
│
├── Operations_Data_Analysis_Project_Dataset.xlsx
├── operations_tasks_all.csv
│
├── operations_db.pbix
│
└── op_db_tasks.sql
```

### 📄 File Descriptions

| File Name | Description |
|---|---|
| `Employee_master.csv` | Employee master data in CSV format, used for database import. |
| `Employee_master.xlsx` | Excel version of the employee master dataset. |
| `Operations_Data_Analysis_Project_Dataset.xlsx` | Main Excel workbook containing the task data, data cleaning, formulas, PivotTables, and Excel dashboard. |
| `operations_tasks_all.csv` | Task dataset in CSV format, used to import the 220 task records into MySQL. |
| `operations_db.pbix` | Power BI report containing the interactive Operations Performance Dashboard and Detailed Operational Analysis pages. |
| `op_db_tasks.sql` | SQL script containing database-related SQL and operational analysis queries. |

---

## 📁 Dataset Description

The project uses two related datasets.

### 1. Task Dataset

The task dataset contains operational records with the following fields:

| Column |            |Description |
|---|---|
| Task_ID |           Unique identifier for each task |
| Employee_ID |       Identifier of the employee assigned to a task |
| Employee_Name |     Name of the assigned employee |
| Department |        Department responsible for the task |
| Task_Type |         Category of the task |
| Priority |          Task priority: High, Medium, or Low |
| Assigned_Date |     Date when the task was assigned |
| Completed_Date |    Date when the task was completed; blank for pending tasks |
| Status |            Current task status: Completed or Pending |
| Expected_TAT_Days | Expected number of days to complete the task |
| Actual_TAT_Days |   Actual turnaround time in days |

### 2. Employee Master Dataset

The employee master dataset contains employee IDs, employee names, and departments. It supports employee-level analysis and SQL joins with the task dataset.
---

## 🧹 Data Cleaning and Preparation in Excel

The initial task dataset contained 222 rows, including two duplicate records. After removing duplicates, the dataset contained **220 unique task records**.

The following steps were performed:

1. **Duplicate removal:** Removed duplicate task records using Excel's Remove Duplicates feature.
2. **Text standardization:** Corrected inconsistent department and status labels and cleaned priority values using `TRIM()` and `PROPER()`.
3. **Data validation:** Checked required fields to identify missing values.
4. **Date validation:** Verified that completed dates were not earlier than assigned dates.
5. **Turnaround-time calculation:** Calculated task duration using assigned and completed dates.
6. **TAT validation:** Compared calculated turnaround time with the recorded actual TAT.
7. **Task classification:** Classified completed tasks as on time or delayed and retained pending tasks as a separate category.
8. **KPI calculation:** Created formulas to calculate task counts, completion rate, average TAT, and on-time rate.

### Key Excel Formulas

**Calculate actual TAT**

```excel
=IF(H2="","",H2-G2)
```
Returns a blank for tasks without a completion date; otherwise, subtracts the assigned date from the completed date.


**Validate actual TAT**

```excel
=IF(I2="Pending","Pending",IF(K2=M2,"Match","Mismatch"))
```
Checks whether the recorded actual TAT matches the calculated TAT for completed tasks.


**Classify task performance**

```excel
=IF(I2="Pending","Pending",IF(M2<=J2,"On Time","Delayed"))
```
Classifies a completed task as on time when its actual TAT is less than or equal to the expected TAT.


**Calculate completion rate**

```excel
=P3/P2
```
Divides completed tasks by total tasks. Format the result as a percentage.

---

## 🗄️ SQL Analysis Using MySQL

The cleaned dataset was imported into MySQL for structured analysis. SQL queries were used to summarize task performance, investigate delays, and combine task records with employee information.

### SQL Concepts Used

- `SELECT` — selects columns to retrieve.
- `WHERE` — filters records based on conditions.
- `GROUP BY` — groups records for aggregate calculations.
- `ORDER BY` — sorts results.
- `COUNT()` — counts records.
- `AVG()` — calculates average turnaround time.
- `SUM()` and `CASE` — calculate conditional task counts.
- `JOIN` — combines task and employee data.
- `ROUND()` — rounds numeric results.
- `LIMIT` — restricts the number of returned rows.

### Example SQL Queries

**Total number of tasks**

```sql
SELECT COUNT(*) AS total_tasks
FROM tasks;
```


**Completed and pending tasks**

```sql
SELECT
    Status,
    COUNT(*) AS task_count
FROM tasks
GROUP BY Status;
```


**Tasks by department**

```sql
SELECT
    Department,
    COUNT(*) AS task_count
FROM tasks
GROUP BY Department
ORDER BY task_count DESC;
```


**Delayed tasks by priority**

```sql
SELECT
    Priority,
    COUNT(*) AS delayed_tasks
FROM tasks
WHERE Status = 'Completed'
  AND Actual_TAT_Days > Expected_TAT_Days
GROUP BY Priority
ORDER BY delayed_tasks DESC;
```


**Average TAT by department**

```sql
SELECT
    Department,
    ROUND(AVG(Actual_TAT_Days), 2) AS average_tat
FROM tasks
WHERE Status = 'Completed'
GROUP BY Department
ORDER BY average_tat DESC;
```


**Completion rate**

```sql
SELECT
    COUNT(*) AS total_tasks,
    SUM(
        CASE WHEN Status = 'Completed'
        THEN 1 ELSE 0 END
    ) AS completed_tasks,
    SUM(
        CASE WHEN Status = 'Pending'
        THEN 1 ELSE 0 END
    ) AS pending_tasks,
    ROUND(
        SUM(CASE WHEN Status = 'Completed'
            THEN 1 ELSE 0 END) * 100.0
        / COUNT(*),
        2
    ) AS completion_rate
FROM tasks;
```

The SQL analysis was used to validate the Excel findings and generate department-level and employee-level summaries.

---

## 📈 Power BI Dashboard

The final Power BI report contains two pages designed to provide both an overview of operational performance and a more detailed departmental analysis.

### Page 1: Operations Performance Dashboard

The overview page includes:

**KPI cards**
- Total Tasks
- Completed Tasks
- Pending Tasks
- On-Time Tasks
- Average TAT
- On-Time Rate
- Completion Rate
- Delayed Tasks

**Visualizations**
- Total Tasks by Department
- Completed vs. Pending Tasks by Department
- Delayed Tasks by Priority
- Delayed Tasks by Department

**Interactive filters**
- Department
- Priority
- Status

These filters allow users to investigate operational performance for specific departments, task priorities, and task statuses.

### Page 2: Detailed Operational Analysis

The detailed page includes:

- Average TAT by Department
- Delayed, On-Time, and Pending Tasks by Department

The report uses synchronized slicer behavior so that users can explore the detailed page while retaining their selected filters.

### DAX Measures

**Total Tasks**

```dax
Total Tasks =
COUNTROWS(op_db_tasks)
```
Counts the task records in the Power BI table.


**Completed Tasks**

```dax
Completed Tasks =
CALCULATE(
    COUNTROWS(op_db_tasks),
    op_db_tasks[Status] = "Completed"
)
```
Counts only records whose status is Completed.


**Pending Tasks**

```dax
Pending Tasks =
CALCULATE(
    COUNTROWS(op_db_tasks),
    op_db_tasks[Status] = "Pending"
)
```
Counts records whose status is Pending.


**Completion Rate**

```dax
Completion Rate =
DIVIDE(
    [Completed Tasks],
    [Total Tasks],
    0
)
```
Divides completed tasks by total tasks. Format this measure as a percentage.


**Average TAT**

```dax
Average TAT =
AVERAGE(op_db_tasks[Actual_TAT_Days])
```
Calculates the average actual turnaround time over records with valid numeric TAT values.


**On-Time Tasks**

```dax
On-Time Tasks =
CALCULATE(
    COUNTROWS(op_db_tasks),
    op_db_tasks[Status] = "Completed",
    op_db_tasks[Actual_TAT_Days]
        <= op_db_tasks[Expected_TAT_Days]
)
```
Counts completed tasks whose actual TAT does not exceed expected TAT.


**On-Time Rate**

```dax
On-Time Rate =
DIVIDE(
    [On-Time Tasks],
    [Completed Tasks],
    0
)
```
Calculates the proportion of completed tasks finished on time.


**Delayed Tasks**

```dax
Delayed Tasks =
CALCULATE(
    COUNTROWS(op_db_tasks),
    op_db_tasks[Status] = "Completed",
    op_db_tasks[Actual_TAT_Days]
        > op_db_tasks[Expected_TAT_Days]
)
```
Counts completed tasks that exceeded their expected turnaround time.

---

## 🔑 Key Findings

| KPI | Result |
|---|---:|
| Total Tasks | 220 |
| Completed Tasks | 180 |
| Pending Tasks | 40 |
| Completion Rate | 81.82% |
| On-Time Tasks | 124 |
| Delayed Tasks | 56 |
| On-Time Rate | 68.89% |
| Average Actual TAT | 3.28 days |

### Department-Level Findings

- **Operations** handled the highest number of tasks, with 63.
- **Operations** also recorded the highest number of delayed tasks, at 15.
- **Quality** had the highest average TAT, at approximately 3.57 days.
- **Finance** had the lowest pending workload, with four pending tasks.
- **Medium-priority tasks** accounted for 22 delayed tasks, compared with 21 High-priority and 13 Low-priority tasks.

These findings describe the observed dataset. They do not, by themselves, establish why delays occurred or prove that one department performed poorly.

---

## 💡 Business Insights

The analysis provides several areas for further operational investigation:

1. **Pending workload:** The 40 pending tasks can be reviewed to determine their age, urgency, and outstanding dependencies.
2. **Delay analysis:** The 56 delayed tasks can be examined by task type, priority, department, and employee to identify recurring patterns.
3. **Departmental turnaround time:** Departments with higher average TAT can be investigated further to understand differences in task complexity and workflow.
4. **Priority management:** Delays across all three priority levels suggest that priority alone may not explain turnaround-time performance.
5. **Performance monitoring:** The dashboard provides a repeatable way to monitor task completion, backlog, and on-time performance as new data becomes available.

---

## 🚀 How to Run the Project

### Excel
1. Download the Excel workbook from the `Excel/` folder.
2. Open it in Microsoft Excel.
3. Review the cleaned task data, formulas, PivotTables, and dashboard.

### MySQL
1. Install MySQL Server and MySQL Workbench.
2. Create the `operations_db` database.
3. Import the task and employee datasets.
4. Run the SQL scripts from the `SQL/` folder.
5. Create the `tasks_powerbi` view if you want to connect Power BI to the typed dataset.

### Power BI
1. Open the `.pbix` file from the `PowerBI/` folder.
2. Configure the MySQL connection if prompted.
3. Refresh the data when the database is available.
4. Use the slicers to explore task performance by department, priority, and status.

---

## 📚 Skills Demonstrated

- Data cleaning and preprocessing
- Data quality validation
- Excel formulas and PivotTables
- Relational database fundamentals
- SQL querying and aggregation
- Conditional calculations using `CASE`
- Joining related datasets
- KPI definition and calculation
- DAX measures
- Interactive dashboard design
- Data interpretation and business reporting
- Technical documentation

---

## 📌 Project Limitations

- The analysis describes the dataset provided and may not represent the performance of a real organization.
- The dashboard is based on a fixed dataset unless new records are imported and the report is refreshed.
- Employee workload counts indicate assigned task volume, not necessarily productivity, effort, or quality.
- Further investigation is required to establish the causes of delays.

---

## 👩‍💻 About This Project

This is a personal portfolio project created to demonstrate practical data analytics skills using Excel, MySQL, SQL, and Power BI. It showcases the complete workflow from raw data preparation to analysis, visualization, and business reporting.

**Author:** Vaishnavi Videm  
**GitHub:** https://github.com/videmvaishnavi
**LinkedIn:** https://www.linkedin.com/in/vaishnavi-videm/

---

⭐ If you find this project useful, feel free to explore the repository and review the dashboard screenshots and SQL analysis.

