# Project_1-DAashboard-README.md
# Excel Salary Analysis Dashboard

## Project Overview

This project focuses on analysing salary trends across data-related job roles and countries using Microsoft Excel.

The dashboard was created as part of an Excel Data Analytics course. It uses a 2023 job-market dataset to compare salaries and present information through charts, formulas and interactive dropdowns.

The main goal was to make salary information easier to explore and understand.

## Dashboard Preview

![Excel Salary Dashboard](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/raw/main/Images/1_Salary_Dashboard_Final_Dashboard.gif)

## Dataset Information

The dataset contains information about data-related job opportunities, including:

- Job titles
- Annual salaries
- Countries
- Employment schedules
- Technical skills

## Excel Skills Used

The following Excel features were used in this project:

- **Formulas and Functions:** MEDIAN, IF, FILTER, SEARCH and ISNUMBER.
- **Charts:** Horizontal Bar Chart and Geographical Map Chart.
- **Data Validation:** Dropdown lists for job title, country and employment type.
- **Excel Tables:** Structured references and dynamic array formulas.
- **Data Organization:** Sorting, filtering and formatting.

## Dashboard Development

### 1. Salary Comparison – Bar Chart

![Salary Bar Chart](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/raw/main/Images/1_Salary_Dashboard_Chart1.png)

- Created a horizontal bar chart to compare salaries across different job roles.
- Sorted salary values in descending order to make comparisons easier.
- Added salary labels and formatting for better readability.
- Used the chart to identify differences in salary levels across job roles.

### 2. Country-wise Salary Comparison – Map Chart

![Country Salary Map](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/raw/main/Images/1_Salary_Dashboard_Country_Map.gif)

- Used Excel Map Charts to display salary information across countries.
- Applied different colours to represent salary levels.
- Displayed median salary values for the countries available in the dataset.
- Made geographical salary comparisons easier to understand.

## Formulas and Functions

### 1. Median Salary Calculation

**Formula:**

```excel
=MEDIAN(
    IF(
        (jobs[job_title_short]=A2)*
        (jobs[job_country]=country)*
        (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
        (jobs[salary_year_avg]<>0),
        jobs[salary_year_avg]
    )
)
```

**How it works:**

- Checks whether the job title matches the selected job title.
- Checks whether the country matches the selected country.
- Uses SEARCH and ISNUMBER to check the selected employment schedule.
- Excludes records where the annual salary is zero.
- Calculates the median salary from the matching records.

This formula helps calculate salary information based on multiple selections.

### 2. Filtering Job Schedule Types

**Formula:**

```excel
=FILTER(
    J2#,
    (NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*
    (J2#<>0)
)
```

**How it works:**

- Uses the spilled range operator (#) to refer to the dynamic list starting at J2.
- Uses SEARCH to identify entries containing specific text.
- Uses ISNUMBER to check whether the searched text is found.
- Uses NOT to reverse the TRUE and FALSE results.
- Uses FILTER to return entries that satisfy the conditions.

This filtered list is used to support the employment schedule dropdown.

## Data Validation

![Data Validation](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/raw/main/Images/1_Salary_Dashboard_Data_Validation.gif)

Data Validation was used to create dropdown lists for:

- Job Title
- Country
- Employment Type

**Benefits:**

- Allows users to select values instead of typing them manually.
- Reduces incorrect or inconsistent inputs.
- Makes the dashboard easier to use.
- Supports interactive salary analysis.

## Dashboard Components

![Job Title Selection](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/raw/main/Images/1_Salary_Dashboard_Job_Title.png)

![Employment Type Selection](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/raw/main/Images/1_Salary_Dashboard_Type.png)

The dashboard allows users to explore salary information using different job titles, countries and employment schedules.

## What I Learned

Through this project, I gained practical experience in:

- Using Excel formulas with multiple conditions.
- Working with Excel Tables and structured references.
- Creating dynamic lists using FILTER.
- Building charts to present data clearly.
- Creating dropdown lists using Data Validation.
- Designing interactive Excel dashboards.

## Project File

[Download/View Excel Salary Dashboard](https://github.com/NabeelRizwan/Excel-Data-Analytics-Projects/blob/main/Project_1-Dashboard/1_Salary_Dashboard.xlsx)

## Conclusion

This project helped me improve my Excel skills through practical work with salary data. I gained experience in using formulas, creating charts, applying Data Validation and building an interactive dashboard to present information in a clear and simple way.
