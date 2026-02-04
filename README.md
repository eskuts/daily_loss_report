# Logistics Reporting Automation (ETL Pipeline)

I built this project to solve a massive manual workload issue in our transport department. We were spending about 5 hours every morning just to collect data and format reports. This tool automates the whole flow, cutting the time down to about 70 minutes.

### The Problem
The main challenge was that the source software (TransNavigation) didn't have an API. I had to build a custom extractor that simulates user actions to get the raw data. Also, the data was messy and came from different sources (fleet monitoring, financial logs, and manual spreadsheets).

### How it works
The system is divided into 4 Python modules that run sequentially:

1. **data_extractor**: A Selenium-based automation script that logs into TransNavigation and pulls raw performance data.
2. **plan_fact**: This is the core module. It cleans the raw data, calculates the "Plan vs Actual" metrics using Pandas, and pushes everything to a Google Sheet so the management can see it immediately.
3. **colon_report**: Takes the processed data and breaks it down into specific views for different column supervisors.
4. **damage_calculator**: A small module I added to handle financial penalties. It calculates damage percentages based on the losses identified in the plan-fact report.

### Tech choices
* **Python (Pandas/NumPy):** For all data transformations.
* **Selenium:** To bypass the lack of an API in our monitoring software.
* **Google Sheets API:** Chosen as the final UI because it was the easiest way for non-technical staff to access the data.
* **Logging & .env:** I used standard logging and environment variables to make the scripts easier to debug and more secure.

### Result
It's been running daily since I deployed it. We've reached 100% data consistency because I removed the human factor from the data entry stage. For the company, it's a saving of ~20 man-hours per week.
