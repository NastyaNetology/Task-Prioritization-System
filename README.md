# Task Prioritization Decision Support System (DSS)

## Business Task
Anastasia and her team were looking for a way to help companies prioritize their projects and tasks when working with limited budgets, strict deadlines, and resource constraints. Often, managers rely on intuition or time-consuming manual methods to decide what gets done first.
They wanted to build a tool that could automate this decision-making by scoring each task based on business-relevant factors like cost, benefit, complexity, and urgency.

## Dataset
To solve this problem the data set from Kaggle was used, which includes:
- Task Details: name, description, department, region, manager;
- Quantitative Metrics: cost, benefit, complexity, duration, completion %;
- Status Information: start and end dates, business criticality, progress.

## Analytical Approach
We applied the WSM method to calculate a total score for each task:
Total Score = Σ (weight_i * score_i)
Where:
weight_i = Importance of each criterion
score_i = Score assigned to a task on that criterion

## Prototype Features
The system was built with Python and UI tools, includes steps:
- CSV Upload: Import task data in a predefined format;
- Criteria Selection: Assign scores (0–4) to multiple task evaluation criteria;
- Automated Scoring: System calculates scores per task;
- Task Ranking: Top 10 tasks displayed or downloaded;
- Graphical Insights: Visualize complexity distribution;
- Resource Filtering: Input budget and developer limits for final recommendations.

## Tech Stack
- Python
- Pandas
- Tkinter / Streamlit (UI)
- Matplotlib / Seaborn (Visuals)
- CSV/XLSX for data import/export
