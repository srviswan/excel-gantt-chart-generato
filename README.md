# Gantt Chart Generator

This project provides a Python tool to generate Gantt charts from Excel data. It's designed to visualize project tasks, resources, and timelines in an interactive HTML Gantt chart.

## Features

- Reads task data from an Excel file
- Generates an interactive HTML Gantt chart
- Displays tasks by resource, location, and business driver
- Shows task duration as bars with task names as labels
- Provides hover information with detailed task data

## Requirements

- Python 3.6+
- Required Python packages (install via `pip install -r requirements.txt`):
  - pandas
  - openpyxl
  - matplotlib
  - plotly

## Input Excel Format

The input Excel file should have the following columns:
- `Task`: The name of the task
- `Task1`: Task category or sub-type (optional)
- `Business Driver`: The business reason for the task
- `Resource`: The person or team responsible for the task
- `Group`: Grouping category (optional)
- `Data`: Additional data about the task (optional)
- `Location`: Where the task is being performed
- Month columns (`Jan` through `Dec`): Mark with any value (e.g., 'X') to indicate which months the task spans

## Usage

1. Install the required packages:
   ```
   python -m pip install -r requirements.txt
   ```

2. Create your Excel file with the required format, or generate a sample file:
   ```
   python create_sample_data.py
   ```

3. Run the Gantt chart generator:
   ```
   python gantt_chart_generator.py --input sample_tasks.xlsx --output gantt_chart.html
   ```

4. Open the generated HTML file in your web browser to view the interactive Gantt chart.

## Example

The repository includes a script to generate sample data:

```
python create_sample_data.py
```

This will create a file called `sample_tasks.xlsx` with example project tasks that you can use to test the Gantt chart generator.
