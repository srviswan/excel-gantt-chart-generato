# Excel Gantt Chart Generator

This project provides a Python tool to generate Gantt charts from Excel data. It's designed to visualize project tasks, resources, and timelines in an Excel-based Gantt chart.

## Features

- Reads task data from an Excel file
- Generates a formatted Excel Gantt chart
- Organizes tasks by Resource, Business Driver, and Location
- Shows task duration as color-coded bars with task names as labels
- Provides detailed summary information

## Requirements

- Python 3.6+
- Required Python packages (install via `pip install -r requirements.txt`):
  - pandas
  - openpyxl

## Input Excel Format

The input Excel file should have the following columns:
- `Task`: The name of the task
- `Task 1`: Task category or sub-type (optional)
- `Business Driver`: The business reason for the task
- `Resources`: The person or team responsible for the task
- `Data`: Additional data about the task (optional)
- `Location`: Where the task is being performed
- Month columns (`January` through `December`): Mark with any value (e.g., 'X') to indicate which months the task spans

## Usage

1. Install the required packages:
   ```
   python -m pip install -r requirements.txt
   ```

2. Create your Excel file with the required format, or generate a sample file:
   ```
   python create_sample_data.py
   ```

3. Run the simplified Excel Gantt chart generator:
   ```
   python simple_excel_gantt.py --input sample_tasks.xlsx --output gantt_chart_export.xlsx
   ```

4. Open the generated Excel file to view the Gantt chart.

## Output Excel Format

The generated Excel file contains three sheets:

1. **Gantt Chart** - Visual representation of tasks organized by Resource, Driver, and Location with color-coded bars for each task
2. **Resource Summary** - Detailed information about tasks grouped by Resource, Driver, and Location
3. **Task Legend** - Color coding reference for each task

## Example

The repository includes a script to generate sample data:

```
python create_sample_data.py
```

This will create a file called `sample_tasks.xlsx` with example project tasks that you can use to test the Gantt chart generator.
