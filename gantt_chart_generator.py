#!/usr/bin/env python3
"""
Gantt Chart Generator

This script reads task data from an Excel file and generates a Gantt chart visualization.
The Excel file should have columns for Task, Task1, Business Driver, Resource, Group, 
Data, Location, and months from Jan to Dec.
"""

import pandas as pd
import plotly.figure_factory as ff
import plotly.express as px
import plotly.graph_objects as go
import argparse
from datetime import datetime, timedelta
import os

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Generate a Gantt chart from Excel data.')
    parser.add_argument('--input', '-i', required=True, help='Path to the input Excel file')
    parser.add_argument('--output', '-o', default='gantt_chart.html', help='Path to save the output Gantt chart HTML file')
    return parser.parse_args()

def read_excel_data(file_path):
    """Read data from the Excel file."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")
    
    try:
        df = pd.read_excel(file_path)
        required_columns = ['Task', 'Resource', 'Location', 'Business Driver']
        
        # Check if the required columns exist
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Check for month columns
        month_columns = [col for col in df.columns if col in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]
        if not month_columns:
            raise ValueError("No month columns found (Jan-Dec)")
        
        return df
    except Exception as e:
        raise Exception(f"Error reading Excel file: {str(e)}")

def process_data_for_gantt(df):
    """Process the Excel data into a format suitable for Gantt chart."""
    current_year = datetime.now().year
    
    # Create a list to store task data
    tasks = []
    
    # Month to number mapping
    month_to_num = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }
    
    # Process each row
    for _, row in df.iterrows():
        task_name = row['Task']
        resource = row['Resource']
        location = row['Location']
        driver = row.get('Business Driver', '')  # Use get to handle missing column
        group = row.get('Group', resource)  # Default to resource if Group is missing
        
        # Find start and end months
        start_month = None
        end_month = None
        duration = 0
        
        for month in month_to_num.keys():
            if month in df.columns and pd.notna(row[month]) and row[month]:
                if start_month is None:
                    start_month = month_to_num[month]
                end_month = month_to_num[month]
                duration += 1
        
        # Skip if no duration
        if start_month is None or end_month is None:
            continue
        
        # Create start and end dates
        start_date = datetime(current_year, start_month, 1)
        # End date is the last day of the end month
        if end_month == 12:
            end_date = datetime(current_year, 12, 31)
        else:
            end_date = datetime(current_year, end_month + 1, 1) - timedelta(days=1)
        
        # Create a combined key for Resource, Driver, Location
        resource_key = f"{resource} | {driver} | {location}"
        
        # Add task to the list
        tasks.append({
            'Task': task_name,
            'Resource': resource,
            'Location': location,
            'Driver': driver,
            'ResourceKey': resource_key,  # Combined key for grouping
            'Group': group,
            'Start': start_date,
            'Finish': end_date,
            'Duration': duration,
            'StartMonth': start_month,
            'EndMonth': end_month
        })
    
    return pd.DataFrame(tasks)

def create_gantt_chart(df_tasks, output_file):
    """Create and save the Gantt chart."""
    if df_tasks.empty:
        raise ValueError("No valid task data found for creating Gantt chart")
    
    # Sort by Resource, Driver, Location
    df_tasks = df_tasks.sort_values(by=['Resource', 'Driver', 'Location', 'Start'])
    
    # Create a color map for tasks
    tasks = df_tasks['Task'].unique()
    colors = px.colors.qualitative.Plotly[:len(tasks)]
    task_color_map = dict(zip(tasks, colors))
    
    # Group by ResourceKey (Resource | Driver | Location)
    resource_keys = df_tasks['ResourceKey'].unique()
    
    # Create a figure with subplots - one timeline per month
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # Create the main figure
    fig = go.Figure()
    
    # Add a separate trace for each task
    for idx, (resource_key, group) in enumerate(df_tasks.groupby('ResourceKey')):
        # Split the resource key back into components
        resource, driver, location = resource_key.split(' | ')
        
        # For each task in this resource group
        for _, task in group.iterrows():
            # Create a list of month positions where this task is active
            month_positions = []
            for i, month in enumerate(months):
                month_num = i + 1  # 1-based month number
                if month_num >= task['StartMonth'] and month_num <= task['EndMonth']:
                    month_positions.append(i)
            
            # Add a bar for this task spanning the correct months
            if month_positions:
                fig.add_trace(go.Bar(
                    x=month_positions,  # X positions are month indices
                    y=[idx],  # Y position is the resource group index
                    width=0.8,  # Width of the bar
                    marker=dict(color=task_color_map[task['Task']]),
                    name=task['Task'],  # Use task name for the legend
                    text=task['Task'],  # Show task name on the bar
                    textposition='inside',
                    insidetextanchor='middle',
                    textfont=dict(color='white', size=12),
                    hoverinfo='text',
                    hovertext=f"<b>Task:</b> {task['Task']}<br>"
                              f"<b>Resource:</b> {resource}<br>"
                              f"<b>Driver:</b> {driver}<br>"
                              f"<b>Location:</b> {location}<br>"
                              f"<b>Duration:</b> {task['Duration']} month(s)<br>"
                              f"<b>Start:</b> {task['Start'].strftime('%b %Y')}<br>"
                              f"<b>End:</b> {task['Finish'].strftime('%b %Y')}"
                ))
    
    # Add resource labels on the y-axis
    y_labels = []
    y_positions = []
    
    for idx, resource_key in enumerate(resource_keys):
        resource, driver, location = resource_key.split(' | ')
        y_labels.append(f"{resource} | {driver} | {location}")
        y_positions.append(idx)
    
    # Customize the layout
    fig.update_layout(
        title={
            'text': "<b>Project Gantt Chart</b>",
            'font': {'size': 24}
        },
        xaxis=dict(
            title="<b>Months</b>",
            tickmode='array',
            tickvals=list(range(len(months))),
            ticktext=months,
            showgrid=True,
        ),
        yaxis=dict(
            title="<b>Resource | Driver | Location</b>",
            tickmode='array',
            tickvals=y_positions,
            ticktext=y_labels,
            showgrid=True,
        ),
        barmode='overlay',
        height=max(600, len(resource_keys) * 50),  # Adjust height based on number of resource groups
        margin=dict(l=300, r=50, t=100, b=100),  # Increased left margin for resource labels
        legend_title="<b>Tasks</b>",
        hovermode="closest",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    # Create a second figure for resource details
    resource_details_file = output_file.replace('.html', '_resources.html')
    
    # Create a summary table by resource, driver, location
    resource_summary = df_tasks.groupby(['Resource', 'Driver', 'Location']).agg(
        Tasks=('Task', lambda x: ', '.join(sorted(set(x)))),
        Task_Count=('Task', 'count'),
        Total_Duration=('Duration', 'sum'),
        Months=('Task', lambda x: ', '.join([f"{months[m-1]}" for m in sorted(set(
            [month for task_idx, task in df_tasks[df_tasks['Task'].isin(x)].iterrows() 
             for month in range(task['StartMonth'], task['EndMonth']+1)]))]))
    ).reset_index()
    
    fig_resources = go.Figure(data=[go.Table(
        header=dict(
            values=["<b>Resource</b>", "<b>Driver</b>", "<b>Location</b>", "<b>Tasks</b>", "<b>Task Count</b>", "<b>Total Duration</b>", "<b>Months</b>"],
            fill_color='royalblue',
            align='left',
            font=dict(color='white', size=14)
        ),
        cells=dict(
            values=[
                resource_summary['Resource'],
                resource_summary['Driver'],
                resource_summary['Location'],
                resource_summary['Tasks'],
                resource_summary['Task_Count'],
                resource_summary['Total_Duration'],
                resource_summary['Months']
            ],
            fill_color='lavender',
            align='left',
            font=dict(size=12)
        )
    )])
    
    fig_resources.update_layout(
        title="<b>Resource Summary</b>",
        height=max(400, len(resource_summary) * 50)
    )
    
    # Save the resource details to an HTML file
    fig_resources.write_html(resource_details_file)
    print(f"Resource summary saved to {resource_details_file}")
    
    # Combine both visualizations into a single HTML file with tabs
    with open(output_file, 'w') as f:
        f.write(f'''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Project Gantt Chart</title>
            <style>
                body {{font-family: Arial, sans-serif; margin: 0; padding: 0;}}
                .tab {{overflow: hidden; border: 1px solid #ccc; background-color: #f1f1f1;}}
                .tab button {{background-color: inherit; float: left; border: none; outline: none; cursor: pointer; padding: 14px 16px; transition: 0.3s; font-size: 17px;}}
                .tab button:hover {{background-color: #ddd;}}
                .tab button.active {{background-color: #ccc;}}
                .tabcontent {{display: none; padding: 6px 12px; border: 1px solid #ccc; border-top: none;}}
                #GanttChart {{display: block;}}
            </style>
            <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        </head>
        <body>
            <div class="tab">
                <button class="tablinks active" onclick="openTab(event, 'GanttChart')">Gantt Chart</button>
                <button class="tablinks" onclick="openTab(event, 'ResourceSummary')">Resource Summary</button>
            </div>
            
            <div id="GanttChart" class="tabcontent">
                {fig.to_html(full_html=False, include_plotlyjs=False)}
            </div>
            
            <div id="ResourceSummary" class="tabcontent">
                {fig_resources.to_html(full_html=False, include_plotlyjs=False)}
            </div>
            
            <script>
            function openTab(evt, tabName) {{
                var i, tabcontent, tablinks;
                tabcontent = document.getElementsByClassName("tabcontent");
                for (i = 0; i < tabcontent.length; i++) {{
                    tabcontent[i].style.display = "none";
                }}
                tablinks = document.getElementsByClassName("tablinks");
                for (i = 0; i < tablinks.length; i++) {{
                    tablinks[i].className = tablinks[i].className.replace(" active", "");
                }}
                document.getElementById(tabName).style.display = "block";
                evt.currentTarget.className += " active";
            }}
            </script>
        </body>
        </html>
        ''')
    
    return fig

def main():
    """Main function to run the script."""
    args = parse_arguments()
    
    try:
        # Read data from Excel
        print(f"Reading data from {args.input}...")
        df = read_excel_data(args.input)
        
        # Process data for Gantt chart
        print("Processing data...")
        df_tasks = process_data_for_gantt(df)
        
        # Create and save Gantt chart
        print(f"Creating Gantt chart and saving to {args.output}...")
        fig = create_gantt_chart(df_tasks, args.output)
        
        print("Done!")
    except Exception as e:
        print(f"Error: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
