#!/usr/bin/env python3
"""
Create Sample Excel Data

This script creates a sample Excel file with the structure required for the Gantt chart generator.
"""

import pandas as pd
import numpy as np

def create_sample_data(output_file="sample_tasks.xlsx"):
    """Create a sample Excel file with task data."""
    # Define sample data
    data = {
        'Task': [
            'Project Planning', 'Requirements Gathering', 'System Design', 
            'Development Phase 1', 'Development Phase 2', 'Testing', 
            'Deployment', 'Training', 'Documentation', 'Post-Launch Review'
        ],
        'Task1': [
            'Planning', 'Analysis', 'Design', 'Development', 'Development',
            'QA', 'Operations', 'Training', 'Documentation', 'Review'
        ],
        'Business Driver': [
            'Strategic', 'Strategic', 'Technical', 'Technical', 'Technical',
            'Quality', 'Operational', 'Adoption', 'Knowledge', 'Improvement'
        ],
        'Resource': [
            'Project Manager', 'Business Analyst', 'System Architect', 
            'Developer Team A', 'Developer Team B', 'QA Team', 
            'DevOps', 'Trainer', 'Technical Writer', 'Project Manager'
        ],
        'Group': [
            'Management', 'Analysis', 'Design', 'Development', 'Development',
            'Testing', 'Operations', 'Training', 'Documentation', 'Management'
        ],
        'Data': [
            'High', 'Medium', 'Medium', 'High', 'High',
            'Medium', 'High', 'Low', 'Medium', 'Low'
        ],
        'Location': [
            'New York', 'New York', 'San Francisco', 'San Francisco', 'Bangalore',
            'New York', 'London', 'London', 'Bangalore', 'New York'
        ],
    }
    
    # Add month columns (Jan to Dec)
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # Initialize all month values to empty
    for month in months:
        data[month] = ['' for _ in range(len(data['Task']))]
    
    # Set task durations across months
    # Project Planning: Jan-Feb
    data['Jan'][0] = 'X'
    data['Feb'][0] = 'X'
    
    # Requirements Gathering: Feb-Mar
    data['Feb'][1] = 'X'
    data['Mar'][1] = 'X'
    
    # System Design: Mar-Apr
    data['Mar'][2] = 'X'
    data['Apr'][2] = 'X'
    
    # Development Phase 1: Apr-Jun
    data['Apr'][3] = 'X'
    data['May'][3] = 'X'
    data['Jun'][3] = 'X'
    
    # Development Phase 2: Jun-Aug
    data['Jun'][4] = 'X'
    data['Jul'][4] = 'X'
    data['Aug'][4] = 'X'
    
    # Testing: Aug-Sep
    data['Aug'][5] = 'X'
    data['Sep'][5] = 'X'
    
    # Deployment: Oct
    data['Oct'][6] = 'X'
    
    # Training: Oct-Nov
    data['Oct'][7] = 'X'
    data['Nov'][7] = 'X'
    
    # Documentation: Sep-Nov
    data['Sep'][8] = 'X'
    data['Oct'][8] = 'X'
    data['Nov'][8] = 'X'
    
    # Post-Launch Review: Dec
    data['Dec'][9] = 'X'
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Sample data saved to {output_file}")

if __name__ == "__main__":
    create_sample_data()
