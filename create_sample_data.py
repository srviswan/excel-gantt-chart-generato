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
        'Task 1': [
            'Planning', 'Analysis', 'Design', 'Development', 'Development',
            'QA', 'Operations', 'Training', 'Documentation', 'Review'
        ],
        'Business Driver': [
            'Strategic', 'Strategic', 'Technical', 'Technical', 'Technical',
            'Quality', 'Operational', 'Adoption', 'Knowledge', 'Improvement'
        ],
        'Resources': [
            'Project Manager', 'Business Analyst', 'System Architect', 
            'Developer Team A', 'Developer Team B', 'QA Team', 
            'DevOps', 'Trainer', 'Technical Writer', 'Project Manager'
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
    
    # Add month columns (January to December)
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    
    # Initialize all month values to empty
    for month in months:
        data[month] = ['' for _ in range(len(data['Task']))]
    
    # Set task durations across months
    # Project Planning: January-February
    data['January'][0] = 'X'
    data['February'][0] = 'X'
    
    # Requirements Gathering: February-March
    data['February'][1] = 'X'
    data['March'][1] = 'X'
    
    # System Design: March-April
    data['March'][2] = 'X'
    data['April'][2] = 'X'
    
    # Development Phase 1: April-June
    data['April'][3] = 'X'
    data['May'][3] = 'X'
    data['June'][3] = 'X'
    
    # Development Phase 2: June-August
    data['June'][4] = 'X'
    data['July'][4] = 'X'
    data['August'][4] = 'X'
    
    # Testing: August-September
    data['August'][5] = 'X'
    data['September'][5] = 'X'
    
    # Deployment: October
    data['October'][6] = 'X'
    
    # Training: October-November
    data['October'][7] = 'X'
    data['November'][7] = 'X'
    
    # Documentation: September-November
    data['September'][8] = 'X'
    data['October'][8] = 'X'
    data['November'][8] = 'X'
    
    # Post-Launch Review: December
    data['December'][9] = 'X'
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Sample data saved to {output_file}")

if __name__ == "__main__":
    create_sample_data()
