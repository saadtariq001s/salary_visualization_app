import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import datetime
import base64
import io

class PayVisualizer:
    def __init__(self):
        # Default data
        self.grade_data = {
            'Grade': [12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1],
            'Minimum': [45000, 30000, 22500, 18000, 12000, 9000, 7500, 4875, 3750, 2625, 1650, 855],
            'Midpoint': [60000, 40000, 30000, 24000, 16000, 12000, 10000, 6500, 5000, 3500, 2200, 1140],
            'Maximum': [75000, 50000, 37500, 30000, 20000, 15000, 12500, 8125, 6250, 4375, 2750, 1425]
        }
        
        # Updated market data based on the provided table
        # In reverse order (12 to 1) for compatibility with visualization
        self.market_data = [
            76200,   # Grade 12
            49800,   # Grade 11
            38100,   # Grade 10
            30936,   # Grade 9
            22678,   # Grade 8
            16555,   # Grade 7
            12390,   # Grade 6
            8443.5,  # Grade 5
            6350,    # Grade 4
            4515,    # Grade 3
            2816,    # Grade 2
            1482     # Grade 1
        ]
        
        # Convert data to pandas DataFrame
        self.grade_df = pd.DataFrame(self.grade_data)
        self.employee_df = None
        
    def load_employee_data(self, uploaded_file):
        """Load employee data from uploaded Excel file with improved error handling"""
        try:
            # First, print debug information about the file
            file_info = f"Loading file: {uploaded_file.name}, Size: {uploaded_file.size} bytes"
            print(file_info)
            
            # Load Excel file - with explicit engine specification
            try:
                # Try with openpyxl engine first (newer Excel formats)
                self.employee_df = pd.read_excel(uploaded_file, engine='openpyxl')
            except Exception as e1:
                try:
                    # Fall back to xlrd engine (older Excel formats)
                    self.employee_df = pd.read_excel(uploaded_file, engine='xlrd')
                except Exception as e2:
                    return False, f"Failed to read Excel file with either engine. Error 1: {str(e1)}, Error 2: {str(e2)}"
            
            # Print column information for debugging
            columns_info = f"Columns found in file: {', '.join(self.employee_df.columns.tolist())}"
            print(columns_info)
            
            # Check if DataFrame is empty
            if self.employee_df.empty:
                return False, "The uploaded Excel file doesn't contain any data"
                
            # Check if required columns exist (case-insensitive check)
            required_columns = ['EMP ID', 'EMP NAME', 'GRADE', 'TOTAL']
            df_columns_upper = [col.upper() for col in self.employee_df.columns]
            
            missing_columns = []
            column_mapping = {}  # To map required column names to actual column names
            
            for req_col in required_columns:
                # Try to find a match (case-insensitive)
                found = False
                for i, col in enumerate(self.employee_df.columns):
                    if col.upper() == req_col.upper() or col.upper().replace(' ', '') == req_col.upper().replace(' ', ''):
                        column_mapping[req_col] = col
                        found = True
                        break
                
                if not found:
                    missing_columns.append(req_col)
            
            if missing_columns:
                return False, f"Missing required columns: {', '.join(missing_columns)}"
            
            # Rename columns to expected format if needed
            if column_mapping:
                self.employee_df = self.employee_df.rename(columns={column_mapping[k]: k for k in column_mapping})
            
            # Handle GRADE column - more flexible extraction with detailed error reporting
            try:
                # First, check if GRADE is already numeric
                if pd.api.types.is_numeric_dtype(self.employee_df['GRADE']):
                    # If already numeric, just ensure it's an integer
                    self.employee_df['GRADE'] = self.employee_df['GRADE'].astype(int)
                else:
                    # Try multiple extraction patterns for text-based grades
                    patterns = [
                        r'Grade\s*(\d+)',  # "Grade 12" format
                        r'G\s*(\d+)',      # "G 12" format
                        r'(\d+)',          # Just the number
                    ]
                    
                    # Try each pattern
                    for pattern in patterns:
                        extracted_grades = self.employee_df['GRADE'].astype(str).str.extract(pattern, expand=False)
                        if not extracted_grades.isna().all():
                            # Found a pattern that works
                            self.employee_df['GRADE'] = pd.to_numeric(extracted_grades, errors='coerce')
                            break
                
                # Check if we have any valid grades after extraction
                if self.employee_df['GRADE'].isna().all():
                    sample_grades = self.employee_df['GRADE'].head(5).tolist()
                    return False, f"Could not extract numeric grade values. Sample values: {sample_grades}"
            except Exception as e:
                return False, f"Error processing GRADE column: {str(e)}"
            
            # Filter out any rows with NaN or 0 grade values
            original_count = len(self.employee_df)
            self.employee_df = self.employee_df[self.employee_df['GRADE'].notna() & (self.employee_df['GRADE'] > 0)]
            filtered_count = len(self.employee_df)
            
            if self.employee_df.empty:
                return False, f"No valid data rows remaining after filtering invalid grades. Started with {original_count} rows."
            
            # Convert GRADE to integer type after filtering
            self.employee_df['GRADE'] = self.employee_df['GRADE'].astype(int)
            
            # Handle TOTAL column - convert to numeric
            try:
                # Handle numeric columns that might be formatted as strings
                if len(self.employee_df) > 0:
                    if isinstance(self.employee_df['TOTAL'].iloc[0], str):
                        self.employee_df['TOTAL'] = self.employee_df['TOTAL'].astype(str).replace({',': '', 'AED': '', ' ': ''}, regex=True)
                        self.employee_df['TOTAL'] = pd.to_numeric(self.employee_df['TOTAL'], errors='coerce')
                    
                    # In case it's still not numeric, force conversion
                    self.employee_df['TOTAL'] = pd.to_numeric(self.employee_df['TOTAL'], errors='coerce')
                    
                    # Check if we have valid salary data
                    if self.employee_df['TOTAL'].isna().all():
                        sample_total = [str(x) for x in self.employee_df['TOTAL'].head(5).tolist()]
                        return False, f"Could not convert TOTAL column to numeric values. Sample values: {sample_total}"
            except Exception as e:
                return False, f"Error processing TOTAL column: {str(e)}"
                
            # Flag outliers (employees outside their grade's salary range)
            self.employee_df['IS_OUTLIER'] = False
            for index, row in self.employee_df.iterrows():
                grade = row['GRADE']
                salary = row['TOTAL']
                
                # Get min and max for this grade, with error checking
                grade_row = self.grade_df[self.grade_df['Grade'] == grade]
                if not grade_row.empty:
                    grade_min = grade_row['Minimum'].values[0]
                    grade_max = grade_row['Maximum'].values[0]
                    
                    # Check if salary is outside the range
                    if salary < grade_min or salary > grade_max:
                        self.employee_df.at[index, 'IS_OUTLIER'] = True
            
            # Print summary for debugging
            summary = (
                f"Successfully processed {len(self.employee_df)} rows of employee data.\n"
                f"Grades range from {self.employee_df['GRADE'].min()} to {self.employee_df['GRADE'].max()}.\n"
                f"Filtered out {original_count - filtered_count} rows with invalid grades."
            )
            print(summary)
            
            return True, f"Successfully loaded {len(self.employee_df)} employee records"
            
        except Exception as e:
            # Get more detailed error information
            import traceback
            error_details = traceback.format_exc()
            print(f"Detailed error: {error_details}")
            return False, f"Failed to load employee data: {str(e)}"
    
    def update_grade_data(self, new_grade_data):
        """Update grade data with new values"""
        try:
            for i, row in new_grade_data.iterrows():
                grade = row['Grade']
                self.grade_df.loc[self.grade_df['Grade'] == grade, 'Minimum'] = row['Minimum']
                self.grade_df.loc[self.grade_df['Grade'] == grade, 'Midpoint'] = row['Midpoint']
                self.grade_df.loc[self.grade_df['Grade'] == grade, 'Maximum'] = row['Maximum']
            
            # If employee data is loaded, recompute outliers
            if self.employee_df is not None:
                # Flag outliers with the updated grade ranges
                for index, row in self.employee_df.iterrows():
                    grade = row['GRADE']
                    salary = row['TOTAL']
                    
                    # Get min and max for this grade
                    grade_min = self.grade_df.loc[self.grade_df['Grade'] == grade, 'Minimum'].values[0]
                    grade_max = self.grade_df.loc[self.grade_df['Grade'] == grade, 'Maximum'].values[0]
                    
                    # Check if salary is outside the range
                    if salary < grade_min or salary > grade_max:
                        self.employee_df.at[index, 'IS_OUTLIER'] = True
                    else:
                        self.employee_df.at[index, 'IS_OUTLIER'] = False
            
            return True, "Grade data updated successfully"
        except Exception as e:
            return False, f"Failed to update grade data: {str(e)}"
    
    def update_market_data(self, new_market_data):
        """Update market data with new values"""
        try:
            # Create a new market data array based on the data editor values
            updated_market_data = []
            
            # Create a dictionary mapping grades to their market values from the edited data
            grade_to_market = {}
            for i, row in enumerate(new_market_data.itertuples()):
                grade_to_market[row.Grade] = row.Market_50th_Percentile
                
            # Rebuild the market_data array in the proper order (12 down to 1)
            # This ensures compatibility with the visualization function
            for grade in range(12, 0, -1):  # From grade 12 down to grade 1
                if grade in grade_to_market:
                    updated_market_data.append(grade_to_market[grade])
                else:
                    # Use a default or existing value if available
                    grade_index = 12 - grade
                    if 0 <= grade_index < len(self.market_data):
                        updated_market_data.append(self.market_data[grade_index])
                    else:
                        updated_market_data.append(0)
            
            # Update the market data array
            self.market_data = updated_market_data
            
            return True, "Market data updated successfully"
        except Exception as e:
            return False, f"Failed to update market data: {str(e)}"
    
    def set_predefined_market_data(self):
        """Set the market data to predefined values from the table"""
        # These values match the "Market Mid Point" column from the table
        self.market_data = [
            76200,   # Grade 12
            49800,   # Grade 11
            38100,   # Grade 10
            30936,   # Grade 9
            22678,   # Grade 8
            16555,   # Grade 7
            12390,   # Grade 6
            8443.5,  # Grade 5
            6350,    # Grade 4
            4515,    # Grade 3
            2816,    # Grade 2
            1482     # Grade 1
        ]
        return True, "Market data updated with predefined values"
    
    def generate_visualization(self):
        """Generate the salary visualization based on current data"""
        # Create the figure
        fig = go.Figure()
        
        # Sort grade data to ensure proper order
        self.grade_df = self.grade_df.sort_values('Grade', ascending=True)
        
        # Ensure grades are integers
        grades = [int(g) for g in self.grade_df['Grade'].tolist()]
        min_values = self.grade_df['Minimum'].tolist()
        mid_values = self.grade_df['Midpoint'].tolist()
        max_values = self.grade_df['Maximum'].tolist()
        
        # Reorder market data to match the sorted grade order
        sorted_market_data = []
        for grade in grades:
            # Calculate the index in the original market_data array
            # Original data is for grades 12 down to 1, so we need to adjust the index
            original_index = 12 - grade  # If grade is 1, we need index 11 (last item)
            if 0 <= original_index < len(self.market_data):
                sorted_market_data.append(self.market_data[original_index])
            else:
                # Fallback if grade is out of range
                sorted_market_data.append(0)
        
        # Layer 1: Vertical bars for salary ranges
        for i, grade in enumerate(grades):
            # Create bar for each grade's salary range
            fig.add_trace(go.Bar(
                x=[grade],
                y=[max_values[i] - min_values[i]],  # Height of bar is max-min
                base=min_values[i],  # Start bar at minimum value
                width=0.8,  # Increased width for better visibility
                marker=dict(
                    color='rgba(176, 196, 222, 0.8)',  # Light steel blue, more professional
                    line=dict(color='rgba(70, 130, 180, 1)', width=1.5)  # Steel blue border
                ),
                name=f'Grade {grade} Range',
                hovertemplate=
                    "<b>Grade %{x} Salary Range</b><br><br>" +
                    "Minimum: AED %{customdata[0]:,.0f}<br>" +
                    "Midpoint: AED %{customdata[1]:,.0f}<br>" +
                    "Maximum: AED %{customdata[2]:,.0f}<br>" +
                    "<extra></extra>",
                customdata=np.column_stack((min_values[i], mid_values[i], max_values[i])),
                showlegend=False
            ))
            
            # Add minimum marker (small line)
            fig.add_trace(go.Scatter(
                x=[grade-0.3, grade+0.3],
                y=[min_values[i], min_values[i]],
                mode='lines',
                line=dict(color='rgba(70, 130, 180, 0.8)', width=2, dash='dot'),
                name=f'Min - Grade {grade}',
                hovertemplate="<b>Minimum Salary</b><br>Grade %{x}<br>AED %{y:,.0f}<extra></extra>",
                showlegend=False
            ))
            
            # Add maximum marker (small line)
            fig.add_trace(go.Scatter(
                x=[grade-0.3, grade+0.3],
                y=[max_values[i], max_values[i]],
                mode='lines',
                line=dict(color='rgba(70, 130, 180, 0.8)', width=2, dash='dot'),
                name=f'Max - Grade {grade}',
                hovertemplate="<b>Maximum Salary</b><br>Grade %{x}<br>AED %{y:,.0f}<extra></extra>",
                showlegend=False
            ))
            
            # Add midpoint marker as horizontal line spanning the bar width
            fig.add_trace(go.Scatter(
                x=[grade-0.35, grade+0.35],
                y=[mid_values[i], mid_values[i]],
                mode='lines',
                line=dict(
                    color='rgba(46, 139, 87, 0.95)',  # Sea green, more professional
                    width=2.5  # Slightly thicker for visibility
                ),
                name=f'Grade {grade} Midpoint',
                hovertemplate="<b>Midpoint Salary</b><br>Grade %{x}<br>AED %{y:,.0f}<extra></extra>",
                showlegend=False
            ))
        
        # Layer 2: Market 50th percentile line - Enhanced style
        # Use the sorted market data instead of the original
        fig.add_trace(go.Scatter(
            x=grades,
            y=sorted_market_data,  # Use our reordered market data
            mode='lines+markers',
            line=dict(
                color='rgba(25, 25, 112, 0.95)', 
                width=4,
                dash='solid'
            ),
            marker=dict(
                size=12, 
                color='rgba(25, 25, 112, 0.95)',
                symbol='circle',
                line=dict(
                    color='white',
                    width=2
                )
            ),
            name='Market 50th Percentile',
            hovertemplate="<b>Market 50th Percentile</b><br>Grade %{x}<br>AED %{y:,.0f}<extra></extra>"
        ))
            
        # Layer 3: Employee salary data points if available
        if self.employee_df is not None:
            # Group employees by grade
            for grade in grades:
                # Normal employees (within range)
                grade_employees = self.employee_df[(self.employee_df['GRADE'] == grade) & (~self.employee_df['IS_OUTLIER'])]
                
                if not grade_employees.empty:
                    # Plot employee salaries as scatter points
                    # Safely check for presence of optional columns
                    designation_col = 'DESIGNATION' if 'DESIGNATION' in grade_employees.columns else None
                    department_col = 'DEPARTMENT' if 'DEPARTMENT' in grade_employees.columns else None
                    doj_col = 'DOJ' if 'DOJ' in grade_employees.columns else None
                    nationality_col = 'NATIONALITY' if 'NATIONALITY' in grade_employees.columns else None
                    basic_col = 'BASIC' if 'BASIC' in grade_employees.columns else None
                    
                    # Prepare customdata with fallbacks for missing columns
                    customdata_list = [grade_employees['EMP ID'].tolist()]
                    
                    if designation_col:
                        customdata_list.append(grade_employees[designation_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(grade_employees))
                        
                    if department_col:
                        customdata_list.append(grade_employees[department_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(grade_employees))
                        
                    if doj_col:
                        customdata_list.append(grade_employees[doj_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(grade_employees))
                        
                    if nationality_col:
                        customdata_list.append(grade_employees[nationality_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(grade_employees))
                        
                    if basic_col:
                        customdata_list.append(grade_employees[basic_col].tolist())
                        customdata_list.append((grade_employees['TOTAL'] - grade_employees[basic_col]).tolist())
                    else:
                        customdata_list.append([0] * len(grade_employees))
                        customdata_list.append([0] * len(grade_employees))
                    
                    fig.add_trace(go.Scatter(
                        x=[grade] * len(grade_employees),
                        y=grade_employees['TOTAL'].tolist(),
                        mode='markers',
                        marker=dict(
                            color='rgba(178, 34, 34, 0.8)',  # Firebrick red for normal employees
                            size=8,
                            symbol='circle'
                        ),
                        name=f'Grade {grade} Employees',
                        text=grade_employees['EMP NAME'].tolist(),
                        customdata=np.stack(customdata_list, axis=1),
                        hovertemplate=(
                            '<b>%{text}</b><br>' +
                            'ID: %{customdata[0]}<br>' +
                            'Designation: %{customdata[1]}<br>' +
                            'Department: %{customdata[2]}<br>' +
                            'Joined: %{customdata[3]}<br>' +
                            'Nationality: %{customdata[4]}<br>' +
                            '<br>' +
                            'Basic Salary: AED %{customdata[5]:,.2f}<br>' +
                            'Allowances: AED %{customdata[6]:,.2f}<br>' +
                            'Total Salary: AED %{y:,.2f}' +
                            '<extra></extra>'
                        ),
                        showlegend=False
                    ))
                
                # Outlier employees (outside range)
                outlier_employees = self.employee_df[(self.employee_df['GRADE'] == grade) & (self.employee_df['IS_OUTLIER'])]
                
                if not outlier_employees.empty:
                    # Plot outlier employee salaries as scatter points with different color
                    # Safely check for presence of optional columns (same as above)
                    designation_col = 'DESIGNATION' if 'DESIGNATION' in outlier_employees.columns else None
                    department_col = 'DEPARTMENT' if 'DEPARTMENT' in outlier_employees.columns else None
                    doj_col = 'DOJ' if 'DOJ' in outlier_employees.columns else None
                    nationality_col = 'NATIONALITY' if 'NATIONALITY' in outlier_employees.columns else None
                    basic_col = 'BASIC' if 'BASIC' in outlier_employees.columns else None
                    
                    # Prepare customdata with fallbacks for missing columns
                    customdata_list = [outlier_employees['EMP ID'].tolist()]
                    
                    if designation_col:
                        customdata_list.append(outlier_employees[designation_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(outlier_employees))
                        
                    if department_col:
                        customdata_list.append(outlier_employees[department_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(outlier_employees))
                        
                    if doj_col:
                        customdata_list.append(outlier_employees[doj_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(outlier_employees))
                        
                    if nationality_col:
                        customdata_list.append(outlier_employees[nationality_col].tolist())
                    else:
                        customdata_list.append(["N/A"] * len(outlier_employees))
                        
                    if basic_col:
                        customdata_list.append(outlier_employees[basic_col].tolist())
                        customdata_list.append((outlier_employees['TOTAL'] - outlier_employees[basic_col]).tolist())
                    else:
                        customdata_list.append([0] * len(outlier_employees))
                        customdata_list.append([0] * len(outlier_employees))
                    
                    fig.add_trace(go.Scatter(
                        x=[grade] * len(outlier_employees),
                        y=outlier_employees['TOTAL'].tolist(),
                        mode='markers',
                        marker=dict(
                            color='rgba(255, 140, 0, 0.9)',  # Dark orange for outliers
                            size=10,  # Slightly larger for emphasis
                            symbol='circle-open',  # Open circles for outliers
                            line=dict(width=2, color='rgba(255, 140, 0, 1)')  # Darker border
                        ),
                        name=f'Grade {grade} Outliers',
                        text=outlier_employees['EMP NAME'].tolist(),
                        customdata=np.stack(customdata_list, axis=1),
                        hovertemplate=(
                            '<b>%{text} (OUTLIER)</b><br>' +
                            'ID: %{customdata[0]}<br>' +
                            'Designation: %{customdata[1]}<br>' +
                            'Department: %{customdata[2]}<br>' +
                            'Joined: %{customdata[3]}<br>' +
                            'Nationality: %{customdata[4]}<br>' +
                            '<br>' +
                            'Basic Salary: AED %{customdata[5]:,.2f}<br>' +
                            'Allowances: AED %{customdata[6]:,.2f}<br>' +
                            'Total Salary: AED %{y:,.2f}' +
                            '<extra></extra>'
                        ),
                        showlegend=False
                    ))
        
        # Create legends for the different elements with enhanced professional styling
        fig.add_trace(go.Scatter(
            x=[None], y=[None], 
            mode='lines',
            line=dict(color='rgba(46, 139, 87, 0.95)', width=2.5),
            name='Midpoint'
        ))
        
        fig.add_trace(go.Scatter(
            x=[None], y=[None], mode='markers',
            marker=dict(
                size=8, 
                color='rgba(178, 34, 34, 0.9)',
                line=dict(width=1, color='white')
            ),
            name='Employee Salary'
        ))
        
        # Add outlier legend only if employee data exists
        if self.employee_df is not None:
            fig.add_trace(go.Scatter(
                x=[None], y=[None], mode='markers',
                marker=dict(
                    size=10, 
                    color='rgba(255, 140, 0, 0.9)',
                    symbol='circle-open',
                    line=dict(width=2, color='rgba(255, 140, 0, 1)')
                ),
                name='Salary Outliers'
            ))
        
        fig.add_trace(go.Bar(
            x=[None], y=[None],
            marker=dict(
                color='rgba(176, 196, 222, 0.8)', 
                line=dict(color='rgba(70, 130, 180, 1)', width=1.5)
            ),
            name='Salary Range (Min-Max)'
        ))
        
        fig.add_trace(go.Scatter(
            x=[None], y=[None], 
            mode='lines',
            line=dict(color='rgba(70, 130, 180, 0.8)', width=2, dash='dot'),
            name='Min/Max Indicators'
        ))
        
        # Update layout with professional styling
        fig.update_layout(
            title={
                'text': 'Salary Structure Analysis by Job Grade',
                'font': {'size': 26, 'color': '#2F4F4F', 'family': 'Helvetica, Arial, sans-serif'},
                'x': 0.5,  # Center the title
                'xanchor': 'center',
                'y': 0.95
            },
            xaxis=dict(
                title={
                    'text': 'Job Grade',
                    'font': {'size': 18, 'family': 'Helvetica, Arial, sans-serif', 'color': '#2F4F4F'}
                },
                tickmode='array',
                tickvals=grades,
                ticktext=[f'Grade {int(g)}' for g in grades],  # Ensure grades are displayed as integers
                gridcolor='rgba(200, 200, 200, 0.3)',
                gridwidth=1,
                showgrid=True,
                zeroline=False,
                showline=True,
                linecolor='rgba(150, 150, 150, 0.5)',
                linewidth=1
            ),
            yaxis=dict(
                title={
                    'text': 'Salary',
                    'font': {'size': 18, 'family': 'Helvetica, Arial, sans-serif', 'color': '#2F4F4F'}
                },
                autorange=True,
                gridcolor='rgba(200, 200, 200, 0.7)',
                gridwidth=1,
                showgrid=True,
                zeroline=True,
                zerolinecolor='rgba(150, 150, 150, 0.5)',
                zerolinewidth=1,
                showline=True,
                linecolor='rgba(150, 150, 150, 0.5)',
                linewidth=1,
                tickformat=',d',  # Add thousands separators to y-axis labels
                tickprefix='AED '  # Add AED currency symbol to y-axis values
            ),
            hovermode='closest',
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=0.01,
                bgcolor='rgba(255, 255, 255, 0.9)',
                bordercolor='rgba(120, 120, 120, 0.5)',
                borderwidth=1,
                font=dict(
                    family="Helvetica, Arial, sans-serif",
                    size=14,
                    color="#2F4F4F"
                )
            ),
            margin=dict(l=60, r=60, t=100, b=60),
            height=800,
            paper_bgcolor='white',  # White paper background
            plot_bgcolor='rgba(245, 245, 250, 0.9)',  # Very light background for professional look
        )
        
        # Add subtitle and date stamp
        fig.add_annotation(
            text="Comparing Internal Salary Structure with Market Benchmarks",
            xref="paper", yref="paper",
            x=0.5, y=0.89,
            showarrow=False,
            font=dict(
                family="Helvetica, Arial, sans-serif",
                size=22,
                color="#000000",
                weight="bold"
            ),
            align="center",
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="#000000",
            borderwidth=2,
            borderpad=4
        )
        
        # Add date stamp
        current_date = datetime.now().strftime("%B %d, %Y")
        fig.add_annotation(
            text=f"Report Generated: {current_date}",
            xref="paper", yref="paper",
            x=0.98, y=0.02,
            showarrow=False,
            font=dict(
                family="Helvetica, Arial, sans-serif",
                size=16,
                color="#000000",
                weight="bold"
            ),
            align="right",
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="#000000",
            borderwidth=1,
            borderpad=4
        )
        
        return fig
    
    def generate_download_link(self, fig):
        """Generate a download link for the visualization"""
        # Create a copy of the figure to ensure we don't modify the original
        download_fig = fig
        
        # Ensure the y-axis has proper formatting for the download version
        download_fig.update_layout(
            yaxis=dict(
                title={
                    'text': 'SALARY (AED)',
                    'font': {'size': 24, 'family': 'Helvetica, Arial, sans-serif', 'color': '#000000', 'weight': 'bold'},
                    'standoff': 25
                },
                autorange=True,
                gridcolor='rgba(0, 0, 0, 0.3)',
                gridwidth=2,
                showgrid=True,
                zeroline=True,
                zerolinecolor='#000000',
                zerolinewidth=3,
                showline=True,
                linecolor='#000000',
                linewidth=3,
                tickformat=',d',
                tickprefix='AED ',
                tickfont=dict(
                    family="Helvetica, Arial, sans-serif",
                    size=18,
                    color="#000000"
                ),
                nticks=15,
                showticklabels=True
            ),
            margin=dict(l=140, r=80, t=120, b=120),  # Increase left margin even more for download version
        )
        
        # Force the figure to render all y-axis labels
        download_fig.update_yaxes(
            showticklabels=True,
            automargin=True,
        )
        
        # Write to HTML with full labels
        buffer = io.StringIO()
        download_fig.write_html(
            buffer,
            include_plotlyjs='cdn',
            full_html=True,
            config={'displayModeBar': True, 'responsive': True}
        )
        html_bytes = buffer.getvalue().encode()
        encoded = base64.b64encode(html_bytes).decode()
        
        href = f'<a href="data:text/html;base64,{encoded}" download="payvisualizer_report.html" class="download-button">Download HTML File</a>'
        return href

def display_guide():
    """Display user guide"""
    st.title("Welcome to PayVisualizer")
    
    st.subheader("This tool helps you see how employee salaries compare with your company's salary ranges and market rates.")
    
    st.markdown("### How to Use This Tool:")
    
    st.markdown("""
    1️⃣ **Load Your Data** - Upload your Excel file with employee information. The file should have columns for employee ID, name, grade, and salary.
    
    2️⃣ **Check Salary Ranges (Optional)** - If needed, adjust the minimum, middle, or maximum salary for each job grade.
    
    3️⃣ **Update Market Rates (Optional)** - If you want to change the market comparison values, update the market data.
    
    4️⃣ **Create Your Chart** - Generate the salary chart. Each employee will show as a dot, with outliers highlighted in orange.
    
    5️⃣ **Save Your Work** - Download the chart as an HTML file you can open later in any web browser.
    """)
    
    st.markdown("### Understanding Your Chart:")
    
    st.markdown("""
    • **Blue Bars** - The salary range for each job grade (from minimum to maximum)
    • **Green Lines** - The midpoint salary for each job grade
    • **Red Dots** - Employee salaries within the grade range
    • **Orange Circles** - Outliers (employee salaries outside their grade range)
    • **Blue Line** - Market comparison rates showing what other companies pay
    • **Hover Details** - Move your mouse over any part of the chart to see more information
    """)
    
    st.markdown("### Required File Format:")
    
    st.markdown("""
    Your Excel file needs these columns:
    • **EMP ID** - Employee ID number
    • **EMP NAME** - Employee name
    • **GRADE** - Job grade (like 'Grade 1', 'Grade 2')
    • **TOTAL** - Total salary amount

    Other helpful columns (if available):
    • **DESIGNATION** - Job title
    • **DEPARTMENT** - Department name
    • **DOJ** - Date of joining
    • **NATIONALITY** - Employee nationality
    • **BASIC** - Basic salary
    """)

def main():
    # Set page configuration
    st.set_page_config(
        page_title="PayVisualizer",
        page_icon="💰",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Add custom CSS for styling
    st.markdown("""
    <style>
    .download-button {
        color: #0366d6;
        text-decoration: underline;
        cursor: pointer;
        background: none;
        border: none;
        padding: 0;
        font-size: 16px;
        margin: 10px 0;
    }
    .download-button:hover {
        color: #0056b3;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Initialize the tool
    if 'tool' not in st.session_state:
        st.session_state.tool = PayVisualizer()
    
    # Initialize session state variables
    if 'show_guide' not in st.session_state:
        st.session_state.show_guide = True
    if 'visualization_generated' not in st.session_state:
        st.session_state.visualization_generated = False
    
    # Sidebar for navigation
    st.sidebar.title("Navigation")
    
    # Navigation buttons
    page = st.sidebar.radio(
        "Go to:",
        ["Guide", "Data Management", "Visualization"],
        index=0 if st.session_state.show_guide else 1
    )
    
    # Update the show guide state based on navigation
    st.session_state.show_guide = (page == "Guide")
    
    # Display the selected page
    if page == "Guide":
        display_guide()
        
    elif page == "Data Management":
        st.title("Data Management")
        
        # Employee data section
        st.header("Employee Data")
        uploaded_file = st.file_uploader("Upload employee data Excel file", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            if st.button("Load Employee Data"):
                success, message = st.session_state.tool.load_employee_data(uploaded_file)
                if success:
                    st.success(message)
                    if st.session_state.tool.employee_df is not None:
                        st.dataframe(st.session_state.tool.employee_df)
                else:
                    st.error(message)
        
        # Grade data section
        st.header("Salary Grade Data")
        
        grade_data_df = st.session_state.tool.grade_df.copy()
        edited_grade_data = st.data_editor(
            grade_data_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True
        )
        
        if st.button("Update Grade Data"):
            success, message = st.session_state.tool.update_grade_data(edited_grade_data)
            if success:
                st.success(message)
            else:
                st.error(message)
        
        # Market data section
        st.header("Market Data")
        
        market_data_df = pd.DataFrame({
            'Grade': st.session_state.tool.grade_df['Grade'],
            'Market 50th Percentile': [
                1482,    # Grade 1
                2816,    # Grade 2
                4515,    # Grade 3
                6350,    # Grade 4
                8443.5,  # Grade 5
                12390,   # Grade 6
                16555,   # Grade 7
                22678,   # Grade 8
                30936,   # Grade 9
                38100,   # Grade 10
                49800,   # Grade 11
                76200    # Grade 12
            ][:len(st.session_state.tool.grade_df)]
        })
        
        edited_market_data = st.data_editor(
            market_data_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True
        )
        
        if st.button("Update Market Data"):
            new_market_data = edited_market_data
            success, message = st.session_state.tool.update_market_data(new_market_data)
            if success:
                st.success(message)
            else:
                st.error(message)
                
        if st.button("Reset to Predefined Market Data"):
            success, message = st.session_state.tool.set_predefined_market_data()
            if success:
                st.success(message)
                # Update the displayed data
                st.experimental_rerun()
            else:
                st.error(message)
    
    elif page == "Visualization":
        st.title("Pay Visualization")
        
        # Always show visualization - with or without employee data
        if st.button("Generate Visualization") or st.session_state.visualization_generated:
            st.session_state.visualization_generated = True
            
            with st.spinner("Generating visualization..."):
                fig = st.session_state.tool.generate_visualization()
                st.plotly_chart(fig, use_container_width=True)
                
                # Add download button
                download_link = st.session_state.tool.generate_download_link(fig)
                st.markdown(download_link, unsafe_allow_html=True)
                
                # Add info text about employee data
                if st.session_state.tool.employee_df is None:
                    st.info("📊 This visualization shows only grade ranges and market data. Upload employee data in the Data Management section to see employee salaries.")
                else:
                    outlier_count = st.session_state.tool.employee_df['IS_OUTLIER'].sum()
                    if outlier_count > 0:
                        st.warning(f"⚠️ Found {outlier_count} employee(s) with salaries outside their grade ranges (shown as orange circles).")

if __name__ == "__main__":
    main()
