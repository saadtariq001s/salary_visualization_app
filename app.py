import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import datetime
import base64
import io

class SalaryVisualizationTool:
    def __init__(self):
        # Default data
        self.grade_data = {
            'Grade': [12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1],
            'Minimum': [45000, 30000, 22500, 18000, 12000, 9000, 7500, 4900, 3800, 2600, 1700, 900],
            'Midpoint': [60000, 40000, 30000, 24000, 16000, 12000, 10000, 6500, 5000, 3500, 2200, 1100],
            'Maximum': [75000, 50000, 37500, 30000, 20000, 15000, 12500, 8100, 6300, 4400, 2800, 1400]
        }
        
        self.market_data = [5000, 26300, 46300, 56300, 56300, 56300, 56300, 56300, 56300, 56300, 56300, 65900]
        
        # Convert data to pandas DataFrame
        self.grade_df = pd.DataFrame(self.grade_data)
        self.employee_df = None
        
    def load_employee_data(self, uploaded_file):
        """Load employee data from uploaded Excel file"""
        try:
            # Load Excel file
            self.employee_df = pd.read_excel(uploaded_file)
            
            # Check if required columns exist
            required_columns = ['EMP ID', 'EMP NAME', 'GRADE', 'TOTAL']
            missing_columns = [col for col in required_columns if col not in self.employee_df.columns]
            
            if missing_columns:
                return False, f"Missing required columns: {', '.join(missing_columns)}"
            
            # Process the data - Use more robust conversion for GRADE
            # Extract grade numbers and handle missing/invalid values
            extracted_grades = self.employee_df['GRADE'].str.extract(r'Grade\s*(\d+)', expand=False)
            # Convert to numeric, coerce errors to NaN, fill NaN with 0, then convert to int
            self.employee_df['GRADE'] = pd.to_numeric(extracted_grades, errors='coerce')
            
            # Filter out any rows with NaN or 0 grade values
            self.employee_df = self.employee_df[self.employee_df['GRADE'].notna() & (self.employee_df['GRADE'] > 0)]
            
            # Convert to integer type after filtering
            self.employee_df['GRADE'] = self.employee_df['GRADE'].astype(int)
            
            # Handle numeric columns that might be formatted as strings
            if len(self.employee_df) > 0 and isinstance(self.employee_df['TOTAL'].iloc[0], str):
                self.employee_df['TOTAL'] = self.employee_df['TOTAL'].replace({',': ''}, regex=True).astype(float)
            
            return True, f"Successfully loaded {len(self.employee_df)} employee records"
            
        except Exception as e:
            return False, f"Failed to load employee data: {str(e)}"
    
    def update_grade_data(self, new_grade_data):
        """Update grade data with new values"""
        try:
            for i, row in new_grade_data.iterrows():
                grade = row['Grade']
                self.grade_df.loc[self.grade_df['Grade'] == grade, 'Minimum'] = row['Minimum']
                self.grade_df.loc[self.grade_df['Grade'] == grade, 'Midpoint'] = row['Midpoint']
                self.grade_df.loc[self.grade_df['Grade'] == grade, 'Maximum'] = row['Maximum']
            
            return True, "Grade data updated successfully"
        except Exception as e:
            return False, f"Failed to update grade data: {str(e)}"
    
    def update_market_data(self, new_market_data):
        """Update market data with new values"""
        try:
            for i, value in enumerate(new_market_data):
                if i < len(self.market_data):
                    self.market_data[i] = value
            
            return True, "Market data updated successfully"
        except Exception as e:
            return False, f"Failed to update market data: {str(e)}"
    
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
        fig.add_trace(go.Scatter(
            x=grades,
            y=self.market_data[:len(grades)],
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
                grade_employees = self.employee_df[self.employee_df['GRADE'] == grade]
                
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
                            color='rgba(178, 34, 34, 0.8)',  # Firebrick red, more professional
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
        
        href = f'<a href="data:text/html;base64,{encoded}" download="salary_visualization.html" class="download-button">Download HTML File</a>'
        return href

def display_guide():
    """Display user guide"""
    st.title("Welcome to the Salary Visualization Tool")
    
    st.subheader("This tool helps you see how employee salaries compare with your company's salary ranges and market rates.")
    
    st.markdown("### How to Use This Tool:")
    
    st.markdown("""
    1️⃣ **Load Your Data** - Upload your Excel file with employee information. The file should have columns for employee ID, name, grade, and salary.
    
    2️⃣ **Check Salary Ranges (Optional)** - If needed, adjust the minimum, middle, or maximum salary for each job grade.
    
    3️⃣ **Update Market Rates (Optional)** - If you want to change the market comparison values, update the market data.
    
    4️⃣ **Create Your Chart** - Generate the salary chart. Each employee will show as a red dot, salary ranges as blue bars, and market rates as a blue line.
    
    5️⃣ **Save Your Work** - Download the chart as an HTML file you can open later in any web browser.
    """)
    
    st.markdown("### Understanding Your Chart:")
    
    st.markdown("""
    • **Blue Bars** - The salary range for each job grade (from minimum to maximum)
    • **Green Lines** - The midpoint salary for each job grade
    • **Red Dots** - Each employee's actual salary
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
        page_title="Salary Visualization Tool",
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
        st.session_state.tool = SalaryVisualizationTool()
    
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
            'Market 50th Percentile': st.session_state.tool.market_data[:len(st.session_state.tool.grade_df)]
        })
        
        edited_market_data = st.data_editor(
            market_data_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True
        )
        
        if st.button("Update Market Data"):
            new_market_data = edited_market_data['Market 50th Percentile'].tolist()
            success, message = st.session_state.tool.update_market_data(new_market_data)
            if success:
                st.success(message)
            else:
                st.error(message)
    
    elif page == "Visualization":
        st.title("Salary Visualization")
        
        # Check if employee data is loaded
        if st.session_state.tool.employee_df is None:
            st.warning("Please load employee data first. Go to Data Management to upload your data.")
        else:
            # Generate or refresh visualization
            if st.button("Generate Visualization") or st.session_state.visualization_generated:
                st.session_state.visualization_generated = True
                
                with st.spinner("Generating visualization..."):
                    fig = st.session_state.tool.generate_visualization()
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Add download button
                    download_link = st.session_state.tool.generate_download_link(fig)
                    st.markdown(download_link, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
