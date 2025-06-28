import streamlit as st
import pandas as pd
import json
import io
from datetime import datetime
from textblob import TextBlob
import re

# Initialize session state
if 'template_columns' not in st.session_state:
    st.session_state.template_columns = []
if 'current_data' not in st.session_state:
    st.session_state.current_data = None
if 'show_help' not in st.session_state:
    st.session_state.show_help = False

# Enhanced NLP Processing Function with 50+ cases
def process_nlp_query(query):
    query = query.lower().strip()
    blob = TextBlob(query)
    words = [str(word).lower() for word in blob.words]
    response = []
    
    # Greetings (5 cases)
    if any(word in ['hi', 'hello', 'hey', 'greetings', 'hola'] for word in words):
        response.append("Hello! How can I assist you with the data conversion tool today?")
    
    # Help requests (8 cases)
    if any(word in ['help', 'assist', 'support', 'guide', 'tutorial', 'howto', 'how to', 'documentation'] for word in words):
        response.append("""
        I can help you with:
        - Creating custom Excel templates (add columns, set formats)
        - Converting Excel/CSV to JSON
        - Flattening nested data structures
        - Troubleshooting conversion issues
        - Understanding template requirements
        """)
    
    # Template creation (12 cases)
    if any(word in ['template', 'create', 'make', 'excel', 'sheet', 'spreadsheet', 'design', 
                   'build', 'generate', 'new', 'column', 'header'] for word in words):
        response.append("""
        **Template Creation Guide:**
        1. Go to 'Create Template' section
        2. Add your columns with names and optional sample values
        3. Generate and download the template
        4. Fill it with your data
        5. Upload back to convert to JSON
        
        Pro Tip: Use descriptive column names like 'PolicyId' or 'MonthlyAllowance'
        """)
    
    # File conversion (10 cases)
    if any(word in ['convert', 'excel', 'csv', 'json', 'upload', 'download', 'transform', 
                   'change', 'export', 'import'] for word in words):
        response.append("""
        **File Conversion Guide:**
        1. Upload your Excel/CSV file in 'Convert File' section
        2. View the preview in table or JSON format
        3. Download the converted JSON or CSV
        
        Supported formats: .xlsx, .xls, .csv
        Max file size: 200MB
        """)
    
    # Technical issues (8 cases)
    if any(word in ['error', 'problem', 'issue', 'bug', 'fix', 'trouble', 'not working', 'fail'] for word in words):
        response.append("""
        **Troubleshooting Guide:**
        Common issues and solutions:
        - File not uploading: Check format (.xlsx, .csv) and size
        - Encoding problems: Save CSV as UTF-8
        - Blank output: Check your file has data below headers
        - Column mismatch: Keep header row unchanged
        
        For persistent issues, contact support@excel2json.com
        """)
    
    # Data structure (7 cases)
    if any(word in ['structure', 'format', 'schema', 'layout', 'design', 'flatten', 'nested'] for word in words):
        response.append("""
        **Data Structure Information:**
        - The converter creates flattened JSON (no nested structures)
        - Each Excel row becomes a JSON object
        - Column headers become JSON keys
        - Empty cells become null values
        
        Example: Excel becomes {"PolicyId":"01","Role":"ResearchAssistant"}
        """)
    
    # Advanced features (6 cases)
    if any(word in ['advanced', 'special', 'custom', 'validation', 'dropdown', 'formula'] for word in words):
        response.append("""
        **Advanced Features:**
        - Add data validation in Excel (dropdowns, number ranges)
        - Use formulas for calculated fields
        - Format dates consistently (YYYY-MM-DD recommended)
        - For complex structures, pre-flatten your data
        
        Note: Formulas are calculated before conversion
        """)
    
    # Thank you (4 cases)
    if any(word in ['thanks', 'thank', 'appreciate', 'grateful'] for word in words):
        response.append("You're welcome! Let me know if you need anything else.")
    
    # Fallback if no matches
    if not response:
        # Try to detect similar words
        similar_words = {
            'convert': ['change', 'transform', 'export'],
            'template': ['form', 'sheet', 'design'],
            'error': ['problem', 'issue', 'bug']
        }
        
        suggestions = []
        for word in words:
            for key, variants in similar_words.items():
                if word in variants:
                    suggestions.append(f"Try asking about '{key}'")
        
        if suggestions:
            response.append("I'm not sure I understand. " + " ".join(set(suggestions)))
        else:
            response.append("""
            I'm not sure I understand. Try asking about:
            - Creating templates
            - Converting files
            - Troubleshooting issues
            - Data formatting requirements
            """)
    
    return response

# Helper function to convert CSV to flattened JSON
def csv_to_flattened_json(csv_content):
    try:
        # Read CSV with proper encoding handling
        df = pd.read_csv(io.StringIO(csv_content.decode('utf-8')))
        
        # Convert to list of dictionaries (one per row)
        json_data = df.to_dict(orient='records')
        
        # For display purposes, we'll show both table and JSON views
        return json_data, df, None
    except Exception as e:
        return None, None, str(e)

# Enhanced template creator
def create_excel_template(columns):
    # Create a DataFrame with sample data
    data = {col['name']: [col.get('sample_value', '')] for col in columns}
    df = pd.DataFrame(data)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        
        # Add formatting
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        # Header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1
        })
        
        # Write column headers with format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, max(15, len(value) * 1.2))
        
        # Add instructions sheet
        instructions = [
            ["Template Instructions"],
            ["1. Fill data in the 'Data' sheet below the headers"],
            ["2. Do not modify or delete the header row"],
            ["3. Save the file when finished"],
            ["4. Upload back to this tool for conversion to JSON"],
            ["", ""],
            ["Template Info:"],
            [f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"],
            [f"Total columns: {len(columns)}"]
        ]
        
        instructions_df = pd.DataFrame(instructions)
        instructions_df.to_excel(writer, index=False, header=False, sheet_name='Instructions')
    
    output.seek(0)
    return output

# Streamlit UI
st.set_page_config(
    page_title="Excel to JSON Converter",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Main title
st.title("🔄 Excel/CSV to JSON Converter")
st.markdown("""
**Create custom templates or convert your existing files**  
*Design Excel templates with your preferred structure and convert them to flattened JSON*
""")

# Help button in top right
if st.sidebar.button("❓ Get Help"):
    st.session_state.show_help = not st.session_state.show_help

# Sidebar navigation
st.sidebar.title("Navigation")
app_mode = st.sidebar.radio(
    "Choose mode:",
    ["📝 Create Template", "🔄 Convert File"],
    index=0
)

# Help section (shown when help button is clicked)
if st.session_state.show_help:
    with st.container():
        st.header("❓ Help & Assistance")
        col1, col2 = st.columns([4, 1])
        with col1:
            user_query = st.text_input("Ask your question here and press Enter:", key="help_query")
        with col2:
            st.write("")  # Spacer
            st.write("")  # Spacer
            if st.button("Close Help", key="close_help"):
                st.session_state.show_help = False
                st.rerun()
        
        if user_query:
            st.markdown("### 🤖 Response:")
            responses = process_nlp_query(user_query)
            for resp in responses:
                st.info(resp)
        
        st.markdown("### 📚 Common Questions:")
        with st.expander("Template Creation"):
            st.markdown("""
            - **How do I create a template?**  
              Go to 'Create Template' and add your columns with optional sample values.
              
            - **Can I add data validation?**  
              Yes, set up validation rules in Excel after downloading the template.
              
            - **What column names should I use?**  
              Use descriptive names like "PolicyId", "Role", "MonthlyAllowance".
            """)
        
        with st.expander("File Conversion"):
            st.markdown("""
            - **How do I convert Excel to JSON?**  
              Upload your file in 'Convert File' section and download the JSON output.
              
            - **What file formats are supported?**  
              Excel (.xlsx, .xls) and CSV files.
              
            - **How is nested data handled?**  
              The converter creates flattened structures (no nested objects/arrays).
            """)
        
        with st.expander("Troubleshooting"):
            st.markdown("""
            - **My file isn't uploading**  
              Check the file format and size (max 200MB). For CSV, ensure UTF-8 encoding.
              
            - **The output looks wrong**  
              Verify your Excel headers match exactly with the template columns.
              
            - **Dates aren't converting properly**  
              Format dates consistently in Excel before uploading.
            """)

# Main content based on navigation
if app_mode == "📝 Create Template":
    st.header("📝 Create Your Custom Template")
    
    with st.expander("➕ Add New Column", expanded=True):
        col1, col2 = st.columns([3, 1])
        with col1:
            new_col_name = st.text_input(
                "Column Name",
                placeholder="e.g., PolicyId, Role, MonthlyAllowance",
                key="new_col_name"
            )
        with col2:
            add_sample = st.checkbox("Add sample value", value=True)
        
        if add_sample:
            sample_value = st.text_input(
                "Sample value for this column",
                placeholder="e.g., 01, ResearchAssistant, PKR 8,000",
                key="sample_value"
            )
        else:
            sample_value = ""
        
        if st.button("Add Column", type="primary"):
            if new_col_name:
                st.session_state.template_columns.append({
                    'name': new_col_name,
                    'sample_value': sample_value
                })
                st.success(f"Column '{new_col_name}' added!")
                st.rerun()
            else:
                st.warning("Please enter a column name")

    # Display current columns
    if st.session_state.template_columns:
        st.subheader("Your Template Columns")
        
        # Create a table to show columns
        cols_table = []
        for idx, col in enumerate(st.session_state.template_columns):
            cols_table.append({
                "Column #": idx + 1,
                "Name": col['name'],
                "Sample Value": col.get('sample_value', '')
            })
        
        st.dataframe(
            pd.DataFrame(cols_table),
            use_container_width=True,
            hide_index=True
        )
        
        # Column management buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Clear All Columns", type="secondary"):
                st.session_state.template_columns = []
                st.rerun()
        with col2:
            if st.button("Generate Template", type="primary"):
                with st.spinner("Creating your template..."):
                    excel_file = create_excel_template(st.session_state.template_columns)
                    
                    st.success("Template created successfully!")
                    st.download_button(
                        label="⬇️ Download Excel Template",
                        data=excel_file,
                        file_name="custom_template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Show preview
                    st.subheader("Template Preview")
                    preview_df = pd.DataFrame({
                        col['name']: [col.get('sample_value', '')] 
                        for col in st.session_state.template_columns
                    })
                    st.dataframe(preview_df, hide_index=True)
    else:
        st.info("ℹ️ Add columns to start building your template")

# File Converter Section
else:
    st.header("🔄 Convert Excel/CSV to JSON")
    
    uploaded_file = st.file_uploader(
        "Upload your Excel or CSV file",
        type=["xlsx", "xls", "csv"],
        help="Supported formats: Excel (.xlsx, .xls) or CSV"
    )
    
    if uploaded_file:
        # Process the file
        with st.spinner("Processing your file..."):
            if uploaded_file.name.endswith('.csv'):
                # For CSV files
                json_data, df, error = csv_to_flattened_json(uploaded_file.getvalue())
            else:
                # For Excel files
                try:
                    df = pd.read_excel(uploaded_file)
                    json_data = df.to_dict(orient='records')
                    error = None
                except Exception as e:
                    json_data, df, error = None, None, str(e)
        
        if error:
            st.error(f"Error processing file: {error}")
        else:
            st.session_state.current_data = json_data
            
            st.success("✅ File processed successfully!")
            st.subheader("Data Preview")
            
            # Show data in tabs
            tab1, tab2 = st.tabs(["📊 Table View", "📄 JSON View"])
            
            with tab1:
                st.dataframe(df, use_container_width=True)
                
                # Download as CSV
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="⬇️ Download as CSV",
                    data=csv,
                    file_name="converted_data.csv",
                    mime="text/csv"
                )
            
            with tab2:
                st.json(json_data, expanded=False)
                
                # Download as JSON
                st.download_button(
                    label="⬇️ Download as JSON",
                    data=json.dumps(json_data, indent=2),
                    file_name="converted_data.json",
                    mime="application/json"
                )
            
            # Show conversion stats
            with st.expander("🔍 Conversion Details"):
                st.write(f"**File Name:** {uploaded_file.name}")
                st.write(f"**Total Rows:** {len(df)}")
                st.write(f"**Total Columns:** {len(df.columns)}")
                st.write("**Columns:**", ", ".join(df.columns.tolist()))

# Footer
st.sidebar.markdown("---")
st.sidebar.info("""
**Quick Help:**
- Type "how to create template"
- Ask "what file formats supported"
- Try "help with conversion error"
""")

st.markdown("---")
st.caption("© 2024 Excel to JSON Converter | For support contact: tahir12721@gmail.com")