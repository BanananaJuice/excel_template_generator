import streamlit as st
from pathlib import Path
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from typing import Union, List
from frontend.components import (
    setup_page,
    action_button,
    success_message,
    error_message,
    info_card
)
from utils.excel_ops import ExcelHandler

def initialize_session_state():
    """Initialize session state variables if they don't exist."""
    if "template_file" not in st.session_state:
        st.session_state.template_file = None
    if "report_files" not in st.session_state:
        st.session_state.report_files = []
    if "mappings" not in st.session_state:
        st.session_state.mappings = {}
    if "report_mappings" not in st.session_state:
        st.session_state.report_mappings = {}
    if "current_sheet" not in st.session_state:
        st.session_state.current_sheet = None
    if "selected_cell" not in st.session_state:
        st.session_state.selected_cell = None
    if "temp_dir" not in st.session_state:
        # Create a temporary directory for uploaded files
        temp_dir = Path("temp_files")
        if not temp_dir.exists():
            temp_dir.mkdir(parents=True)
        st.session_state.temp_dir = temp_dir
    
    # Create assets directory if it doesn't exist
    assets_dir = Path("assets")
    if not assets_dir.exists():
        assets_dir.mkdir(parents=True)
    st.session_state.assets_dir = assets_dir
    
    # Create audio directory if it doesn't exist
    audio_dir = Path("audio")
    if not audio_dir.exists():
        audio_dir.mkdir(parents=True)
    st.session_state.audio_dir = audio_dir

def save_uploaded_file(uploaded_file, directory: Path) -> Path:
    """Save an uploaded file to the specified directory and return its path."""
    save_path = directory / uploaded_file.name
    with open(save_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return save_path

def get_excel_cell_reference(row_index: int, col_name: str, df: pd.DataFrame) -> str:
    """Convert row index and column name to Excel cell reference."""
    try:
        # Find the column index by matching the column name
        column_names = df.columns.tolist()
        column_index = column_names.index(col_name)
        # Convert to Excel column (A, B, C, ...)
        excel_column = ""
        while column_index >= 0:
            excel_column = chr(65 + (column_index % 26)) + excel_column
            column_index = (column_index // 26) - 1
        return f"{excel_column}{row_index + 1}"
    except Exception as e:
        st.error(f"Error calculating cell reference: {str(e)}")
        st.write("Debug info:")
        st.write("Row index:", row_index)
        st.write("Column name:", col_name)
        st.write("Available columns:", df.columns.tolist())
        return "A1"  # Default to A1 if there's an error

def get_all_fund_codes():
    """Get all unique fund codes from all reports."""
    all_fund_codes = set()
    for report_path in st.session_state.report_files:
        try:
            codes = ExcelHandler.get_fund_codes(report_path)
            if codes:
                st.write(f"Found {len(codes)} fund codes in {report_path.name}")
                all_fund_codes.update(codes)
            else:
                st.warning(f"No fund codes found in {report_path.name}")
        except Exception as e:
            st.error(f"Error processing {report_path.name}: {str(e)}")
            continue
    return sorted(list(all_fund_codes))

def main():
    setup_page()
    initialize_session_state()
    
    st.title("📊 Excel Template Inserter")
    
    # Sidebar for navigation
    with st.sidebar:
        st.header("Navigation")
        page = st.radio(
            "Choose a function:",
            ["File Management", "Report Management", "Mapping", "Generate Files", "About the Developer"]
        )
    
    if page == "File Management":
        st.header("File Management")
        
        # Template file upload
        template_file = st.file_uploader(
            "Upload Template File",
            type=["xlsx"],
            key="template_uploader"
        )
        
        if template_file:
            # Save template file and store its path
            template_path = save_uploaded_file(template_file, st.session_state.temp_dir)
            st.session_state.template_file = template_path
            st.success(f"✅ Template uploaded: {template_file.name}")
        
        st.markdown("---")
        
        # Report files upload
        report_files = st.file_uploader(
            "Upload Report Files",
            type=["xlsx"],
            accept_multiple_files=True,
            key="report_uploader"
        )
        
        if report_files:
            # Save report files and store their paths
            for file in report_files:
                report_path = save_uploaded_file(file, st.session_state.temp_dir)
                if not any(str(r) == str(report_path) for r in st.session_state.report_files):
                    st.session_state.report_files.append(report_path)
            st.success(f"✅ {len(report_files)} report(s) uploaded")

    elif page == "Report Management":
        st.header("Report Management")
        
        if not st.session_state.report_files:
            info_card(
                "No Reports",
                "Please upload report files in the File Management section."
            )
            return
        
        # Display report list
        st.subheader("Uploaded Reports")
        for idx, report in enumerate(st.session_state.report_files, 1):
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"📊 {idx}. {report.name}")
            with col2:
                if st.button("Remove", key=f"remove_{idx}"):
                    st.session_state.report_files.remove(report)
                    st.rerun()
        
        # Analyze fund codes
        st.subheader("Fund Code Analysis")
        try:
            fund_codes = ExcelHandler.analyze_fund_codes(st.session_state.report_files)
            if fund_codes:
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Fund Code**")
                    for code in fund_codes.keys():
                        st.write(code)
                with col2:
                    st.write("**Count**")
                    for count in fund_codes.values():
                        st.write(count)
                st.write(f"Total unique fund codes: **{len(fund_codes)}**")
            else:
                st.write("No fund codes found in the reports.")
        except Exception as e:
            error_message(f"Error analyzing fund codes: {str(e)}")

    elif page == "Mapping":
        st.header("Template Mapping")
        
        if not st.session_state.template_file:
            info_card(
                "No Template",
                "Please upload a template file in the File Management section."
            )
            return
            
        if not st.session_state.report_files:
            info_card(
                "No Reports",
                "Please upload report files in the File Management section."
            )
            return
        
        try:
            # Get available sheets and let user select one
            sheet_names = ExcelHandler.get_sheet_names(st.session_state.template_file)
            
            # Sheet selection at top of mapping page
            col1, col2 = st.columns([2, 1])
            with col1:
                st.subheader("Select Worksheet")
                selected_sheet = st.selectbox(
                    "Choose worksheet to map",
                    sheet_names,
                    index=sheet_names.index(st.session_state.current_sheet) if st.session_state.current_sheet in sheet_names else 0,
                    key="sheet_selector"
                )
            
            if selected_sheet != st.session_state.current_sheet:
                st.session_state.current_sheet = selected_sheet
                st.rerun()

            # Display Report List
            st.write("### Available Reports")
            report_names = [report.name for report in st.session_state.report_files]
            for idx, name in enumerate(report_names, 1):
                st.write(f"📊 {idx}. {name}")
            
            st.markdown("---")
            
            # Get template preview
            df_preview = ExcelHandler.get_template_preview(
                st.session_state.template_file,
                sheet_name=st.session_state.current_sheet,
                extra_space=5
            )
            
            # Display template preview
            st.write(f"### Template Preview - Sheet: {st.session_state.current_sheet}")
            st.dataframe(df_preview, use_container_width=True)
            
            st.markdown("---")
            
            # Add new mapping section
            st.write("### Add New Mapping")
            st.write("Enter a cell reference (e.g., A1, B2) and select a report to insert at that position.")
            
            col1, col2, col3 = st.columns([1, 2, 1])
            
            with col1:
                cell_ref = st.text_input(
                    "Cell Reference",
                    key="new_cell_ref",
                    help="Enter cell reference (e.g., A1, B2)"
                ).strip().upper()
            
            with col2:
                selected_report = st.selectbox(
                    "Select Report",
                    ["-- Select a Report --"] + report_names,
                    key="new_report_select"
                )
            
            with col3:
                if st.button("Add Mapping", key="add_mapping"):
                    if cell_ref and selected_report != "-- Select a Report --":
                        # Save report mapping
                        mapping_key = f"{st.session_state.current_sheet}!{cell_ref}"
                        st.session_state.report_mappings[mapping_key] = {
                            "report": selected_report,
                            "insert_position": cell_ref
                        }
                        st.success(f"✅ Report '{selected_report}' will be inserted at cell {cell_ref}")
                        st.rerun()
                    else:
                        st.error("Please enter both a cell reference and select a report.")
            
            # Display current mappings
            st.markdown("---")
            current_report_mappings = {k: v for k, v in st.session_state.report_mappings.items() 
                                     if k.startswith(f"{st.session_state.current_sheet}!")}
            
            if current_report_mappings:
                st.write("### 📋 Current Report Mappings")
                st.write(f"Sheet: **{st.session_state.current_sheet}**")
                
                # Create a table for mappings
                mapping_data = []
                for cell, mapping in current_report_mappings.items():
                    cell_ref = cell.split('!')[1]
                    mapping_data.append({
                        "Position": cell_ref,
                        "Report": mapping['report']
                    })
                
                if mapping_data:
                    st.dataframe(
                        pd.DataFrame(mapping_data),
                        use_container_width=True,
                        hide_index=True
                    )
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("🗑️ Clear Sheet Mappings"):
                        st.session_state.report_mappings = {k: v for k, v in st.session_state.report_mappings.items() 
                                                          if not k.startswith(f"{st.session_state.current_sheet}!")}
                        st.rerun()
                with col2:
                    if st.button("🗑️ Clear All Mappings"):
                        st.session_state.report_mappings = {}
                        st.rerun()
            else:
                st.info("No reports have been mapped yet. Add a mapping using the form above.")
                    
        except Exception as e:
            error_message(f"Error in mapping interface: {str(e)}")
    
    elif page == "Generate Files":
        st.header("Generate Files")
        if not st.session_state.template_file:
            info_card(
                "No Template",
                "Please upload a template file in the File Management section."
            )
            return
        
        if not st.session_state.report_files:
            info_card(
                "No Reports",
                "Please upload report files in the File Management section."
            )
            return
        
        if not st.session_state.report_mappings:
            info_card(
                "No Mappings",
                "Please create report mappings in the Mapping section first."
            )
            return

        try:
            st.write(f"Template: **{st.session_state.template_file.name}**")
            st.write(f"Number of reports: **{len(st.session_state.report_files)}**")
            
            # Get all fund codes
            all_fund_codes = get_all_fund_codes()
            
            if not all_fund_codes:
                st.error("No fund codes found in the reports. Please ensure your reports contain fund codes.")
                return
            
            # Fund code selection
            st.write("### Select Fund Codes")
            st.write("Choose which fund codes to generate templates for:")
            
            # Initialize session state for selected funds if not exists
            if "selected_funds" not in st.session_state:
                st.session_state.selected_funds = set(all_fund_codes)
            
            # Create columns for fund code selection
            cols = st.columns(3)
            fund_selections = {}
            
            # Select/Deselect All buttons
            col1, col2, _ = st.columns(3)
            with col1:
                if st.button("Select All"):
                    st.session_state.selected_funds = set(all_fund_codes)
                    st.rerun()
            with col2:
                if st.button("Deselect All"):
                    st.session_state.selected_funds = set()
                    st.rerun()
            
            # Display fund codes in columns with checkboxes
            for idx, fund_code in enumerate(sorted(all_fund_codes)):
                col_idx = idx % 3
                with cols[col_idx]:
                    is_selected = fund_code in st.session_state.selected_funds
                    if st.checkbox(fund_code, value=is_selected, key=f"fund_{fund_code}"):
                        st.session_state.selected_funds.add(fund_code)
                    else:
                        st.session_state.selected_funds.discard(fund_code)
            
            st.markdown("---")
            
            # Display generation options
            st.write("### Output Settings")
            
            col1, col2 = st.columns(2)
            with col1:
                output_dir = st.text_input(
                    "Output Directory",
                    value="output",
                    help="Name of the directory where generated files will be saved"
                )
            
            with col2:
                file_prefix = st.text_input(
                    "File Prefix",
                    value="populated_template",
                    help="Prefix for generated files (will be followed by fund code)"
                )
            
            # Show summary before generation
            st.write("### Generation Summary")
            st.write(f"- Selected Funds: **{len(st.session_state.selected_funds)}** of **{len(all_fund_codes)}**")
            st.write(f"- Output Directory: **{output_dir}**")
            st.write(f"- File Names: **{file_prefix}_[FUND_CODE].xlsx**")
            
            # Display mappings that will be used
            with st.expander("View Current Mappings"):
                for sheet in set(k.split('!')[0] for k in st.session_state.report_mappings.keys()):
                    st.write(f"**Sheet: {sheet}**")
                    sheet_mappings = {k.split('!')[1]: v for k, v in st.session_state.report_mappings.items() 
                                    if k.startswith(f"{sheet}!")}
                    for cell, mapping in sheet_mappings.items():
                        st.write(f"- Cell {cell}: {mapping['report']}")
            
            # Generate button
            if st.button("🔄 Generate Files", type="primary"):
                if not st.session_state.selected_funds:
                    st.error("Please select at least one fund code.")
                    return
                
                try:
                    # Create output directory
                    output_path = Path(output_dir)
                    if not output_path.exists():
                        output_path.mkdir(parents=True)
                    
                    # Initialize progress tracking
                    st.info("Starting file generation...")
                    progress_bar = st.progress(0)
                    status_container = st.empty()
                    error_container = st.empty()
                    
                    total_files = len(st.session_state.selected_funds)
                    successful_files = 0
                    failed_files = 0
                    
                    # Initialize report cache
                    report_cache = {}
                    
                    # Create mapping dictionary with actual file paths once
                    processed_mappings = {}
                    for key, mapping in st.session_state.report_mappings.items():
                        report_name = mapping['report']
                        report_path = st.session_state.temp_dir / report_name
                        if report_path.exists():
                            processed_mappings[key] = {
                                "report": str(report_path),
                                "insert_position": mapping['insert_position']
                            }
                    
                    # Process each fund code
                    for idx, fund_code in enumerate(sorted(st.session_state.selected_funds)):
                        try:
                            # Update status
                            status_container.write(f"Processing fund code: **{fund_code}** ({idx + 1}/{total_files})")
                            
                            # Generate filename
                            output_file = output_path / f"{file_prefix}_{fund_code}.xlsx"
                            
                            # Generate populated template with cache
                            report_cache = ExcelHandler.generate_populated_template(
                                template_path=st.session_state.template_file,
                                output_path=output_file,
                                report_mappings=processed_mappings,
                                fund_code=fund_code,
                                report_cache=report_cache
                            )
                            
                            successful_files += 1
                            
                        except Exception as e:
                            error_msg = f"Error generating file for fund {fund_code}: {str(e)}"
                            error_container.error(error_msg)
                            failed_files += 1
                        
                        # Update progress
                        progress = (idx + 1) / total_files
                        progress_bar.progress(progress)
                    
                    # Final status update
                    status_container.empty()
                    if successful_files > 0:
                        st.success(f"""✅ Generation Complete:
                        - Successfully generated: {successful_files} files
                        - Failed: {failed_files} files
                        - Output directory: '{output_dir}'""")
                    
                    if failed_files == 0:
                        st.balloons()
                    
                except Exception as e:
                    st.error(f"Error during file generation: {str(e)}")
                    st.write("Please check the error messages above and try again.")
        
        except Exception as e:
            error_message(f"Error in generate files: {str(e)}")
    
    elif page == "About the Developer":
        st.header("About the Developer")
        
        # Background Music
        audio_path = st.session_state.audio_dir / "background_music.mp3"
        if audio_path.exists():
            st.markdown(
                f"""
                <audio id="background-music" autoplay loop>
                    <source src="data:audio/mp3;base64,{open(audio_path, 'rb').read().hex()}" type="audio/mp3">
                    Your browser does not support the audio element.
                </audio>
                <script>
                    var audio = document.getElementById('background-music');
                    audio.volume = 0.5;  // Set volume to 50%
                </script>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info("🎵 Add 'background_music.mp3' to the audio directory to enable background music.")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            # Display the developer photo
            photo_path = st.session_state.assets_dir / "developer_photo.jpg"
            if photo_path.exists():
                st.image(photo_path, caption="Shailen Desai", use_container_width=True)
            else:
                st.error("Developer photo not found. Please ensure 'developer_photo.jpg' is in the assets directory.")
        
        with col2:
            st.markdown("""
            ### Shailen Desai
            
            Former EY Professional who completed his articles in 2024. With a passion for automation and 
            technology, Shailen developed this Excel Template Inserter to streamline financial reporting processes.
            
            #### Connect with Shailen:
            - 📸 Instagram: [@shailendesai1](https://instagram.com/shailendesai1)
            - 💼 LinkedIn: [Shailen Desai](https://www.linkedin.com/in/shailen-d-572300120/)
            
            #### About This Project
            This application was developed to automate the process of populating Excel templates with data 
            from multiple report files. It maintains all formatting while allowing for efficient handling of 
            multiple fund codes and reports.
            
            Feel free to reach out for any questions or collaboration opportunities!
            """)
            
            # Add direct social media buttons
            col1, col2 = st.columns(2)
            with col1:
                st.link_button("🔗 LinkedIn", "https://www.linkedin.com/in/shailen-d-572300120/")
            with col2:
                st.link_button("📸 Instagram", "https://instagram.com/shailendesai1")

if __name__ == "__main__":
    main() 