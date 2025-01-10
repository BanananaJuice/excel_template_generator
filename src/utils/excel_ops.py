import pandas as pd
import xlwings as xw
from pathlib import Path
from typing import Union, List, Dict, Tuple
from collections import Counter
import string
import shutil

class ExcelHandler:
    @staticmethod
    def _get_excel_col_name(n: int) -> str:
        """Convert column number to Excel column name (A, B, C, ..., Z, AA, AB, ...)."""
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    @staticmethod
    def read_excel(file_path: Union[str, Path], sheet_name: str = None) -> pd.DataFrame:
        """Read an Excel file and return a DataFrame."""
        try:
            if sheet_name is None:
                # Get first sheet if none specified
                xl = pd.ExcelFile(file_path)
                sheet_name = xl.sheet_names[0]
            return pd.read_excel(file_path, sheet_name=sheet_name)
        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")

    @staticmethod
    def get_sheet_names(file_path: Union[str, Path]) -> List[str]:
        """Get all sheet names from an Excel file."""
        try:
            xl = pd.ExcelFile(file_path)
            return xl.sheet_names
        except Exception as e:
            raise Exception(f"Error getting sheet names: {str(e)}")

    @staticmethod
    def write_excel(df: pd.DataFrame, file_path: Union[str, Path]) -> None:
        """Write a DataFrame to an Excel file."""
        try:
            df.to_excel(file_path, index=False)
        except Exception as e:
            raise Exception(f"Error writing Excel file: {str(e)}")

    @staticmethod
    def get_template_preview(template_path: Union[str, Path], sheet_name: str = None, extra_space: int = 5) -> pd.DataFrame:
        """Get template preview with additional columns and rows at the bottom and right."""
        try:
            df = ExcelHandler.read_excel(template_path, sheet_name=sheet_name)
            if df.empty:
                return pd.DataFrame()
            
            # Get original dimensions
            num_cols = len(df.columns)
            num_rows = len(df)
            
            # Create column names for original and extra columns
            all_cols = []
            
            # Add original columns with Excel-style names
            for i in range(num_cols):
                col_name = ExcelHandler._get_excel_col_name(i + 1)
                all_cols.append(col_name)
            
            # Add extra columns with Excel-style names
            for i in range(extra_space):
                col_name = ExcelHandler._get_excel_col_name(num_cols + i + 1)
                all_cols.append(col_name)
            
            # Create empty dataframe with extra rows and columns
            total_rows = num_rows + extra_space
            result = pd.DataFrame('', index=range(total_rows), columns=all_cols)
            
            # Fill in the original data
            original_data = pd.DataFrame(df.values, columns=all_cols[:num_cols])
            result.iloc[:num_rows, :num_cols] = original_data
            
            return result
        except Exception as e:
            raise Exception(f"Error creating template preview: {str(e)}")

    @staticmethod
    def get_column_names(file_path: Union[str, Path], sheet_name: str = None) -> List[str]:
        """Get column names from an Excel file."""
        try:
            df = ExcelHandler.read_excel(file_path, sheet_name=sheet_name)
            return df.columns.tolist()
        except Exception as e:
            raise Exception(f"Error getting column names: {str(e)}")

    @staticmethod
    def get_template_dimensions(template_path: Union[str, Path], sheet_name: str = None) -> Tuple[int, int]:
        """Get the dimensions of the template file."""
        try:
            df = ExcelHandler.read_excel(template_path, sheet_name=sheet_name)
            return len(df), len(df.columns)
        except Exception as e:
            raise Exception(f"Error getting template dimensions: {str(e)}")

    @staticmethod
    def get_fund_codes(file_path: Union[str, Path], sheet_name: str = None) -> List[str]:
        """Extract unique fund codes from the first column of an Excel file."""
        try:
            # For reports, always read the first (and only) sheet
            df = pd.read_excel(file_path)
            
            # Debug information
            print(f"Reading file: {file_path}")
            print(f"Columns found: {df.columns.tolist()}")
            
            if df.empty:
                print("DataFrame is empty")
                return []
            
            # Get the first column values, skipping the header row
            first_col = df.iloc[1:, 0]  # Skip first row as it's likely a header
            
            # Debug the first few values
            print("First few values in first column:")
            for idx, val in first_col[:5].items():
                print(f"Row {idx}: {val} (type: {type(val)})")
            
            # Convert to strings and clean up
            fund_codes = []
            for code in first_col:
                if pd.notna(code):  # Check if not NaN
                    str_code = str(code).strip()
                    if str_code and str_code.lower() != 'nan':
                        fund_codes.append(str_code)
            
            # Get unique values
            unique_codes = list(set(fund_codes))
            
            print(f"Found {len(unique_codes)} unique fund codes: {unique_codes}")
            
            return sorted(unique_codes)
        except Exception as e:
            print(f"Error in get_fund_codes: {str(e)}")
            raise Exception(f"Error extracting fund codes: {str(e)}")

    @staticmethod
    def analyze_fund_codes(files: List[Union[str, Path]], sheet_name: str = None) -> Dict[str, int]:
        """Analyze fund codes across multiple files and return counts."""
        all_fund_codes = []
        for file in files:
            try:
                # For reports, we don't need to check multiple sheets
                fund_codes = ExcelHandler.get_fund_codes(file)
                if fund_codes:
                    print(f"Found fund codes in {file}: {fund_codes}")
                    all_fund_codes.extend(fund_codes)
            except Exception as e:
                print(f"Error processing file {file}: {str(e)}")
                continue
        return dict(Counter(all_fund_codes))

    @staticmethod
    def generate_populated_template(
        template_path: Union[str, Path],
        output_path: Union[str, Path],
        report_mappings: Dict,
        fund_code: str,
        report_cache: Dict = None
    ) -> None:
        """
        Generate a populated template for a specific fund code while preserving formatting.
        Uses caching to avoid reading the same report multiple times.
        
        Args:
            template_path: Path to the template Excel file
            output_path: Path where the populated file should be saved
            report_mappings: Dictionary of mappings (sheet!cell -> report)
            fund_code: The fund code to filter data by
            report_cache: Optional cache of report DataFrames
        """
        # First, copy the template to preserve all formatting
        shutil.copy2(template_path, output_path)
        
        try:
            # Initialize cache if not provided
            if report_cache is None:
                report_cache = {}
            
            # Group mappings by report to minimize file reads
            report_to_mappings = {}
            for mapping_key, mapping_info in report_mappings.items():
                report_path = mapping_info['report']
                if report_path not in report_to_mappings:
                    report_to_mappings[report_path] = []
                report_to_mappings[report_path].append((mapping_key, mapping_info))
            
            # Use xlwings to work with the Excel file
            with xw.App(visible=False) as app:
                wb = app.books.open(str(output_path))
                
                # Process each report
                for report_path, mappings in report_to_mappings.items():
                    try:
                        # Use cached data if available
                        if report_path in report_cache:
                            report_df = report_cache[report_path]
                        else:
                            report_df = pd.read_excel(report_path)
                            report_cache[report_path] = report_df
                        
                        # Filter data once per report
                        filtered_df = report_df[report_df.iloc[:, 0] == fund_code]
                        
                        if not filtered_df.empty:
                            # Convert filtered data to values once
                            values = filtered_df.values.tolist()
                            
                            # Group mappings by sheet
                            sheet_mappings = {}
                            for mapping_key, mapping_info in mappings:
                                sheet_name, cell_ref = mapping_key.split('!')
                                if sheet_name not in sheet_mappings:
                                    sheet_mappings[sheet_name] = []
                                sheet_mappings[sheet_name].append((cell_ref, values))
                            
                            # Process each sheet
                            for sheet_name, sheet_mappings in sheet_mappings.items():
                                sheet = wb.sheets[sheet_name]
                                
                                # Process all mappings for this sheet
                                for cell_ref, values in sheet_mappings:
                                    paste_range = sheet.range(cell_ref).resize(len(values), len(values[0]))
                                    paste_range.clear_contents()
                                    paste_range.value = values
                                    
                    except Exception as e:
                        raise Exception(f"Error processing report {report_path}: {str(e)}")
                
                # Save and close
                wb.save()
                wb.close()
                
        except Exception as e:
            # Clean up the output file if there was an error
            if Path(output_path).exists():
                Path(output_path).unlink()
            raise Exception(f"Error generating populated template: {str(e)}")
        
        return report_cache  # Return the cache for reuse

    @staticmethod
    def filter_report_by_fund(report_path: Union[str, Path], fund_code: str) -> pd.DataFrame:
        """Filter a report's data by fund code."""
        try:
            df = pd.read_excel(report_path)
            return df[df.iloc[:, 0] == fund_code]
        except Exception as e:
            raise Exception(f"Error filtering report: {str(e)}") 