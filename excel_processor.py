
import pandas as pd
from pathlib import Path
from typing import List, Dict
import re
import logging

class ExcelProcessor:
    def __init__(self, file_path: Path):
        """
        Initialize ExcelProcessor with the path to the Excel file.
        
        Args:
            file_path (Path): Path to the Excel file.
        
        Raises:
            FileNotFoundError: If the Excel file doesn't exist.
            ValueError: If the file is not a valid Excel file.
        """
        self.logger = logging.getLogger(__name__)
        if not file_path.exists():
            raise FileNotFoundError(f"Excel file not found at {file_path}")
        
        try:
            self.df = pd.read_excel(file_path, sheet_name="CapBudgWS", header=None)
            self.tables = self._parse_tables()
        except Exception as e:
            raise ValueError(f"Error reading Excel file: {str(e)}")

    def _parse_tables(self) -> Dict[str, pd.DataFrame]:
        """
        Parse the Excel sheet to identify and extract tables.
        
        Returns:
            Dict[str, pd.DataFrame]: Dictionary of table names and their corresponding DataFrames.
        """
        tables = {}
        current_table = None
        start_row = None
        table_pattern = re.compile(r'^(INITIAL INVESTMENT|SALVAGE VALUE|OPERATING CASHFLOWS|BOOK VALUE & DEPRECIATION)$')
        
        for idx, row in self.df.iterrows():
            first_cell = str(row[0]).strip() if pd.notnull(row[0]) else ""
            
            # Check for table headers
            if table_pattern.match(first_cell):
                if current_table and start_row is not None:
                    # Save the previous table
                    tables[current_table] = self.df.iloc[start_row:idx].reset_index(drop=True)
                current_table = first_cell
                start_row = idx + 1  # Start after the header row
            elif current_table and first_cell == "":
                # End of current table
                if start_row is not None:
                    tables[current_table] = self.df.iloc[start_row:idx].reset_index(drop=True)
                current_table = None
                start_row = None
        
        # Save the last table
        if current_table and start_row is not None:
            tables[current_table] = self.df.iloc[start_row:].reset_index(drop=True)
        
        return tables

    def get_table_names(self) -> List[str]:
        """
        Get the list of table names in the Excel sheet.
        
        Returns:
            List[str]: List of table names.
        """
        return list(self.tables.keys())

    def get_row_names(self, table_name: str) -> List[str]:
        """
        Get the row names for a specific table.
        
        Args:
            table_name (str): Name of the table.
        
        Returns:
            List[str]: List of row names from the first column.
        
        Raises:
            ValueError: If the table name is invalid.
        """
        if table_name not in self.tables:
            raise ValueError(f"Table '{table_name}' not found")
        
        table_df = self.tables[table_name]
        # Extract first column, remove NaN, and convert to strings
        row_names = [str(row).strip() for row in table_df[0] if pd.notnull(row)]
        return row_names

    def calculate_row_sum(self, table_name: str, row_name: str) -> float:
        """
        Calculate the sum of numerical values in the specified row.
        
        Args:
            table_name (str): Name of the table.
            row_name (str): Name of the row to sum.
        
        Returns:
            float: Sum of numerical values in the row.
        
        Raises:
            ValueError: If table or row is not found or if no numerical values exist.
        """
        if table_name not in self.tables:
            raise ValueError(f"Table '{table_name}' not found")
        
        table_df = self.tables[table_name]
        row_names = [str(row).strip() for row in table_df[0] if pd.notnull(row)]
        
        if row_name not in row_names:
            raise ValueError(f"Row '{row_name}' not found in table '{table_name}'")
        
        row_idx = row_names.index(row_name)
        row_data = table_df.iloc[row_idx, 1:]  # Skip the first column (row name)
        
        # Convert to numeric, ignoring non-numeric values
        numeric_values = pd.to_numeric(row_data, errors='coerce')
        sum_value = numeric_values.sum()
        
        if pd.isna(sum_value):
            raise ValueError(f"No numerical values found in row '{row_name}'")
        
        return float(sum_value)
