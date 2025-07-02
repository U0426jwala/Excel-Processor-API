from fastapi import FastAPI, HTTPException
from typing import List, Dict
from excel_processor import ExcelProcessor
from pathlib import Path
import logging

app = FastAPI(title="Excel Processor API")

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize ExcelProcessor with the path to the Excel file
excel_file_path = Path("Data/capbudg.xlsx")  # Updated to .xlsx
processor = ExcelProcessor(excel_file_path)

@app.get("/list_tables")
async def list_tables() -> Dict[str, List[str]]:
    """
    Returns a list of table names present in the Excel sheet.
    """
    try:
        tables = processor.get_table_names()
        return {"tables": tables}
    except Exception as e:
        logger.error(f"Error listing tables: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error listing tables: {str(e)}")

@app.get("/get_table_details")
async def get_table_details(table_name: str) -> Dict[str, str | List[str]]:
    """
    Returns the row names for the specified table.
    
    Args:
        table_name (str): Name of the table to retrieve row names from.
    """
    try:
        row_names = processor.get_row_names(table_name)
        if not row_names:
            raise HTTPException(status_code=404, detail="Table not found or no row names available")
        return {"table_name": table_name, "row_names": row_names}
    except ValueError as ve:
        logger.error(f"Invalid table name: {str(ve)}")
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        logger.error(f"Error getting table details: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error getting table details: {str(e)}")

@app.get("/row_sum")
async def row_sum(table_name: str, row_name: str) -> Dict[str, str | float]:
    """
    Calculates the sum of numerical values in the specified row of the given table.
    
    Args:
        table_name (str): Name of the table.
        row_name (str): Name of the row to sum.
    """
    try:
        sum_value = processor.calculate_row_sum(table_name, row_name)
        return {"table_name": table_name, "row_name": row_name, "sum": sum_value}
    except ValueError as ve:
        logger.error(f"Invalid input: {str(ve)}")
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        logger.error(f"Error calculating row sum: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error calculating row sum: {str(e)}")