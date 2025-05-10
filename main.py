from fastapi import FastAPI, HTTPException, Query, Depends
from typing import List, Dict, Any, Optional
import pandas as pd
import numpy as np
import uvicorn
import os
from pydantic import BaseModel

app = FastAPI(
    title="Excel Data Processor API",
    description="API for processing capital budgeting Excel data",
    version="1.0.0"
)

class TableResponse(BaseModel):
    """
    Pydantic model defining the structure for table metadata responses.
    
    Attributes:
        table_name (str): The name of the table being described.
                         Example: "INITIAL INVESTMENT"
                         
        row_names (List[str]): List of all available row/rows  within the table.
                              Example: ["Initial Investment", "Opportunity Cost", ...]
    """
    table_name: str
    row_names: List[str]

class RowSumResponse(BaseModel):
    """
    Pydantic model defining the structure for numeric summation responses.
    
    Attributes:
        table_name (str): The source table containing the analyzed row.
                         Example: "INITIAL INVESTMENT"
                         
        row_name (str): The specific row that was summed.
                        Example: "Opportunity Cost (if any)"
                        
        sum (float): The calculated sum of all numeric values in the row.
                     
    """
    table_name: str
    row_name: str
    sum: float

def get_parsed_data():
    """
    Dependency function that provides parsed Excel data to route handlers.
    
    Returns:
        Dict[str, Any]: Nested dictionary containing all parsed financial data.
        
    Raises:
        HTTPException: 500 status if file parsing fails, with error details.
    """
    file_path = os.path.join("Data", "capbudg.xls")
    try:
        return parse_capbudg(file_path)
    except Exception as e:
        raise HTTPException(
            status_code = 500,
            detail = f"Failed to parse Excel file: {str(e)}"
        )

@app.get("/list_tables", response_model=Dict[str, List[str]])
def list_tables(data: Dict[str, Any] = Depends(get_parsed_data)):
    """
    Retrieve all available table names from the parsed dataset.
    
    Args:
        data (Dict): Injected parsed data
        
    Returns:
        Dict: {"tables": List[str]} containing all top-level table names.
        
    Raises:
        HTTPException: 404 if no data is available
        
    Example Response:
        {
            "tables": [
                "INITIAL INVESTMENT",
                "CASHFLOW DETAILS",
                "DISCOUNT RATE",
                ...
            ]
        }
    """
    if not data:
        raise HTTPException(
            status_code=404,
            detail="No data available"
        )
    return {"tables": list(data.keys())}

@app.get("/get_table_details", response_model=TableResponse)
def get_table_details(
    table_name: str = Query(..., description="Name of the table to inspect"),
    data: Dict[str, Any] = Depends(get_parsed_data)
):
    """
    Retrieve adata for a specific table.
    
    Args:
        table_name: Name of target table (case-sensitive)
        data: Injected parsed data
        
    Returns:
        TableResponse: Structured response containing:
                      - table_name: Confirmed table identifier
                      - row_names: Available rows/columns in table
        
    Raises:
        HTTPException: 
            - 404 if table doesn't exist
            - 404 if no data loaded
            
    Example Usage:
        GET /get_table_details?table_name=DISCOUNT_RATE
        
    Example Response:
        {
            "table_name": "DISCOUNT RATE",
            "row_names": [
                "Approach",
                "Discount rate",
                "Beta",
                ...
            ]
        }
    """
    if not data:
        raise HTTPException(
            status_code=404,
            detail="No data available"
        )
    
    table_data = data.get(table_name)
    if not table_data:
        available_tables = list(data.keys())
        raise HTTPException(
            status_code=404,
            detail=(
                f"Table '{table_name}' not found. "
                f"Available tables: {available_tables}"
            )
        )
    
    # Extract row names based on data structure
    row_names = []
    if isinstance(table_data, dict):
        row_names = list(table_data.keys())
    elif isinstance(table_data, list) and table_data and isinstance(table_data[0], dict):
        row_names = list(table_data[0].keys())
    
    return TableResponse(
        table_name = table_name,
        row_names = row_names
    )

@app.get("/row_sum", response_model=RowSumResponse)
def row_sum(
    table_name: str = Query(..., description="Name of the table"),
    row_name: str = Query(..., description="Name of the row to sum"),
    data: Dict[str, Any] = Depends(get_parsed_data)
):
    """
    Calculates the sum of all numerical values in a specified financial data row.
    
    Processes both scalar values and lists of values, automatically handling:
    - Regular numbers (int/float)
    - Percentage values (e.g., '10%' → 0.10)
    - String-encoded numbers (e.g., '15.75' → 15.75)
    - Numpy numeric types

    Args:
        table_name: Identifier of the financial table (case-sensitive)
        row_name: Target row containing numeric data
        data: Injected parsed dataset from Excel

    Returns:
        RowSumResponse: Structured response containing:
            - table_name: Verified source table
            - row_name: Verified target row
            - sum: Calculated total (rounded to 2 decimal places)

    Raises:
        HTTPException: 
            - 404 if dataset not loaded
            - 404 if table doesn't exist
            - 404 if row doesn't exist
            - 400 if no numeric values found

    Example Request:
        GET /row_sum?table_name=INITIAL INVESTMENT&row_name=Tax Credit

    Example Response:
        {
            "table_name": "INITIAL INVESTMENT",
            "row_name": "Tax Credit",
            "sum": 10.0
        }
    """
    if not data:
        raise HTTPException(
            status_code=404,
            detail="No financial data available. Please verify Excel file is loaded."
        )
    
    # Validate table exists
    table_data = data.get(table_name)
    if not table_data:
        available_tables = list(data.keys())
        raise HTTPException(
            status_code=404,
            detail=(
                f"Financial table '{table_name}' not found. "
                f"Available tables: {available_tables}"
            )
        )
    
    numeric_sum = 0.0
    found = False  # Track if any numeric values were processed
    
    # Handle dictionary-style tables or dict within dict (e.g., INITIAL INVESTMENT)
    if isinstance(table_data, dict):
        row_value = table_data.get(row_name)
        if row_value is None:
            raise HTTPException(
                status_code=404,
                detail =(
                    f"Row '{row_name}' not found in table '{table_name}'. "
                    f"Use /get_table_details to view available rows."
                )
            )
        # Process list values within the outer dict(e.g., GROWTH RATES)
        if isinstance(row_value, list):
            for v in row_value:
                coerced = coerce_to_float(v)
                if coerced is not None:
                    numeric_sum += coerced
                    found = True
        # Process single value within the dict
        else:
            coerced = coerce_to_float(row_value)
            if coerced is not None:
                numeric_sum += coerced
                found = True
    
    # Handle list-of-dicts tables (e.g., OPERATING CASHFLOWS)
    elif isinstance(table_data, list):
        for record in table_data:
            if row_name in record:
                coerced = coerce_to_float(record[row_name])
                if coerced is not None:
                    numeric_sum += coerced
                    found = True
    
    if not found:
        raise HTTPException(
            status_code=400,
            detail =(
                f"No numeric values found in row '{row_name}' of table '{table_name}'. "
                "Please verify the row contains numerical data."
            )
        )
    
    return RowSumResponse(
        table_name=table_name,
        row_name=row_name,
        sum=round(numeric_sum, 2)
    )

def coerce_to_float(value) -> Optional[float]:
    """
    Safely converts diverse input types to float, with percentage handling.
    
    Supports:
    - Native Python numbers (int, float)
    - Numpy numeric types
    - Percentage strings (e.g., "15%")
    - Regular numeric strings (e.g., "3.14")
    - None values (returns None)

    Args:
        value: Input value to convert

    Returns:
        float: Converted numeric value
        None: If conversion fails or input is None

    Examples:
        >>> coerce_to_float(10)
        10.0
        >>> coerce_to_float("5.5%")
        0.055
        >>> coerce_to_float("N/A")
        None
    """
    if value is None:
        return None
    
    # Handle native numeric types
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    try:
        str_val = str(value).strip()

        # Convert string percentages to decimal (e.g., "10%" → 0.10)
        if str_val.endswith('%'):
            return float(str_val.strip('%')) / 100
        return float(str_val)
    except (ValueError, TypeError):
        return None

def parse_capbudg(file_path: str) -> Dict:
    """
    Parses data from an Excel file and extracts financial metrics into a structured dictionary.
    
    This function reads a specifically formatted Excel file containing capital budgeting information,
    extracts various financial data points from predefined cell locations, and organizes them into
    a nested dictionary structure for easy API consumption.

    Args:
        file_path (str): Absolute or relative path to the Excel file (.xls or .xlsx format)

    Returns:
        Dict: A nested dictionary containing all extracted financial data organized by category:
              - Initial investment details
              - Cash flow parameters
              - Discount rate calculations
              - Working capital information
              - Growth rate projections
              - Cash flow tables
              - Investment performance metrics
              - Asset depreciation schedules

    Raises:
        HTTPException: 404 if file not found, 500 for any other parsing error

    Notes:
        - The function expects the Excel file to follow a specific format with data in exact cell locations
        - Uses xlrd engine for .xls files and openpyxl for .xlsx files
        - Percentages are automatically converted to float representations (e.g., 10.00% → 0.1)
        - All monetary values are preserved as their raw numeric values
    """
    try:
        # Determine appropriate Excel engine based on file extension
        if file_path.endswith('.xls'):
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        # Load the Excel file
        xls = pd.ExcelFile(file_path, engine=engine)
        df = pd.read_excel(xls, sheet_name='CapBudgWS', header=None)

        data = {} # Master dictionary to hold all parsed data

        # --- Extract INITIAL INVESTMENT ---
        investment_data = {
            'Initial Investment': df.iloc[2, 2],
            'Opportunity Cost (if any)': df.iloc[3, 2],
            'Lifetime of the investment': df.iloc[4, 2],
            'Salvage Value at end of project': df.iloc[5, 2],
            'Deprec. method(1.St. line; 2.DDB)': df.iloc[6, 2],
            'Tax Credit (if any)': df.iloc[7, 2],
            'Other Invest.(non-depreciable)': df.iloc[8, 2],
        }
        data['INITIAL INVESTMENT'] = investment_data

        # --- Extract CASHFLOW DETAILS ---
        cashflow_details_data = {
            'Revenues in year 1': df.iloc[2, 6],
            'Var. Expenses as % of Rev': df.iloc[3, 6],
            'Fixed expenses in year 1': df.iloc[4, 6],
            'Tax rate on net income': df.iloc[5, 6],
        }
        data['CASHFLOW DETAILS'] = cashflow_details_data

        # --- Extract DISCOUNT RATE ---
        discount_rate_data = {
            'Approach': df.iloc[2, 10],
            'Discount rate': f"{df.iloc[3, 10] * 100:.2f}%" if isinstance(df.iloc[3, 10], (int, float)) else df.iloc[3, 10],
            'Beta': df.iloc[4, 10],
            'Riskless rate': f"{df.iloc[5, 10] * 100:.2f}%" if isinstance(df.iloc[5, 10], (int, float)) else df.iloc[5, 10],
            'Market risk premium': f"{df.iloc[6, 10] * 100:.2f}%" if isinstance(df.iloc[6, 10], (int, float)) else df.iloc[6, 10],
            'Debt Ratio': f"{df.iloc[7, 10] * 100:.2f}%" if isinstance(df.iloc[7, 10], (int, float)) else df.iloc[7, 10],
            'Cost of Borrowing': f"{df.iloc[8, 10] * 100:.2f}%" if isinstance(df.iloc[8, 10], (int, float)) else df.iloc[8, 10],
            'Discount rate used': f"{df.iloc[9, 10] * 100:.2f}%" if isinstance(df.iloc[9, 10], (int, float)) else df.iloc[9, 10],
        }
        data['DISCOUNT RATE'] = discount_rate_data

        # --- Extract WORKING CAPITAL ---
        working_capital_data = {
            'Initial WC': df.iloc[11, 2],
            'WC as % of Revenue': df.iloc[12, 2],
            'Salvageable Fraction': df.iloc[13, 2],
        }
        data['WORKING CAPITAL'] = working_capital_data

        # --- Extract GROWTH RATES ---
        growth_rates = {
            'Revenue Growth': df.iloc[17, 3:].to_list(),
            'Fixed Expense Growth': df.iloc[18, 3:].to_list()
        }
        data['GROWTH RATES'] = growth_rates

        # --- Extract INITIAL INVESTMENT (Lower Section) ---
        lower_investment_data = {
            'Investment': df.iloc[23, 1],
            'Tax Credit': df.iloc[24, 1],
            'Net Investment': df.iloc[25, 1],
            'Working Cap': df.iloc[26, 1],
            'Opp Cost': df.iloc[27, 1],
            'Other Invest': df.iloc[28, 1],
            'Initial Investment': df.iloc[29, 1],
        }
        data['INITIAL INVESTMENT (Lower)'] = lower_investment_data

        # --- Extract SALVAGE VALUE (Lower Section) ---
        salvage_value = {
            'Equipment': df.iloc[32, 3:].to_list(),
            'Working Capital': df.iloc[33, 3:].to_list(),
        }
        data['SALVAGE VALUE'] = salvage_value

        # --- Extract CASH FLOWS TABLE ---
        cashflows = df.iloc[36:50, 1:].transpose()
        cashflows.columns = [
            'Lifetime_Index', 'Revenues', 'Var_Expenses', 'Fixed_Expenses',
            'EBITDA', 'Depreciation', 'EBIT', 'Tax', 'EBIT_after_tax',
            'Add_Depreciation', 'Delta_Working_Capital', 'NATCF', 'Discount_Factor', 'Discounted_CF'
        ]
        data['OPERATING CASHFLOWS'] = cashflows.to_dict('records')

        # --- Extract INVESTMENT MEASURES ---
        investment_measures_data = {
            'NPV': df.iloc[52, 2],
            'IRR': df.iloc[53, 2],
            'ROC': df.iloc[54, 2],
        }
        data['INVESTMENT MEASURES'] = investment_measures_data

        # --- Extract BOOK VALUE AND DEPRECIATION ---
        bv_dep = df.iloc[58:61, 1:].transpose()
        bv_dep.columns = ['Book_Value_Beginning', 'Depreciation', 'Book_Value_Ending']
        data['BOOK VALUE AND DEPRECIATION'] = bv_dep.to_dict('index')

        return data
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"File not found at '{file_path}'")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred: {e}")

# Load the data when the application starts
file_path = os.path.join(os.getcwd(), "Data", "capbudg.xls")
try:
    parsed_data = parse_capbudg(file_path)
except Exception as e:
    # Log the error and exit
    print(f"Failed to load data: {e}")
    parsed_data = {}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=9090)