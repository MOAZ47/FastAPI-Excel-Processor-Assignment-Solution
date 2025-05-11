# FastAPI-Excel-Processor-Assignment-Solution

## Insights
### Potential Improvements
<ul>
  <li><b>Caching</b></li>
  <p> The Parsing mechanism in the code is computationally intensive. Caching would improve performance.</p>
  <li><b>Frontend Integration</b></li>
  <p>Adding a simple Streamlit or Dash interface would allow non-technical users to upload Excel files and view results interactively</p>
  <li><b>ETL</b></li>
  <p>A pure Pythonic ETL can be built around this, where the data is uploaded on database after extraction and various transformations.</p>
</ul>

### Missed Edge Cases
<ul>
  <li><b>Mixed Data Types in Columns</b></li>
  <p>Columns with mixed data types (e.g., strings and numbers) could lead to unexpected results during type-based analysis or aggregation.</p>
  <li><b>Data Parsing</b></li>
  <p>Current Implementation expects a fixed format, and will fail if the format of data file changes</p>
</ul>

## Testing
Import the Postman collection from:
`ExcelProcessor_API.postman_collection.json`

Or run these curl commands:
```bash
# List tables
curl http://localhost:9090/list_tables

# Get table details
curl "http://localhost:9090/get_table_details?table_name=INITIAL%20INVESTMENT"

# Calculate sum
curl "http://localhost:9090/row_sum?table_name=INITIAL%20INVESTMENT&row_name=Tax%20Credit%20(if%20any)"
