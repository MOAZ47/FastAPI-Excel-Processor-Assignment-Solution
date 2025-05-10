# FastAPI-Excel-Processor-Assignment-Solution

## Insights
### Potential Improvements
<ul>
  <li><b>UI Integration</b></li>
  <p>Adding a simple Streamlit or Dash interface would allow non-technical users to upload Excel files and view results interactively</p>
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
