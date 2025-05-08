# Excel Visualization MCP for Claude Desktop 📊

## 1. Introduction 📜

This document presents an Excel visualization extension developed for Claude Desktop, built using Anthropic's Model Context Protocol (MCP). This extension enables seamless integration between Claude's conversational interface and Excel data visualization capabilities.

The Excel MCP delivers:

* Conversational interface for Excel data analysis via Claude Desktop 🗣️
* Text-based reports and statistical summaries ✍️
* Dynamic data visualizations (bar charts, pie charts, line graphs, scatter plots) 📈
* Performance optimizations for handling files of various sizes, up to 100,000 rows 🚀
* Export capabilities to PowerPoint and Word via VBA integration 📤

**Important Note**: This MCP works *exclusively* with Claude Desktop. The Claude API is not yet supported. 🚫

## 2. Core Capabilities ✨

### Data Analysis
* Excel file reading and sheet listing
* Statistical summaries and data exploration
* Customizable data queries with filtering capabilities
* Support for files up to 100,000 rows

### Visualization
* Dynamic chart generation directly in the conversation
* Multiple chart types: bar, line, pie, scatter
* Customizable titles, labels, and aggregation methods
* Performance optimizations for large datasets

### Data Management
* Append data to existing Excel files
* Update individual cells (values and formulas)
* Insert formatted data with table styling options
* Test file modification capabilities

### Export Integration
* VBA code generation for PowerPoint/Word exports
* Dynamic chart creation with export buttons
* Chart refreshing capabilities

## 3. Technical Architecture 🏗️

The MCP consists of:

1. **Python Server**: Handles Excel processing using:
   - Pandas (for standard files)
   - Polars (automatically used for files over 50MB)
   - Matplotlib (for chart generation)
   - FastMCP (Anthropic's server framework for MCPs)

2. **Claude Desktop Integration**: 
   - Configured via the `claude_desktop_config.json` file
   - Communication through Anthropic's Model Context Protocol

3. **Performance Optimizations**:
   - Automatic library switching based on file size
   - Sampling for large datasets
   - Custom aggregation methods
   - Chunked processing for memory efficiency

## 4. Installation and Configuration 🛠️

1. Clone the repository to your preferred directory.
2. Ensure you have Python 3.8+ and the following packages installed:
   ```bash
   pip install pandas polars matplotlib openpyxl fastmcp
   ```
3. In the `%APPDATA%\Claude\` folder (Windows) or equivalent on macOS, create a file named `claude_desktop_config.json` with the following content:

   ```json
   {
     "mcpServers": {
       "excel-viz": {
         "command": "python",
         "args": ["PATH\\TO\\YOUR\\excel-mcp\\py\\excel_viz_server.py"]
       }
     }
   }
   ```
   
   Replace `PATH\\TO\\YOUR` with the actual path to your installation directory.

4. Restart Claude Desktop for the configuration to take effect.

## 5. Usage Guide ▶️

The MCP is automatically loaded when you start Claude Desktop. To use it:

1. Open Claude Desktop (Windows or macOS).
2. In the sidebar or plugins menu, select `excel-viz`.
3. Send your requests directly using natural language.

### Example Prompts

#### Basic Data Exploration
```
Can you list all sheets in the Excel file "budget.xlsx" in my Documents folder?
```

```
Show me the content of the "Revenue" sheet in "finances2024.xlsx"
```

```
Generate a statistical summary of the data in the "Sales" sheet of "quarterly_report.xlsx"
```

#### Creating Visualizations
```
Create a bar chart from "sales_data.xlsx" with "Month" on the x-axis and "Revenue" on the y-axis. Use the title "Monthly Revenue 2024".
```

```
Make a line chart from "performance.xlsx" showing "Metric1,Metric2,Metric3" over time with "Date" as the x-axis. Please aggregate by mean and limit to 300 points.
```

```
Create a scatter plot from "correlation_study.xlsx" with "Age" on the x-axis, "Income" on the y-axis, and use "Region" for color coding.
```

#### Advanced Queries and Aggregations
```
In the file "employees.xlsx", find all rows where the "Salary" column is greater than 50000 and "Department" is "Marketing"
```

```
Create an aggregated chart from "transactions.xlsx" that shows the sum of "Amount" by "Category", displayed as a pie chart with the top 8 categories.
```

#### Excel Modifications
```
Add the following data to "inventory.xlsx" in the "Products" sheet: 
Product001,Laptop,1299.99,15
Product002,Monitor,349.99,32
Product003,Keyboard,89.99,47
```

```
Update cell C10 in the "Budget" sheet of "financial_plan.xlsx" to contain the formula "=SUM(C2:C9)"
```

#### Chart Export
```
Generate VBA code to export charts from "quarterly_report.xlsx" to PowerPoint
```

```
Create dynamic chart VBA for "sales_data.xlsx" with source sheet "Data", chart sheet "Chart", and data range "A1:D25"
```

## 6. Important File Handling Notes ⚠️

### File Paths
You can use relative paths (starting from your Documents folder) or absolute paths:

```
// Relative path (from Documents folder)
budget.xlsx

// Subfolder in Documents
Finance/budget.xlsx

// Absolute path
D:/MyData/Excel/budget.xlsx
```

### Query Syntax
For the `excel_query` function, use syntax similar to pandas:

```
// Valid query examples
"Age > 30"
"Department == 'Marketing'"
"Sales > 1000 and Region == 'North'"
```

### Sheet Handling
If you don't specify a sheet name, the server will use the first sheet by default.

### Performance Considerations
- Files larger than 50MB will automatically use Polars instead of Pandas for better performance
- For large files, specify the sheet and use targeted queries to avoid displaying huge tables
- Consider using aggregation options (`sum`, `mean`, `count`, `min`, `max`) when creating charts for large datasets
- The current implementation handles files up to approximately 100,000 rows efficiently

## 7. Troubleshooting 🔧

### Access Issues
If Claude indicates difficulty accessing a file:
- Verify the file path is correct
- Ensure the file isn't open in Excel (which can block access)
- Check that the sheet name is spelled correctly and exists in the file
- For files with spaces in the name, use quotes around the path

### Performance Issues
- For slow performance with large files, try:
  - Using `read_excel_optimized` instead of `read_excel`
  - Specifying only the columns you need
  - Increasing the chunk size for larger files
  - Using `create_aggregated_chart` instead of standard chart functions

### VBA Export Problems
- Make sure to save the file as .xlsm format (Excel with macros)
- Enable macros in Excel and authorize content
- If buttons don't appear, run the `AddExportButtons` macro manually

## 8. Development and Extension 👨‍💻

This MCP can be extended with new functionality:

1. Add new tool functions to `excel_viz_server.py` using the `@mcp.tool()` decorator
2. Implement additional chart types or data processing capabilities
3. Create new export options or integrations with other Office applications

To run the server in development mode:
```bash
python py/run.py
```

## 9. Future Enhancements 🔮

Planned improvements include:
- Direct image generation and manipulation for chart editing
- Support for more complex query operations
- Interactive dashboard creation
- Multi-file comparison capabilities
- API support for integration with the Claude API
- Improved error handling and diagnostic tools
- Support for additional file formats (CSV, Google Sheets)

## 10. License and Attribution 📝

This project is provided as-is without warranty. When using and modifying this code, please maintain attribution to the original authors.

---

For questions, bug reports, or feature requests, please contact the development team or submit an issue on the project repository.
