# 📊 Excel-Processor-API

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![FastAPI](https://img.shields.io/badge/FastAPI-0.68+-green.svg)
![Pandas](https://img.shields.io/badge/Pandas-Latest-orange.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

*A powerful FastAPI-based service for processing and analyzing Excel files with capital budgeting data*

[Features](#-features) • [Quick Start](#-quick-start) • [API Documentation](#-api-documentation) • [Examples](#-examples) • [Contributing](#-contributing)

</div>

---

## 🌟 Features

- 🚀 **Fast API Processing** - Built with FastAPI for high-performance Excel data processing
- 📈 **Capital Budgeting Analysis** - Specialized for financial worksheets and investment analysis
- 🔍 **Smart Table Detection** - Automatically identifies and parses structured tables in Excel sheets
- 📊 **Row-wise Calculations** - Perform sum calculations on specific rows across multiple columns
- 🛡️ **Robust Error Handling** - Comprehensive validation and error reporting
- 📝 **Interactive Documentation** - Auto-generated Swagger/OpenAPI documentation
- 🧪 **Postman Ready** - Includes collection for easy API testing

## 🏗️ Architecture

```mermaid
graph TD
    A[Excel File] --> B[ExcelProcessor Class]
    B --> C[Table Parser]
    C --> D[Data Extraction]
    D --> E[FastAPI Endpoints]
    E --> F[JSON Response]
    
    G[Client Request] --> E
    E --> H[/list_tables]
    E --> I[/get_table_details]
    E --> J[/row_sum]
```

## 🚀 Quick Start

### Prerequisites

- Python 3.8+
- pip package manager

### Installation

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd excel-processor-api
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Prepare your Excel file**
   - Place your Excel file at `Data/capbudg.xlsx`
   - Ensure it has a sheet named "CapBudgWS"

4. **Run the application**
   ```bash
   uvicorn main:app --host 0.0.0.0 --port 9090 --reload
   ```

5. **Access the API**
   - API: `http://localhost:9090`
   - Interactive Docs: `http://localhost:9090/docs`
   - Redoc: `http://localhost:9090/redoc`

## 📋 Required Dependencies

Add these to your `requirements.txt`:

```txt
fastapi>=0.68.0
uvicorn[standard]>=0.15.0
pandas>=1.3.0
openpyxl>=3.0.0
python-multipart>=0.0.5
```

## 🔧 API Documentation

### Endpoints Overview

| Endpoint | Method | Description | Parameters |
|----------|--------|-------------|------------|
| `/list_tables` | GET | Get all available tables | None |
| `/get_table_details` | GET | Get row names for a table | `table_name` |
| `/row_sum` | GET | Calculate sum of a specific row | `table_name`, `row_name` |

### 📊 Supported Table Types

The API automatically detects these table types in your Excel sheet:

- **INITIAL INVESTMENT** - Capital investment data
- **SALVAGE VALUE** - Asset salvage values
- **OPERATING CASHFLOWS** - Operational cash flow projections
- **BOOK VALUE & DEPRECIATION** - Depreciation schedules

## 💡 Examples

### List All Tables

```bash
curl -X GET "http://localhost:9090/list_tables"
```

**Response:**
```json
{
  "tables": [
    "INITIAL INVESTMENT",
    "SALVAGE VALUE", 
    "OPERATING CASHFLOWS",
    "BOOK VALUE & DEPRECIATION"
  ]
}
```

### Get Table Details

```bash
curl -X GET "http://localhost:9090/get_table_details?table_name=INITIAL%20INVESTMENT"
```

**Response:**
```json
{
  "table_name": "INITIAL INVESTMENT",
  "row_names": ["Investment", "Setup Costs", "Working Capital"]
}
```

### Calculate Row Sum

```bash
curl -X GET "http://localhost:9090/row_sum?table_name=INITIAL%20INVESTMENT&row_name=Investment"
```

**Response:**
```json
{
  "table_name": "INITIAL INVESTMENT",
  "row_name": "Investment", 
  "sum": 250000.0
}
```

## 🧪 Testing with Postman

1. Import the provided `postman_collection.json` file
2. The collection includes pre-configured requests for all endpoints
3. Update the base URL if running on a different host/port

## 📁 Project Structure

```
excel-processor-api/
├── 📄 main.py                 # FastAPI application
├── 📄 excel_processor.py      # Core Excel processing logic
├── 📄 requirements.txt        # Python dependencies
├── 📄 postman_collection.json # Postman test collection
├── 📄 README.md              # This file
└── 📁 Data/
    └── 📄 capbudg.xlsx       # Excel data file
```

## 🔧 Configuration

### Excel File Requirements

Your Excel file must meet these criteria:

- **Format**: .xlsx (Excel 2007+)
- **Sheet Name**: "CapBudgWS"
- **Structure**: Tables should have clear headers matching supported types
- **Layout**: First column contains row names, subsequent columns contain data

### Environment Variables

You can customize the application using these environment variables:

```bash
export EXCEL_FILE_PATH="Data/capbudg.xlsx"
export API_HOST="0.0.0.0"
export API_PORT="9090"
```

## 🚨 Error Handling

The API provides comprehensive error handling:

- **400 Bad Request** - Invalid table/row names
- **404 Not Found** - Table not found or empty
- **500 Internal Server Error** - File processing errors

## 🤝 Contributing

We welcome contributions! Here's how you can help:

1. 🍴 Fork the repository
2. 🌿 Create a feature branch (`git checkout -b feature/amazing-feature`)
3. 💾 Commit your changes (`git commit -m 'Add amazing feature'`)
4. 📤 Push to the branch (`git push origin feature/amazing-feature`)
5. 🔄 Open a Pull Request

### Development Setup

1. Clone and install in development mode:
   ```bash
   pip install -e .
   ```

2. Run tests:
   ```bash
   pytest tests/
   ```

3. Format code:
   ```bash
   black . && isort .
   ```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- 📧 **Email**: [ujwalakusma26@gmail.com]
- 🐛 **Issues**: [GitHub Issues](https://github.com/U0426jwala/Excel-Processor-API.git)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/U0426jwala)

## 🙏 Acknowledgments

- Built with [FastAPI](https://fastapi.tiangolo.com/)
- Data processing powered by [Pandas](https://pandas.pydata.org/)
- Excel integration via [OpenPyXL](https://openpyxl.readthedocs.io/)

---

<div align="center">

**Made with ❤️ for financial data analysis**

⭐ Star this repo if you find it helpful!

</div>