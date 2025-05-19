# Automation Tool

A Python automation tool for connecting to Supabase PostgreSQL databases and fetching table data.

## Features

- Connect to Supabase PostgreSQL databases
- List all tables in a database
- Fetch data from specific tables
- Execute custom SQL queries
- Save query results to CSV files

## Installation

1. Clone the repository:
```
git clone https://github.com/yourusername/supabase-automation.git
cd supabase-automation
```

2. Create a virtual environment and install dependencies:
```
python -m venv venv
venv\Scripts\activate  # On Windows
source venv/bin/activate  # On Linux/Mac
pip install -r requirements.txt
```

3. Create a `.env` file with your Supabase credentials:
```
# Supabase credentials
SUPABASE_URL=https://your-project-url.supabase.co
SUPABASE_KEY=your-supabase-anon-key
SUPABASE_SERVICE_KEY=your-supabase-service-key

# Database connection
DB_HOST=db.your-project-url.supabase.co
DB_PORT=5432
DB_NAME=postgres
DB_USER=postgres
DB_PASSWORD=your-database-password
```

## Usage

### List all tables

```
python -m src.main --list-tables
```

### Fetch data from a specific table

```
python -m src.main --table your_table_name
```

With row limit:
```
python -m src.main --table your_table_name --limit 500
```

### Execute a custom SQL query

```
python -m src.main --query "SELECT * FROM your_table WHERE column = 'value' LIMIT 10"
```

### Save results to CSV

```
python -m src.main --table your_table_name --save-csv output.csv
```

## Using as a Python Library

You can also use this tool as a Python library in your own scripts:

```python
from src.database import SupabaseClient

# Initialize client
client = SupabaseClient()

# List tables
tables = client.list_tables()
print(tables)

# Fetch data from a table
df = client.get_table_data("your_table_name", limit=100)
print(df.head())

# Execute a custom query
df = client.execute_query("SELECT * FROM your_table WHERE column = 'value' LIMIT 10")
print(df.head())
```

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 
