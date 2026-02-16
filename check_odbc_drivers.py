import pyodbc

print("Available ODBC Drivers:")
print("=" * 50)
drivers = pyodbc.drivers()
for driver in drivers:
    print(f"  - {driver}")
print("=" * 50)
print(f"\nTotal drivers found: {len(drivers)}")
