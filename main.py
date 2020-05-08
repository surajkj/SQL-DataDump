import datetime
from sqlalchemy import create_engine
from sqlalchemy import inspect
from xlwt import Workbook
from dotenv import load_dotenv
import os

# Load environments variables
load_dotenv()

# Create SQL Connection
engine = create_engine(os.getenv("DB_CONNECTION"))

# Load sql inspector for reading tables & columns
inspector = inspect(engine)

# Create data directory. This is where all excel files will be written
if not os.path.exists('data'):
    os.makedirs('data')

# Read table names
for table_name in inspector.get_table_names():
    # Initialize new workbook for every table
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    # Connect to database
    with engine.connect() as con:
        # Get data from table
        rs = con.execute('SELECT * FROM '+table_name)

        # First row should be coulmn name
        header_column_count = 0
        for column in inspector.get_columns(table_name):
            sheet1.write(0, header_column_count, column['name'])
            header_column_count += 1
        # Rows parsing & writing to workbook
        row_count = 1
        for row in rs:
            column_count = 0
            for data in row:
                if type(data) is datetime.datetime:
                    data = str(data)
                sheet1.write(row_count, column_count, data)
                column_count += 1
            row_count += 1
        # save data to workbook
        wb.save('data/'+table_name+'.xls')
