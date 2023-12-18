from bottle import Bottle, route, run, template, request, static_file, HTTPError, redirect
import subprocess
import sqlite3
import sys
import os
import numpy as np
import pandas as pd
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle, Border, Side, Color
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE
from openpyxl.utils import get_column_letter

app = Bottle()

# Set up SQLite database
db_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'users.db')
conn = sqlite3.connect(db_path)
cursor = conn.cursor()
cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL,
        password TEXT NOT NULL,
        email TEXT NOT NULL
    )
''')
conn.commit()
conn.close()

#Routes

@app.route('/')
def index():
    return template('index', message=None)

@app.route('/report')
def input():
    return template('report')

@app.route('/login', method='POST')
def login():
    username = request.forms.get('username')
    password = request.forms.get('password')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM users WHERE username=?', (username,))
    user = cursor.fetchone()
    conn.close()

    if user and user[2] == password:  # In a real application, use password hashing
        return template('welcome', message=f"Welcome, {username}!")
    else:
        return template('index', message="Invalid login credentials. Please try again.")

@app.route('/signup', method='POST')
def signup():
    username = request.forms.get('username')
    password = request.forms.get('password')
    email = request.forms.get('email')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('INSERT INTO users (username, password, email) VALUES (?, ?, ?)', (username, password, email))
    conn.commit()
    conn.close()

    return template('welcome', message=f"Account created for {username}!")

@app.route('/generate_file', method='POST')
def generate_file():
    # Get user input from the form
    ticker = request.forms.get('ticker')
    start_year = request.forms.get('start_year')
    end_year = request.forms.get('end_year')
    statement_types = request.forms.getall('statement_types')

   # Validate start year and end year as integers
    if not (start_year.isdigit() and end_year.isdigit()):
        raise HTTPError(400, "Start Year and End Year must be integers.")
    if not statement_types:
        raise HTTPError(400, "Select one to multiple statement types.")
    
    # Call the existing Python script with the user input
    script_path = 'GenerateReport.py'  # Update this with the correct path
    # Build the command to run the script
    command = [
        sys.executable,  # Use the Python executable from the current environment
        '-c', 'import sys, os; sys.path.append(os.path.dirname(sys.argv[1])); ' +
        f'from {os.path.basename(script_path).replace(".py", "")} import main; ' +
        'main()',  # Import and run the main function in the script
        script_path,  # Path to the script
        ticker,
        start_year,
        end_year
    ] + statement_types

    result = subprocess.run(command, capture_output=True, text=True)
    output = result.stdout.strip()
    # Return a link to the generated file for downloading
    return f'File generated. <a href="/download_file?filename={output}">Download File: {output}</a>'

@app.route('/download_file')
def download_file():
    # Specify the path to the generated file
    file_path = 'C://Development/PY/reports/'  # Update this with the correct path
    filename = request.query.get('filename')
    # Return the file for download
    try:
        return static_file(filename, root=file_path, download= filename)
    except Exception as e:
        print(f"Error: {e}")
    #return static_file('file.xlsx', root=file_path, download='file.xlsx')

if __name__ == '__main__':
    run(app, host='localhost', port=8080, debug=True)
