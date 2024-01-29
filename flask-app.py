from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory
import subprocess
import sqlite3
import sys
import os
from dotenv import load_dotenv

# Generate a secure random key with 32 bytes
#secret_key = secrets.token_hex(32)

app = Flask(__name__)
load_dotenv(dotenv_path='config.env')
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'default_secret_key')
  # Replace with a secret key for session security

# Set up SQLite database
db_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'reports.db')
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
cursor.execute('''
    CREATE TABLE IF NOT EXISTS Files (
        fileId INTEGER PRIMARY KEY AUTOINCREMENT,
        ticker TEXT,
        filename TEXT
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS DownloadHistory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
        userid INTEGER,
        fileId INTEGER,
        FOREIGN KEY (fileId) REFERENCES Files (fileId)
    )
''')

conn.commit()
conn.close()

# Routes
@app.route('/')
def index():
    return render_template('header.html', title='Homepage') + render_template('index.html', message='')

@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username')
    password = request.form.get('password')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM users WHERE username=?', (username,))
    user = cursor.fetchone()
    conn.close()

    if user and user[2] == password:  # In a real application, use password hashing
        # Store userid in session
        session['userid'] = username
        return render_template('header.html', title='Welcome') + render_template('welcome.html', username=username, message=f"Welcome, {username}!")
    else:
        return render_template('header.html', title='Homepage') + render_template('index.html', message="Invalid login credentials. Please try again.")

@app.route('/logout')
def logout():
    # Expire the session to logout
    session.pop('userid', None)
    return redirect(url_for('login'))

@app.route('/signup', methods=['POST'])
def signup():
    username = request.form.get('username')
    password = request.form.get('password')
    email = request.form.get('email')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('INSERT INTO users (username, password, email) VALUES (?, ?, ?)', (username, password, email))
    conn.commit()
    conn.close()

    return render_template('header.html', title='Welcome') + render_template('welcome.html', username=username, message=f"Account created for {username}!")

@app.route('/welcome')
def welcome():
    userid = session.get('userid', None)

    if userid:
        return render_template('header.html', title='Welcome') + render_template('welcome.html', username=userid)
    else:
        # Redirect to login if userid is not in session
        return redirect(url_for('login'))

@app.route('/report/<userid>')
def report(userid):
    return render_template('header.html', title='Report') + render_template('report.html', userid=userid)

@app.route('/generate_file/<userid>', methods=['POST'])
def generate_file(userid):
    # Get user input from the form
    ticker = request.form.get('ticker')
    start_year = request.form.get('start_year')
    end_year = request.form.get('end_year')
    statement_types = request.form.getlist('statement_types')

    # Validate start year and end year as integers
    if not (start_year.isdigit() and end_year.isdigit()):
        return "Start Year and End Year must be integers.", 400

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
        end_year,
        userid
    ]

    result = subprocess.run(command, capture_output=True, text=True)
    output = result.stdout.strip()
    # Return a link to the generated file for downloading
    return f'File generated. <a href="/download_file?filename={output}">Download File: {output}</a>'

@app.route('/download_file')
def download_file():
    # Specify the path to the generated file
    file_path = 'C://Development/PY/reports/'  # Update this with the correct path
    filename = request.args.get('filename')
    # Return the file for download
    try:
        return send_from_directory("reports", filename)
        #(file_path, filename=filename, as_attachment=True)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == '__main__':
    app.run(debug=True)
