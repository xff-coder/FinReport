from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, Response
import subprocess
import sqlite3
import sys
import os
from datetime import datetime
import requests
import json
import pandas as pd
import numpy as np
import statementFunct
import styleModule
from main_excel import StartReport

#from dotenv import load_dotenv

# Generate a secure random key with 32 bytes
#secret_key = secrets.token_hex(32)

app = Flask(__name__)
#load_dotenv(dotenv_path='config.env')
#app.secret_key = os.getenv('FLASK_SECRET_KEY', 'default_secret_key')
app.secret_key = '13245634287' 
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

@app.route('/about')
def about():
    return render_template('header.html', title='Homepage') + render_template('homepage.html')

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
    return redirect(url_for('index'))

@app.route('/signup', methods=['POST'])
def signup():
    username = request.form.get('username')
    password = request.form.get('password')
    email = request.form.get('email')
    invitationCode = request.form.get('invitationcode')
    if invitationCode != '0209dr@gon2024' :
        return render_template('header.html', title='Homepage') + render_template('index.html', message="Invalid invitation code. Access denied.")

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

    app.logger.info(f"User {userid} generated a file for ticker {ticker} from {start_year} to {end_year}")

    # Validate start year and end year as integers
    if not (start_year.isdigit() and end_year.isdigit()):
        return "Start Year and End Year must be integers.", 400

    # Call the existing Python script with the user input   
    output = StartReport(ticker, int(start_year), int(end_year), userid)
    filename = output.split('/')[-1]
    # Return a link to the generated file for downloading
    
    return render_template('header.html', title='Report') + render_template('download.html', userid=userid, output=output, filename=filename)

@app.route('/download/<path:filename>')
def download_file(filename):
    # Specify the path to the generated file
    directory = 'reports'  # Update this with the correct path
    #filename = request.args.get('filename')
    # Return the file for download
    try:
        return send_from_directory(directory, filename, as_attachment=True)
        #(file_path, filename=filename, )
    except Exception as e:
        print(f"Error: {e}")

if __name__ == '__main__':
    app.run(host='localhost', port=5000, debug=True)
