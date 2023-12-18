<!DOCTYPE html>
<html>
<head>
    <title>Financial Data Generator</title>
    <style>
        form {
            max-width: 300px; /* Adjust the maximum width as needed */
            margin: auto; /* Center the form on the page */
        }

        label {
            display: block; /* Make labels block elements for vertical alignment */
            margin-bottom: 5px; /* Add some space between labels */
        }

        input[type="text"],
        select {
            width: 100%; /* Make input fields and select full width */
            box-sizing: border-box; /* Include padding and border in the width */
            margin-bottom: 10px; /* Add some space between input fields */
        }

        input[type="submit"] {
            background-color: #4CAF50; /* Green submit button color */
            color: white; /* Text color */
            padding: 10px 15px; /* Padding for a better appearance */
            border: none; /* No border */
            border-radius: 4px; /* Rounded corners */
            cursor: pointer; /* Pointer cursor on hover */
        }

        input[type="submit"]:hover {
            background-color: #45a049; /* Darker green color on hover */
        }
    </style>
</head>
<body>
    <h2 style="text-align: center;">Generate Financial Data</h2>
    <form action="/generate_file" method="post">
        <label>Ticker: <input type="text" name="ticker" required></label>
        <label>Start Year: <input type="text" name="start_year" required></label>
        <label>End Year: <input type="text" name="end_year" required></label>
        <br>
        <label>Statement Types:</label>
        <select name="statement_types" multiple>
            <option value="BS">Balance Sheet</option>
            <option value="CF">Cash Flow</option>
            <option value="PL">Financial Statement</option>
        </select>
        <br><br>
        <input type="submit" value="Generate File">
    </form>
</body>
</html>
