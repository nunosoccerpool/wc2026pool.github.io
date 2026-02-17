import pandas as pd
from datetime import datetime

# 1. Load the data
# Make sure this filename matches yours exactly!
file_path = 'YourActualFileName.xlsm' 
df = pd.read_excel(file_path, engine='openpyxl')

# 2. Generate the HTML Table
# 'table-style' is a custom class we'll define below
table_html = df.to_html(index=False, classes='standings-table')

# 3. Create the Full Webpage Template
html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WC 2026 Pool Standings</title>
    <style>
        :root {{
            --pitch-green: #1b4332;
            --grass-light: #2d6a4f;
            --trophy-gold: #ffca3a;
            --white: #ffffff;
        }}

        body {{
            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background-color: #e9ecef;
            margin: 0;
            padding: 20px;
            color: #333;
        }}

        .container {{
            max-width: 900px;
            margin: auto;
            background: var(--white);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }}

        header {{
            text-align: center;
            border-bottom: 3px solid var(--pitch-green);
            margin-bottom: 20px;
            padding-bottom: 10px;
        }}

        h1 {{
            color: var(--pitch-green);
            margin: 0;
            text-transform: uppercase;
            letter-spacing: 2px;
        }}

        .update-time {{
            font-size: 0.9em;
            color: #666;
            font-style: italic;
        }}

        /* Table Styling */
        .standings-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 16px;
        }}

        .standings-table th {{
            background-color: var(--pitch-green);
            color: var(--white);
            padding: 15px;
            text-align: center;
            text-transform: uppercase;
        }}

        .standings-table td {{
            padding: 12px;
            text-align: center;
            border-bottom: 1px solid #dee2e6;
        }}

        /* Highlight the Leader */
        .standings-table tr:first-child td {{
            background-color: rgba(255, 202, 58, 0.2);
            font-weight: bold;
        }}

        /* Zebra Striping */
        .standings-table tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}

        .standings-table tr:hover {{
            background-color: #f1f1f1;
        }}

        /* Mobile Friendly */
        @media (max-width: 600px) {{
            .container {{ padding: 10px; }}
            .standings-table {{ font-size: 14px; }}
            .standings-table th, .standings-table td {{ padding: 8px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>🏆 WC 2026 Pool Standings</h1>
            <p class="update-time">Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </header>
        
        {table_html}

        <footer style="margin-top: 30px; text-align: center; font-size: 0.8em; color: #888;">
            <p>Automated by Python & GitHub Actions</p>
        </footer>
    </div>
</body>
</html>
"""

# 4. Save the file
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("index.html created successfully with pro styling!")