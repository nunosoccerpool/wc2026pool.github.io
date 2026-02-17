import pandas as pd
from datetime import datetime

# 1. Load the Excel file
# Use 'openpyxl' for .xlsm files
file_path = 'standings.xlsm'
df = pd.read_excel(file_path, engine='openpyxl')

# 2. Get the current timestamp
# Format: Month Day, Year at Hour:Minute
last_updated = datetime.now().strftime("%B %d, %Y at %I:%M %p")

# 3. Define the HTML Template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>2026 World Cup Pool</title>
    <style>
        body {{ font-family: 'Segoe UI', sans-serif; background-color: #f0f2f0; margin: 0; }}
        .header {{ background: linear-gradient(rgba(0,0,0,0.6), rgba(0,0,0,0.6)), url('https://images.unsplash.com/photo-1551958219-acbc608c6377?auto=format&fit=crop&q=80&w=1000'); background-size: cover; color: white; text-align: center; padding: 40px 20px; }}
        
        .soccer-nav {{ background-color: #2e7d32; display: flex; justify-content: center; position: sticky; top: 0; z-index: 1000; }}
        .soccer-nav a {{ color: white; padding: 16px 25px; text-decoration: none; font-weight: bold; text-transform: uppercase; }}
        .soccer-nav a:hover {{ background-color: #1b5e20; color: #ffff00; }}

        .container {{ max-width: 1000px; margin: 30px auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }}
        
        /* Timestamp Styling */
        .timestamp {{ background-color: #fff9c4; border-left: 5px solid #fbc02d; padding: 10px; margin-bottom: 20px; font-style: italic; color: #555; }}

        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th {{ background-color: #011f4b; color: white; padding: 12px; }}
        td {{ border: 1px solid #eee; padding: 10px; text-align: center; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        
        a {{ color: #d32f2f; text-decoration: none; font-weight: bold; }}
    </style>
</head>
<body>

<div class="header">
    <h1>2026 World Cup Pool</h1>
</div>

<div class="soccer-nav">
    <a href="#standings">Standings</a>
    <a href="#stats">Pool Statistics</a>
    <a href="#bonus">Bonus</a>
    <a href="#points">Points</a>
</div>

<div class="container" id="standings">
    <div class="timestamp">
        <strong>Last Updated:</strong> {update_time}
    </div>
    <h2>Current Standings</h2>
    {table_content}
</div>

</body>
</html>
"""

# 4. Convert to HTML
# render_links=True attempts to keep clickable URLs
table_html = df.to_html(index=False, render_links=True)

# 5. Write the file
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_template.format(update_time=last_updated, table_content=table_html))

print(f"Successfully updated index.html at {last_updated}")