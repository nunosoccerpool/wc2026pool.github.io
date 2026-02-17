import pandas as pd
from datetime import datetime
import os

# 1. Configuration - Change 'YourActualFileName.xlsm' to your file's name
excel_file = 'index.xlsm' 
sheet_to_read = 'index'

def generate_standings():
    # Check if file exists to avoid crashes
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found.")
        return

    # 2. Read the specific "index" sheet
    # We use engine='openpyxl' for .xlsm files
    df = pd.read_excel(excel_file, sheet_name=sheet_to_read, engine='openpyxl')

    # 3. Data Cleaning
    # Removes rows and columns that are completely empty
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    # Fills any remaining empty cells with a blank space so the table stays aligned
    df = df.fillna('')

    # 4. Convert Dataframe to HTML
    table_html = df.to_html(index=False, classes='standings-table')

    # 5. The Design (CSS) and Layout
    html_template = f"""
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
                --gold: #ffca3a;
                --white: #ffffff;
                --dark-text: #212529;
            }}

            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f0f2f5;
                margin: 0;
                padding: 20px;
                display: flex;
                justify-content: center;
            }}

            .dashboard {{
                width: 100%;
                max-width: 1000px;
                background: var(--white);
                padding: 30px;
                border-radius: 20px;
                box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            }}

            header {{
                text-align: center;
                margin-bottom: 30px;
                border-bottom: 4px solid var(--pitch-green);
                padding-bottom: 15px;
            }}

            h1 {{
                color: var(--pitch-green);
                font-size: 2.5em;
                margin: 0;
                letter-spacing: -1px;
            }}

            .timestamp {{
                color: #6c757d;
                font-size: 0.9em;
                margin-top: 5px;
            }}

            /* Soccer Table Styling */
            .standings-table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 10px;
            }}

            .standings-table th {{
                background-color: var(--pitch-green);
                color: var(--white);
                padding: 18px;
                text-align: center;
                text-transform: uppercase;
                font-size: 0.85em;
                letter-spacing: 1px;
            }}

            .standings-table td {{
                padding: 15px;
                text-align: center;
                border-bottom: 1px solid #eee;
                color: var(--dark-text);
                font-size: 1.05em;
            }}

            /* Highlight the Leader */
            .standings-table tr:first-child td {{
                background-color: rgba(255, 202, 58, 0.15);
                font-weight: bold;
                border-left: 5px solid var(--gold);
            }}

            /* Zebra Stripes */
            .standings-table tr:nth-child(even) {{
                background-color: #fcfcfc;
            }}

            .standings-table tr:hover {{
                background-color: #f1f8f5;
            }}

            /* Mobile Responsiveness */
            @media (max-width: 768px) {{
                body {{ padding: 10px; }}
                .dashboard {{ padding: 15px; }}
                h1 {{ font-size: 1.8em; }}
                .standings-table td, .standings-table th {{
                    padding: 10px 5px;
                    font-size: 0.9em;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="dashboard">
            <header>
                <h1>🏆 World Cup 2026 Pool</h1>
                <p class="timestamp">Live Leaderboard • Last Updated: {datetime.now().strftime('%b %d, %Y | %H:%M:%S')}</p>
            </header>
            
            <div style="overflow-x: auto;">
                {table_html}
            </div>

            <footer style="text-align: center; margin-top: 40px; color: #adb5bd; font-size: 0.8em;">
                Proudly powered by Python & GitHub Pages
            </footer>
        </div>
    </body>
    </html>
    """

    # 6. Save the final HTML
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_template)
    
    print("Success: index.html has been updated from the 'index' sheet!")

if __name__ == "__main__":
    generate_standings()