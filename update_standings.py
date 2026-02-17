import pandas as pd
from datetime import datetime
import os

# --- CONFIGURATION ---
excel_file = 'index.xlsm' # <--- MAKE SURE THIS MATCHES YOUR FILENAME
sheet_name = 'index'
# Row 20 in Excel is index 19 in Python
start_row_index = 19 
total_rows = 181
# Column K in Excel is index 10 in Python
start_col_index = 10 

def generate_standings():
    if not os.path.exists(excel_file):
        print("Excel file not found!")
        return

    # Load the sheet with no header logic to maintain coordinate control
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    rows_html = ""

    # Loop from Row 20 to Row 200 (181 rows total)
    for i in range(start_row_index, start_row_index + total_rows):
        # Extract data based on relative column positions from K (index 10)
        rank      = df.iloc[i, start_col_index]     # K (10)
        player    = df.iloc[i, start_col_index + 1] # L (11)
        blank     = ""                              # M (12) - Keeping it blank as requested
        total     = df.iloc[i, start_col_index + 3] # N (13)
        group     = df.iloc[i, start_col_index + 4] # O (14)
        bracket   = df.iloc[i, start_col_index + 5] # P (15)
        possible  = df.iloc[i, start_col_index + 6] # Q (16)
        bonus     = df.iloc[i, start_col_index + 7] # R (17)

        # Build the HTML row
        rows_html += f"""
        <tr>
            <td>{rank}</td>
            <td style="text-align: left; font-weight: bold;">{player}</td>
            <td>{blank}</td>
            <td style="font-weight: bold; color: #1b4332;">{total}</td>
            <td>{group}</td>
            <td>{bracket}</td>
            <td>{possible}</td>
            <td>{bonus}</td>
        </tr>"""

    # --- HTML TEMPLATE ---
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; padding: 20px; }}
            .container {{ max-width: 1100px; margin: auto; background: white; padding: 20px; border-radius: 12px; shadow: 0 4px 15px rgba(0,0,0,0.1); }}
            h1 {{ text-align: center; color: #1b4332; text-transform: uppercase; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background: #1b4332; color: white; padding: 12px; font-size: 0.8em; text-transform: uppercase; }}
            td {{ padding: 10px; border-bottom: 1px solid #eee; text-align: center; font-size: 0.95em; }}
            tr:nth-child(even) {{ background: #fafafa; }}
            tr:hover {{ background: #f1f8f5; }}
            .update-tag {{ text-align: center; font-size: 0.8em; color: #888; margin-top: 20px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🏆 World Cup 2026 Standings</h1>
            <table>
                <thead>
                    <tr>
                        <th>Rank</th>
                        <th>Participant</th>
                        <th></th>
                        <th>Total Points</th>
                        <th>Group Stage</th>
                        <th>Bracket Stage</th>
                        <th>Possible Left</th>
                        <th>Bonus Game</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
            <div class="update-tag">Last Sync: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
        </div>
    </body>
    </html>
    """

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(full_html)
    print("Webpage successfully mapped to K20:R200!")

if __name__ == "__main__":
    generate_standings()