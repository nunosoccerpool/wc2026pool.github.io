import pandas as pd
from datetime import datetime
import os

# --- CONFIGURATION ---
excel_file = 'index.xlsm' 
sheet_name = 'index'
start_row_index = 19  # Excel Row 20
total_rows = 100      # Limit to 100 rows
start_col_index = 10  # Excel Column K

def generate_standings():
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found.")
        return

    # Load the sheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    rows_html = ""
    available_rows = df.shape[0]
    rows_to_process = min(total_rows, available_rows - start_row_index)

    for i in range(start_row_index, start_row_index + rows_to_process):
        # Extract data from Columns K through R (10-17)
        try:
            rank      = df.iloc[i, 10] # K
            player    = df.iloc[i, 11] # L
            total     = df.iloc[i, 13] # N
            group     = df.iloc[i, 14] # O
            bracket   = df.iloc[i, 15] # P
            possible  = df.iloc[i, 16] # Q
            bonus     = df.iloc[i, 17] # R

            # Clean NaN values
            r, p, t, g, b, ps, bn = [v if pd.notna(v) else "" for v in [rank, player, total, group, bracket, possible, bonus]]

            rows_html += f"""
            <tr>
                <td class="col-rank">{r}</td>
                <td class="col-name">{p}</td>
                <td class="col-total">{t}</td>
                <td class="col-grey">{g}</td>
                <td class="col-grey">{b}</td>
                <td class="col-grey">{ps}</td>
                <td class="col-grey">{bn}</td>
            </tr>"""
        except Exception as e:
            continue

    # --- HTML TEMPLATE (CSS Braces are escaped with {{ }}) ---
    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>WC 2026 Standings</title>
        <style>
            :root {{
                --maroon: #800000;
                --excel-green: #c6efce;
                --border-color: #ddd;
            }}
            body {{
                font-family: -apple-system, system-ui, sans-serif;
                background-color: #f4f4f9;
                margin: 0;
                padding: 10px;
            }}
            .table-title {{
                text-align: center;
                color: var(--maroon);
                font-family: 'Arial Black', sans-serif;
                font-size: 2.5rem;
                margin: 10px 0;
            }}
            .table-container {{
                width: 100%;
                overflow-x: auto;
                background: white;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                min-width: 800px;
            }}
            thead th {{
                position: sticky;
                top: 0;
                background-color: #eee;
                color: #333;
                font-size: 11px;
                font-weight: bold;
                text-transform: uppercase;
                padding: 12px 8px;
                border-bottom: 2px solid var(--maroon);
                z-index: 10;
            }}
            td {{
                padding: 10px;
                border-bottom: 1px solid var(--border-color);
                text-align: center;
                font-size: 14px;
            }}
            .col-rank {{ width: 50px; font-weight: bold; color: #666; }}
            .col-name {{ 
                text-align: left; 
                background-color: var(--excel-green); 
                font-weight: bold;
                position: sticky;
                left: 0;
                z-index: 5;
                border-right: 1px solid var(--border-color);
            }}
            .col-total {{ font-weight: bold; color: var(--maroon); }}
            .col-grey {{ color: #888; }}
            tr:nth-child(even) {{ background-color: #fafafa; }}
            tr:hover {{ background-color: #f1f8f5; }}
            .update-time {{ text-align: center; font-size: 11px; color: #999; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <h1 class="table-title">STANDINGS</h1>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>RANK</th>
                        <th>Participant</th>
                        <th>TOTAL POINTS</th>
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
        </div>
        <div class="update-time">Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
    </body>
    </html>
    """

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(full_html)

if __name__ == "__main__":
    generate_standings()