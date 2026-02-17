import pandas as pd
from datetime import datetime
import os

# --- CONFIGURATION ---
excel_file = 'index.xlsm' 
sheet_name = 'index'
start_row_index = 19  # Excel Row 20
total_rows = 100      
start_col_index = 10  # Excel Column K

def generate_standings():
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found.")
        return

    # Load data
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    rows_html = ""
    available_rows = df.shape[0]
    rows_to_process = min(total_rows, available_rows - start_row_index)

    for i in range(start_row_index, start_row_index + rows_to_process):
        try:
            rank   = df.iloc[i, 10] 
            player = df.iloc[i, 11] 
            total  = df.iloc[i, 13] 
            group  = df.iloc[i, 14] 
            bkout  = df.iloc[i, 15] 
            poss   = df.iloc[i, 16] 
            bonus  = df.iloc[i, 17] 

            # Clean NaN and convert to strings
            vals = [rank, player, total, group, bkout, poss, bonus]
            r, p, t, g, b, ps, bn = [str(v) if pd.notna(v) else "" for v in vals]

            rows_html += f"<tr><td class='col-rank'>{r}</td><td class='col-name'>{p}</td><td class='col-data bold-points'>{t}</td><td class='col-data'>{g}</td><td class='col-data'>{b}</td><td class='col-data'>{ps}</td><td class='col-data'>{bn}</td></tr>"
        except:
            continue

    # Define the template with escaped double braces for CSS
    template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WC 2026 Standings</title>
    <style>
        :root {{
            --maroon: #800000;
            --excel-green: rgba(198, 239, 206, 0.9);
            --border-color: #ccc;
        }}
        body {{
            font-family: -apple-system, system-ui, sans-serif;
            margin: 0;
            padding: 20px 10px;
            display: flex;
            flex-direction: column;
            align-items: center;
            background: linear-gradient(rgba(0,0,0,0.6), rgba(0,0,0,0.6)), 
                        url('https://images.unsplash.com/photo-1508098682722-e99c43a406b2?auto=format&fit=crop&w=1500');
            background-size: cover;
            background-attachment: fixed;
            background-position: center;
            min-height: 100vh;
        }}
        .table-title {{
            text-align: center;
            color: #ffffff;
            font-family: 'Arial Black', sans-serif;
            font-size: clamp(2rem, 8vw, 4rem);
            margin: 10px 0 20px 0;
            text-shadow: 3px 3px 10px rgba(0,0,0,0.8);
        }}
        .table-container {{
            width: 100%;
            max-width: 950px;
            overflow-x: auto;
            background: rgba(255, 255, 255, 0.92);
            border-radius: 12px;
            backdrop-filter: blur(8px);
            box-shadow: 0 10px 30px rgba(0,0,0,0.5);
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
        }}
        thead th {{
            position: sticky;
            top: 0;
            background-color: #222;
            color: white;
            font-size: 10px;
            font-weight: bold;
            text-transform: uppercase;
            padding: 15px 5px;
            border-bottom: 3px solid var(--maroon);
            z-index: 10;
        }}
        td {{
            padding: 12px 5px;
            border-bottom: 1px solid var(--border-color);
            text-align: center;
            font-size: 14px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: #111;
        }}
        .col-rank {{ width: 50px; font-weight: bold; color: #555; }}
        .col-name {{ 
            width: 180px; 
            text-align: left; 
            background-color: var(--excel-green); 
            font-weight: bold;
            padding-left: 10px;
        }}
        .col-data {{ width: 80px; }}
        .bold-points {{ font-weight: bold; color: var(--maroon); font-size: 16px; }}
        tr:nth-child(even) {{ background-color: rgba(245, 245, 245, 0.6); }}
        tr:hover {{ background-color: rgba(198, 239, 206, 0.5); }}
        .update-time {{ 
            text-align: center; 
            font-size: 12px; 
            color: #fff; 
            margin-top: 20px; 
            font-weight: bold;
            text-shadow: 1px 1px 5px rgba(0,0,0,0.5);
        }}
    </style>
</head>
<body>
    <h1 class="table-title">STANDINGS</h1>
    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th class="col-rank">RANK</th>
                    <th class="col-name">Participant</th>
                    <th class="col-data">TOTAL PTS</th>
                    <th class="col-data">Group</th>
                    <th class="col-data">Knockout</th>
                    <th class="col-data">Possible</th>
                    <th class="col-data">Bonus</th>
                </tr>
            </thead>
            <tbody>
                {content}
            </tbody>
        </table>
    </div>
    <div class="update-time">Tournament Data Live • Updated: {time}</div>
</body>
</html>
"""

    # Final string assembly
    final_html = template.format(content=rows_html, time=datetime.now().strftime('%b %d, %Y | %H:%M:%S'))

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(final_html)

if __name__ == "__main__":
    generate_standings()