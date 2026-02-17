import pandas as pd
from datetime import datetime
import os

# --- CONFIGURATION ---
excel_file = 'index.xlsm' 
sheet_name = 'index'
start_row_index = 19  
total_rows = 181
header_row_index = 18

def generate_standings():
    if not os.path.exists(excel_file):
        print(f"File {excel_file} not found.")
        return

    # Load the sheet - usecols forces Pandas to read at least up to Column Z (index 25)
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    # Safety check: If Excel is smaller than expected, pad it with empty columns
    # We need index 25 (Column Z)
    while df.shape[1] < 26:
        df[df.shape[1]] = ""

    # --- PULL TEAM NAMES DYNAMICALLY ---
    teams = []
    for col in range(18, 26): # S to Z
        val = df.iloc[header_row_index, col]
        teams.append(str(val).strip() if pd.notna(val) and str(val).lower() != 'nan' else "")

    rows_html = ""

    # --- LOOP PLAYERS ---
    for i in range(start_row_index, start_row_index + total_rows):
        # Using a helper to avoid crashes if rows are missing
        def get_val(r, c):
            try: return df.iloc[r, c]
            except: return ""

        rank      = get_val(i, 10) # K
        player    = get_val(i, 11) # L
        flag_code = str(get_val(i, 12)).lower().strip() # M
        total     = get_val(i, 13) # N
        group     = get_val(i, 14) # O
        bracket   = get_val(i, 15) # P
        possible  = get_val(i, 16) # Q
        bonus     = get_val(i, 17) # R

        # Predictions (S-Z)
        s = [get_val(i, col) for col in range(18, 26)]
        s = [val if pd.notna(val) and str(val).lower() != 'nan' else "" for val in s]

        flag_html = f'<img src="https://flagcdn.com/w20/{flag_code}.png" width="20">' if len(flag_code) == 2 else ""

        rows_html += f"""
        <tr>
            <td class="excel-cell bold center">{rank}</td>
            <td class="excel-cell player-bg">{player}</td>
            <td class="excel-cell center">{flag_html}</td>
            <td class="excel-cell bold center">{total}</td>
            <td class="excel-cell grey-text center">{group}</td>
            <td class="excel-cell grey-text center">{bracket}</td>
            <td class="excel-cell grey-text center">{possible}</td>
            <td class="excel-cell grey-text center border-right-maroon">{bonus}</td>
            
            <td class="pred-cell">{s[0]}</td><td class="pred-cell border-right-grey">{s[1]}</td>
            <td class="pred-cell">{s[2]}</td><td class="pred-cell border-right-grey">{s[3]}</td>
            <td class="pred-cell">{s[4]}</td><td class="pred-cell border-right-grey">{s[5]}</td>
            <td class="pred-cell">{s[6]}</td><td class="pred-cell">{s[7]}</td>
        </tr>
        """

    # --- HTML TEMPLATE (remains same as previous) ---
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; background: white; color: #333; }}
            .container {{ width: fit-content; margin: auto; text-align: center; }}
            h1 {{ color: #800000; font-size: 72px; margin: 10px 0 0 0; font-family: 'Arial Black'; }}
            .last-update {{ font-style: italic; color: #666; font-size: 14px; margin-bottom: 20px; }}
            table {{ border-collapse: collapse; margin: auto; border: 2px solid #800000; }}
            th {{ border: 1px solid #800000; padding: 8px; font-size: 11px; background: #e9e9e9; }}
            .vertical-text {{ writing-mode: vertical-rl; transform: rotate(180deg); font-size: 11px; padding: 8px 2px; width: 22px; height: 50px; text-align: left; }}
            .excel-cell {{ border: 1px solid #800000; padding: 6px; font-size: 14px; }}
            .player-bg {{ background-color: #c6efce; text-align: left; padding-left: 10px; min-width: 150px; }}
            .pred-cell {{ border: 1px solid #ccc; width: 25px; text-align: center; color: #555; font-size: 13px; }}
            .center {{ text-align: center; }}
            .bold {{ font-weight: bold; }}
            .grey-text {{ color: #999; }}
            .border-right-maroon {{ border-right: 2px solid #800000; }}
            .border-right-grey {{ border-right: 2px solid #999; }}
            .trophy {{ width: 60px; margin: 5px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>STANDINGS</h1>
            <img src="https://upload.wikimedia.org/wikipedia/en/thumb/e/e3/2026_FIFA_World_Cup_Logo.svg/1200px-2026_FIFA_World_Cup_Logo.svg.png" class="trophy">
            <div class="last-update">Last updated on:<br><b>{datetime.now().strftime('%b %d, %Y  %I:%M %p')}</b></div>
            <table>
                <thead>
                    <tr>
                        <th colspan="8" style="border:none;"></th>
                        <th colspan="8" style="color:#999; font-weight:normal; border:none; text-align:right;">Upcoming Match Predictions</th>
                    </tr>
                    <tr>
                        <th>RANK</th><th>PARTICIPANT</th><th></th><th>TOTAL POINTS</th>
                        <th style="color:#800000;">Group<br>Stage<br>Points</th>
                        <th style="color:#800000;">Bracket<br>Stage<br>Points</th>
                        <th style="color:#800000;">Possible<br>Points<br>Left</th>
                        <th style="color:#800000;">Bonus<br>Game</th>
                        <th class="vertical-text">{teams[0]}</th><th class="vertical-text">{teams[1]}</th>
                        <th class="vertical-text">{teams[2]}</th><th class="vertical-text">{teams[3]}</th>
                        <th class="vertical-text">{teams[4]}</th><th class="vertical-text">{teams[5]}</th>
                        <th class="vertical-text">{teams[6]}</th><th class="vertical-text">{teams[7]}</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
    </body>
    </html>
    """
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(full_html)

if __name__ == "__main__":
    generate_standings()