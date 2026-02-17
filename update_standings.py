import pandas as pd
from datetime import datetime
import os

# --- CONFIGURATION ---
excel_file = 'index.xlsm' 
sheet_name = 'index'
start_row_index = 19  # Excel Row 20
total_rows = 181
start_col_index = 10  # Excel Column K

def generate_standings():
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found.")
        return

    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')
    rows_html = ""

    for i in range(start_row_index, start_row_index + total_rows):
        # Data Extraction
        rank      = df.iloc[i, 10]
        player    = df.iloc[i, 11]
        total     = df.iloc[i, 13]
        group_p   = df.iloc[i, 14]
        bracket_p = df.iloc[i, 15]
        possible  = df.iloc[i, 16]
        bonus     = df.iloc[i, 17]

        scores = [df.iloc[i, col] for col in range(18, 26)]
        scores = [s if pd.notna(s) else '-' for s in scores]

        rows_html += f"""
        <tr>
            <td class="main-cell rank">{rank}</td>
            <td class="main-cell player sticky-col">{player}</td>
            <td class="main-cell bold">{total}</td>
            <td class="main-cell small">{group_p}</td>
            <td class="main-cell small">{bracket_p}</td>
            <td class="main-cell small">{possible}</td>
            <td class="main-cell small">{bonus}</td>
            
            <td class="gap"></td> 

            <td class="score-cell left">{scores[0]}</td><td class="score-cell right">{scores[1]}</td>
            <td class="mini-gap"></td>
            <td class="score-cell left">{scores[2]}</td><td class="score-cell right">{scores[3]}</td>
            <td class="mini-gap"></td>
            <td class="score-cell left">{scores[4]}</td><td class="score-cell right">{scores[5]}</td>
            <td class="mini-gap"></td>
            <td class="score-cell left">{scores[6]}</td><td class="score-cell right">{scores[7]}</td>
        </tr>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f8f9fa; margin: 0; padding: 20px; }}
            
            /* The fix for the horizontal explosion */
            .outer-wrapper {{
                max-width: 100%;
                overflow-x: auto; /* Forces scrollbar inside the container */
                border-radius: 8px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            }}

            table {{ border-collapse: collapse; background: white; white-space: nowrap; min-width: 1000px; }}
            
            th {{ background: #1b4332; color: white; padding: 12px; font-size: 11px; position: sticky; top: 0; z-index: 10; }}
            .upcoming-header {{ background: #2c3e50; border-left: 2px solid #fff; }}
            
            .main-cell {{ padding: 12px 10px; border-bottom: 1px solid #eee; text-align: center; }}
            
            /* Sticky Column for Participant Name */
            .sticky-col {{
                position: sticky;
                left: 0;
                background: white;
                z-index: 5;
                border-right: 2px solid #1b4332 !important;
            }}
            th.sticky-col {{ z-index: 11; background: #1b4332; }}

            .rank {{ color: #888; width: 40px; }}
            .player {{ text-align: left; font-weight: bold; min-width: 160px; }}
            .bold {{ font-weight: bold; color: #1b4332; }}
            .score-cell {{ width: 30px; font-weight: bold; text-align: center; background: #fdfdfd; border-bottom: 1px solid #eee; }}
            .left {{ color: #0056b3; border-left: 1px solid #ddd; }}
            .right {{ color: #c82333; border-right: 1px solid #ddd; }}
            
            .gap {{ width: 20px; background: #f8f9fa; }}
            .mini-gap {{ width: 4px; background: #f8f9fa; }}
            
            tr:hover .main-cell, tr:hover .score-cell {{ background-color: #f1f8f5; }}
            tr:hover .sticky-col {{ background-color: #f1f8f5; }}
        </style>
    </head>
    <body>
        <h1 style="text-align:center; color:#1b4332;">🏆 World Cup 2026 Pool</h1>
        <div class="outer-wrapper">
            <table>
                <thead>
                    <tr>
                        <th colspan="7">Tournament Standings</th>
                        <th class="gap"></th>
                        <th colspan="11" class="upcoming-header">Upcoming Match Predictions</th>
                    </tr>
                    <tr>
                        <th class="rank">#</th>
                        <th class="player sticky-col">Participant</th>
                        <th>Total</th>
                        <th>Grp</th><th>Bkt</th><th>Poss</th><th>Bon</th>
                        <th class="gap"></th>
                        <th colspan="2">Match 1</th><th class="mini-gap"></th>
                        <th colspan="2">Match 2</th><th class="mini-gap"></th>
                        <th colspan="2">Match 3</th><th class="mini-gap"></th>
                        <th colspan="2">Match 4</th>
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