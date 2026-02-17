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

    # Load data
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    rows_html = ""

    # Loop through the 181 participants
    for i in range(start_row_index, start_row_index + total_rows):
        # 1. Main Standings Data (K-R)
        rank      = df.iloc[i, 10] # K
        player    = df.iloc[i, 11] # L
        total     = df.iloc[i, 13] # N
        group_p   = df.iloc[i, 14] # O
        bracket_p = df.iloc[i, 15] # P
        possible  = df.iloc[i, 16] # Q
        bonus     = df.iloc[i, 17] # R

        # 2. Predictions Data (S-Z)
        # S(18), T(19) | U(20), V(21) | W(22), X(23) | Y(24), Z(25)
        scores = [df.iloc[i, col] for col in range(18, 26)]
        scores = [s if pd.notna(s) else '-' for s in scores]

        rows_html += f"""
        <tr>
            <td class="main-cell rank">{rank}</td>
            <td class="main-cell player">{player}</td>
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

    # --- HTML & CSS ---
    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: 'Segoe UI', Helvetica, Arial, sans-serif; background: #f8f9fa; color: #212529; padding: 20px; }}
            .wrapper {{ overflow-x: auto; }}
            table {{ border-collapse: collapse; background: white; margin: auto; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
            
            /* Headers */
            th {{ background: #1b4332; color: white; padding: 10px; font-size: 11px; text-transform: uppercase; border: 1px solid #143225; }}
            .upcoming-header {{ background: #2c3e50; font-size: 13px; letter-spacing: 1px; }}
            .match-label {{ background: #495057; font-size: 10px; }}
            
            /* Cells */
            .main-cell {{ padding: 12px 10px; border-bottom: 1px solid #eee; text-align: center; }}
            .rank {{ color: #888; font-weight: bold; width: 40px; }}
            .player {{ text-align: left; font-weight: bold; min-width: 160px; border-right: 1px solid #eee; }}
            .bold {{ font-weight: bold; color: #1b4332; font-size: 1.1em; }}
            .small {{ font-size: 11px; color: #777; }}

            /* Score Boxes */
            .score-cell {{ width: 28px; border-bottom: 1px solid #eee; font-weight: bold; text-align: center; background: #fff; }}
            .left {{ border-left: 1px solid #ddd; color: #0056b3; }}
            .right {{ border-right: 1px solid #ddd; color: #c82333; }}
            
            /* Gaps */
            .gap {{ width: 25px; background: #f8f9fa; border: none; }}
            .mini-gap {{ width: 6px; background: #f8f9fa; border: none; }}
            
            tr:nth-child(even) {{ background: #fdfdfd; }}
            tr:hover {{ background: #f1f8f5; }}
        </style>
    </head>
    <body>
        <div class="wrapper">
            <h1 style="text-align:center; color:#1b4332;">🏆 World Cup 2026 Pool</h1>
            <table>
                <thead>
                    <tr>
                        <th colspan="7">Current Standings</th>
                        <th class="gap"></th>
                        <th colspan="11" class="upcoming-header">Upcoming Match Predictions</th>
                    </tr>
                    <tr>
                        <th>Rank</th><th>Participant</th><th>Total</th>
                        <th>Group</th><th>Bkt</th><th>Poss</th><th>Bonus</th>
                        <th class="gap"></th>
                        <th colspan="2" class="match-label">Match 1</th>
                        <th class="mini-gap"></th>
                        <th colspan="2" class="match-label">Match 2</th>
                        <th class="mini-gap"></th>
                        <th colspan="2" class="match-label">Match 3</th>
                        <th class="mini-gap"></th>
                        <th colspan="2" class="match-label">Match 4</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
            <p style="text-align:center; font-size:11px; color:#aaa; margin-top:20px;">
                Generated from index.xlsm • {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </p>
        </div>
    </body>
    </html>
    """

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(full_html)

if __name__ == "__main__":
    generate_standings()