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
        # Column M is blank
        total     = df.iloc[i, 13] # N
        group_p   = df.iloc[i, 14] # O
        bracket_p = df.iloc[i, 15] # P
        possible  = df.iloc[i, 16] # Q
        bonus     = df.iloc[i, 17] # R

        # 2. Predictions Data (S-Z) - 4 pairs of scores
        # Game 1: S(18), T(19) | Game 2: U(20), V(21) | Game 3: W(22), X(23) | Game 4: Y(24), Z(25)
        scores = [df.iloc[i, col] for col in range(18, 26)]
        # Replace empty scores with '-'
        scores = [s if pd.notna(s) else '-' for s in scores]

        rows_html += f"""
        <tr>
            <td class="main-cell rank">{rank}</td>
            <td class="main-cell player">{player}</td>
            <td class="main-cell">{total}</td>
            <td class="main-cell small">{group_p}</td>
            <td class="main-cell small">{bracket_p}</td>
            <td class="main-cell small">{possible}</td>
            <td class="main-cell small">{bonus}</td>
            
            <td class="gap"></td> <td class="score-cell left">{scores[0]}</td><td class="score-cell right">{scores[1]}</td>
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
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f4f7f6; color: #333; }}
            .wrapper {{ overflow-x: auto; padding: 20px; }}
            table {{ border-collapse: collapse; background: white; margin: auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            
            /* Header Styles */
            th {{ background: #1b4332; color: white; padding: 12px 8px; font-size: 11px; text-transform: uppercase; }}
            .prediction-header {{ background: #495057; }}
            
            /* Main Table Cells */
            .main-cell {{ padding: 10px; border-bottom: 1px solid #eee; text-align: center; border-right: 1px solid #f9f9f9; }}
            .rank {{ font-weight: bold; color: #666; }}
            .player {{ text-align: left; font-weight: bold; min-width: 150px; }}
            .small {{ font-size: 12px; color: #666; }}

            /* Score Box Styles */
            .score-cell {{ width: 30px; padding: 8px 0; text-align: center; border-bottom: 1px solid #eee; font-weight: bold; background: #fdfdfd; }}
            .left {{ border-left: 1px solid #ddd; border-right: 1px solid #eee; color: #007bff; }}
            .right {{ border-right: 1px solid #ddd; color: #dc3545; }}
            
            /* Spacing */
            .gap {{ width: 30px; background: #f4f7f6; border: none; }}
            .mini-gap {{ width: 8px; background: #f4f7f6; border: none; }}
            
            tr:hover {{ background-color: #f1f8f5; }}
        </style>
    </head>
    <body>
        <div class="wrapper">
            <h1 style="text-align:center; color:#1b4332;">World Cup 2026 Pool Dashboard</h1>
            <table>
                <thead>
                    <tr>
                        <th colspan="7">Tournament Standings</th>
                        <th class="gap"></th>
                        <th colspan="2" class="prediction-header">Match 1</th>
                        <th class="mini-gap"></th>
                        <th colspan="2" class="prediction-header">Match 2</th>
                        <th class="mini-gap"></th>
                        <th colspan="2" class="prediction-header">Match 3</th>
                        <th class="mini-gap"></th>
                        <th colspan="2" class="prediction-header">Match 4</th>
                    </tr>
                    <tr>
                        <th>Rank</th><th>Participant</th><th>Total</th>
                        <th>Grp</th><th>Bkt</th><th>Poss</th><th>Bon</th>
                        <th class="gap"></th>
                        <th>T1</th><th>T2</th><th></th><th>T1</th><th>T2</th><th></th><th>T1</th><th>T2</th><th></th><th>T1</th><th>T2</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
            <p style="text-align:center; font-size:12px; color:#888;">Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </body>
    </html>
    """

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(full_html)

if __name__ == "__main__":
    generate_standings()