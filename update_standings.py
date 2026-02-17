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

    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    rows_html = ""
    available_rows = df.shape[0]
    rows_to_process = min(total_rows, available_rows - start_row_index)

    for i in range(start_row_index, start_row_index + rows_to_process):
        try:
            rank      = df.iloc[i, 10] 
            player    = df.iloc[i, 11] 
            total     = df.iloc[i, 13] 
            group     = df.iloc[i, 14] 
            bracket   = df.iloc[i, 15] 
            possible  = df.iloc[i, 16] 
            bonus     = df.iloc[i, 17] 

            r, p, t, g, b, ps, bn = [v if pd.notna(v) else "" for v in [rank, player, total, group, bracket, possible, bonus]]

            rows_html += f"""
            <tr>
                <td class="col-rank">{r}</td>
                <td class="col-name">{p}</td>
                <td class="col-data bold-points">{t}</td>
                <td class="col-data">{g}</td>
                <td class="col-data">{b}</td>
                <td class="col-data">{ps}</td>
                <td class="col-data">{bn}</td>
            </tr>"""
        except Exception:
            continue

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
                --excel-green: rgba(198, 239, 206, 0.9); /* Semi-transparent green */
                --border-color: #ccc;
            }}
            body {{
                font-family: -apple-system, system-ui, sans-serif;
                margin: 0;
                padding: 20px 10px;
                display: flex;
                flex-direction: column;
                align-items: center;
                /* Soccer Pitch Background */
                background: linear-gradient(rgba(0,0,0,0.5), rgba(0,0,0,0.5)), 
                            url('https://images.unsplash.com/photo-1508098682722-e99c43a406b2?q=80&w=2070&auto=format&fit=crop');
                background-size: cover;
                background-attachment: fixed;
                background-position: center;
                min-height: 100vh;
            }}
            .table-title {{
                text-align: center;
                color: #ffffff;
                font-family: 'Arial Black', sans-serif;
                font-size: clamp(2rem, 8vw, 4rem); /* Responsive sizing */
                margin: