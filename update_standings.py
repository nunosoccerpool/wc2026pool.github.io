import pandas as pd
from datetime import datetime
import os

# --- CONFIGURATION ---
excel_file = 'index.xlsm' 
sheet_name = 'index'
start_row_index = 19  # Excel Row 20
total_rows = 100      # Limit to 100 rows as requested
start_col_index = 10  # Excel Column K

def generate_standings():
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found.")
        return

    # Load the sheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    # Safety check: ensure we don't go out of bounds if the sheet is small
    available_rows = df.shape[0]
    rows_to_process = min(total_rows, available_rows - start_row_index)

    rows_html = ""

    # Loop to generate exactly 100 rows
    for i in range(start_row_index, start_row_index + rows_to_process):
        # Extract data from Columns K through R
        rank      = df.iloc[i, 10] # K
        player    = df.iloc[i, 11] # L
        # Column M (12) is skipped/blank per your layout
        total     = df.iloc[i, 13] # N
        group     = df.iloc[i, 14] # O
        bracket   = df.iloc[i, 15] # P
        possible  = df.iloc[i, 16] # Q
        bonus     = df.iloc[i, 17] # R

        # Clean NaN values
        fields = [rank, player, total, group, bracket, possible, bonus]
        r, p, t, g, b, ps, bn = [v if pd.notna(v) else "" for v in fields]

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

    # --- MOBILE-FRIENDLY HTML5 TEMPLATE ---
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