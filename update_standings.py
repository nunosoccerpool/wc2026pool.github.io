import pandas as pd
from datetime import datetime

def generate_html():
    # Load the specific Excel file and sheet
    df = pd.read_excel('index.xlsm', sheet_name='index', header=None, engine='openpyxl')

    rows_html = ""
    # Range: Excel Row 20 (index 19) to Row 200 (181 rows total)
    for i in range(19, 19 + 181):
        # Mapping Column K(10) to R(17)
        rows_html += f"""
        <tr>
            <td>{df.iloc[i, 10]}</td> <td class="player-name">{df.iloc[i, 11]}</td> <td></td> <td class="bold">{df.iloc[i, 13]}</td> <td>{df.iloc[i, 14]}</td> <td>{df.iloc[i, 15]}</td> <td>{df.iloc[i, 16]}</td> <td>{df.iloc[i, 17]}</td> </tr>"""

    # HTML Template with CSS Styling
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>World Cup Pool Standings</title>
        <style>
            body {{ font-family: sans-serif; background: #f4f7f6; padding: 20px; }}
            .table-container {{ max-width: 1000px; margin: auto; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }}
            table {{ width: 100%; border-collapse: collapse; }}
            th {{ background: #1b4332; color: white; padding: 12px; font-size: 0.8em; text-transform: uppercase; }}
            td {{ padding: 10px; border-bottom: 1px solid #eee; text-align: center; }}
            .player-name {{ text-align: left; font-weight: bold; }}
            .bold {{ font-weight: bold; color: #1b4332; }}
            tr:hover {{ background: #f1f8f5; }}
        </style>
    </head>
    <body>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Rank</th><th>Participant</th><th></th><th>Total Points</th>
                        <th>Group</th><th>Bracket</th><th>Possible</th><th>Bonus</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        <p style="text-align:center; color:#888; font-size:12px;">Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
    </body>
    </html>"""

    with open('index.html', 'w') as f:
        f.write(html_content)

if __name__ == "__main__":
    generate_html()