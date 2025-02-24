import pandas as pd
import pdfkit
from pathlib import Path

def excel_to_pdf(excel_file):
    # Excelファイルを読み込む
    df = pd.read_excel(excel_file)
    
    # HTMLに変換
    html_content = df.to_html(index=False)
    
    # スタイルを追加
    styled_html = f"""
    <html>
        <head>
            <style>
                table {{
                    border-collapse: collapse;
                    width: 100%;
                }}
                th, td {{
                    border: 1px solid black;
                    padding: 8px;
                    text-align: left;
                }}
                th {{
                    background-color: #f2f2f2;
                }}
            </style>
            <meta charset="UTF-8">
        </head>
        <body style="font-family: IPAGothic, IPA Mincho;">
            <h1 style="text-align: center;">従業員情報一覧</h1>
            <p style="text-align: right;">作成日: 2025年2月22日</p>
            {html_content}
            <p style="margin-top: 20px;">※この文書は自動生成されています。</p>
        </body>
    </html>
    """

    # 出力ファイル名を設定（/app/outputディレクトリに保存）
    output_pdf = f"/app/output/{Path(excel_file).stem}.pdf"

    # PDFに変換（オプションでフォントを指定）
    options = {
        'encoding': "UTF-8",
        'custom-header': [
            ('Content-Encoding', 'utf-8')
        ],
    }
    pdfkit.from_string(styled_html, output_pdf, options=options)
    print(f"Created PDF: {output_pdf}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python excel_to_pdf.py <excel_file>")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    excel_to_pdf(excel_file)
