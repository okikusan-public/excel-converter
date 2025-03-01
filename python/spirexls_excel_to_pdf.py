import logging
from pathlib import Path
from datetime import datetime
import time
import traceback
import glob
import sys
import os
import platform

# Spire.XLSのインポート（エラーハンドリング付き）
try:
    from spire.xls import Workbook, FileFormat, PageOrientationType
except ImportError as e:
    print("Spire.XLSのインポートに失敗しました。")
    print(f"エラー: {e}")
    
    # Linuxでlibgdiplusの問題が発生した場合のヘルプを表示
    if platform.system() == "Linux" and "libgdiplus" in str(e):
        print("\n===== Linux環境でのSpire.XLS使用に関する注意 =====")
        print("Spire.XLS for Pythonは、内部でSystem.Drawingライブラリを使用しており、")
        print("Linux上での動作には'libgdiplus'というネイティブライブラリが必要です。")
        print("\n以下のコマンドでlibgdiplusをインストールしてください:")
        print("  sudo apt-get update")
        print("  sudo apt-get install -y libgdiplus libc6-dev")
        print("\nそれでも問題が解決しない場合は、以下のコマンドを試してください:")
        print("  sudo ln -s /usr/lib/libgdiplus.so /usr/lib/gdiplus.dll")
        print("===============================================")
    
    sys.exit(1)

def add_custom_page_breaks(worksheet):
    """
    ユーザーが指定したカスタム改ページを追加する
    """
    while True:
        print("\nカスタム改ページの追加:")
        print("  1: 水平改ページ（行の後）を追加")
        print("  2: 垂直改ページ（列の後）を追加")
        print("  3: 完了")
        choice = input("選択してください: ").strip()
        
        if choice == "1":
            try:
                row = int(input("水平改ページを挿入する行番号（1から始まる）: ").strip())
                if row < 1:
                    print("無効な行番号です。1以上の値を入力してください。")
                    continue
                    
                # Spire.XLSでの水平改ページの追加
                worksheet.HPageBreaks.Add(row)
                print(f"行 {row} の後に水平改ページを追加しました。")
            except ValueError:
                print("有効な数値を入力してください。")
                
        elif choice == "2":
            try:
                col = int(input("垂直改ページを挿入する列番号（1から始まる）: ").strip())
                if col < 1:
                    print("無効な列番号です。1以上の値を入力してください。")
                    continue
                    
                # Spire.XLSでの垂直改ページの追加
                worksheet.VPageBreaks.Add(col)
                print(f"列 {col} の後に垂直改ページを追加しました。")
            except ValueError:
                print("有効な数値を入力してください。")
                
        elif choice == "3":
            break
        else:
            print("無効な選択です。1、2、または3を入力してください。")

    # 現在の改ページ設定を表示
    print("\n現在の改ページ設定:")
    print("水平改ページ:")
    if worksheet.HPageBreaks.Count == 0:
        print("  なし")
    else:
        for i in range(worksheet.HPageBreaks.Count):
            print(f"  行 {worksheet.HPageBreaks[i].Location.Row} の後")
            
    print("垂直改ページ:")
    if worksheet.VPageBreaks.Count == 0:
        print("  なし")
    else:
        for i in range(worksheet.VPageBreaks.Count):
            print(f"  列 {worksheet.VPageBreaks[i].Location.Column} の後")

def excel_to_pdf_spire(excel_file: str) -> None:
    """
    Converts an Excel file to PDF using Spire.XLS for Python.
    Excel上の設定（改ページ、印刷設定など）をそのままPDFに反映します。
    """
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)
    
    # 時間計測用
    start_time = time.time()

    try:
        # 出力ディレクトリの確認と作成
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
    
        print(f"Excelファイル '{excel_file}' をPDFに変換しています...")
        
        # Excelファイルを読み込む
        load_start_time = time.time()
        workbook = Workbook()
        
        # フォントパスを設定（Linux環境でのフォント問題対策）
        if platform.system() == "Linux":
            # 一般的なLinux環境でのフォントパス
            font_paths = [
                "/usr/share/fonts/truetype/ipafont/ipag.ttf",  # IPAゴシック
                "/usr/share/fonts/truetype/ipafont/ipam.ttf",  # IPA明朝
                "/usr/share/fonts/truetype/ipaexfont/ipaexg.ttf",  # IPAexゴシック
                "/usr/share/fonts/truetype/ipaexfont/ipaexm.ttf",  # IPAex明朝
                "/usr/share/fonts/truetype/vlgothic/VL-Gothic-Regular.ttf",  # VLゴシック
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",  # Noto Sans CJK
            ]
            
            # 存在するフォントパスのみを設定
            existing_font_paths = [path for path in font_paths if Path(path).exists()]
            if existing_font_paths:
                print(f"カスタムフォントパスを設定: {len(existing_font_paths)}個のフォントを使用")
                workbook.CustomFontFilePaths = existing_font_paths
            else:
                print("警告: システム上に適切なフォントが見つかりません。PDF変換が失敗する可能性があります。")
        
        workbook.LoadFromFile(excel_file)
        load_end_time = time.time()
        load_duration = load_end_time - load_start_time
        print(f"Excelファイルの読み込み完了: {load_duration:.2f}秒")
        logger.info(f"Excelファイルの読み込み時間: {load_duration:.2f}秒")

        # ワークシートを取得（ここでは最初のワークシート）
        worksheet = workbook.Worksheets[0]

        # 既存の印刷設定を表示（情報提供のため）
        print(f"印刷設定 - 用紙サイズ: {worksheet.PageSetup.PaperSize}")
        print(f"印刷設定 - 向き: {'横' if worksheet.PageSetup.Orientation == PageOrientationType.Landscape else '縦'}")
        print(f"印刷設定 - フィット設定: 幅={worksheet.PageSetup.FitToPagesWide}, 高さ={worksheet.PageSetup.FitToPagesTall}")

        # 改ページ情報を表示
        print("\n既存の改ページ情報:")
        try:
            if worksheet.HPageBreaks.Count > 0:
                print("水平改ページ:")
                for i in range(worksheet.HPageBreaks.Count):
                    print(f"  行 {worksheet.HPageBreaks[i].Location.Row} の後")
            else:
                print("水平改ページはありません")
                
            if worksheet.VPageBreaks.Count > 0:
                print("垂直改ページ:")
                for i in range(worksheet.VPageBreaks.Count):
                    print(f"  列 {worksheet.VPageBreaks[i].Location.Column} の後")
            else:
                print("垂直改ページはありません")
        except Exception as page_break_error:
            logger.error(f"改ページ情報の取得中にエラーが発生しました: {page_break_error}")
            print("改ページ情報の取得に失敗しました。処理を続行します。")

        # 出力ファイル名を設定
        output_pdf = f"output/{Path(excel_file).stem}_spire.pdf"

        # PDFの変換開始時間を記録
        pdf_start_time = time.time()
        
        # PDFに変換
        try:
            print("\nPDFに変換中...")
            logger.info("Spire.XLSでPDFに保存しています...")
            save_start_time = time.time()
            workbook.SaveToFile(output_pdf, FileFormat.PDF)
            save_end_time = time.time()
            save_duration = save_end_time - save_start_time
            pdf_end_time = time.time()
            pdf_duration = pdf_end_time - pdf_start_time
            total_duration = pdf_end_time - start_time
            print(f"PDFファイルが作成されました: {output_pdf} (保存時間: {save_duration:.2f}秒, PDF変換合計: {pdf_duration:.2f}秒, 総処理時間: {total_duration:.2f}秒)")
            logger.info(f"PDF変換・保存時間: {save_duration:.2f}秒")
        except Exception as pdf_error:
            logger.error(f"PDF変換中にエラーが発生しました: {pdf_error}")
            print(f"PDF変換に失敗しました: {pdf_error}")
                
        # 処理時間のサマリーを表示
        end_time = time.time()
        total_duration = end_time - start_time
        print(f"\n総処理時間: {total_duration:.2f}秒")

    except Exception as e:
        logger.error(f"Excel→PDF変換中にエラーが発生しました: {e}")
        # スタックトレースを表示（デバッグ用）
        logger.error(traceback.format_exc())
        
        # エラー時も処理時間を表示
        end_time = time.time()
        error_duration = end_time - start_time
        print(f"エラーが発生しました。エラーまでの処理時間: {error_duration:.2f}秒")

def create_excel_file(filename: str) -> None:
    """
    Creates a new Excel file with a sample sheet and data.
    """
    try:
        # Create a new workbook
        workbook = Workbook()

        # Get the first worksheet
        worksheet = workbook.Worksheets[0]

        # Add some data
        worksheet.Range["A1"].Text = "Name"
        worksheet.Range["B1"].Text = "Age"
        worksheet.Range["A2"].Text = "John Doe"
        worksheet.Range["B2"].NumberValue = 30
        worksheet.Range["A3"].Text = "Jane Smith"
        worksheet.Range["B3"].NumberValue = 25

        # Save the workbook
        input_dir = Path("input")
        input_dir.mkdir(exist_ok=True)
        workbook.SaveToFile(str(input_dir / f"{filename}.xlsx"))
        print(f"Created new Excel file: input/{filename}.xlsx")

    except Exception as e:
        print(f"Error creating Excel file: {e}")

def select_file(files):
    """
    Presents a list of files to the user and returns the selected file path.
    """
    print("Available Excel files:")
    for i, file in enumerate(files):
        print(f"{i + 1}. {file}")
    while True:
        try:
            choice = int(input("Enter the number of the file to convert: "))
            if 1 <= choice <= len(files):
                return files[choice - 1]
            else:
                print("Invalid choice. Please enter a number from the list.")
        except ValueError:
            print("Invalid input. Please enter a number.")

if __name__ == "__main__":
    while True:
        print("\nメニュー:")
        print("  1: Excelファイルを'input'フォルダからPDFに変換")
        print("  2: 新しいExcelファイルを作成")
        print("  q: 終了")
        action = input("選択してください: ").strip()

        if action == '1':
            excel_files = glob.glob("input/*.xls*")
            if not excel_files:
                print("'input'フォルダにExcelファイルが見つかりません。")
            else:
                selected_file = select_file(excel_files)
                if selected_file:  # Check if a file was actually selected
                    excel_to_pdf_spire(selected_file)

        elif action == '2':
            new_filename = input("新しいExcelファイルの名前を入力してください（拡張子なし）: ").strip()
            create_excel_file(new_filename)

        elif action.lower() == 'q':
            break
        else:
            print("無効な選択です。")
