import logging
from pathlib import Path
from datetime import datetime
from typing import Optional
import jpype
from aspose import cells
from aspose.pydrawing import Color

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
                row = int(input("水平改ページを挿入する行番号（1から始まる）: ").strip()) - 1
                if row < 0:
                    print("無効な行番号です。0以上の値を入力してください。")
                    continue
                    
                # 既存の改ページかどうかチェック
                exists = False
                for page_break in worksheet.horizontal_page_breaks:
                    if page_break.row == row:
                        exists = True
                        break
                        
                if exists:
                    print(f"行 {row + 1} の後には既に改ページが設定されています。")
                else:
                    worksheet.horizontal_page_breaks.add(row, 0)
                    print(f"行 {row + 1} の後に水平改ページを追加しました。")
            except ValueError:
                print("有効な数値を入力してください。")
                
        elif choice == "2":
            try:
                col = int(input("垂直改ページを挿入する列番号（1から始まる）: ").strip()) - 1
                if col < 0:
                    print("無効な列番号です。0以上の値を入力してください。")
                    continue
                    
                # 既存の改ページかどうかチェック
                exists = False
                for page_break in worksheet.vertical_page_breaks:
                    if page_break.column == col:
                        exists = True
                        break
                        
                if exists:
                    print(f"列 {col + 1} の後には既に改ページが設定されています。")
                else:
                    worksheet.vertical_page_breaks.add(col, 0)
                    print(f"列 {col + 1} の後に垂直改ページを追加しました。")
            except ValueError:
                print("有効な数値を入力してください。")
                
        elif choice == "3":
            break
        else:
            print("無効な選択です。1、2、または3を入力してください。")

    # 現在の改ページ設定を表示
    print("\n現在の改ページ設定:")
    print("水平改ページ:")
    try:
        # countはメソッドなので()を付けて呼び出す
        if worksheet.horizontal_page_breaks.count() == 0:
            print("  なし")
        else:
            for page_break in worksheet.horizontal_page_breaks:
                print(f"  行 {page_break.row + 1} の後")
    except Exception:
        # lenを使用する代替法
        try:
            if len(worksheet.horizontal_page_breaks) == 0:
                print("  なし")
            else:
                for page_break in worksheet.horizontal_page_breaks:
                    print(f"  行 {page_break.row + 1} の後")
        except Exception:
            print("  改ページ情報の取得に失敗しました。")
            
    print("垂直改ページ:")
    try:
        # countはメソッドなので()を付けて呼び出す
        if worksheet.vertical_page_breaks.count() == 0:
            print("  なし")
        else:
            for page_break in worksheet.vertical_page_breaks:
                print(f"  列 {page_break.column + 1} の後")
    except Exception:
        # lenを使用する代替法
        try:
            if len(worksheet.vertical_page_breaks) == 0:
                print("  なし")
            else:
                for page_break in worksheet.vertical_page_breaks:
                    print(f"  列 {page_break.column + 1} の後")
        except Exception:
            print("  改ページ情報の取得に失敗しました。")

def excel_to_pdf_aspose(excel_file: str) -> None:
    """
    Converts an Excel file to PDF using Aspose.Cells.
    Excel上の設定（改ページ、印刷設定など）をそのままPDFに反映します。
    """
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)
    
    # 時間計測用
    import time
    start_time = time.time()

    try:
        # 出力ディレクトリの確認と作成
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
    
        print(f"Excelファイル '{excel_file}' をPDFに変換しています...")
        
        # Excelファイルを読み込む
        load_start_time = time.time()
        workbook = cells.Workbook(excel_file)
        load_end_time = time.time()
        load_duration = load_end_time - load_start_time
        print(f"Excelファイルの読み込み完了: {load_duration:.2f}秒")
        logger.info(f"Excelファイルの読み込み時間: {load_duration:.2f}秒")

        # ワークシートを取得（ここでは最初のワークシート）
        worksheet = workbook.worksheets[0]

        # 既存の印刷設定を表示（情報提供のため）
        print(f"印刷設定 - 用紙サイズ: {worksheet.page_setup.paper_size}")
        print(f"印刷設定 - 向き: {'横' if worksheet.page_setup.orientation == cells.PageOrientationType.LANDSCAPE else '縦'}")
        print(f"印刷設定 - フィット設定: 幅={worksheet.page_setup.fit_to_pages_wide}, 高さ={worksheet.page_setup.fit_to_pages_tall}")

        # 改ページ情報を表示
        print("\n既存の改ページ情報:")
        try:
            # countはメソッドなので()を付けて呼び出す
            horizontal_breaks_count = worksheet.horizontal_page_breaks.count()
            if horizontal_breaks_count > 0:
                print("水平改ページ:")
                for page_break in worksheet.horizontal_page_breaks:
                    print(f"  行 {page_break.row + 1} の後")
            else:
                print("水平改ページはありません")
                
            vertical_breaks_count = worksheet.vertical_page_breaks.count()
            if vertical_breaks_count > 0:
                print("垂直改ページ:")
                for page_break in worksheet.vertical_page_breaks:
                    print(f"  列 {page_break.column + 1} の後")
            else:
                print("垂直改ページはありません")
        except Exception as page_break_error:
            logger.error(f"改ページ情報の取得中にエラーが発生しました: {page_break_error}")
            # 代替方法：lengthプロパティを試す
            try:
                logger.info("代替方法で改ページ情報を取得します...")
                # lenまたはlengthプロパティを試みる
                if len(worksheet.horizontal_page_breaks) > 0:
                    print("水平改ページ:")
                    for page_break in worksheet.horizontal_page_breaks:
                        print(f"  行 {page_break.row + 1} の後")
                else:
                    print("水平改ページはありません")
                    
                if len(worksheet.vertical_page_breaks) > 0:
                    print("垂直改ページ:")
                    for page_break in worksheet.vertical_page_breaks:
                        print(f"  列 {page_break.column + 1} の後")
                else:
                    print("垂直改ページはありません")
            except Exception as alt_page_break_error:
                logger.error(f"代替方法での改ページ情報取得中にエラーが発生しました: {alt_page_break_error}")
                print("改ページ情報の取得に失敗しました。処理を続行します。")

        # 出力ファイル名を設定
        output_pdf = f"output/{Path(excel_file).stem}_aspose.pdf"

        # PDFの変換開始時間を記録
        pdf_start_time = time.time()
        
        # PDFに直接変換する（PdfSaveOptionsを使わない方法）
        # シンプルな方法でプロパティエラーを回避
        try:
            print("\nPDFに変換中...")
            logger.info("標準的な方法でPDFに保存しています...")
            # 最もシンプルな方法 - セーブオプションなしで保存
            save_start_time = time.time()
            workbook.save(output_pdf)
            save_end_time = time.time()
            save_duration = save_end_time - save_start_time
            pdf_end_time = time.time()
            pdf_duration = pdf_end_time - pdf_start_time
            total_duration = pdf_end_time - start_time
            print(f"PDFファイルが作成されました: {output_pdf} (保存時間: {save_duration:.2f}秒, PDF変換合計: {pdf_duration:.2f}秒, 総処理時間: {total_duration:.2f}秒)")
            logger.info(f"PDF変換・保存時間: {save_duration:.2f}秒")
        except Exception as pdf_error:
            logger.error(f"標準的な方法でのPDF変換中にエラーが発生しました: {pdf_error}")
            print(f"標準的な方法でのPDF変換に失敗しました: {pdf_error}")
            
            # 代替方法
            try:
                print("\n代替方法でPDFに変換中...")
                logger.info("代替方法でPDFに保存しています...")
                # 別の方法 - セーブオプションを最小限に設定
                save_options = cells.PdfSaveOptions()
                # 特に設定を追加せず、デフォルト設定で保存
                save_start_time = time.time()
                workbook.save(output_pdf, save_options)
                save_end_time = time.time()
                save_duration = save_end_time - save_start_time
                pdf_end_time = time.time()
                pdf_duration = pdf_end_time - pdf_start_time
                total_duration = pdf_end_time - start_time
                print(f"PDFファイルが作成されました: {output_pdf} (代替方法での保存時間: {save_duration:.2f}秒, PDF変換合計: {pdf_duration:.2f}秒, 総処理時間: {total_duration:.2f}秒)")
                logger.info(f"代替方法でのPDF変換・保存時間: {save_duration:.2f}秒")
            except Exception as alt_error:
                logger.error(f"代替方法でのPDF変換中にエラーが発生しました: {alt_error}")
                print(f"代替方法でのPDF変換に失敗しました: {alt_error}")
                
                # 最終手段
                try:
                    print("\n最終手段でPDFに変換中...")
                    logger.info("最終手段でPDFに保存しています...")
                    import os
                    raw_output = str(Path(output_dir) / f"{Path(excel_file).stem}_raw.pdf")
                    save_start_time = time.time()
                    workbook.save(raw_output)
                    save_end_time = time.time()
                    save_duration = save_end_time - save_start_time
                    pdf_end_time = time.time()
                    pdf_duration = pdf_end_time - pdf_start_time
                    total_duration = pdf_end_time - start_time
                    print(f"PDFファイルが作成されました（最終手段）: {raw_output} (最終手段での保存時間: {save_duration:.2f}秒, PDF変換合計: {pdf_duration:.2f}秒, 総処理時間: {total_duration:.2f}秒)")
                    logger.info(f"最終手段でのPDF変換・保存時間: {save_duration:.2f}秒")
                except Exception as final_error:
                    logger.error(f"最終手段でのPDF変換中にエラーが発生しました: {final_error}")
                    print(f"最終手段でのPDF変換に失敗しました: {final_error}")
                    
        # 処理時間のサマリーを表示
        end_time = time.time()
        total_duration = end_time - start_time
        print(f"\n総処理時間: {total_duration:.2f}秒")

    except Exception as e:
        logger.error(f"Excel→PDF変換中にエラーが発生しました: {e}")
        # スタックトレースを表示（デバッグ用）
        import traceback
        logger.error(traceback.format_exc())
        
        # エラー時も処理時間を表示
        end_time = time.time()
        error_duration = end_time - start_time
        print(f"エラーが発生しました。エラーまでの処理時間: {error_duration:.2f}秒")

def create_excel_file(filename: str, shape_type=None, row=None, column=None, height=None, width=None) -> None:
    """
    Creates a new Excel file with a sample sheet and data.
    Optionally adds a shape if shape_type and position are provided.
    """
    try:
        # Create a new workbook
        workbook = cells.Workbook()

        # Get the first worksheet
        worksheet = workbook.worksheets[0]

        # Add some data
        worksheet.cells.get("A1").put_value("Name")
        worksheet.cells.get("B1").put_value("Age")
        worksheet.cells.get("A2").put_value("John Doe")
        worksheet.cells.get("B2").put_value(30)
        worksheet.cells.get("A3").put_value("Jane Smith")
        worksheet.cells.get("B3").put_value(25)

        # Add shape if provided
        if shape_type and row is not None and column is not None and height is not None and width is not None:
            worksheet.shapes.add_shape(
                shape_type,
                row,
                0,
                column,
                0,
                height,
                width
            )

        # Save the workbook
        input_dir = Path("input")
        input_dir.mkdir(exist_ok=True)
        workbook.save(str(input_dir / f"{filename}.xlsx")) # Convert Path to string
        print(f"Created new Excel file: input/{filename}.xlsx")

    except Exception as e:
        print(f"Error creating Excel file: {e}")

def select_shape():
    """
    Presents a list of AutoShape types to the user and returns the selected type, row, column, height, and width.
    """
    print("Available AutoShape types:")
    print("  1. Rectangle")
    print("  2. Oval")
    print("  3. Line")
    while True:
        try:
            choice = int(input("Enter the number of the shape to insert: "))
            if 1 <= choice <= 3:
                shape_type = None
                if choice == 1:
                    shape_type = cells.drawing.MsoDrawingType.RECTANGLE
                elif choice == 2:
                    shape_type = cells.drawing.MsoDrawingType.OVAL
                elif choice == 3:
                    shape_type = cells.drawing.MsoDrawingType.LINE

                print("Enter the position and size for the shape:")
                row = int(input("  Row (1-based index): "))
                column = int(input("  Column (1-based index): "))
                height = int(input("  Height: "))
                width = int(input("  Width: "))

                print(f"Shape Type: {shape_type}")
                print(f"Position: row={row}, column={column}, height={height}, width={width}")

                return shape_type, row, column, height, width
            else:
                print("Invalid choice. Please enter a number from the list.")
        except ValueError:
            print("Invalid input. Please enter a number or valid cell reference.")
        except Exception as e:
            print(f"Error in select_shape: {e}")

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
    import sys
    import glob
    import jpype

    # Start JVM here, outside the loop
    if not jpype.isJVMStarted():
        jpype.startJVM()

    while True:
        print("\nメニュー:")
        print("  1: Excelファイルを'input'フォルダからPDFに変換")
        print("  2: 新しいExcelファイルを作成")
        print("  3: 図形を含む新しいExcelファイルを作成")
        print("  q: 終了")
        action = input("選択してください: ").strip()

        if action == '1':
            excel_files = glob.glob("input/*.xls*")
            if not excel_files:
                print("'input'フォルダにExcelファイルが見つかりません。")
            else:
                selected_file = select_file(excel_files)
                if selected_file:  # Check if a file was actually selected
                    excel_to_pdf_aspose(selected_file)

        elif action == '2':
            new_filename = input("新しいExcelファイルの名前を入力してください（拡張子なし）: ").strip()
            create_excel_file(new_filename)

        elif action == '3':
            new_filename = input("新しいExcelファイルの名前を入力してください（拡張子なし）: ").strip()
            shape_type, row, column, height, width = select_shape()
            create_excel_file(new_filename, shape_type, row, column, height, width)
        elif action.lower() == 'q':
            break
        else:
            print("無効な選択です。")

    # Shutdown JVM after the loop exits
    jpype.shutdownJVM()
