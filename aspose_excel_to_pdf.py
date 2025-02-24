import logging
from pathlib import Path
from datetime import datetime
from typing import Optional
import jpype
from aspose import cells
from aspose.pydrawing import Color

def excel_to_pdf_aspose(excel_file: str) -> None:
    """
    Converts an Excel file to PDF using Aspose.Cells.
    """
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

    try:
        # 出力ディレクトリの確認と作成
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
    
        # Excelファイルを読み込む
        workbook = cells.Workbook(excel_file)

        # Set default font for the workbook
        # workbook.default_font_name = "MS Gothic"

        # ワークシートを取得（ここでは最初のワークシート）
        worksheet = workbook.worksheets[0]

        # Print the number of rows in the original file
        print(f"Original number of rows: {worksheet.cells.max_row + 1}")

        # フォントとスタイルを設定
        style = workbook.create_style()
        style.font.name = "MS Gothic"
        style.font.size = 11

        # ヘッダーのスタイルを設定
        header_style = workbook.create_style()
        header_style.font.name = "MS Gothic"
        header_style.font.size = 12
        header_style.font.is_bold = True

        # ヘッダー行にスタイルを適用
        for col in range(worksheet.cells.max_column):
            cell = worksheet.cells.get(0, col)
            cell.set_style(header_style)

        # すべてのセルに罫線を設定
        for row in range(worksheet.cells.max_row + 1):
            for col in range(worksheet.cells.max_column + 1):
                cell = worksheet.cells.get(row, col)
                cell_style = cell.get_style()
                cell_style.set_border(cells.BorderType.TOP_BORDER, cells.CellBorderType.THIN, Color.black)
                cell_style.set_border(cells.BorderType.BOTTOM_BORDER, cells.CellBorderType.THIN, Color.black)
                cell_style.set_border(cells.BorderType.LEFT_BORDER, cells.CellBorderType.THIN, Color.black)
                cell_style.set_border(cells.BorderType.RIGHT_BORDER, cells.CellBorderType.THIN, Color.black)
                cell.set_style(cell_style)

        # 出力ファイル名を設定
        output_pdf = f"output/{Path(excel_file).stem}_aspose.pdf"

        # PDFの保存オプションを設定
        save_options = cells.PdfSaveOptions()
        # try:
        #     from aspose.cells import PdfFontEncoding
        #     save_options.font_encoding = PdfFontEncoding.IDENTITY_H
        # except ImportError:
        #     save_options.font_encoding = "UTF-8"

        # ページ設定
        worksheet.page_setup.orientation = cells.PageOrientationType.LANDSCAPE
        worksheet.page_setup.fit_to_pages_wide = 1
        worksheet.page_setup.left_margin = 5.08
        worksheet.page_setup.right_margin = 5.08
        worksheet.page_setup.top_margin = 5.08
        worksheet.page_setup.bottom_margin = 5.08

        # フッターを設定
        worksheet.page_setup.set_footer(0, f"作成日: {datetime.now().strftime('%Y年%m月%d日')}")
        worksheet.page_setup.set_footer(2, "ページ &P / &N")

        # 5行目の後に水平改ページを挿入
        worksheet.horizontal_page_breaks.add(5, 0)

        # Merge cells in the first row (header row)
        worksheet.cells.merge(0, 0, 1, worksheet.cells.max_column + 1)

        # Duplicate data to increase row count
        max_row = worksheet.cells.max_row
        max_col = worksheet.cells.max_column
        data = []
        for row in range(max_row + 1):
            row_data = []
            for col in range(max_col + 1):
                cell = worksheet.cells.get(row, col)
                row_data.append(cell.value)
            data.append(row_data)

        for i in range(10):  # Duplicate 10 times
            for row in range(1, max_row + 1): # Start from row 1 to skip header
                for col in range(max_col + 1):
                    worksheet.cells.get(max_row + 1 + (i * max_row) + row, col).value = data[row][col]

        # PDFとして保存
        workbook.save(output_pdf, save_options)
        print(f"Created PDF: {output_pdf}")

    except Exception as e:
        logger.error(f"Error during Excel to PDF conversion: {e}")

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

if __name__ == "__main__":
    import sys
    import glob
    import jpype

    # Start JVM here, outside the loop
    if not jpype.isJVMStarted():
        jpype.startJVM()

    while True:
        print("Menu:")
        print("  1: Convert Excel file(s) in 'input' folder to PDF")
        print("  2: Create a new Excel file")
        print("  3: Create a new Excel file with an AutoShape")
        print("  q: Quit")
        action = input("Enter your choice: ").strip()

        if action == '1':
            excel_files = glob.glob("input/*.xls*")
            if not excel_files:
                print("No Excel files found in the 'input' folder.")
            else:
                selected_file = select_file(excel_files)
                if selected_file:  # Check if a file was actually selected
                  excel_to_pdf_aspose(selected_file)

        elif action == '2':
            new_filename = input("Enter a name for the new Excel file (without extension): ").strip()
            create_excel_file(new_filename)

        elif action == '3':
            new_filename = input("Enter a name for the new Excel file (without extension): ").strip()
            shape_type, row, column, height, width = select_shape()
            create_excel_file(new_filename, shape_type, row, column, height, width)
        elif action.lower() == 'q':
            break
        else:
            print("Invalid action.")

    # Shutdown JVM after the loop exits
    jpype.shutdownJVM()
