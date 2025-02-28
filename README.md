# Excel to PDF Converter and Shape Creator

This project provides multiple ways to convert Excel files to PDF:
1. Using Aspose.Cells for Python
2. Using Spire.XLS for Python
3. Using pandas and pdfkit (basic conversion)

## Prerequisites

- Docker
- Docker Compose

For local development:
- Python 3.6+
- JPype1 (for Aspose.Cells)
- Spire.XLS Python package

## Usage

1. **Build the Docker image:**

    ```bash
    docker-compose build
    ```

    **注意**: Dockerfileを変更した場合は、以下のコマンドでイメージを再ビルドしてください:

    ```bash
    docker-compose build --no-cache
    ```

2. **Run the script:**

    **Aspose Converter** (デフォルト):

    ```bash
    docker-compose run --service-ports excel-converter
    ```

    **Spire.XLS Converter**:

    ```bash
    docker-compose run --service-ports spire-converter
    ```

    **注意**: 既存のコンテナが残っている場合は、以下のコマンドでクリーンアップしてから再実行してください:
    
    ```bash
    docker-compose down --remove-orphans
    ```

    This will start the script in interactive mode. You will see the following menu:

    ```
    Menu:
      1: Convert Excel file(s) in 'input' folder to PDF
      2: Create a new Excel file
      3: Create a new Excel file with an AutoShape (Aspose only)
      q: Quit
    ```

    - Enter `1` to convert Excel files in the `input/` folder to PDF. You will be prompted to select a file.
    - Enter `2` to create a new Excel file. You will be prompted to enter a filename (without the extension).
    - Enter `3` to create a new Excel file with a shape (only available in Aspose converter). You will be prompted to enter a filename, select a shape type (Rectangle, Oval, or Line), and specify the shape's row, column, height, and width.
    - Enter `q` to quit the script.

3. **Using Converters Directly:**

    You can also run the converter scripts directly without Docker:

    ```bash
    # Aspose converter
    python aspose_excel_to_pdf.py
    
    # Spire.XLS converter
    python spirexls_excel_to_pdf.py
    ```

    This will start a similar interactive menu that allows you to:
    - Convert Excel files to PDF using the selected library
    - Create new Excel files

3. **Input and Output:**

    - Place Excel files to be converted in the `input/` folder.
    - Converted PDF files will be saved in the `output/` folder.
    - Newly created Excel files will be saved in the `input/` folder.

## Notes
- The platform is `linux/amd64`
- Spire.XLS for Python requires additional dependencies on Linux environments:
  - **libgdiplus**: Required for System.Drawing functionality
    ```bash
    sudo apt-get install -y libgdiplus libc6-dev
    sudo ln -s /usr/lib/libgdiplus.so /usr/lib/gdiplus.dll
    ```
  - **Japanese fonts**: Required for proper PDF rendering
    ```bash
    sudo apt-get install -y fonts-noto-cjk fonts-ipafont fonts-ipaexfont fonts-vlgothic
    sudo fc-cache -fv
    ```
  - All these dependencies are automatically installed in the Docker container
