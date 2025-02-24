# Excel to PDF Converter and Shape Creator

This script converts Excel files to PDF and creates new Excel files (with optional shapes) using Aspose.Cells for Python.

## Prerequisites

- Docker
- Docker Compose

## Usage

1.  **Build the Docker image:**

    ```bash
    docker-compose build
    ```

2.  **Run the script:**

    ```bash
    docker-compose run --service-ports excel-converter
    ```

    This will start the script in interactive mode. You will see the following menu:

    ```
    Menu:
      1: Convert Excel file(s) in 'input' folder to PDF
      2: Create a new Excel file
      3: Create a new Excel file with an AutoShape
      q: Quit
    ```

    -   Enter `1` to convert Excel files in the `input/` folder to PDF. You will be prompted to select a file.
    -   Enter `2` to create a new Excel file. You will be prompted to enter a filename (without the extension).
    -   Enter `3` to create a new Excel file with a shape. You will be prompted to enter a filename, select a shape type (Rectangle, Oval, or Line), and specify the shape's row, column, height, and width.
    -   Enter `q` to quit the script.

3.  **Input and Output:**

    -   Place Excel files to be converted in the `input/` folder.
    -   Converted PDF files will be saved in the `output/` folder.
    -   Newly created Excel files will be saved in the `input/` folder.

## Notes
- The platform is `linux/amd64`
