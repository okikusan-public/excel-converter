# Excel to PDF Converter and Shape Creator

This project provides multiple ways to convert Excel files to PDF:
1. Using Aspose.Cells for Python
2. Using Spire.XLS for Python
3. Using pandas and pdfkit (basic conversion)
4. Using Aspose.Cells for Node.js via Java (TypeScript implementation)

## Prerequisites

- Docker
- Docker Compose

For local development:
- Python 3.6+
- JPype1 (for Aspose.Cells Python)
- Spire.XLS Python package
- Node.js 14+ and npm (for TypeScript implementation)
- Java JDK 8+ (for Aspose.Cells Node.js)

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

    **Aspose Converter (Python)** (デフォルト):

    ```bash
    docker-compose run --service-ports excel-converter
    ```

    **Spire.XLS Converter (Python)**:

    ```bash
    docker-compose run --service-ports spire-converter
    ```

    **Aspose Converter (Node.js/TypeScript)**:

    ```bash
    cd nodejs && docker-compose run excel-converter-nodejs
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
    # Aspose converter (Python)
    python python/aspose_excel_to_pdf.py

    # Spire.XLS converter (Python)
    python python/spirexls_excel_to_pdf.py
    
    # Aspose converter (Node.js/TypeScript)
    cd nodejs && npm run dev
    ```

    This will start a similar interactive menu that allows you to:
    - Convert Excel files to PDF using the selected library
    - Create new Excel files
    - Create Excel files with shapes

3. **Input and Output:**

    **Python版**:
    - Place Excel files to be converted in the project root `input/` folder.
    - Converted PDF files will be saved in the project root `output/` folder.
    - Newly created Excel files will be saved in the project root `input/` folder.
    
    **Node.js版**:
    - Place Excel files to be converted in the `nodejs/input/` folder.
    - Converted PDF files will be saved in the `nodejs/output/` folder.
    - Newly created Excel files will be saved in the `nodejs/input/` folder.
    
    **注意**: Docker使用時も同じフォルダが共有されます。ホストマシンのフォルダにExcelファイルを配置すると、Dockerコンテナ内からアクセスできます。同様に、コンテナ内で生成されたPDFファイルはホストマシンのフォルダに保存されます。

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

## Node.js/TypeScript Implementation

The Node.js implementation uses Aspose.Cells for Node.js via Java with TypeScript for type safety. There are two ways to use this implementation:

### Using Docker (Recommended)

This method doesn't require Java to be installed on your local machine:

```bash
cd nodejs
docker-compose build  # 初回または変更時
docker-compose run excel-converter-nodejs
```

Docker環境では以下の共有フォルダが設定されています：
- `nodejs/input`: Excelファイルを配置するフォルダ（ホストとコンテナ間で共有）
- `nodejs/output`: 生成されたPDFファイルが保存されるフォルダ（ホストとコンテナ間で共有）
- `nodejs/aspose.cells`: Aspose.Cellsライブラリファイル

When the Docker container is running, you'll see a menu in the terminal. To interact with the application:

1. Type your selection (1, 2, 3, or q) and press Enter
2. Follow the prompts to select files, enter filenames, etc.

Example interaction:
```
メニュー:
  1: Excelファイルを'input'フォルダからPDFに変換
  2: 新しいExcelファイルを作成
  3: 図形を含む新しいExcelファイルを作成
  q: 終了
選択してください: 2
新しいExcelファイルの名前を入力してください（拡張子なし）: test
新しいExcelファイルが作成されました: input/test.xlsx

メニュー:
  1: Excelファイルを'input'フォルダからPDFに変換
  2: 新しいExcelファイルを作成
  3: 図形を含む新しいExcelファイルを作成
  q: 終了
選択してください: q
プログラムを終了します。
```

To stop the Docker container, press Ctrl+C in the terminal.

### Local Development

For local development without Docker:

1. **Prerequisites:**
   - Node.js 14+
   - Java JDK 8+ with JAVA_HOME environment variable set
   
   To set up Java:
   ```bash
   # macOS (using Homebrew)
   brew install openjdk@11
   echo 'export JAVA_HOME=$(/usr/libexec/java_home)' >> ~/.zshrc
   source ~/.zshrc
   
   # Ubuntu/Debian
   sudo apt install openjdk-11-jdk
   echo 'export JAVA_HOME=/usr/lib/jvm/java-11-openjdk-amd64' >> ~/.bashrc
   source ~/.bashrc
   
   # Windows
   # Install JDK from https://adoptium.net/
   # Set JAVA_HOME in Environment Variables
   ```

2. **Install dependencies:**

    ```bash
    cd nodejs
    npm install
    ```

3. **Build the TypeScript code:**

    ```bash
    npm run build
    ```

4. **Run the application:**

    ```bash
    npm start
    ```

   Or for development with automatic reloading:

    ```bash
    npm run dev
    ```

5. **Troubleshooting:**
   - If you get "Java is not installed or not in the system PATH" error, make sure Java is installed and JAVA_HOME is set correctly.
   - You can verify your Java installation with `java -version` and `echo $JAVA_HOME` (macOS/Linux) or `echo %JAVA_HOME%` (Windows).
   - If you encounter issues with Java setup, consider using the Docker method instead.
