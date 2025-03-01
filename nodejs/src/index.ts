import * as path from 'path';
import * as fs from 'fs-extra';
import * as readline from 'readline';
import { ExcelConverter } from './converter';

// Create input/output directories if they don't exist
fs.ensureDirSync(path.join(__dirname, '..', 'input'));
fs.ensureDirSync(path.join(__dirname, '..', 'output'));

// Create readline interface for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Initialize the Excel converter
const converter = new ExcelConverter();

/**
 * Display the main menu and handle user input
 */
function showMenu(): void {
    // Clear the mapping file at the start of each menu display
    const mappingFilePath = path.join(__dirname, '..', 'output', 'filename_mapping.json');
    try {
      if (fs.existsSync(mappingFilePath)) {
          fs.unlinkSync(mappingFilePath);
          console.log("Deleted filename_mapping.json");
      }
    } catch (err) {
      console.error("Failed to delete filename_mapping.json:", err);
      // Proceed anyway. We'll overwrite it later.
    }

  console.log('\nメニュー:');
  console.log('  1: Excelファイルを\'input\'フォルダからPDFに変換');
  console.log('  2: 新しいExcelファイルを作成');
  console.log('  3: 図形を含む新しいExcelファイルを作成');
  console.log('  q: 終了');

  rl.question('選択してください: ', (answer) => {
    switch (answer.trim()) {
      case '1':
        convertExcelToPdf();
        break;
      case '2':
        createNewExcelFile();
        break;
      case '3':
        createExcelFileWithShape();
        break;
      case 'q':
        console.log('プログラムを終了します。');
        rl.close();
        break;
      default:
        console.log('無効な選択です。');
        showMenu();
        break;
    }
  });
}

/**
 * Convert Excel file to PDF
 */
function convertExcelToPdf(): void {
  const inputDir = path.join(__dirname, '..', 'input');

  // Get all Excel files in the input directory
  const excelFiles = fs.readdirSync(inputDir)
    .filter(file => /\.(xls|xlsx|xlsm|xlsb)$/i.test(file))
    .map(file => path.join(inputDir, file));

  if (excelFiles.length === 0) {
    console.log('\'input\'フォルダにExcelファイルが見つかりません。');
    showMenu();
    return;
  }

  // Display available Excel files
  console.log('利用可能なExcelファイル:');
  excelFiles.forEach((file, index) => {
    console.log(`${index + 1}. ${path.basename(file)}`);
  });

  // Ask user to select a file
  rl.question('変換するファイルの番号を入力してください: ', (answer) => {
    const fileIndex = parseInt(answer.trim()) - 1;

    if (isNaN(fileIndex) || fileIndex < 0 || fileIndex >= excelFiles.length) {
      console.log('無効な選択です。');
      convertExcelToPdf(); // Call recursively to re-prompt
      return;
    }

    const selectedFile = excelFiles[fileIndex];

    // Log file information for debugging
    console.log(`選択されたファイル: ${selectedFile}`);
    console.log(`ファイルの存在確認: ${fs.existsSync(selectedFile) ? '存在します' : '存在しません'}`);

    // Convert the selected file to PDF and then rename files
    converter.convertToPdf(selectedFile)
    .then(() => {
      // Run the rename script after conversion *completes*
      const renameScriptPath = path.join(__dirname, 'rename_pdfs.js');
      if (fs.existsSync(renameScriptPath)) {
        console.log('PDF変換後にファイル名を修正します...');
        const { exec } = require('child_process');
        exec(`node ${renameScriptPath}`, (error: any, stdout: string, stderr: string) => {
          if (error) {
            console.error(`ファイル名修正スクリプトの実行エラー: ${error}`);
          } else {
            console.log(stdout);
            if (stderr) {
              console.error(`ファイル名修正スクリプトのエラー出力: ${stderr}`);
            }
          }
          showMenu(); // Show the menu *after* renaming, and always show it.
        });
      } else {
        console.warn('ファイル名修正スクリプトが見つかりません:', renameScriptPath);
        showMenu(); // Show menu if script not found.
      }
    })
    .catch(error => {
      console.error('変換中にエラーが発生しました:', error);
      showMenu(); // Show menu on error.
    });
  });
}

/**
 * Create a new Excel file
 */
function createNewExcelFile(): void {
  rl.question('新しいExcelファイルの名前を入力してください（拡張子なし）: ', (filename) => {
    if (!filename.trim()) {
      console.log('ファイル名を入力してください。');
      createNewExcelFile();
      return;
    }

    converter.createExcelFile(filename.trim())
      .then(() => {
        showMenu();
      })
      .catch(error => {
        console.error('Excelファイルの作成中にエラーが発生しました:', error);
        showMenu();
      });
  });
}

/**
 * Create a new Excel file with a shape
 */
function createExcelFileWithShape(): void {
  rl.question('新しいExcelファイルの名前を入力してください（拡張子なし）: ', (filename) => {
    if (!filename.trim()) {
      console.log('ファイル名を入力してください。');
      createExcelFileWithShape();
      return;
    }

    console.log('利用可能な図形の種類:');
    console.log('  1. 長方形');
    console.log('  2. 楕円');
    console.log('  3. 直線');

    rl.question('挿入する図形の番号を入力してください: ', (shapeTypeAnswer) => {
      const shapeTypeNum = parseInt(shapeTypeAnswer.trim());

      if (isNaN(shapeTypeNum) || shapeTypeNum < 1 || shapeTypeNum > 3) {
        console.log('無効な選択です。リストから番号を入力してください。');
        createExcelFileWithShape();
        return;
      }

      rl.question('  行（1から始まる）: ', (rowAnswer) => {
        const row = parseInt(rowAnswer.trim());

        if (isNaN(row) || row < 1) {
          console.log('無効な行番号です。');
          createExcelFileWithShape();
          return;
        }

        rl.question('  列（1から始まる）: ', (colAnswer) => {
          const col = parseInt(colAnswer.trim());

          if (isNaN(col) || col < 1) {
            console.log('無効な列番号です。');
            createExcelFileWithShape();
            return;
          }

          rl.question('  高さ: ', (heightAnswer) => {
            const height = parseInt(heightAnswer.trim());

            if (isNaN(height) || height <= 0) {
              console.log('無効な高さです。');
              createExcelFileWithShape();
              return;
            }

            rl.question('  幅: ', (widthAnswer) => {
              const width = parseInt(widthAnswer.trim());

              if (isNaN(width) || width <= 0) {
                console.log('無効な幅です。');
                createExcelFileWithShape();
                return;
              }

              converter.createExcelFileWithShape(filename.trim(), shapeTypeNum, row, col, height, width)
                .then(() => {
                  showMenu();
                })
                .catch(error => {
                  console.error('図形を含むExcelファイルの作成中にエラーが発生しました:', error);
                  showMenu();
                });
            });
          });
        });
      });
    });
  });
}

// Start the application
console.log('Excel to PDF Converter (Node.js/TypeScript版)');
console.log('Aspose.Cells for Node.js via Java を使用');

// Initialize the converter and show the menu
converter.initialize()
  .then(() => {
    showMenu();
  })
  .catch(error => {
    console.error('初期化中にエラーが発生しました:', error);
    process.exit(1);
  });

// Handle application exit
rl.on('close', () => {
  console.log('プログラムを終了します。');
  process.exit(0);
});
