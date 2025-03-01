"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const path = __importStar(require("path"));
const fs = __importStar(require("fs-extra"));
const readline = __importStar(require("readline"));
const converter_1 = require("./converter");
// Create input/output directories if they don't exist
fs.ensureDirSync(path.join(__dirname, '..', 'input'));
fs.ensureDirSync(path.join(__dirname, '..', 'output'));
// Create readline interface for user input
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});
// Initialize the Excel converter
const converter = new converter_1.ExcelConverter();
/**
 * Display the main menu and handle user input
 */
function showMenu() {
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
function convertExcelToPdf() {
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
            convertExcelToPdf();
            return;
        }
        const selectedFile = excelFiles[fileIndex];
        // Log file information for debugging
        console.log(`選択されたファイル: ${selectedFile}`);
        console.log(`ファイルの存在確認: ${fs.existsSync(selectedFile) ? '存在します' : '存在しません'}`);
        // Convert the selected file to PDF
        converter.convertToPdf(selectedFile)
            .then(() => {
            showMenu();
        })
            .catch(error => {
            console.error('変換中にエラーが発生しました:', error);
            showMenu();
        })
            .finally(() => {
            // Run the rename script after conversion
            const renameScriptPath = path.join(__dirname, 'rename_pdfs.js');
            if (fs.existsSync(renameScriptPath)) {
                console.log('PDF変換後にファイル名を修正します...');
                const { exec } = require('child_process');
                exec(`node ${renameScriptPath}`, (error, stdout, stderr) => {
                    if (error) {
                        console.error(`ファイル名修正スクリプトの実行エラー: ${error}`);
                        return;
                    }
                    console.log(stdout);
                    if (stderr) {
                        console.error(`ファイル名修正スクリプトのエラー出力: ${stderr}`);
                    }
                });
            }
            else {
                console.warn('ファイル名修正スクリプトが見つかりません:', renameScriptPath);
            }
        });
    });
}
/**
 * Create a new Excel file
 */
function createNewExcelFile() {
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
function createExcelFileWithShape() {
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
