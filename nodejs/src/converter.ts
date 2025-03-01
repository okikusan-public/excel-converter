import * as path from 'path';
import * as fs from 'fs-extra';
import * as crypto from 'crypto';

/**
 * ExcelConverter class for handling Excel to PDF conversion using Aspose.Cells
 */
export class ExcelConverter {
  private aspose: any;
  private initialized: boolean = false;

  /**
   * Initialize the ExcelConverter by setting up Aspose.Cells
   */
  public async initialize(): Promise<void> {
    try {
      // Initialize aspose.cells using the correct pattern
      const globalAny = global as any;
      const aspose = globalAny.aspose || {};
      aspose.cells = require("aspose.cells");
      this.aspose = aspose;
      
      this.initialized = true;
      console.log('Aspose.Cells for Node.js via Java が初期化されました。');
    } catch (error: any) {
      console.error('初期化中にエラーが発生しました:', error.message);
      throw error;
    }
  }

  /**
   * Convert Excel file to PDF
   * @param excelFilePath Path to the Excel file
   * @returns Promise that resolves when conversion is complete
   */
  public async convertToPdf(excelFilePath: string): Promise<void> {
    if (!this.initialized) {
      await this.initialize();
    }

    return new Promise<void>((resolve, reject) => {
      try {
        console.log(`Excelファイル '${path.basename(excelFilePath)}' をPDFに変換しています...`);

        // Record start time for performance measurement
        const startTime = Date.now();

        // Load Excel file
        const loadStartTime = Date.now();
        console.log(`Excelファイルのパス: ${excelFilePath}`);
        console.log(`ファイルの存在確認 (Node.js): ${fs.existsSync(excelFilePath) ? '存在します' : '存在しません'}`);

        // Try to load the file using a different approach for Japanese filenames
        // Read the file into a buffer
        const fileData = fs.readFileSync(excelFilePath);
        console.log(`ファイルサイズ: ${fileData.length} バイト`);

        // Create a temporary file with a simple name
        const tempDir = path.join(__dirname, '..', 'temp');
        fs.ensureDirSync(tempDir);
        const tempFilePath = path.join(tempDir, 'temp_file' + path.extname(excelFilePath));
        fs.writeFileSync(tempFilePath, fileData);
        console.log(`一時ファイルを作成しました: ${tempFilePath}`);

        // Load the workbook from the temporary file
        const workbook = new this.aspose.cells.Workbook(tempFilePath);
        const loadEndTime = Date.now();
        const loadDuration = (loadEndTime - loadStartTime) / 1000;
        console.log(`Excelファイルの読み込み完了: ${loadDuration.toFixed(2)}秒`);

        // Get the first worksheet
        const worksheet = workbook.getWorksheets().get(0);

        // Display existing print settings
        console.log(`印刷設定 - 用紙サイズ: ${worksheet.getPageSetup().getPaperSize()}`);
        console.log(`印刷設定 - 向き: ${worksheet.getPageSetup().getOrientation() === 1 ? '横' : '縦'}`);
        console.log(`印刷設定 - フィット設定: 幅=${worksheet.getPageSetup().getFitToPagesWide()}, 高さ=${worksheet.getPageSetup().getFitToPagesTall()}`);

        // Display page breaks information
        this.displayPageBreaks(worksheet);

        // Set output PDF path
        const outputDir = path.join(__dirname, '..', 'output');
        fs.ensureDirSync(outputDir);

        // Get the original filename
        const basename = path.basename(excelFilePath);
        const extname = path.extname(excelFilePath);
        const nameWithoutExt = basename.substring(0, basename.length - extname.length);

        // Create a timestamp for uniqueness
        const now = new Date();
        const timestamp = now.toISOString().replace(/[-:]/g, '').replace('T', '_').split('.')[0];

        // Generate a hash of the original filename to preserve it
        const hash = crypto.createHash('md5').update(nameWithoutExt).digest('hex').substring(0, 8);

        // Create a filename that includes the original name (for readability in logs)
        // but uses a hash and timestamp for the actual file to avoid encoding issues
        const outputFilename = `${timestamp}_${hash}.pdf`;

        // Log the filename information
        console.log(`元のファイル名: ${nameWithoutExt}`);
        console.log(`ハッシュ: ${hash}`);
        console.log(`出力PDFファイル名: ${outputFilename}`);

        // Define the mapping interface
        interface FileMapping {
          originalName: string;
          timestamp: string;
          excelPath: string;
        }

        interface FileMappings {
          [key: string]: FileMapping;
        }

        // Create a mapping file to track original filenames to generated filenames
        const mappingFilePath = path.join(outputDir, 'filename_mapping.json');
        let mappings: FileMappings = {};

        // Load existing mappings if available
        if (fs.existsSync(mappingFilePath)) {
          try {
            mappings = JSON.parse(fs.readFileSync(mappingFilePath, 'utf8')) as FileMappings;
          } catch (e) {
            console.error('マッピングファイルの読み込みに失敗しました:', e);
          }
        }

        // Add the new mapping
        mappings[outputFilename] = {
          originalName: nameWithoutExt,
          timestamp: now.toISOString(),
          excelPath: excelFilePath
        };

        // Save the updated mappings
        fs.writeFileSync(mappingFilePath, JSON.stringify(mappings, null, 2), 'utf8');

        const outputPdf = path.join(outputDir, outputFilename);

        // Record PDF conversion start time
        const pdfStartTime = Date.now();

        try {
          console.log('\nPDFに変換中...');

          // Create PDF save options
          const saveOptions = new this.aspose.cells.PdfSaveOptions();
          // Ensure that all columns are NOT forced onto one page; respect existing breaks.
          saveOptions.setAllColumnsInOnePagePerSheet(false);
          // Ensure PDF compliance doesn't interfere with page breaks
          saveOptions.setCompliance(this.aspose.cells.PdfCompliance.NONE);

          // Save as PDF
          const saveStartTime = Date.now();
          workbook.save(outputPdf, saveOptions);
          const saveEndTime = Date.now();
          const saveDuration = (saveEndTime - saveStartTime) / 1000;
          const pdfEndTime = Date.now();
          const pdfDuration = (pdfEndTime - pdfStartTime) / 1000;
          const totalDuration = (pdfEndTime - startTime) / 1000;

          console.log(`PDFファイルが作成されました: ${outputPdf} (保存時間: ${saveDuration.toFixed(2)}秒, PDF変換合計: ${pdfDuration.toFixed(2)}秒, 総処理時間: ${totalDuration.toFixed(2)}秒)`);
          resolve(); // Resolve the Promise after successful save
        } catch (error: any) {
          console.error(`標準的な方法でのPDF変換に失敗しました: ${error.message}`);

          // Alternative method
          try {
            console.log('\n代替方法でPDFに変換中...');

            // Save without options
            const saveStartTime = Date.now();
            workbook.save(outputPdf);
            const saveEndTime = Date.now();
            const saveDuration = (saveEndTime - saveStartTime) / 1000;
            const pdfEndTime = Date.now();
            const pdfDuration = (pdfEndTime - pdfStartTime) / 1000;
            const totalDuration = (pdfEndTime - startTime) / 1000;

            console.log(`PDFファイルが作成されました: ${outputPdf} (代替方法での保存時間: ${saveDuration.toFixed(2)}秒, PDF変換合計: ${pdfDuration.toFixed(2)}秒, 総処理時間: ${totalDuration.toFixed(2)}秒)`);
            resolve(); // Resolve the Promise after successful save (alternative method)
          } catch (altError: any) {
            console.error(`代替方法でのPDF変換に失敗しました: ${altError.message}`);
            reject(altError); // Reject the Promise if both methods fail
          }
        }
      } catch (error: any) {
        console.error(`Excel→PDF変換中にエラーが発生しました: ${error.message}`);
        console.error(error.stack);

        // Display processing time even in case of error
        const endTime = Date.now();
        const errorDuration = (endTime - Date.now()) / 1000;
        console.log(`エラーが発生しました。エラーまでの処理時間: ${errorDuration.toFixed(2)}秒`);

        reject(error); // Reject the Promise in case of an error
      }
    });
  }

    /**
     * Create a new Excel file with sample data
     * @param filename Name of the file to create (without extension)
     * @returns Promise that resolves when file creation is complete
     */
    public async createExcelFile(filename: string, shapeType?: number, row?: number, column?: number, height?: number, width?: number): Promise<void> {
      if (!this.initialized) {
        await this.initialize();
      }

      try {
        // Create a new workbook with XLSX format
        const workbook = new this.aspose.cells.Workbook(this.aspose.cells.FileFormatType.XLSX);

        // Get the first worksheet
        const worksheet = workbook.getWorksheets().get(0);

        // Add some data
        worksheet.getCells().get("A1").putValue("Name");
        worksheet.getCells().get("B1").putValue("Age");
        worksheet.getCells().get("A2").putValue("John Doe");
        worksheet.getCells().get("B2").putValue(30);
        worksheet.getCells().get("A3").putValue("Jane Smith");
        worksheet.getCells().get("B3").putValue(25);

        // Add shape if provided
        if (shapeType !== undefined && row !== undefined && column !== undefined && height !== undefined && width !== undefined) {
          let msoDrawingType;

          switch (shapeType) {
            case 1:
              msoDrawingType = this.aspose.cells.MsoDrawingType.RECTANGLE;
              break;
            case 2:
              msoDrawingType = this.aspose.cells.MsoDrawingType.OVAL;
              break;
            case 3:
              msoDrawingType = this.aspose.cells.MsoDrawingType.LINE;
              break;
            default:
              throw new Error('無効な図形タイプです。');
          }

          try {
            worksheet.getShapes().addShape(
              msoDrawingType,
              row - 1,
              0,
              column - 1,
              0,
              height,
              width
            );
          } catch (shapeError) {
            console.log('図形の追加中にエラーが発生しました。この機能はサポートされていない可能性があります。');
          }
        }

        // Save the workbook
        const inputDir = path.join(__dirname, '..', 'input');
        fs.ensureDirSync(inputDir);
        const outputPath = path.join(inputDir, `${filename}.xlsx`);
        workbook.save(outputPath);

        console.log(`新しいExcelファイルが作成されました: ${outputPath}`);
      } catch (error: any) {
        console.error(`Excelファイルの作成中にエラーが発生しました: ${error.message}`);
        throw error;
      }
    }

    /**
     * Create a new Excel file with a shape
     * @param filename Name of the file to create (without extension)
     * @param shapeType Type of shape to add (1=Rectangle, 2=Oval, 3=Line)
     * @param row Row position (1-based)
     * @param column Column position (1-based)
     * @param height Height of the shape
     * @param width Width of the shape
     * @returns Promise that resolves when file creation is complete
     */
    public async createExcelFileWithShape(filename: string, shapeType: number, row: number, column: number, height: number, width: number): Promise<void> {
      return this.createExcelFile(filename, shapeType, row, column, height, width);
    }

  /**
   * Display page breaks information for a worksheet
   * @param worksheet Worksheet to display page breaks for
   */
  private displayPageBreaks(worksheet: any): void {
    console.log('\n既存の改ページ情報:');
    
    try {
      // Display horizontal page breaks
      const horizontalPageBreaks = worksheet.getHorizontalPageBreaks();
      
      // Check if count is a function
      if (typeof horizontalPageBreaks.count === 'function') {
        const horizontalBreaksCount = horizontalPageBreaks.count();
        
        if (horizontalBreaksCount > 0) {
          console.log('水平改ページ:');
          for (let i = 0; i < horizontalBreaksCount; i++) {
            const pageBreak = horizontalPageBreaks.get(i);
            console.log(`  行 ${pageBreak.getRow() + 1} の後`);
          }
        } else {
          console.log('水平改ページはありません');
        }
      } else {
        console.log('水平改ページ情報を取得できません');
      }
      
      // Display vertical page breaks
      const verticalPageBreaks = worksheet.getVerticalPageBreaks();
      
      // Check if count is a function
      if (typeof verticalPageBreaks.count === 'function') {
        const verticalBreaksCount = verticalPageBreaks.count();
        
        if (verticalBreaksCount > 0) {
          console.log('垂直改ページ:');
          for (let i = 0; i < verticalBreaksCount; i++) {
            const pageBreak = verticalPageBreaks.get(i);
            console.log(`  列 ${pageBreak.getColumn() + 1} の後`);
          }
        } else {
          console.log('垂直改ページはありません');
        }
      } else {
        console.log('垂直改ページ情報を取得できません');
      }
    } catch (error: any) {
      console.error(`改ページ情報の取得中にエラーが発生しました: ${error.message}`);
      console.log('改ページ情報の取得に失敗しました。処理を続行します。');
    }
  }
}
