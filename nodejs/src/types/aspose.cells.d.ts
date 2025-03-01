declare module 'aspose.cells' {
  interface Workbook {
    save(filename: string, saveOptions?: any): void;
    getWorksheets(): Worksheets;
    calculateFormula(): void;
  }

  interface Worksheets {
    get(index: number): Worksheet;
    add(): number;
  }

  interface Worksheet {
    getCells(): Cells;
    getPageSetup(): PageSetup;
    getHorizontalPageBreaks(): PageBreaks;
    getVerticalPageBreaks(): PageBreaks;
    getShapes(): Shapes;
    getPictures(): Pictures;
  }

  interface PageSetup {
    getPaperSize(): number;
    getOrientation(): number;
    getFitToPagesWide(): number;
    getFitToPagesTall(): number;
  }

  interface PageBreaks {
    count(): number;
    get(index: number): PageBreak;
  }

  interface PageBreak {
    getRow(): number;
    getColumn(): number;
  }

  interface Cells {
    get(cellAddress: string | number, column?: number): Cell;
  }

  interface Cell {
    putValue(value: string | number): void;
    getValue(): any;
    setStyle(style: any): void;
  }

  interface Shapes {
    addShape(shapeType: number, row: number, rowOffset: number, column: number, columnOffset: number, height: number, width: number): void;
  }

  interface Pictures {
    add(row: number, column: number, imagePath: string): void;
  }

  interface PdfSaveOptions {
    setOnePagePerSheet(value: boolean): void;
  }

  const FileFormatType: {
    XLSX: number;
    XLS: number;
  };

  const MsoDrawingType: {
    RECTANGLE: number;
    OVAL: number;
    LINE: number;
  };

  const SaveFormat: {
    PDF: number;
    EXCEL_97_TO_2003: number;
  };

  const Color: {
    getBlue(): any;
  };

  class Workbook {
    constructor(filePathOrFormatType?: string | number);
    createStyle(): any;
  }

  class PdfSaveOptions {
    constructor();
    setOnePagePerSheet(value: boolean): void;
  }
}

declare namespace global {
  interface Global {
    aspose: {
      cells: typeof import('aspose.cells');
    };
  }
}
