Aspose.Cells for Node.js via Java is a scalable and feature-rich API to create, process, manipulate & convert Excel & OpenOffice spreadsheets using Node.js. API offers Excel file generation, conversion, worksheets styling, Pivot Table & chart management & rendering, reliable formula calculation engine and much more - all without any dependency on Office Automation or Microsoft Excel®.

# Note
We have released **[Aspose.Cells for Node.js via C++](https://www.npmjs.com/package/aspose.cells.node)** since v24.7.0. It is highly recommended to use it to replace Aspose.Cells for Node.js via Java. We will gradually reduce the release frequency of Aspose.Cells for Node.js via Java and focus on **[Aspose.Cells for Node.js via C++](https://www.npmjs.com/package/aspose.cells.node)** in the future.

## Node.js Spreadsheet API Features 
- Generate Excel files via API or using templates.
- Create Pivot Tables, charts, sparklines & conditional formatting rules on-the-fly.
- Refresh existing charts & convert charts to images or PDF.
- Create & manipulate comments & hyperlinks.
- Set complex formulas & calculate results via API.
- Set protection on workbooks, worksheets, cells, columns or rows.
- Create & manipulate named ranges.
- Populate worksheets through Smart Markers.
- Manipulate & refresh Pivot Tables via API.
- Convert worksheets to PDF, XPS & SVG formats.
- Inter-convert files to popular Excel formats.

## Read & Write Excel Files
**Microsoft Excel:** XLS, XLSX, XLSB, XLTX, XLTM, XLSM, XML
**OpenOffice:** ODS
**Text:** CSV, Tab-Delimited, TXT, JSON
**Web:** HTML, MHTML

## Save Excel Files As 
**Fixed Layout:** PDF, XPS
**Images:** JPEG, PNG, BMP, SVG, TIFF, GIF, EMF
**Text:** CSV, Tab-Delimited, JSON, SQL, XML

## Getting Started with Aspose.Cells for Nodejs via Java
### Create Excel XLSX File from Scratch using Node.js
``` js
var aspose = aspose || {};
aspose.cells = require("aspose.cells");

var workbook = new aspose.cells.Workbook(aspose.cells.FileFormatType.XLSX);
workbook.getWorksheets().get(0).getCells().get("A1").putValue("testing...");
workbook.save("output.xlsx");
```

### Convert Excel XLSX File to PDF using Node.js
``` js
var aspose = aspose || {};
aspose.cells = require("aspose.cells");

var workbook = new aspose.cells.Workbook("example.xlsx");
var saveOptions = aspose.cells.PdfSaveOptions();
saveOptions.setOnePagePerSheet(true);
workbook.save("example.pdf", saveOptions);
```

### Format Excel Cells via Node.js
```js
var aspose = aspose || {};
aspose.cells = require("aspose.cells");

var excel = new aspose.cells.Workbook();
var style = excel.createStyle();
style.getFont().setName("Times New Roman");
style.getFont().setColor(aspose.cells.Color.getBlue());
for (var i = 0; i < 100; i++)
{
    excel.getWorksheets().get(0).getCells().get(0, i).setStyle(style);
}
```

### Add Picture to Excel Worksheet with Node.js
```js
var aspose = aspose || {};
aspose.cells = require("aspose.cells");

var workbook = new aspose.cells.Workbook();
var sheetIndex = workbook.getWorksheets().add();
var worksheet = workbook.getWorksheets().get(sheetIndex);

// adding a picture at "F6" cell
worksheet.getPictures().add(5, 5, "image.gif");

workbook.save("output.xls", aspose.cells.SaveFormat.EXCEL_97_TO_2003);
```

[Product Page](https://products.aspose.com/cells/nodejs-java) | [Product Documentation](https://docs.aspose.com/display/cellsnodejsjava/Aspose.Cells+for+Node.js+via+Java+Home) | [Blog](https://blog.aspose.com/category/cells/) |[API Reference](https://apireference.aspose.com/cells/nodejs) | [Source Code Samples](https://github.com/aspose-cells/Aspose.Cells-for-Java) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license)
