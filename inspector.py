import jpype
from aspose import cells
from aspose.pydrawing import Color

jpype.startJVM()

workbook = cells.Workbook()
worksheet = workbook.worksheets[0]
style = cells.CellsFactory().create_style()
save_options = cells.PdfSaveOptions()

print("Worksheet methods and properties:", dir(worksheet))
print("Style methods and properties:", dir(style))
print("Font methods and properties:", dir(style.font))
print("Border methods and properties:", dir(style.borders))
print("SaveOptions methods and properties:", dir(save_options))
print("Cells methods and properties", dir(worksheet.cells))
try:
    print("PdfCompliance properties:", dir(cells.PdfCompliance))
except:
    print("PdfCompliance not found")
try:
    print("OptimizationType properties:", dir(cells.OptimizationType))
except:
    print("OptimizationType not found")

jpype.shutdownJVM()
