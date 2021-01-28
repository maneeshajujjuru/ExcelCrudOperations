using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelCrudOperations
{
    public class Delete
    {
            public static void UpdateExcelUsingOpenXMLSDK(string fileName)
            {
                // Open the document for editing.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
                {
                    // Access the main Workbook part, which contains all references.
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                    // get sheet by name
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Sheet1").FirstOrDefault();

                    // get worksheetpart by sheet id
                    WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    // The SheetData object will contain all the data.
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    Cell cell = GetCell(worksheetPart.Worksheet, "B", 4);

                    cell.CellValue = new CellValue("10");
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                    // Save the worksheet.
                    worksheetPart.Worksheet.Save();

                    // for recacluation of formula
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

                }
            }

            private static Cell GetCell(Worksheet worksheet,
                string columnName, uint rowIndex)
            {
                Row row = GetRow(worksheet, rowIndex);

                if (row == null) return null;

                var FirstRow = row.Elements<Cell>().Where(c => string.Compare
                                                               (c.CellReference.Value, columnName +
                                                                                       rowIndex, true) == 0).FirstOrDefault();

                if (FirstRow == null) return null;

                return FirstRow;
            }

            private static Row GetRow(Worksheet worksheet, uint rowIndex)
            {
                Row row = worksheet.GetFirstChild<SheetData>().
                    Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                if (row == null)
                {
                    throw new ArgumentException(String.Format("No row with index {0} found in spreadsheet", rowIndex));
                }
                return row;
            }
    }
}
