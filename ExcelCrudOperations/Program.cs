using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace ExcelCrudOperations
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadExcelFile.ReadExcel();
            //CreateExcelFile.WriteExcelFile();
            //Delete.UpdateExcelUsingOpenXMLSDK();
            //string docName = @"c:\ExcelSample\RK_Excel1.xlsx";
            //string sheetName = "Sheet1";
            //string colName = "D";
            //uint rowIndex = 4;
            //DeleteTextFromCell(docName, sheetName, colName, rowIndex);
        
                //FileInfo fileInfo = new FileInfo(@"c:\ExcelSample\RK_Excel1.xlsx");
                //ExcelPackage p = new ExcelPackage(fileInfo);
                //ExcelWorksheet myWorksheet = p.Workbook.Worksheets["Sheet1"];
                //myWorksheet.Cells[4, 2].Value = 10;
                //p.Save();

        }
        // Given a document, a worksheet name, a column name, and a one-based row index,
        // deletes the text from the cell at the specified column and row on the specified worksheet.
        public static void DeleteTextFromCell(string docName, string sheetName, string colName, uint rowIndex)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist.
                    return;
                }
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

                // Get the cell at the specified column and row.
                Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
                if (cell == null)
                {
                    // The specified cell does not exist.
                    return;
                }

                cell.Remove();
                worksheetPart.Worksheet.Save();
            }
        }

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        private static Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r?.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return null;
            }

            IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return null;
            }

            return cells.First();
        }

        // Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
        // reference the specified SharedStringItem and removes the item.
        private static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
        {
            bool remove = true;

            foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                Worksheet worksheet = part.Worksheet;
                foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                {
                    // Verify if other cells in the document reference the item.
                    if (cell.DataType != null &&
                        cell.DataType.Value == CellValues.SharedString &&
                        cell.CellValue.Text == shareStringId.ToString())
                    {
                        // Other cells in the document still reference the item. Do not remove the item.
                        remove = false;
                        break;
                    }
                }

                if (!remove)
                {
                    break;
                }
            }

            // Other cells in the document do not reference the item. Remove the item.
            if (remove)
            {
                SharedStringTablePart shareStringTablePart = document.WorkbookPart.SharedStringTablePart;
                if (shareStringTablePart == null)
                {
                    return;
                }

                SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
                if (item != null)
                {
                    item.Remove();

                    // Refresh all the shared string references.
                    foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
                    {
                        Worksheet worksheet = part.Worksheet;
                        foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                        {
                            if (cell.DataType != null &&
                                cell.DataType.Value == CellValues.SharedString)
                            {
                                int itemIndex = int.Parse(cell.CellValue.Text);
                                if (itemIndex > shareStringId)
                                {
                                    cell.CellValue.Text = (itemIndex - 1).ToString();
                                }
                            }
                        }
                        worksheet.Save();
                    }

                    document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
                }
            }
        }
        //static void ReadExcelFile()
        //{
        //    try
        //    {
        //        //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
        //        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(@"c:\ExcelSample\RK_Excel1.xlsx", false))
        //        {
        //            //create the object for workbook part  
        //            WorkbookPart workbookPart = doc.WorkbookPart;
        //            Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
        //            StringBuilder excelResult = new StringBuilder();

        //            //using for each loop to get the sheet from the sheetcollection  
        //            foreach (Sheet thesheet in thesheetcollection)
        //            {
        //                excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
        //                excelResult.AppendLine("----------------------------------------------- ");
        //                //statement to get the worksheet object by using the sheet id  
        //                Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

        //                SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
        //                foreach (Row thecurrentrow in thesheetdata)
        //                {
        //                    foreach (Cell thecurrentcell in thecurrentrow)
        //                    {
        //                        //statement to take the integer value  
        //                        string currentcellvalue = string.Empty;
        //                        if (thecurrentcell.DataType != null)
        //                        {
        //                            if (thecurrentcell.DataType == CellValues.SharedString)
        //                            {
        //                                int id;
        //                                if (Int32.TryParse(thecurrentcell.InnerText, out id))
        //                                {
        //                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        //                                    if (item.Text != null)
        //                                    {
        //                                        //code to take the string value  
        //                                        excelResult.Append(item.Text.Text + " ");
        //                                    }
        //                                    else if (item.InnerText != null)
        //                                    {
        //                                        currentcellvalue = item.InnerText;
        //                                    }
        //                                    else if (item.InnerXml != null)
        //                                    {
        //                                        currentcellvalue = item.InnerXml;
        //                                    }
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
        //                        }
        //                    }
        //                    excelResult.AppendLine();
        //                }
        //                excelResult.Append("");
        //                Console.WriteLine(excelResult.ToString());
        //                Console.ReadLine();
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {

        //    }
        //}

        //Instantiate a Workbook object


    }

}
