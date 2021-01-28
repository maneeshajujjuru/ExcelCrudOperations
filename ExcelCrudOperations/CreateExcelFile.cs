﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace ExcelCrudOperations
{
    public class CreateExcelFile
    {
        public static void WriteExcelFile()
        {
            List<UserDetails> persons = new List<UserDetails>()
           {
               new UserDetails() {ID="1001", Name="ABCD", City ="City1", Country="USA"},
               new UserDetails() {ID="1002", Name="PQRS", City ="City2", Country="INDIA"},
               new UserDetails() {ID="1003", Name="XYZZ", City ="City3", Country="CHINA"},
               new UserDetails() {ID="1004", Name="LMNO", City ="City4", Country="UK"},
          };

            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting. 
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(@"c:\ExcelSample\RK_Excel1.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }
    }
}
