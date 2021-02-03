using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
//using OfficeOpenXml;


namespace ExcelCrudOperations
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateExcelFile.WriteExcelFile();
        }
    }
}
