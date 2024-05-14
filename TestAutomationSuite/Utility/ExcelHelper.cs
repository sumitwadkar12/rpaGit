using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using TestAutomationSuite.Utility;

namespace AutomationUAT.Utility
{
    public class ExcelHelper
    {
        public List<string> IsHeaderPresent(string dataFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(dataFilePath)))
                {
                    // Assuming the header is in the first row of the first worksheet
                    var worksheet = package.Workbook.Worksheets[0];
                    var headerCell = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns];
                    List<string> columns = new List<string>();
                    foreach (var cell in headerCell)
                    {
                        columns.Add(cell.Text);
                    }
                    return columns;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }

        private string GetCellValue(WorkbookPart wbPart, List<Cell> theCells, string cellColumnReference)
        {
            Cell theCell = null;
            string value = "";
            foreach (Cell cell in theCells)
            {
                if (cell.CellReference.Value.StartsWith(cellColumnReference))
                {
                    theCell = cell;
                    break;
                }
            }
            if (theCell != null)
            {
                value = theCell.InnerText;
                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code looks up the corresponding value in the shared string table. For Booleans, the code converts the value into the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared strings table.
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something is wrong. Return the index that is in the cell. Otherwise, look up the correct text in the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }

        private string GetCellValue(WorkbookPart wbPart, List<Cell> theCells, int index)
        {
            return GetCellValue(wbPart, theCells, GetExcelColumnName(index));
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }

        public string DataTableToJSONWithJSONNet(DataTable table)
        {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }

        public bool IsExcelFileCorrupted(string filePath)
        {
            try
            {
                FileInfo file = new FileInfo(filePath);
                using (FileStream fileStream = file.OpenRead())
                {
                    using (var package = new ExcelPackage(file))
                    {
                        // Attempt to load the Excel package
                        package.Load(fileStream);

                        // If loading succeeds, the file is not corrupted
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                return true; // Consider it corrupted on any error
            }
        }
        //Only xlsx files
        public DataTable GetDataTableFromExcelFile(string filePath, string sheetName = "")
        {
            LoggingHelper LoggingHelper = new LoggingHelper();
            DataTable dt = new DataTable();
            string campFile = filePath;
            //For comparing strings
            string[] lastFolderName = filePath.Split('\\');
            string dir = lastFolderName[lastFolderName.Length - 3];
            string[] campLastfoldername = campFile.Split('\\');
            string campdir = campLastfoldername[campLastfoldername.Length - 3];
            try
            {

                if (IsExcelFileCorrupted(filePath))
                {
                    // Delete the corrupted file
                    try
                    {
                        if(dir == campdir && lastFolderName[lastFolderName.Length - 3].ToLower().Trim().Contains("data"))
                        {
                            LoggingHelper.insertLog(" The " + System.IO.Path.GetFileName(filePath) + " file that you you have entered in " + lastFolderName[lastFolderName.Length - 2] + " folder is corrupt, please enter a Valid .xlsx file \n ");
                            File.Delete(filePath);
                        }
                        else if(dir == campdir && lastFolderName[lastFolderName.Length - 2].ToLower().Trim().Contains("action"))
                        {
                            LoggingHelper.insertLog(" The " + System.IO.Path.GetFileName(filePath) + " file that you you have entered in " + lastFolderName[lastFolderName.Length - 2] + " folder is corrupt, please enter a Valid .xlsx file \n ");
                            LoggingHelper.insertLog("***BOT IS DEACTIVATED***");
                            Environment.Exit(1);
                        }                                                
                    }
                    catch (IOException ex)
                    {
                        LoggingHelper.insertLog("Error deleting the corrupted file: {ex.Message}");
                    }
                }
                else
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
                    {                      
                        WorkbookPart wbPart = document.WorkbookPart;
                        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                        string sheetId = sheetName != "" ? sheets.Where(q => q.Name == sheetName).First().Id.Value : sheets.First().Id.Value;
                        WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
                        SheetData sheetdata = wsPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                        int totalHeaderCount = sheetdata.Descendants<Row>().ElementAt(0).Descendants<Cell>().Count();
                        //Get the header                    
                        for (int i = 1; i <= totalHeaderCount; i++)
                        {
                            dt.Columns.Add(GetCellValue(wbPart, sheetdata.Descendants<Row>().ElementAt(0).Elements<Cell>().ToList(), i));
                        }
                        foreach (Row r in sheetdata.Descendants<Row>())
                        {
                            if (r.RowIndex > 1)
                            {
                                DataRow tempRow = dt.NewRow();

                                //Always get from the header count, because the index of the row changes where empty cell is not counted
                                for (int i = 1; i <= totalHeaderCount; i++)
                                {
                                    tempRow[i - 1] = GetCellValue(wbPart, r.Elements<Cell>().ToList(), i);
                                }
                                dt.Rows.Add(tempRow);
                            }
                        }
                        //document.WorkbookPart.Workbook.Save();
                        document.Close();
                        document.Dispose();
                    }
                }

            }
            catch (Exception ex)
            {
                //If the sheet name of Input data file is invalid 
                if (ex.Message == "Sequence contains no elements" && dir == campdir && lastFolderName[lastFolderName.Length - 3].ToLower().Trim().Contains("data"))
                {                    
                    LoggingHelper.insertLog(" Change the sheet name of the provided input excel file " + System.IO.Path.GetFileName(filePath) + " present in " + lastFolderName[lastFolderName.Length - 2] + " folder to TestData to start its processing");
                }
                //If the sheet name of Action file is invalid                 
                else if (ex.Message == "Specified argument was out of the range of valid values. (Parameter 'index')" && dir == campdir && lastFolderName[lastFolderName.Length - 3].ToLower().Trim().Contains("data"))
                {
                    LoggingHelper.insertLog(" The provided input excel file " + System.IO.Path.GetFileName(filePath) + " present in " + lastFolderName[lastFolderName.Length - 2] + " folder is blank and contains no data, please enter a valid Input File");
                }
                //If the provided Data file is blank and contains no data
                else if (ex.Message == "Sequence contains no elements" && dir == campdir && lastFolderName[lastFolderName.Length - 2].ToLower().Trim().Contains("action"))
                {
                    LoggingHelper.insertLog(" Change the sheet name of the provided input excel file " + System.IO.Path.GetFileName(filePath) + " present in " + lastFolderName[lastFolderName.Length - 2] + " folder to Actions to start its processing");
                    LoggingHelper.insertLog("***BOT IS DEACTIVATED***");
                    Environment.Exit(1);
                }
                //If the provided Action excel file is blank and contains no data
                else if (ex.Message == "Specified argument was out of the range of valid values. (Parameter 'index')" && dir == campdir && lastFolderName[lastFolderName.Length - 2].ToLower().Trim().Contains("action"))
                {
                    LoggingHelper.insertLog(" The provided input excel file " + System.IO.Path.GetFileName(filePath) + " present in " + lastFolderName[lastFolderName.Length - 2] + " folder is blank and contains no data, please enter a valid Input File");
                    LoggingHelper.insertLog("***BOT IS DEACTIVATED***");
                    Environment.Exit(1);
                }
                else
                {                   
                    LoggingHelper.insertLog(ex.ToString());
                }
            }
            return dt;
        }

        public void UpdateCellold(string docName, string sheetName, string text, uint rowIndex, string columnName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                Row row = null;
                WorkbookPart wbPart = document.WorkbookPart;
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string sheetId = sheetName != "" ? sheets.Where(q => q.Name == sheetName).First().Id.Value : sheets.First().Id.Value;
                WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
                SheetData sheetdata = wsPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                int totalHeaderCount = sheetdata.Descendants<Row>().ElementAt(0).Descendants<Cell>().Count();

                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                if (sheetdata != null)
                {
                    // Insert the text into the SharedStringTablePart.
                    int index = InsertSharedStringItem(text, shareStringPart);

                    // Insert cell A1 into the new worksheet.
                    Cell cell = InsertCellInWorksheet(columnName, rowIndex, wsPart);

                    // Set the value of cell A1.
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    // Save the new worksheet.
                    wsPart.Worksheet.Save();
                }
                document.WorkbookPart.Workbook.Save();
                document.Close();
            }

        }

        public void UpdateCellsold(string docName, string sheetName, string text, uint rowIndex, string columnName, string textColor = "000000")
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                Row row = null;
                WorkbookPart wbPart = document.WorkbookPart;
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string sheetId = sheetName != "" ? sheets.Where(q => q.Name == sheetName).First().Id.Value : sheets.First().Id.Value;
                WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
                SheetData sheetdata = wsPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                int totalHeaderCount = sheetdata.Descendants<Row>().ElementAt(0).Descendants<Cell>().Count();

                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                if (sheetdata != null)
                {
                    // Insert the text into the SharedStringTablePart.
                    int index = InsertSharedStringItem(text, shareStringPart);

                    // Insert cell A1 into the new worksheet.
                    Cell cell = InsertCellInWorksheet(columnName, rowIndex, wsPart);

                    // Set the value of cell A1.
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    // Set text color
                    Font font = new Font();
                    Color color = new Color() { Rgb = new HexBinaryValue() { Value = textColor } };
                    font.Color = color; // Set the color directly
                    cell.StyleIndex = InsertCellFormatInStylesheet(wbPart, font);

                    // Save the new worksheet.
                    wsPart.Worksheet.Save();
                }
                document.WorkbookPart.Workbook.Save();
                document.Close();
            }
        }

        private static uint InsertCellFormatInStylesheet(WorkbookPart workbookPart, Font font)
        {
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
            CellFormat cellFormat = new CellFormat();
            cellFormat.FontId = InsertFontInStylesheet(workbookPart, font);
            stylesheet.CellFormats.AppendChild(cellFormat);
            return (uint)stylesheet.CellFormats.Count - 1;
        }

        private static uint InsertFontInStylesheet(WorkbookPart workbookPart, Font font)
        {
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
            stylesheet.Fonts.AppendChild(font);
            return (uint)stylesheet.Fonts.Count - 1;
        }



        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    //return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }

        public int GetColumnCount(string filePath, string sheetName)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    int colCount = worksheet.Dimension.End.Column;
                    return colCount;
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions accordingly (e.g., file not found, invalid sheet name, etc.)
                Console.WriteLine($"Error: {ex.Message}");
                return -1; // Return -1 to indicate an error
            }
        }

        public bool IsHeaderPresent(string dataFilePath, string headerName)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(dataFilePath)))
                {
                    // Assuming the header is in the first row of the first worksheet
                    var worksheet = package.Workbook.Worksheets[0];
                    var headerCell = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns];

                    foreach (var cell in headerCell)
                    {
                        if (cell.Text == headerName)
                        {
                            return true;
                        }
                    }

                    return false;
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions accordingly (e.g., file not found, invalid sheet name, etc.)
                Console.WriteLine($"Error: {ex.Message}");
                return false; // Return -1 to indicate an error
            }            
        }

        public string getActionFileFromPath(string FilePath, out string exception)
        {
            string filesResult = "";
            bool status = true;
            exception = "";
            try
            {
                string folderPath = FilePath; // Replace with your folder path
                // Get first one file
                filesResult = Directory.GetFiles(folderPath, "*.xlsx").FirstOrDefault();
                if (!string.IsNullOrEmpty(filesResult))
                {
                    filesResult = System.IO.Path.GetFileName(filesResult);
                }
                else
                {
                    filesResult = "";
                    status = false;
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();

            }
            return filesResult;
        }

        public List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        public static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }
    }
}
