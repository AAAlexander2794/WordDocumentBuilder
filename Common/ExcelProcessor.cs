using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace WordDocumentBuilder
{
    /// <summary>
    /// Класс работы с Экселем.
    /// </summary>
    /// <remarks>
    /// Не должен ничего знать про бизнес-логику.
    /// </remarks>
    public static class ExcelProcessor
    {
        /// <summary>
        /// Чтение листа Экселя
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="firstRowIsHeader"></param>
        /// <param name="sheetNumber">Номер листа</param>
        /// <returns></returns>
        public static DataTable ReadExcelSheet(string filename, bool firstRowIsHeader = true, int sheetNumber = 0)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false))
                {
                    //Read the first Sheets 
                    Sheet sheet = doc.WorkbookPart.Workbook.Sheets.ChildElements[sheetNumber] as Sheet;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                    int counter = 0;
                    foreach (Row row in rows)
                    {
                        counter = counter + 1;
                        // Счетчик столбцов
                        int columnNumber = 0;
                        //Read the first row as header
                        if (counter == 1)
                        {
                            var j = 1;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                                dt.Columns.Add(colunmName);
                                columnNumber++;
                            }
                        }
                        else
                        {
                            dt.Rows.Add();
                            // Почти полностью рабочий вариант (пропускает пустые ячейки)
                            int i = 0;
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                                i++;
                            }

                            //// Нерабочий вариант
                            //var cells = row.Descendants<Cell>().ToList();
                            //for (int j = 0; j <= columnNumber; j++)
                            //{
                            //    dt.Rows[dt.Rows.Count - 1][j] = GetCellValue(doc, cells[j]);
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"ReadExcelSheet\r\n{ex.Message}");
            }
            return dt;
        }

        public static DataTable ReadExcelSheetClosedXML(string filename, int sheetNumber = 1, bool firstRowIsHeader = true)
        {
            //
            var workbook = new XLWorkbook(filename);
            var ws = workbook.Worksheet(sheetNumber);
            return ReadExcelSheetClosedXML(ws, firstRowIsHeader);
        }

        public static DataTable ReadExcelSheetClosedXML(string filename, string sheetName, bool firstRowIsHeader = true)
        {
            //
            var workbook = new XLWorkbook(filename);
            var ws = workbook.Worksheet(sheetName);
            return ReadExcelSheetClosedXML(ws, firstRowIsHeader);
        }

        public static DataTable ReadExcelSheetClosedXML(IXLWorksheet ws, bool firstRowIsHeader = true)
        {
            DataTable dt = new DataTable();
            // Счетчик количества столбцов
            int count = 0;
            //
            if (firstRowIsHeader)
            {
                //
                var rowHead = ws.Row(1);
                var cells = rowHead.Cells();
                foreach (var cell in cells)
                {
                    if (cell.IsEmpty()) break;
                    // Создаем столбец
                    dt.Columns.Add(cell.GetValue<string>());
                    count++;
                }
            }
            //
            foreach (var row in ws.Rows())
            {
                // Первую пропускаем
                if (row == ws.Row(1)) continue;
                // На первой пустой строке прекращаем
                if (row.IsEmpty()) break;
                //
                dt.Rows.Add();
                int i = 0;
                //
                foreach (var cell in row.Cells())
                {
                    // Если ячеек больше, чем заголовков, прекращаем
                    if (i >= count) break;
                    dt.Rows[dt.Rows.Count - 1][i] = cell.Value;
                    i++;
                }
            }
            //
            return dt;
        }

        /// <summary>
        /// Из интернета, ругается, что метод не найден
        /// </summary>
        /// <param name="path"></param>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static DataTable GetDataFromExcel(string path, dynamic worksheet)
        {
            //Save the uploaded Excel file.


            DataTable dt = new DataTable();
            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(path))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(worksheet);

                //Create a new DataTable.

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            if (!string.IsNullOrEmpty(cell.Value.ToString()))
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            else
                            {
                                break;
                            }
                        }
                        firstRow = false;
                    }
                    else
                    {
                        int i = 0;
                        DataRow toInsert = dt.NewRow();
                        foreach (IXLCell cell in row.Cells(1, dt.Columns.Count))
                        {
                            try
                            {
                                toInsert[i] = cell.Value.ToString();
                            }
                            catch (Exception ex)
                            {

                            }
                            i++;
                        }
                        dt.Rows.Add(toInsert);
                    }
                }
                return dt;
            }
        }

            /// <summary>
            /// Чтение ячейки листа, используется в <see cref="ReadExcelSheet(string, bool, int)"/>.
            /// </summary>
            /// <param name="doc"></param>
            /// <param name="cell"></param>
            /// <returns></returns>
            private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.CellValue == null) return "";
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }

        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="text"></param>
        public static void InsertText(string docName, string text)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                // Insert the text into the SharedStringTablePart.
                int index = InsertSharedStringItem(text, shareStringPart);

                // Insert a new worksheet.
                WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

                // Insert cell A1 into the new worksheet.
                Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

                // Set the value of cell A1.
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
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
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
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

    }
}
