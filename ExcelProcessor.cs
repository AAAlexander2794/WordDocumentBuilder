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
                    //Read the first row as header
                    if (counter == 1)
                    {
                        var j = 1;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                            dt.Columns.Add(colunmName);
                        }
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                            i++;
                        }
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// Чтение ячейки листа, используется в <see cref="ReadExcelSheet(string, bool, int)"/>.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }

        /// <summary>
        /// Альтернативное чтение листа Экселя, не используется
        /// </summary>
        /// <returns></returns>
        static string ReadExcelAlt()
        {
            string result = "";
            string fileName = "data.xlsm";

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    // Книга
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    // 
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;
                    // 
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();

                    result += $"Row count = {rows.LongCount()}\n";
                    result += $"Cell count = {cells.LongCount()}\n";

                    // One way: go through each cell in the sheet
                    foreach (Cell cell in cells)
                    {
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            int ssid = int.Parse(cell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            result += $"Shared string {ssid}: {str}\n";
                        }
                        else if (cell.CellValue != null)
                        {
                            result += $"Cell contents: {cell.CellValue.Text}";
                        }
                    }

                    // Or... via each row
                    foreach (Row row in rows)
                    {
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            if (c.DataType != null && c.DataType == CellValues.SharedString)
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                result += $"Shared string {ssid}: {str}\n";
                            }
                            else if (c.CellValue != null)
                            {
                                result += $"Cell contents: {c.CellValue.Text}";
                            }
                        }
                    }
                }
            }
            return result;

        }

    }
}
