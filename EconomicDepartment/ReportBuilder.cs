using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    public class ReportBuilder
    {
        public void BuildTotalReports()
        {
            //
            string subCatalog = $"{DateTime.Now.ToShortDateString()} {DateTime.Now.Hour}.{DateTime.Now.Minute}.{DateTime.Now.Second}";
            //
            CreateTotalReport("Маяк", subCatalog);
            CreateTotalReport("Радио России", subCatalog);
            CreateTotalReport("Вести ФМ", subCatalog);
            CreateTotalReport("Россия 1", subCatalog);
            CreateTotalReport("Россия 24", subCatalog);
        }

        public void CreateTotalReport(string mediaResource, string subCatalog)
        {


            string recordsFilePath = $@"./Настройки/Учет вещания/{mediaResource}.xlsx";
            // Список строк вещания одной СМИ
            var broadcastRecords = ReadBroadcastRecordsFromExcel(recordsFilePath);
            // На основе строк одной СМИ делаем таблицу
            DataTable dt = BuildTotalReport(broadcastRecords);

            // Запись в файл Excel
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "Отчет");
            wb.SaveAs($@"./Документы/Отчеты/{subCatalog}/{mediaResource}.xlsx");

        }

        public DataTable BuildTotalReport(List<BroadcastRecord> records)
        {
            DataTable dt = new DataTable();
            // Создаем 7 столбцов
            dt.Columns.Add("№ п/п");
            dt.Columns.Add("Ф.И.О. зарегистрированного кандидата");
            dt.Columns.Add("Форма предвыборной агитации");
            dt.Columns.Add("Наименование теле-, радиоматериала");
            dt.Columns.Add("Дата и время выхода в эфир");
            dt.Columns.Add("Объем фактически использованного эфирного времени (час:мин:сек)");
            dt.Columns.Add("Основание предоставления (дата заключения и номер договора)");
            //
            int i = 1;
            //
            foreach (var record in records)
            {
                dt.Rows.Add();
                dt.Rows[dt.Rows.Count - 1][0] = i;
                dt.Rows[dt.Rows.Count - 1][1] = record.ClientName;
                dt.Rows[dt.Rows.Count - 1][2] = record.BroadcastType;
                dt.Rows[dt.Rows.Count - 1][3] = record.BroadcastCaption;
                dt.Rows[dt.Rows.Count - 1][4] = $"{record.Date} {record.Time}";
                dt.Rows[dt.Rows.Count - 1][5] = record.DurationActual;
                dt.Rows[dt.Rows.Count - 1][6] = "";
                //
                i++;
            }
            //
            return dt;
        }

        List<BroadcastRecord> ReadBroadcastRecordsFromExcel(string filePath)
        {
            List<BroadcastRecord> broadcastRecords;
            //
            var dt = ExcelProcessor.ReadExcelSheet(filePath);
            //
            broadcastRecords = BuildBroadcastRecords(dt);
            //
            return broadcastRecords;
        }

        List<BroadcastRecord> BuildBroadcastRecords(DataTable dt)
        {
            List<BroadcastRecord> broadcastRecords = new List<BroadcastRecord>();
            //
            foreach (DataRow row in dt.Rows)
            {
                //
                var record = new BroadcastRecord(
                    row.ItemArray[0].ToString(),
                    row.ItemArray[1].ToString(),
                    row.ItemArray[2].ToString(),
                    row.ItemArray[3].ToString(),
                    row.ItemArray[4].ToString(),
                    row.ItemArray[5].ToString(),
                    row.ItemArray[6].ToString(),
                    row.ItemArray[7].ToString(),
                    row.ItemArray[8].ToString(),
                    row.ItemArray[9].ToString());
                //
                broadcastRecords.Add(record);
            }
            //
            return broadcastRecords;
        }
    }
}
