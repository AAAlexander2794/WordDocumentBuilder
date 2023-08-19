using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.EconomicDepartment
{
    public class Builder
    {
        public void BuildTable()
        {
            //
            var builder = new ElectionContracts.Builder();
            //
            var candidates = builder.BuildCandidates("1");
            //
            var parties = builder.BuildParties("1");
            //
            var broadcastRecords = BuildBroadcastRecords(parties, candidates);
            //
            string subCatalog = $"{DateTime.Now.ToShortDateString()} {DateTime.Now.Hour}.{DateTime.Now.Minute}.{DateTime.Now.Second}";
            //
            WriteBroadcastRecordsToExcel(broadcastRecords, $@"./Документы/Рабочие/{subCatalog}/Маяк.xlsx", "Маяк");
            WriteBroadcastRecordsToExcel(broadcastRecords, $@"./Документы/Рабочие/{subCatalog}/Радио России.xlsx", "Радио России");
            WriteBroadcastRecordsToExcel(broadcastRecords, $@"./Документы/Рабочие/{subCatalog}/Вести ФМ.xlsx", "Вести ФМ");
            WriteBroadcastRecordsToExcel(broadcastRecords, $@"./Документы/Рабочие/{subCatalog}/Россия 1.xlsx", "Россия 1");
            WriteBroadcastRecordsToExcel(broadcastRecords, $@"./Документы/Рабочие/{subCatalog}/Россия 24.xlsx", "Россия 24");
        }

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

        DataTable WriteBroadcastRecordsToExcel(List<BroadcastRecord> records, string filePath, string mediaResource)
        {
            //
            DataTable dt = new DataTable();
            // Заголовки таблицы
            dt.Columns.Add("Канал");
            dt.Columns.Add("Дата");
            dt.Columns.Add("Отрезок");
            dt.Columns.Add("Хрон");
            dt.Columns.Add("Округ");
            dt.Columns.Add("Партия/кандидат");
            dt.Columns.Add("Название партии/ФИО кандидата");
            dt.Columns.Add("Факт время");
            dt.Columns.Add("Форма выступления");
            dt.Columns.Add("Название ролика/тема дебатов");
            // Оставляем записи только заданного медиаресурса
            records = records.Where(x => x.MediaResource == mediaResource).ToList();
            //
            foreach (var record in records)
            {
                dt.Rows.Add();
                dt.Rows[dt.Rows.Count - 1][0] = record.MediaResource;
                dt.Rows[dt.Rows.Count - 1][1] = record.Date;
                dt.Rows[dt.Rows.Count - 1][2] = record.Time;
                dt.Rows[dt.Rows.Count - 1][3] = record.DurationNominal;
                dt.Rows[dt.Rows.Count - 1][4] = record.RegionNumber;
                dt.Rows[dt.Rows.Count - 1][5] = record.ClientType;
                dt.Rows[dt.Rows.Count - 1][6] = record.ClientName;
                dt.Rows[dt.Rows.Count - 1][7] = "";
                dt.Rows[dt.Rows.Count - 1][8] = "";
                dt.Rows[dt.Rows.Count - 1][9] = "";
            }
            // Запись в файл Excel
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "Отчет");
            wb.SaveAs(filePath);           
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

        /// <summary>
        /// 
        /// </summary>
        /// <remarks>
        /// Будущий я, прости меня
        /// </remarks>
        /// <param name="parties"></param>
        /// <param name="candidates"></param>
        /// <returns></returns>
        List<BroadcastRecord> BuildBroadcastRecords(List<Party> parties, List<Candidate> candidates)
        {
            List<BroadcastRecord> records = new List<BroadcastRecord>();
            foreach (var party in parties)
            {
                var list1 = BuildBroadcastRecords(party, party.Талон_Россия_1);
                var list2 = BuildBroadcastRecords(party, party.Талон_Россия_24);
                var list3 = BuildBroadcastRecords(party, party.Талон_Радио_России);
                var list4 = BuildBroadcastRecords(party, party.Талон_Маяк);
                var list5 = BuildBroadcastRecords(party, party.Талон_Вести_ФМ);
                foreach (var r in list1) records.Add(r);
                foreach (var r in list2) records.Add(r);
                foreach (var r in list3) records.Add(r);
                foreach (var r in list4) records.Add(r);
                foreach (var r in list5) records.Add(r);
            }
            foreach (var candidate in candidates)
            {
                var list1 = BuildBroadcastRecords(candidate, candidate.Талон_Россия_1);
                var list2 = BuildBroadcastRecords(candidate, candidate.Талон_Россия_24);
                var list3 = BuildBroadcastRecords(candidate, candidate.Талон_Радио_России);
                var list4 = BuildBroadcastRecords(candidate, candidate.Талон_Маяк);
                var list5 = BuildBroadcastRecords(candidate, candidate.Талон_Вести_ФМ);
                foreach (var r in list1) records.Add(r);
                foreach (var r in list2) records.Add(r);
                foreach (var r in list3) records.Add(r);
                foreach (var r in list4) records.Add(r);
                foreach (var r in list5) records.Add(r);
            }
            //
            return records;
        }

        List<BroadcastRecord> BuildBroadcastRecords(Party party, Talon talon)
        {
            List<BroadcastRecord> broadcastRecords = new List<BroadcastRecord>();
            if (talon == null) return broadcastRecords;
            foreach (var talonRecord in talon.TalonRecords)
            {
                BroadcastRecord record = new BroadcastRecord()
                {
                    MediaResource = talonRecord.MediaResource,
                    Date = talonRecord.Date,
                    Time = talonRecord.Time,
                    DurationNominal = talonRecord.Duration,
                    ClientType = "партия",
                    ClientName = party.Info.Партия_Название_Рабочее
                };
                broadcastRecords.Add(record);
            }
            return broadcastRecords;
        }

        List<BroadcastRecord> BuildBroadcastRecords(Candidate candidate, Talon talon)
        {
            List<BroadcastRecord> broadcastRecords = new List<BroadcastRecord>();
            if (talon == null) return broadcastRecords;
            foreach (var talonRecord in talon.TalonRecords)
            {
                BroadcastRecord record = new BroadcastRecord()
                {
                    MediaResource = talonRecord.MediaResource,
                    Date = talonRecord.Date,
                    Time = talonRecord.Time,
                    DurationNominal = talonRecord.Duration,
                    ClientType = "кандидат",
                    RegionNumber = candidate.Info.Округ_Номер,
                    ClientName = $"{candidate.Info.Фамилия} " +
                        $"{candidate.Info.Имя} " +
                        $"{candidate.Info.Отчество} " +
                        $"({candidate.Info.Округ_Номер})"
                };
                broadcastRecords.Add(record);
            }
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
