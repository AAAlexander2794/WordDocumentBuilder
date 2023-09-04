using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentBuilder.ElectionContracts;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.EconomicDepartment
{
    public class ReportBuilder
    {
        public void BuildTotalReports()
        {
            //
            string subCatalog = $"{DateTime.Now.ToShortDateString()} {DateTime.Now.Hour}.{DateTime.Now.Minute}.{DateTime.Now.Second}";
            //
            BuildTotalReport("Маяк", subCatalog);
            BuildTotalReport("Радио России", subCatalog);
            BuildTotalReport("Вести ФМ", subCatalog);
            BuildTotalReport("Россия 1", subCatalog);
            BuildTotalReport("Россия 24", subCatalog);
        }



        public void BuildTotalReport(string mediaResource, string subCatalog)
        {
            string recordsFilePath = $@"./Настройки/Учет вещания/{mediaResource}.xlsx";
            // Список строк вещания одной СМИ
            var broadcastRecords = ReadBroadcastRecordsFromExcel(recordsFilePath);
            // На основе строк одной СМИ строим блоки для таблицы
            var blocks = BuildTotalReport(broadcastRecords);
            // Добавляем к клиентам данные о договоре
            var builder = new ElectionContracts.Builder();
            var candidates = builder.BuildCandidates("1");
            foreach (var candidate in candidates)
            {
                bool isFound = false;
                // Поиск только по блокам, где указан округ, то есть только по кандидатам
                foreach (var block in blocks.Where(b => b.RegionNumber.All(Char.IsDigit)))
                {
                    foreach (var client in block.ClientBlocks)
                    {
                        // Если нашли этого кандидата, добавляем информацию о договоре и переходим сразу к следующему
                        if (client.ClientName == $"{candidate.Info.Фамилия} {candidate.Info.Имя} {candidate.Info.Отчество}")
                        {
                            isFound = true;
                            client.ClientContract = $"Договор № {candidate.Info.Номер_договора} от {candidate.Info.Дата_договора} г.";
                            break;
                        }
                    }
                    if (isFound) break;
                }
            }
            var parties = builder.BuildParties("1");
            foreach (var party in parties)
            {
                bool isFound = false;
                foreach (var block in blocks.Where(b => b.RegionNumber == "–"))
                {
                    foreach (var client in block.ClientBlocks)
                    {
                        // Если нашли партию, добавляем информацию о договоре и переходим сразу к следующему
                        if (party.Info.Партия_Название_Рабочее.Contains(client.ClientName))
                        {
                            isFound = true;
                            client.ClientContract = $"Договор № {party.Info.Номер_договора} от {party.Info.Дата_договора} г.";
                            break;
                        }
                    }
                    if (isFound) break;
                }
            }
            // Строим таблицу из блоков
            DataTable dt = BuildTotalReport(blocks);
            // Запись в файл Excel
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "Отчет");
            wb.SaveAs($@"./Документы/Отчеты/{subCatalog}/{mediaResource}.xlsx");

        }

        public List<ReportRegionBlock> BuildTotalReport(List<BroadcastRecord> broadcastRecords)
        {
            List<ReportRegionBlock> regionBlocks = new List<ReportRegionBlock>();
            // По каждой строке вещания
            foreach (var broadcastRecord in broadcastRecords)
            {
                bool regionCreated = false;
                foreach (var region in regionBlocks)
                {
                    // Если такой регион уже есть
                    if (region.RegionNumber == broadcastRecord.RegionNumber)
                    {
                        regionCreated = true;
                        bool clientCreated = false;
                        // По каждому клиенту в регионе
                        foreach (var client in region.ClientBlocks)
                        {
                            // Если такой клиент уже есть
                            if (client.ClientName == broadcastRecord.ClientName)
                            {
                                clientCreated = true;
                                client.BroadcastRecords.Add(broadcastRecord);
                                region.TotalDuration += broadcastRecord.DurationActual;
                            }
                        }
                        // Если клиента еще не создали
                        if (!clientCreated)
                        {
                            var newClient = new ReportClientBlock()
                            {
                                ClientName = broadcastRecord.ClientName
                            };
                            newClient.BroadcastRecords.Add(broadcastRecord);
                            region.ClientBlocks.Add(newClient);
                            region.TotalDuration += broadcastRecord.DurationActual;
                        }
                    }
                }
                // Если регион еще не создали
                if (!regionCreated)
                {
                    var newRegion = new ReportRegionBlock();
                    newRegion.RegionNumber = broadcastRecord.RegionNumber;
                    regionBlocks.Add(newRegion);
                    // Создаем нового клиента (регион новый, значит, клиента не было)
                    var newClient = new ReportClientBlock()
                    {
                        ClientName = broadcastRecord.ClientName
                    };
                    newClient.BroadcastRecords.Add(broadcastRecord);
                    newRegion.ClientBlocks.Add(newClient);
                    newRegion.TotalDuration += broadcastRecord.DurationActual;
                }
            }
            return regionBlocks;
        }

        public DataTable BuildTotalReport(List<ReportRegionBlock> blocks)
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
            // Счетчик п/п
            int i = 1;
            // Суммарный объем фактически выделенного времени в СМИ
            TimeSpan sumDuration = TimeSpan.Zero;
            // По каждому округу
            foreach (var block in blocks)
            {
                // Если не было вещания фактического, пропускаем.
                if (block.TotalDuration == TimeSpan.Zero) continue;
                //
                dt.Rows.Add();
                dt.Rows[dt.Rows.Count - 1][0] = $"Округ № {block.RegionNumber}";
                // По каждому клиенту
                foreach (var client in block.ClientBlocks)
                {
                    // Если не было вещания фактического, пропускаем.
                    if (client.TotalDuration == TimeSpan.Zero) continue;
                    //
                    bool firstRow = true;
                    // По каждой записи вещания клиента
                    foreach (var record in client.BroadcastRecords)
                    {
                        // Если не было вещания фактического, пропускаем.
                        if (record.DurationActual == TimeSpan.Zero) continue;
                        //
                        dt.Rows.Add();
                        if (firstRow)
                        {
                            dt.Rows[dt.Rows.Count - 1][0] = i;
                            dt.Rows[dt.Rows.Count - 1][1] = client.ClientName;
                            dt.Rows[dt.Rows.Count - 1][6] = client.ClientContract;
                        }
                        //
                        dt.Rows[dt.Rows.Count - 1][2] = record.BroadcastType;
                        dt.Rows[dt.Rows.Count - 1][3] = record.BroadcastCaption;
                        dt.Rows[dt.Rows.Count - 1][4] = $"{record.Date} {record.Time}";
                        dt.Rows[dt.Rows.Count - 1][5] = record.DurationActual;
                        //
                        firstRow = false;
                    }
                    //
                    if (client.TotalDuration != TimeSpan.Zero)
                    {
                        //
                        dt.Rows.Add();
                        dt.Rows[dt.Rows.Count - 1][4] = $"Итого";
                        dt.Rows[dt.Rows.Count - 1][5] = client.TotalDuration;
                        //
                        sumDuration += client.TotalDuration;
                        // Увеличения счетчика п/п
                        i++;
                    }
                }
            }
            //
            dt.Rows.Add();
            dt.Rows[dt.Rows.Count - 1][4] = $"Итого";
            dt.Rows[dt.Rows.Count - 1][5] = sumDuration;
            //
            return dt;
        }

        List<BroadcastRecord> ReadBroadcastRecordsFromExcel(string filePath)
        {
            List<BroadcastRecord> broadcastRecords;
            //
            var dt = ExcelProcessor.ReadExcelSheetClosedXML(filePath, "Отчет");
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
