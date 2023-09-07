using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        //private void CreateProtocol(List<ReportRegionBlock> blocks, string templatePath, string resultPath, string mediaresource)
        //{
        //    //
        //    string fieldMedia = "";
        //    string fileName = "_";
        //    switch (mediaresource)
        //    {
        //        case "Маяк":
        //            fieldMedia = settings.Наименование_СМИ_Маяк;
        //            fileName = "Маяк.docx";
        //            break;
        //        case "Вести ФМ":
        //            fieldMedia = settings.Наименование_СМИ_Вести_ФМ;
        //            fileName = "Вести ФМ.docx";
        //            break;
        //        case "Радио России":
        //            fieldMedia = settings.Наименование_СМИ_Радио_России;
        //            fileName = "Радио России.docx";
        //            break;
        //        case "Россия 1":
        //            fieldMedia = settings.Наименование_СМИ_Россия_1;
        //            fileName = "Россия 1.docx";
        //            break;
        //        case "Россия 24":
        //            fieldMedia = settings.Наименование_СМИ_Россия_24;
        //            fileName = "Россия 24.docx";
        //            break;
        //    }
        //    // Новый протокол
        //    var document = new WordDocument(templatePath);
        //    // Заполняем поля слияния
        //    document.SetMergeFieldText("Наименование_СМИ", $"{fieldMedia}");
        //    document.SetMergeFieldText("ИО_Фамилия_предст_СМИ", $"{settings.Кандидаты_ИО_Фамилия_предст_СМИ}");
        //    document.SetMergeFieldText("Дата", $"{settings.Кандидаты_Дата}");
        //    document.SetMergeFieldText("ИО_Фамилия_члена_изб_ком", $"{settings.Кандидаты_ИО_Фамилия_члена_изб_ком}");
        //    //
        //    try
        //    {
        //        document.SetBookmarkText($"Талон", "");
        //        var table = CreateTableProtocolCandidates(protocol, mediaresource);
        //        document.SetBookmarkTable($"Талон", table);
        //    }
        //    catch { }
        //    //
        //    document.Save(resultPath + $"{fileName}");
        //    document.Close();
        //}


        ///// <summary>
        ///// Создает таблицу для отчета в ворде
        ///// </summary>
        ///// <param name="candidates"></param>
        ///// <param name="parties"></param>
        ///// <param name="mediaresource"></param>
        ///// <returns></returns>
        //Table CreateTableReport(List<Candidate> candidates, List<Party> parties, string mediaresource)
        //{
        //    // 
        //    Table table = new Table();
        //    //
        //    TableProperties tblProp = new TableProperties();
        //    TableBorders tblBorders = new TableBorders()
        //    {
        //        BottomBorder = new BottomBorder()
        //        {
        //            Size = 4,
        //            Val = BorderValues.Single
        //        },
        //        TopBorder = new TopBorder()
        //        {
        //            Size = 4,
        //            Val = BorderValues.Single
        //        },
        //        LeftBorder = new LeftBorder()
        //        {
        //            Size = 4,
        //            Val = BorderValues.Single
        //        },
        //        RightBorder = new RightBorder()
        //        {
        //            Size = 4,
        //            Val = BorderValues.Single
        //        },
        //        InsideHorizontalBorder = new InsideHorizontalBorder()
        //        {
        //            Size = 4,
        //            Val = BorderValues.Single
        //        },
        //        InsideVerticalBorder = new InsideVerticalBorder()
        //        {
        //            Size = 4,
        //            Val = BorderValues.Single
        //        }
        //    };
        //    tblProp.Append(tblBorders);
        //    //// Now we create a new layout and make it "fixed".
        //    //TableLayout tl = new TableLayout() { Type = TableLayoutValues.Fixed };
        //    //tblProp.TableLayout = tl;
        //    //
        //    table.Append(tblProp);
        //    // Заголовки таблицы
        //    TableRow trHead = new TableRow();
        //    var tcH4 = new TableCell(CreateParagraph($"Даты и время\r\n" +
        //        $"выхода в эфир предвыборных\r\n" +
        //        $"агитационных материалов\r\n" +
        //        $"(число, месяц, год; время;\r\n" +
        //        $"количество\r\n" +
        //        $"минут/секунд"));
        //    tcH4.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
        //    trHead.Append(
        //        new TableCell(CreateParagraph($"№ п/п")),
        //        new TableCell(CreateParagraph($"Фамилия, инициалы\r\n" +
        //        $"зарегистрированного кандидата,\r\n" +
        //        $"№ одномандатного\r\n" +
        //        $"избирательного округа, по\r\n" +
        //        $"которому он зарегистрирован")),
        //        new TableCell(CreateParagraph($"Даты и время выхода в эфир\r\n" +
        //        $"совместных агитационных\r\n" +
        //        $"мероприятий\r\n" +
        //        $"(число, месяц, год; время;\r\n" +
        //        $"количество\r\n" +
        //        $"минут/секунд")),
        //        tcH4,
        //        new TableCell(CreateParagraph($"Фамилия, инициалы представителя\r\n" +
        //        $"зарегистрированного кандидата,\r\n" +
        //        $"участвовавшего\r\n" +
        //        $"в жеребьевке (члена\r\n" +
        //        $"соответствующей\r\n" +
        //        $"избирательной комиссии с\r\n" +
        //        $"правом решающего голоса)")),
        //        new TableCell(CreateParagraph($"Подпись зарегистрированного кандидата,\r\n" +
        //        $"участвовавшего в жеребьевке\r\n" +
        //        $"(члена соответствующей\r\n" +
        //        $"избирательной комиссии с\r\n" +
        //        $"правом решающего голоса)\r\n" +
        //        $"и дата подписания"))
        //        );
        //    // Добавляем заголовок к таблице
        //    table.Append(trHead);
        //    // Добавляем строчку с нумерованием столбцов
        //    TableRow tr = new TableRow();
        //    TableCell tc1 = new TableCell(CreateParagraph($"1"));
        //    TableCell tc2 = new TableCell(CreateParagraph($"2"));
        //    TableCell tc3 = new TableCell(CreateParagraph($"3"));
        //    TableCell tc4 = new TableCell(CreateParagraph($"4"));
        //    TableCell tc5 = new TableCell(CreateParagraph($"5"));
        //    TableCell tc6 = new TableCell(CreateParagraph($"6"));
        //    tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
        //    table.Append(tr);
        //    // Добавляем округ одной строкой
        //    tr = CreateRowMergedCells(protocol.Округ, 6);
        //    table.Append(tr);
        //    // Для строки Итого со всех кандидатов длительность берем
        //    TimeSpan duration = TimeSpan.Zero;
        //    // По каждому кандидату из протокола
        //    for (int i = 0; i < protocol.Candidates.Count; i++)
        //    {
        //        //
        //        var c = protocol.Candidates[i];
        //        //
        //        string cell5Text = "";
        //        if (c.Info.Явка_кандидата == "1")
        //        {
        //            // Если кандидат внесен (вообще и так должен быть внесен)
        //            if (c.Info.Фамилия.Length > 0 &&
        //            c.Info.Имя.Length > 0 &&
        //            c.Info.Отчество.Length > 0)
        //            {
        //                cell5Text = $"{c.Info.Фамилия} {c.Info.Имя[0]}. {c.Info.Отчество[0]}.";
        //            }
        //        }
        //        else if (c.Info.Явка_представителя == "1")
        //        {
        //            // Если Представитель внесен
        //            if (c.Info.Представитель_Фамилия.Length > 0 &&
        //            c.Info.Представитель_Имя.Length > 0 &&
        //            c.Info.Представитель_Отчество.Length > 0)
        //            {
        //                cell5Text = $"{c.Info.Представитель_Фамилия} {c.Info.Представитель_Имя[0]}. {c.Info.Представитель_Отчество[0]}.";
        //            }
        //        }
        //        else
        //        {
        //            cell5Text = $"{protocol.Изб_ком_Фамилия_ИО}";
        //        }
        //        //
        //        Talon talon = null;
        //        // Определяем, какой из талонов надо использовать
        //        switch (mediaresource)
        //        {
        //            case "Маяк":
        //                talon = c.Талон_Маяк;
        //                break;
        //            case "Вести ФМ":
        //                talon = c.Талон_Вести_ФМ;
        //                break;
        //            case "Радио России":
        //                talon = c.Талон_Радио_России;
        //                break;
        //            case "Россия 1":
        //                talon = c.Талон_Россия_1;
        //                break;
        //            case "Россия 24":
        //                talon = c.Талон_Россия_24;
        //                break;
        //        }
        //        // Для общей длительности в Итого
        //        if (talon != null && talon.TotalDuration != null) duration += talon.TotalDuration;
        //        // Делаем строку кандидата
        //        tr = CreateRowProtocolCandidates(c, talon, mediaresource, i, cell5Text);
        //        //
        //        if (tr == null) continue;
        //        // Добавляем к таблице
        //        table.Append(tr);
        //    }
        //    // Строка "Итого"
        //    tr = new TableRow();
        //    tc1 = new TableCell(CreateParagraph($"Итого"));
        //    tc2 = new TableCell(CreateParagraph($""));
        //    tc3 = new TableCell(CreateParagraph($""));
        //    if (duration != TimeSpan.Zero)
        //    {
        //        tc4 = new TableCell(CreateParagraph($"{duration}"));
        //    }
        //    else
        //    {
        //        tc4 = new TableCell(CreateParagraph($""));
        //    }
        //    tc5 = new TableCell(CreateParagraph($""));
        //    tc6 = new TableCell(CreateParagraph($""));
        //    tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
        //    table.Append(tr);
        //    // Возвращаем
        //    return table;
        //}

        ///// <summary>
        ///// Создает строку таблицы в ворде
        ///// </summary>
        ///// <param name="candidate"></param>
        ///// <param name="talon"></param>
        ///// <param name="mediaresource"></param>
        ///// <param name="i"></param>
        ///// <param name="row5Text"></param>
        ///// <returns></returns>
        //private TableRow CreateRowProtocolCandidates(Candidate candidate, Talon talon, string mediaresource, int i, string row5Text)
        //{
        //    var tr = new TableRow();

        //    //
        //    if (candidate.Info.Фамилия == "") return null;
        //    // Формируем текст ячейки с талоном
        //    List<string> lines = new List<string>();
        //    //
        //    if (talon != null)
        //    {
        //        // Добавляем номер талона
        //        lines.Add($"Талон № {talon.Id}");
        //        //
        //        foreach (var row in talon.TalonRecords)
        //        {
        //            if (talon.MediaResource == "Вести ФМ")
        //            {
        //                lines.Add($"{row.Date} {row.Time}:{row.Time.Second} {row.Duration} {row.Description}");
        //            }
        //            else
        //            {
        //                lines.Add($"{row.Date} {row.Time} {row.Duration} {row.Description}");
        //            }
        //        }
        //    }
        //    else
        //    {
        //        lines.Add("");
        //    }
        //    // Строка с данными
        //    tr = new TableRow();
        //    //// Чтобы не разделялась при переходе на другую страницу
        //    //var rowProp = new TableRowProperties(new CantSplit());
        //    //tr.Append(rowProp);
        //    // 
        //    var tc1 = new TableCell(CreateParagraph($"{i + 1}"));
        //    var tc2 = new TableCell(CreateParagraph($"{candidate.Info.Фамилия} {candidate.Info.Имя} {candidate.Info.Отчество}, {candidate.Info.Округ_Номер}"));
        //    var tc3 = new TableCell(CreateParagraph($""));
        //    var tc4 = new TableCell(CreateParagraph(lines));
        //    //tc4.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "60" }));
        //    var tc5 = new TableCell(CreateParagraph($"{row5Text}"));
        //    var tc6 = new TableCell(CreateParagraph($""));
        //    tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
        //    //
        //    return tr;
        //}

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
            // Находим блок с партиями (номер региона - не цифры)
            var partyBlock = regionBlocks.Where(x => !Regex.IsMatch(x.RegionNumber, @"^\d+$")).First();
            // Перемещаем блок в конец
            regionBlocks.Remove(partyBlock);
            regionBlocks.Add(partyBlock);
            //
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
