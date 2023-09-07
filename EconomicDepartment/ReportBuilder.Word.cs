using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    public partial class ReportBuilder
    {
        private void CreateReport(List<ReportRegionBlock> blocks, string templatePath, string resultPath, string mediaresource)
        {
            //
            string fieldMedia = "";
            string fileName = "_";
            switch (mediaresource)
            {
                case "Маяк":
                    //fieldMedia = settings.Наименование_СМИ_Маяк;
                    fileName = "Маяк.docx";
                    break;
                case "Вести ФМ":
                    //fieldMedia = settings.Наименование_СМИ_Вести_ФМ;
                    fileName = "Вести ФМ.docx";
                    break;
                case "Радио России":
                    //fieldMedia = settings.Наименование_СМИ_Радио_России;
                    fileName = "Радио России.docx";
                    break;
                case "Россия 1":
                    //fieldMedia = settings.Наименование_СМИ_Россия_1;
                    fileName = "Россия 1.docx";
                    break;
                case "Россия 24":
                    //fieldMedia = settings.Наименование_СМИ_Россия_24;
                    fileName = "Россия 24.docx";
                    break;
            }
            // Новый 
            var document = new WordDocument(templatePath);
            //// Заполняем поля слияния
            //document.SetMergeFieldText("Наименование_СМИ", $"{fieldMedia}");
            //document.SetMergeFieldText("ИО_Фамилия_предст_СМИ", $"{settings.Кандидаты_ИО_Фамилия_предст_СМИ}");
            //document.SetMergeFieldText("Дата", $"{settings.Кандидаты_Дата}");
            //document.SetMergeFieldText("ИО_Фамилия_члена_изб_ком", $"{settings.Кандидаты_ИО_Фамилия_члена_изб_ком}");
            //
            document.SetBookmarkText($"Учет", "");
            var table = CreateTableReport(blocks);
            document.SetBookmarkTable($"Учет", table);
            // Создает путь для документов, если вдруг каких-то папок нет
            Directory.CreateDirectory(resultPath);
            //
            document.Save(resultPath + $"{fileName}");
            document.Close();
        }


        /// <summary>
        /// Создает таблицу для отчета в ворде
        /// </summary>
        /// <param name="candidates"></param>
        /// <param name="parties"></param>
        /// <param name="mediaresource"></param>
        /// <returns></returns>
        Table CreateTableReport(List<ReportRegionBlock> blocks)
        {
            // 
            Table table = new Table();
            //
            TableProperties tblProp = new TableProperties();
            TableBorders tblBorders = new TableBorders()
            {
                BottomBorder = new BottomBorder()
                {
                    Size = 4,
                    Val = BorderValues.Single
                },
                TopBorder = new TopBorder()
                {
                    Size = 4,
                    Val = BorderValues.Single
                },
                LeftBorder = new LeftBorder()
                {
                    Size = 4,
                    Val = BorderValues.Single
                },
                RightBorder = new RightBorder()
                {
                    Size = 4,
                    Val = BorderValues.Single
                },
                InsideHorizontalBorder = new InsideHorizontalBorder()
                {
                    Size = 4,
                    Val = BorderValues.Single
                },
                InsideVerticalBorder = new InsideVerticalBorder()
                {
                    Size = 4,
                    Val = BorderValues.Single
                }
            };
            tblProp.Append(tblBorders);
            //// Now we create a new layout and make it "fixed".
            //TableLayout tl = new TableLayout() { Type = TableLayoutValues.Fixed };
            //tblProp.TableLayout = tl;
            //
            table.Append(tblProp);
            // Заголовки таблицы
            TableRow trHead = new TableRow();
            var tcH4 = new TableCell(CreateParagraph($"Наименование теле -, радиоматериала"));
            tcH4.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
            trHead.Append(
                new TableCell(CreateParagraph($"№ п/п")),
                new TableCell(CreateParagraph($"Ф.И.О.\r\n" +
                $"зарегистрированного кандидата,\r\n" +
                $"наименование избирательного объединения, зарегистрировавшего областной список кандидатов")),
                new TableCell(CreateParagraph($"Форма предвыборной агитации ")),
                tcH4,
                new TableCell(CreateParagraph($"Дата и время выхода в эфир")),
                new TableCell(CreateParagraph($"Объем фактически предоставленного эфирного времени (час: мин:сек)")),
                new TableCell(CreateParagraph($"Основание предоставления\r\n(дата заключения\r\nи номер договора)\r\n"))
                );
            // Добавляем заголовок к таблице
            table.Append(trHead);
            // Добавляем строчку с нумерованием столбцов
            TableRow tr = new TableRow();
            TableCell tc1 = new TableCell(CreateParagraph($"1"));
            TableCell tc2 = new TableCell(CreateParagraph($"2"));
            TableCell tc3 = new TableCell(CreateParagraph($"3"));
            TableCell tc4 = new TableCell(CreateParagraph($"4"));
            TableCell tc5 = new TableCell(CreateParagraph($"5"));
            TableCell tc6 = new TableCell(CreateParagraph($"6"));
            TableCell tc7 = new TableCell(CreateParagraph($"7"));
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6, tc7);
            table.Append(tr);
            // Счетчик п/п
            int i = 1;
            // Суммарный объем фактически выделенного времени в СМИ
            TimeSpan sumDuration = TimeSpan.Zero;
            // По региону
            foreach (var region in blocks)
            {
                // Если не было вещания фактического, пропускаем.
                if (region.TotalDuration == TimeSpan.Zero) continue;
                //
                string regionCaption = "";
                // Если номер округа есть, то указываем номер и название
                if (Regex.IsMatch(region.RegionNumber, @"^\d+$"))
                {
                    regionCaption = $"{region.RegionCaption} округ № {region.RegionNumber}";
                }
                else
                {
                    regionCaption = $"Партии";
                }
                // Добавляем округ одной строкой
                tr = CreateRowMergedCells(regionCaption, 7);
                table.Append((TableRow)tr.CloneNode(true));
                // По клиенту
                foreach (var client in region.ClientBlocks)
                {
                    // Если не было вещания фактического, пропускаем.
                    if (client.TotalDuration == TimeSpan.Zero) continue;
                    //
                    bool firstRow = true;
                    //
                    // По каждой записи вещания клиента
                    foreach (var record in client.BroadcastRecords)
                    {
                        // Если не было вещания фактического, пропускаем.
                        if (record.DurationActual == TimeSpan.Zero) continue;
                        //
                        if (firstRow)
                        {
                            tc1 = new TableCell(CreateParagraph($"{i}"));
                            tc2 = new TableCell(CreateParagraph($"{client.ClientName}"));
                            tc7 = new TableCell(CreateParagraph($"{client.ClientContract}"));
                        }
                        else
                        {
                            tc1 = new TableCell(CreateParagraph($""));
                            tc2 = new TableCell(CreateParagraph($""));
                            tc7 = new TableCell(CreateParagraph($""));
                        }
                        //
                        tc3 = new TableCell(CreateParagraph($"{record.BroadcastType}"));
                        tc4 = new TableCell(CreateParagraph($"{record.BroadcastCaption}"));
                        tc5 = new TableCell(CreateParagraph($"{record.Date} {record.Time}"));
                        tc6 = new TableCell(CreateParagraph($"{record.DurationActual}"));
                        //
                        firstRow = false;
                        //
                        tr = new TableRow();
                        tr.Append((TableCell)tc1.CloneNode(true), 
                            (TableCell)tc2.CloneNode(true),
                            (TableCell)tc3.CloneNode(true),
                            (TableCell)tc4.CloneNode(true),
                            (TableCell)tc5.CloneNode(true),
                            (TableCell)tc6.CloneNode(true), 
                            (TableCell)tc7.CloneNode(true));
                        table.Append((TableRow)tr.CloneNode(true));
                    }
                    //
                    if (client.TotalDuration != TimeSpan.Zero)
                    {
                        tc1 = new TableCell(CreateParagraph($""));
                        tc2 = new TableCell(CreateParagraph($""));
                        tc3 = new TableCell(CreateParagraph($""));
                        tc4 = new TableCell(CreateParagraph($""));
                        tc5 = new TableCell(CreateParagraph($"Итого"));
                        tc6 = new TableCell(CreateParagraph($"{client.TotalDuration}"));
                        tc7 = new TableCell(CreateParagraph($""));
                        //
                        tr = new TableRow();
                        tr.Append((TableCell)tc1.CloneNode(true),
                            (TableCell)tc2.CloneNode(true),
                            (TableCell)tc3.CloneNode(true),
                            (TableCell)tc4.CloneNode(true),
                            (TableCell)tc5.CloneNode(true),
                            (TableCell)tc6.CloneNode(true),
                            (TableCell)tc7.CloneNode(true));
                        table.Append((TableRow)tr.CloneNode(true));
                        //
                        sumDuration += client.TotalDuration;
                        // Увеличения счетчика п/п
                        i++;
                    }
                }
            }
            //
            tc1 = new TableCell(CreateParagraph($""));
            tc2 = new TableCell(CreateParagraph($""));
            tc3 = new TableCell(CreateParagraph($""));
            tc4 = new TableCell(CreateParagraph($""));
            tc5 = new TableCell(CreateParagraph($"Итого"));
            tc6 = new TableCell(CreateParagraph($"{sumDuration}"));
            tc7 = new TableCell(CreateParagraph($""));
            //
            tr = new TableRow();
            tr.Append((TableCell)tc1.CloneNode(true),
                            (TableCell)tc2.CloneNode(true),
                            (TableCell)tc3.CloneNode(true),
                            (TableCell)tc4.CloneNode(true),
                            (TableCell)tc5.CloneNode(true),
                            (TableCell)tc6.CloneNode(true),
                            (TableCell)tc7.CloneNode(true));
            table.Append((TableRow)tr.CloneNode(true));
            // Возвращаем
            return table;
        }

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

        /// <summary>
        /// Объединяет указанное количество ячеек в строке вместе, также вставляет туда текст. (в теории, пока нет)
        /// </summary>
        /// <param name="text"></param>
        /// <param name="cellsNumber"></param>
        /// <returns></returns>
        private TableRow CreateRowMergedCells(string text, int cellsNumber)
        {
            var tr = new TableRow();
            // Создаем свойства ячейки для начала объединения
            TableCellProperties propStart = new TableCellProperties();
            propStart.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Restart,
            });
            // Делаем ячейку с текстом и добавляем ей свойство начала объединения
            var tc = new TableCell(CreateParagraph($"{text}", "alignmentCenter"));
            tc.Append(propStart);
            tr.Append(tc);
            // Цикл по количеству ячеек, которые надо объединить
            for (int i = 1; i < cellsNumber; i++)
            {
                // Создаем свойства ячейки для продолжения объединения
                var prop = new TableCellProperties();
                prop.Append(new HorizontalMerge()
                {
                    Val = MergedCellValues.Continue
                });
                // Создаем новую ячейку
                var tcNext = new TableCell(CreateParagraph($""));
                // Прикрепляем к новой ячейке свойства продолжения объединения
                tcNext.Append(prop);
                // Добавляем ячейку к строке
                tr.Append(tcNext);
            };
            //
            return tr;
        }

        /// <summary>
        /// Создает новый абзац текста
        /// </summary>
        /// <param name="text"></param>
        /// <param name="style">Для выбора различных дополнений текста типа выравнивания по центру</param>
        /// <returns></returns>
        Paragraph CreateParagraph(string text, string style = "default")
        {
            var paragraph = new Paragraph();
            var run = new Run();
            var runText = new Text($"{text}");
            //
            RunProperties runProperties = new RunProperties();
            FontSize size = new FontSize();
            size.Val = StringValue.FromString("18");
            runProperties.Append(size);
            //
            run.Append(runProperties);
            run.Append(runText);
            //
            if (style == "alignmentCenter")
            {
                Justification justification = new Justification()
                {
                    Val = JustificationValues.Center
                };
                var prProp = new ParagraphProperties();
                prProp.Append(justification);
                paragraph.Append(prProp);
            }
            //
            paragraph.Append(run);
            //
            return paragraph;
        }
    }
}
