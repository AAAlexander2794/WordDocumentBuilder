using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.ElectionContracts
{
    public class Builder
    {
        //string _templatePath = Settings.Default.TemplateFilePath;
        //string _dataFilepath = Settings.Default.CandidatesFilePath;
        // С подпапкой времени
        //string _dateTimeForDirectory = DateTime.Now.ToString().Replace(":", "_");
        //string _contractsFolderPath = $"{Settings.Default.ContractsFolderPath}{DateTime.Now.ToString().Replace(":", "_")}\\";
        

        public DataTable Do(string talonVariant = "default")
        {
            // test
            DataTable dt = ExcelProcessor.ReadExcelSheet(Settings.Default.CandidatesFilePath, sheetNumber: 0);
            //
            string _contractsFolderPath = $"{Settings.Default.ContractsFolderPath}{DateTime.Now.ToString().Replace(":", "_")}\\";
            
            var talons = TalonBuilder.BuildTalons(talonVariant);
            
            //
            var candidatesInfos = ReadCandidates(Settings.Default.CandidatesFilePath);
            //
            var candidates = BuildCandidates(candidatesInfos, talons);
            // Создает путь для документов, если вдруг каких-то папок нет
            Directory.CreateDirectory(_contractsFolderPath);
            //
            foreach (var candidate in candidates) 
            {
                Debug.WriteLine(candidate.Info.На_печать);
                Debug.WriteLine($"{_contractsFolderPath}{candidate.Округ_для_создания_каталога}\\");
                // Если не отмечено на печать, пропускаем
                if (candidate.Info.На_печать == "") continue;
                // Создает подпапку округа
                Directory.CreateDirectory($"{_contractsFolderPath}{candidate.Округ_для_создания_каталога}\\");
                // Создаем договор РВ
                var document = new WordDocument(Settings.Default.TemplateFilePath_РВ);
                try
                {
                    
                    var resultPath = $"{_contractsFolderPath}{candidate.Округ_для_создания_каталога}\\" +
                        $"{candidate.Info.Фамилия} {candidate.Info.Имя} {candidate.Info.Отчество}";
                    try
                    {
                        // Устанавливаем значения текста для полей документа, кроме закладок
                        SetMergeFields(document, candidate);
                    }
                    catch { }
                    //
                    try
                    {
                        SetTables(document, candidate, "radio");
                    }
                    catch { }
                    try
                    {
                        // Сохраняем и закрываем
                        document.Save(resultPath + "_радио.docx");
                        document.Close();

                        // Повторяем создание документа для договора ТВ
                        document = new WordDocument(Settings.Default.TemplateFilePath_ТВ);
                    }
                    catch { }
                    try
                    {
                        SetMergeFields(document, candidate);
                    }
                    catch { }
                    try
                    {
                        SetTables(document, candidate, "tele");
                    }
                    catch
                    {

                    }
                    try
                    {
                        document.Save(resultPath + "_ТВ.docx");
                        document.Close();
                    }
                    catch { }
                }
                catch { }
            }
                
                return dt;
        }

        List<CandidateInfo> ReadCandidates(string dataFilePath)
        {
            var dt = ExcelProcessor.ReadExcelSheet(dataFilePath, sheetNumber: 0);
            var candidates = new List<CandidateInfo>();
            // По строкам
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                candidates.Add(new CandidateInfo()
                {
                    Фамилия = dt.Rows[i].Field<string>(0),
                    Имя = dt.Rows[i].Field<string>(1),
                    Отчество = dt.Rows[i].Field<string>(2),
                    Талон_Маяк = dt.Rows[i].Field<string>(3),
                    Талон_Вести_ФМ = dt.Rows[i].Field<string>(4),
                    Талон_Радио_России = dt.Rows[i].Field<string>(5),
                    Талон_Россия_1 = dt.Rows[i].Field<string>(6),
                    Талон_Россия_24 = dt.Rows[i].Field<string>(7),
                    Номер_договора = dt.Rows[i].Field<string>(8),
                    Постановление_ТИК = dt.Rows[i].Field<string>(9),
                    Фамилия_представителя = dt.Rows[i].Field<string>(10),
                    Имя_представителя = dt.Rows[i].Field<string>(11),
                    Отчество_представителя = dt.Rows[i].Field<string>(12),
                    Дата_договора = dt.Rows[i].Field<string>(13),
                    Доверенность_на_представителя = dt.Rows[i].Field<string>(14),
                    ИНН = dt.Rows[i].Field<string>(15),
                    Спец_изб_счет_номер = dt.Rows[i].Field<string>(16),
                    Округ_дат_падеж = dt.Rows[i].Field<string>(17),
                    На_печать = dt.Rows[i].Field<string>(18)
                }) ;
            }
            return candidates;
        }

        

        List<Candidate> BuildCandidates(List<CandidateInfo> infos, List<Talon> talons)
        {
            var candidates = new List<Candidate>();
            foreach (var info in infos)
            {
                //var talon = talons.FirstOrDefault(t => t.Id.ToString() == info.Номер_талона_1);
                //if (talon == null) continue;
                var candidate = new Candidate(info, talons);
                candidates.Add(candidate);
            }
            return candidates;
        }

        /// <summary>
        /// Захардкоженное присваивание значений местам в документе.
        /// </summary>
        private void SetMergeFields(WordDocument doc, Candidate c)
        {
            doc.SetMergeFieldText("Фамилия", $"{c.Info.Фамилия}");
            doc.SetMergeFieldText("Имя", $"{c.Info.Имя}");
            doc.SetMergeFieldText("Отчество", $"{c.Info.Отчество}");
            doc.SetMergeFieldText("Номер_договора", $"{c.Info.Номер_договора}");
            doc.SetMergeFieldText("Дата_договора", $"{c.Info.Дата_договора}");
            doc.SetMergeFieldText("Постановление_ТИК", $"{c.Info.Постановление_ТИК}");
            doc.SetMergeFieldText("Фамилия_представителя_род_падеж", $"{c.Info.Фамилия_представителя}");
            doc.SetMergeFieldText("Имя_представителя_род_падеж", $"{c.Info.Имя_представителя}");
            doc.SetMergeFieldText("Отчество_представителя_род_падеж", $"{c.Info.Отчество_представителя}");
            doc.SetMergeFieldText("ИО_Фамилия", $"{c.ИО_Фамилия}");
            doc.SetMergeFieldText("ИО_Фамилия_предст", $"{c.ИО_Фамилия_представителя}");
            doc.SetMergeFieldText("Доверенность_на_представителя", $"{c.Info.Доверенность_на_представителя}");
            doc.SetMergeFieldText("ИНН", $"{c.Info.ИНН}");
            doc.SetMergeFieldText("Спец_изб_счет", $"{c.Info.Спец_изб_счет_номер}");
            doc.SetMergeFieldText("Округ_дат_падеж", $"{c.Info.Округ_дат_падеж}");
        }

        /// <summary>
        /// Захардкоженное присваивание таблиц заладкам в документе
        /// </summary>
        /// <remarks>Так как в двух местах будут таблицы размещаться, 
        /// а закладка только на одно место, то сделаны дубликаты закладок 
        /// (надо потом в MergeField переделать)</remarks>
        /// <param name="doc"></param>
        /// <param name="c"></param>
        /// <param name="mode"></param>
        private void SetTables(WordDocument doc, Candidate c, string mode = "both")
        {
            Table table;
            Table table2;
            //
            doc.SetBookmarkText($"Талон_1", "");
            doc.SetBookmarkText($"Талон_2", "");
            doc.SetBookmarkText($"Талон_3", "");
            doc.SetBookmarkText($"Талон_4", "");
            doc.SetBookmarkText($"Талон_5", "");
            doc.SetBookmarkText($"Талон_1_2", "");
            doc.SetBookmarkText($"Талон_2_2", "");
            doc.SetBookmarkText($"Талон_3_2", "");
            doc.SetBookmarkText($"Талон_4_2", "");
            doc.SetBookmarkText($"Талон_5_2", "");
            //
            if (mode == "both" || mode == "radio")
            {
                //
                table = CreateTable(c.Талон_Маяк);
                table2 = CreateTable(c.Талон_Маяк);
                doc.SetBookmarkTable($"Талон_1", table);
                doc.SetBookmarkTable($"Талон_1_2", table2);
                //
                table = CreateTable(c.Талон_Радио_России);
                table2 = CreateTable(c.Талон_Радио_России);
                doc.SetBookmarkTable($"Талон_2", table);
                doc.SetBookmarkTable($"Талон_2_2", table2);
                //
                table = CreateTable(c.Талон_Вести_ФМ);
                table2 = CreateTable(c.Талон_Вести_ФМ);
                doc.SetBookmarkTable($"Талон_3", table);
                doc.SetBookmarkTable($"Талон_3_2", table2);
            }
            //
            if (mode == "both" || mode == "tele")
            {
                //
                table = CreateTable(c.Талон_Россия_1);
                table2 = CreateTable(c.Талон_Россия_1);
                doc.SetBookmarkTable($"Талон_4", table);
                doc.SetBookmarkTable($"Талон_4_2", table2);
                //
                table = CreateTable(c.Талон_Россия_24);
                table2 = CreateTable(c.Талон_Россия_24);
                doc.SetBookmarkTable($"Талон_5", table);
                doc.SetBookmarkTable($"Талон_5_2", table2);
            }
        }

        /// <summary>
        /// Захардкоженная таблица талона
        /// </summary>
        /// <param name="talon"></param>
        /// <returns></returns>
        Table CreateTable(Talon talon)
        {
            if (talon == null) return null;
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
            table.Append(tblProp);
            //
            TableRow trHead = new TableRow();
            trHead.Append(
                new TableCell(CreateParagraph($"Название радиоканала")),
                new TableCell(CreateParagraph($"Дата выхода в эфир")),
                new TableCell(CreateParagraph($"Время выхода \r\nв эфир")),
                new TableCell(CreateParagraph($"Хронометраж")),
                new TableCell(CreateParagraph($"Вид (форма) предвыборной агитации\r\n" +
                $"(Материалы, Совместные агитационные мероприятия)"))
                );
            //
            table.Append(trHead);
            //
            foreach (var row in talon.TalonRecords)
            {
                //
                TableRow tr = new TableRow();
                //
                TableCell tc1 = new TableCell(CreateParagraph($"{row.MediaResource}"));
                TableCell tc2 = new TableCell(CreateParagraph($"{row.Date}"));
                TableCell tc3 = new TableCell(CreateParagraph($"{row.Time}"));
                TableCell tc4 = new TableCell(CreateParagraph($"{row.Duration}"));
                TableCell tc5 = new TableCell(CreateParagraph($""));
                //
                tr.Append(tc1, tc2, tc3, tc4, tc5);
                //
                table.Append(tr);
            }
            return table;
        }

        Paragraph CreateParagraph(string text)
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
            paragraph.Append(run);
            //
            return paragraph;
        }



    }
    
}
