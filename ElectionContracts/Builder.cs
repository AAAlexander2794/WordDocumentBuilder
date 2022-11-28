using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.ElectionContracts
{
    public class Builder
    {
        const string _templatePath = "Шаблон Договора.dotx";
        const string _dataFilepath = "data.xlsm";
        

        public void Do(string talonVariant = "1")
        {
            var talons = new List<Talon>();
            if (talonVariant == "1")
            {
                // Вариант 1
                talons = TalonBuilder.BuildTalonsVariant1();
            }
            else
            {
                // Вариант 2
                talons = TalonBuilder.BuildTalonsVariant2();
            }
            //
            var candidatesInfos = ReadCandidates();
            //
            var candidates = BuildCandidates(candidatesInfos, talons);
            // 
            foreach (var candidate in candidates)
            {
                var document = new WordDocument(_templatePath);
                var resultPath = $"{candidate.Info.Фамилия} {candidate.Info.Имя} {candidate.Info.Отчество}.docx";
                // Устанавливаем начения текста для закладок документа
                SetValues(document, candidate);
                // Сохраняем и закрываем
                document.Save(resultPath);
                document.Close();
            }
            //// test
            //ExcelProcessor.InsertText("test.xlsx", "some text");
        }

        List<CandidateInfo> ReadCandidates()
        {
            var dt = ExcelProcessor.ReadExcelSheet(_dataFilepath, sheetNumber: 0);
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
                    Округ_дат_падеж = dt.Rows[i].Field<string>(17)
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
        private void SetValues(WordDocument doc, Candidate c)
        {
            var table = CreateTable(c.Талон_Маяк);
            doc.SetBookmarkText($"Талон_1", "");
            doc.SetBookmarkTable($"Талон_1", table);
            //
            table = CreateTable(c.Талон_Радио_России);
            doc.SetBookmarkText($"Талон_2", "");
            doc.SetBookmarkTable($"Талон_2", table);
            //
            table = CreateTable(c.Талон_Вести_ФМ);
            doc.SetBookmarkText($"Талон_3", "");
            doc.SetBookmarkTable($"Талон_3", table);
            //
            table = CreateTable(c.Талон_Россия_1);
            doc.SetBookmarkText($"Талон_4", "");
            doc.SetBookmarkTable($"Талон_4", table);
            //
            table = CreateTable(c.Талон_Россия_24);
            doc.SetBookmarkText($"Талон_5", "");
            doc.SetBookmarkTable($"Талон_5", table);


            //
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
