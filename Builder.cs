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
using WordDocumentBuilder.Entities;

namespace WordDocumentBuilder
{
    public class Builder
    {
        const string _templatePath = "Шаблон Договора.dotx";
        const string _dataFilepath = "data.xlsm";

        public void Do()
        {
            //
            var talonRecordInfos = ReadTalonRecords();
            //
            var talonRecords = BuildTalonRecords(talonRecordInfos);
            //
            var talons = BuildTalons(talonRecords);
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
                //
                document.SetMergeFieldText("Фамилия", "Это просто фамилия");
                // Сохраняем и закрываем
                document.Save(resultPath);
                document.Close();
            }
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
                    Номер_талона = dt.Rows[i].Field<string>(3),
                    Номер_договора = dt.Rows[i].Field<string>(4),
                    Постановление_ТИК = dt.Rows[i].Field<string>(5),
                    Фамилия_представителя = dt.Rows[i].Field<string>(6),
                    Имя_представителя = dt.Rows[i].Field<string>(7),
                    Отчество_представителя = dt.Rows[i].Field<string>(8),
                    Дата_договора = dt.Rows[i].Field<string>(9)
                });
            }
            return candidates;
        }


        List<TalonRecordInfo> ReadTalonRecords()
        {
            var dt = ExcelProcessor.ReadExcelSheet(_dataFilepath, sheetNumber: 1);
            var records = new List<TalonRecordInfo>();
            // По строкам
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                records.Add(new TalonRecordInfo(
                    dt.Rows[i].Field<string>(0),
                    dt.Rows[i].Field<string>(1),
                    dt.Rows[i].Field<string>(2),
                    dt.Rows[i].Field<string>(3),
                    dt.Rows[i].Field<string>(4)));
            }
            return records;
        }

        List<TalonRecord> BuildTalonRecords(List<TalonRecordInfo> infos)
        {
            var records = new List<TalonRecord>();
            foreach (var inf in infos)
            {
                records.Add(new TalonRecord(inf));
            }
            return records;
        }

        List<Talon> BuildTalons(List<TalonRecord> talonRecords)
        {
            var talons = new List<Talon>();
            //
            var ids = new List<int>();
            foreach (var rec in talonRecords)
            {
                ids.Add(rec.Id);
            }
            //
            var uniqIds = ids.Distinct().ToList();
            //
            foreach (var id in uniqIds)
            {
                var talon = new Talon()
                {
                    Id = id,
                    TalonRecords = talonRecords.Where(rec => rec.Id == id).ToList()
                };
                talons.Add(talon);
            }
            return talons;
        }

        List<Candidate> BuildCandidates(List<CandidateInfo> infos, List<Talon> talons)
        {
            var candidates = new List<Candidate>();
            foreach (var info in infos)
            {
                var talon = talons.FirstOrDefault(t => t.Id.ToString() == info.Номер_талона);
                if (talon == null) continue;
                var candidate = new Candidate(info, talon);
                candidates.Add(candidate);
            }
            return candidates;
        }

        /// <summary>
        /// Захардкоженное присваивание значений закладкам.
        /// </summary>
        private void SetValues(WordDocument doc, Candidate c)
        {
            var table = CreateTable(c.Talon);
            doc.SetBookmarkText("Фамилия", $"{c.Info.Фамилия}");
            doc.SetBookmarkText("Фамилия1", $"{c.Info.Фамилия}");
            doc.SetBookmarkText("Имя", $"{c.Info.Имя}");
            doc.SetBookmarkText("Имя1", $"{c.Info.Имя}");
            doc.SetBookmarkText("Отчество", $"{c.Info.Отчество}");
            doc.SetBookmarkText("Отчество1", $"{c.Info.Отчество}");
            doc.SetBookmarkText("Номер_договора", $"{c.Info.Номер_договора}");
            doc.SetBookmarkText("Номер_договора1", $"{c.Info.Номер_договора}");
            doc.SetBookmarkText("Номер_договора2", $"{c.Info.Номер_договора}");
            doc.SetBookmarkText("Дата_договора", $"{c.Info.Дата_договора}");
            doc.SetBookmarkText("Постановление_ТИК", $"{c.Info.Постановление_ТИК}");
            doc.SetBookmarkText("ФИО_представителя", $"{c.Info.Фамилия_представителя} " +
                $"{c.Info.Имя_представителя} {c.Info.Отчество_представителя}");
            doc.SetBookmarkText("Талон", $"{c.Talon.Id}");
            doc.SetBookmarkTable("Талон", table);
        }

        /// <summary>
        /// Захардкоженная таблица талона
        /// </summary>
        /// <param name="talon"></param>
        /// <returns></returns>
        Table CreateTable(Talon talon)
        {
            // 
            Table table = new Table();
            ////
            //TableProperties tPr = new TableProperties();
            //tPr.TableIndentation = new TableIndentation() { Firs };
            //table.Append(tPr);
            //
            foreach (var row in talon.TalonRecords)
            {
                //
                TableRow tr = new TableRow();
                //
                TableCell tc1 = new TableCell(new Paragraph(new Run(new Text($"{row.MediaResource}")), 
                    new ParagraphProperties(new Indentation() { FirstLine = "0" })));
                TableCell tc2 = new TableCell(new Paragraph(new Run(new Text($"{row.Date}"))));
                TableCell tc3 = new TableCell(new Paragraph(new Run(new Text($"{row.Time}"))));
                TableCell tc4 = new TableCell(new Paragraph(new Run(new Text($"{row.Duration}"))));
                //
                tr.Append(tc1, tc2, tc3, tc4);
                //
                table.Append(tr);
            }
            return table;
        }

    }
    
}
