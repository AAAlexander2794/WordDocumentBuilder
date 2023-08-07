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
    /// <summary>
    /// Отдельный файл для всего, что по договорам кандидатов
    /// </summary>
    public partial class Builder
    {
        public DataTable BuildContractsCandidates(string talonVariant = "default")
        {
            // test
            DataTable dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_FilePath, sheetNumber: 0);
            //
            string _contractsFolderPath = $"{Settings.Default.ContractsFolderPath}{DateTime.Now.ToString().Replace(":", "_")}\\";
            // Получаем список кандидатов
            var candidates = BuildCandidates(talonVariant);
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
                var document = new WordDocument(Settings.Default.Candidates_TemplateFilePath_РВ);
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
                        document = new WordDocument(Settings.Default.Candidates_TemplateFilePath_ТВ);
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
                    На_печать = dt.Rows[i].Field<string>(0),
                    Округ_Номер = dt.Rows[i].Field<string>(1),
                    Округ_Название_падеж_им = dt.Rows[i].Field<string>(2).Trim(),
                    Фамилия = dt.Rows[i].Field<string>(3),
                    Имя = dt.Rows[i].Field<string>(4),
                    Отчество = dt.Rows[i].Field<string>(5),
                    Талон_Россия_1 = dt.Rows[i].Field<string>(6),
                    Талон_Россия_24 = dt.Rows[i].Field<string>(7),
                    Талон_Радио_России = dt.Rows[i].Field<string>(8),
                    Талон_Маяк = dt.Rows[i].Field<string>(9),
                    Талон_Вести_ФМ = dt.Rows[i].Field<string>(10),                    
                    Явка_кандидата = dt.Rows[i].Field<string>(11),
                    Явка_представителя = dt.Rows[i].Field<string>(12),
                    Представитель_Фамилия = dt.Rows[i].Field<string>(13),
                    Представитель_Имя = dt.Rows[i].Field<string>(14),
                    Представитель_Отчество = dt.Rows[i].Field<string>(15),
                    Постановление = dt.Rows[i].Field<string>(16),
                    Номер_договора = dt.Rows[i].Field<string>(17),
                    Дата_договора = dt.Rows[i].Field<string>(18),
                    Округ_Название_падеж_дат = dt.Rows[i].Field<string>(19),
                    Представитель_Доверенность = dt.Rows[i].Field<string>(20),
                    ИНН = dt.Rows[i].Field<string>(21),
                    Спец_изб_счет_номер = dt.Rows[i].Field<string>(22),
                    Партия = dt.Rows[i].Field<string>(23)
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

        List<Candidate> BuildCandidates(string talonVariant = "default")
        {
            //
            var talons = TalonBuilder.BuildTalonsCandidates(talonVariant);
            //
            var candidatesInfos = ReadCandidates(Settings.Default.Candidates_FilePath);
            //
            var candidates = BuildCandidates(candidatesInfos, talons);
            //
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
            doc.SetMergeFieldText("Округ_Название", $"{c.Info.Округ_Название_падеж_дат}");
            doc.SetMergeFieldText("Округ_Номер", $"{c.Info.Округ_Номер}");
            doc.SetMergeFieldText("Постановление", $"{c.Info.Постановление}");
            doc.SetMergeFieldText("Фамилия_представителя_род_падеж", $"{c.Info.Представитель_Фамилия}");
            doc.SetMergeFieldText("Имя_представителя_род_падеж", $"{c.Info.Представитель_Имя}");
            doc.SetMergeFieldText("Отчество_представителя_род_падеж", $"{c.Info.Представитель_Отчество}");
            doc.SetMergeFieldText("ИО_Фамилия", $"{c.ИО_Фамилия}");
            doc.SetMergeFieldText("ИО_Фамилия_предст", $"{c.ИО_Фамилия_представителя}");
            doc.SetMergeFieldText("Доверенность_на_представителя", $"{c.Info.Представитель_Доверенность}");
            doc.SetMergeFieldText("ИНН", $"{c.Info.ИНН}");
            doc.SetMergeFieldText("Спец_изб_счет", $"{c.Info.Спец_изб_счет_номер}");
            
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

    }
    
}
