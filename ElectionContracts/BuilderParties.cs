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
    /// Отдельный файл для всего, что по договорам партий
    /// </summary>
    public partial class Builder
    {
        public DataTable BuildContractsParties(string talonVariant = "default")
        {
            // test
            DataTable dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_FilePath, sheetNumber: 0);
            // Берем путь к каталогу договоров
            string _folderPath = $"{Settings.Default.ContractsFolderPath}{DateTime.Now.ToString().Replace(":", "_")}\\";
            // Получаем список партий
            var parties = BuildParties(talonVariant);
            // Создает путь для документов, если вдруг каких-то папок нет
            Directory.CreateDirectory(_folderPath);
            // Для каждой партии
            foreach (var party in parties)
            {
                // Если не отмечено на печать, пропускаем
                if (party.Info.На_печать == "") continue;
                // Создаем договор РВ
                var document = new WordDocument(Settings.Default.Parties_TemplateFilePath_РВ);
                // Формируем има файла договора
                var resultPath = $"{_folderPath}" + $"{party.Info.Партия_Название_Краткое}";
                // Устанавливаем значения текста для полей документа, кроме закладок (талонов)
                SetMergeFields(document, party);
                try
                {
                    // Устанавливаем таблицы талонов по закладкам
                    SetTables(document, party, "radio");
                }
                catch { };
                // Сохраняем и закрываем
                document.Save(resultPath + "_радио.docx");
                document.Close();
                // Повторяем для договора ТВ
                document = new WordDocument(Settings.Default.Parties_TemplateFilePath_ТВ);
                //
                SetMergeFields(document, party);
                try
                {
                    //
                    SetTables(document, party, "tele");
                }
                catch { };
                //
                document.Save(resultPath + "_ТВ.docx");
                document.Close();
            }
            //
            return dt;
        }

        List<PartyInfo> ReadParties(string dataFilePath)
        {
            var dt = ExcelProcessor.ReadExcelSheet(dataFilePath, sheetNumber: 0);
            var parties = new List<PartyInfo>();
            //
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                parties.Add(new PartyInfo()
                {
                    //
                    На_печать = dt.Rows[i].Field<string>(0),
                    //
                    Партия_Название_Полное = dt.Rows[i].Field<string>(1),
                    Партия_Название_Краткое = dt.Rows[i].Field<string>(2),
                    //
                    Представитель_Фамилия = dt.Rows[i].Field<string>(3),
                    Представитель_Имя = dt.Rows[i].Field<string>(4),
                    Представитель_Отчество = dt.Rows[i].Field<string>(5),
                    //
                    Талон_Маяк = dt.Rows[i].Field<string>(6),
                    Талон_Вести_ФМ = dt.Rows[i].Field<string>(7),
                    Талон_Радио_России = dt.Rows[i].Field<string>(8),
                    Талон_Россия_1 = dt.Rows[i].Field<string>(9),
                    Талон_Россия_24 = dt.Rows[i].Field<string>(10),
                    Явка_представителя = dt.Rows[i].Field<string>(11),
                    //
                    Постановление = dt.Rows[i].Field<string>(12),
                    //
                    Номер_договора = dt.Rows[i].Field<string>(13),
                    Дата_договора = dt.Rows[i].Field<string>(14),
                    //
                    Представитель_Доверенность = dt.Rows[i].Field<string>(15),
                    //
                    Нотариус_Город = dt.Rows[i].Field<string>(16),
                    Нотариус_Фамилия = dt.Rows[i].Field<string>(17),
                    Нотариус_Имя = dt.Rows[i].Field<string>(18),
                    Нотариус_Отчество = dt.Rows[i].Field<string>(19),
                    Нотариус_Реестр = dt.Rows[i].Field<string>(20),
                    //
                    ОГРН = dt.Rows[i].Field<string>(21),
                    ИНН = dt.Rows[i].Field<string>(22),
                    КПП = dt.Rows[i].Field<string>(23),
                    Спец_изб_счет_номер = dt.Rows[i].Field<string>(24),
                    Место_нахождения = dt.Rows[i].Field<string>(25),
                    Партия_Название_Рабочее = dt.Rows[i].Field<string>(26)
                });
            }
            return parties;
        }

        List<Party> BuildParties(List<PartyInfo> infos, List<Talon> talons)
        {
            var parties = new List<Party>();
            foreach (var info in infos)
            {
                try
                {
                    var party = new Party(info, talons);
                    parties.Add(party);
                }
                catch
                {
                    var text = "";
                    foreach (var t in talons)
                    {
                        text += t.GetTalonText() + "\r\n";
                    }
                    throw new Exception($"Ошибка с {info.Партия_Название_Краткое}\r\n{text}");
                }
            }
            return parties;
        }

        public List<Party> BuildParties(string talonVariant = "default")
        {
            List<Talon> talons;
            try
            {
                // Формируем талоны
                talons = TalonBuilder.BuildTalonsParties(talonVariant);
            }
            catch(Exception ex)
            {
                throw new Exception($"BuildTalonParties\r\n{ex.Message}");
            }
            // Читаем партии
            var partiesInfos = ReadParties(Settings.Default.Parties_FilePath);
            List<Party> parties;
            try
            {
                // Создаем сущности партий
                parties = BuildParties(partiesInfos, talons);
            }
            catch(Exception ex)
            {
                throw ex;
            }
            //
            return parties;
        }

        /// <summary>
        /// Захардкоженное присваивание значений местам в документе для партий.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="p"></param>
        private void SetMergeFields(WordDocument doc, Party p)
        {
            //
            doc.SetMergeFieldText("Номер_договора", $"{p.Info.Номер_договора}");
            doc.SetMergeFieldText("Дата_договора", $"{p.Info.Дата_договора}");
            //
            doc.SetMergeFieldText("Название", $"{p.Info.Партия_Название_Полное}");
            doc.SetMergeFieldText("Постановление", $"{p.Info.Постановление}");
            //
            doc.SetMergeFieldText("Представитель_Фамилия", $"{p.Info.Представитель_Фамилия}");
            doc.SetMergeFieldText("Представитель_Имя", $"{p.Info.Представитель_Имя}");
            doc.SetMergeFieldText("Представитель_Отчество", $"{p.Info.Представитель_Отчество}");
            doc.SetMergeFieldText("Представитель_Доверенность", $"{p.Info.Представитель_Доверенность}");
            doc.SetMergeFieldText("Представитель_ИО_Фамилия", $"{p.Представитель_ИО_Фамилия}");
            //
            doc.SetMergeFieldText("Нотариус_Фамилия", $"{p.Info.Нотариус_Фамилия}");
            doc.SetMergeFieldText("Нотариус_Имя", $"{p.Info.Нотариус_Имя}");
            doc.SetMergeFieldText("Нотариус_Отчество", $"{p.Info.Нотариус_Отчество}");
            doc.SetMergeFieldText("Нотариус_Реестр", $"{p.Info.Нотариус_Реестр}");
            doc.SetMergeFieldText("Нотариус_Город", $"{p.Info.Нотариус_Город}");
            //
            doc.SetMergeFieldText("ИНН", $"{p.Info.ИНН}");
            doc.SetMergeFieldText("КПП", $"{p.Info.КПП}");
            doc.SetMergeFieldText("ОГРН", $"{p.Info.ОГРН}");
            doc.SetMergeFieldText("Счет", $"{p.Info.Спец_изб_счет_номер}");
            doc.SetMergeFieldText("Место_нахождения", $"{p.Info.Место_нахождения}");
        }

        /// <summary>
        /// Захардкоженное присваивание таблиц заладкам в документе
        /// </summary>
        /// <remarks>Так как в двух местах будут таблицы размещаться, 
        /// а закладка только на одно место, то сделаны дубликаты закладок 
        /// (надо потом в MergeField переделать)</remarks>
        /// <param name="doc"></param>
        /// <param name="p"></param>
        /// <param name="mode"></param>
        private void SetTables(WordDocument doc, Party p, string mode = "both")
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
                table = CreateTable(p.Талон_Маяк);
                table2 = CreateTable(p.Талон_Маяк);
                doc.SetBookmarkTable($"Талон_1", table);
                doc.SetBookmarkTable($"Талон_1_2", table2);
                //
                table = CreateTable(p.Талон_Радио_России);
                table2 = CreateTable(p.Талон_Радио_России);
                doc.SetBookmarkTable($"Талон_2", table);
                doc.SetBookmarkTable($"Талон_2_2", table2);
                //
                table = CreateTable(p.Талон_Вести_ФМ);
                table2 = CreateTable(p.Талон_Вести_ФМ);
                doc.SetBookmarkTable($"Талон_3", table);
                doc.SetBookmarkTable($"Талон_3_2", table2);
            }
            //
            if (mode == "both" || mode == "tele")
            {
                //
                table = CreateTable(p.Талон_Россия_1);
                table2 = CreateTable(p.Талон_Россия_1);
                doc.SetBookmarkTable($"Талон_4", table);
                doc.SetBookmarkTable($"Талон_4_2", table2);
                //
                table = CreateTable(p.Талон_Россия_24);
                table2 = CreateTable(p.Талон_Россия_24);
                doc.SetBookmarkTable($"Талон_5", table);
                doc.SetBookmarkTable($"Талон_5_2", table2);
            }
        }
    }
}
