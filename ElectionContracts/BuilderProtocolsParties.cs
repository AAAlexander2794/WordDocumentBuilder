using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.ElectionContracts
{
    /// <summary>
    /// 
    /// </summary>
    /// <remarks>
    /// Здесь используются методы этого класса из других файлов.
    /// </remarks>
    public partial class Builder
    {
        /// <summary>
        /// 
        /// </summary>
        /// <remarks>
        /// По каждому СМИ по каждой партии делаем отдельный протокол.
        /// Сортируем в подкаталоги партии, в нем по 5 протоколов для каждой СМИ.
        /// </remarks>
        /// <param name="talonVariant"></param>
        /// <returns></returns>
        public DataTable BuildProtocolsParties(string talonVariant = "default")
        {
            var _folderPath = $"{Settings.Default.Protocols_FolderPath}{DateTime.Now.ToString().Replace(":", "_")}\\";
            var _templatePath = Settings.Default.Protocols_TemplateFilePath_Parties;
            var _protocolsFilePath = Settings.Default.Protocols_FilePath;
            // test
            DataTable dt;
            try
            {
                dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_FilePath, sheetNumber: 0);
            }
            catch
            {
                throw new Exception("ошибка с чтением таблицы");
            }
            // Настройки текущей жеребьевки
            ProtocolsInfo settings;
            try
            {
                settings = ReadProtocols(_protocolsFilePath);
            }
            catch { throw new Exception("Не читает настройки протоколов."); }
            // Получаем список партий
            List<Party> parties;
            try
            {
                parties = BuildParties(talonVariant);
            }
            catch(Exception ex)
            {
                throw new Exception($"BuildParties\r\n{ex.Message}");
            }
            // По каждой партии
            foreach (var party in parties)
            {
                try
                {
                    // Если не отмечено на печать, пропускаем
                    if (party.Info.На_печать == "") continue;
                    // Формируем путь к документу
                    var resultPath = $"{_folderPath}" + $"{party.Info.Партия_Название_Краткое}\\";
                    // Создает путь для документов, если вдруг каких-то папок нет
                    Directory.CreateDirectory(resultPath);
                    // По каждому СМИ
                    CreateProtocol(party, settings, _templatePath, resultPath, "Маяк");
                    CreateProtocol(party, settings, _templatePath, resultPath, "Вести ФМ");
                    CreateProtocol(party, settings, _templatePath, resultPath, "Радио России");
                    CreateProtocol(party, settings, _templatePath, resultPath, "Россия 1");
                    CreateProtocol(party, settings, _templatePath, resultPath, "Россия 24");
                }
                catch
                {
                    throw new Exception($"Ошибка с {party.Info.Партия_Название_Полное}");
                }
            }
            //
            return dt;
        }

        /// <summary>
        /// Формирует файл протокола партии
        /// </summary>
        /// <param name="party"></param>
        /// <param name="templatePath"></param>
        /// <param name="resultPath"></param>
        /// <param name="mediaresource"></param>
        private void CreateProtocol(Party party, ProtocolsInfo settings, string templatePath, string resultPath, string mediaresource)
        {
            //
            var partyName = $"{party.Info.Партия_Название_Полное}";
            // Фамилия И.О. человека, который подписывает талон в протоколе
            string personName = "";
            if (party.Info.Явка_представителя == "1")
            {
                if (party.Info.Представитель_Фамилия.Length > 0 &&
                party.Info.Представитель_Имя.Length > 0 &&
                party.Info.Представитель_Отчество.Length > 0)
                {
                    personName = $"{party.Info.Представитель_Фамилия} {party.Info.Представитель_Имя[0]}. {party.Info.Представитель_Отчество[0]}.";
                }
            }
            else
            {
                personName = $"{settings.Партии_Фамилия_ИО_члена_изб_ком}";
            }
            //
            string fieldMedia = "";
            string fileName = "_";
            //
            Table table = null;
            //
            switch (mediaresource)
            {
                case "Маяк":
                    fieldMedia = settings.Наименование_СМИ_Маяк;
                    fileName = "Маяк.docx";
                    table = CreateTableParty(party.Талон_Маяк, partyName, personName, GetCustomCommonLines_Маяк(), "00:30:00");
                    break;
                case "Вести ФМ":
                    fieldMedia = settings.Наименование_СМИ_Вести_ФМ;
                    fileName = "Вести ФМ.docx";
                    table = CreateTableParty(party.Талон_Вести_ФМ, partyName, personName, GetCustomCommonLines_Вести_ФМ(), "01:00:00");
                    break;
                case "Радио России":
                    fieldMedia = settings.Наименование_СМИ_Радио_России;
                    fileName = "Радио России.docx";
                    table = CreateTableParty(party.Талон_Радио_России, partyName, personName, GetCustomCommonLines_Радио_России(), "00:17:45");
                    break;
                case "Россия 1":
                    fieldMedia = settings.Наименование_СМИ_Россия_1;
                    fileName = "Россия 1.docx";
                    table = CreateTableParty(party.Талон_Россия_1, partyName, personName, GetCustomCommonLines_Россия_1(), "00:23:45");
                    break;
                case "Россия 24":
                    fieldMedia = settings.Наименование_СМИ_Россия_24;
                    fileName = "Россия 24.docx";
                    table = CreateTableParty(party.Талон_Россия_24, partyName, personName, GetCustomCommonLines_Россия_24(), "00:17:45");
                    break;

            }
            // Новый протокол
            var document = new WordDocument(templatePath);
            // Заполняем поля слияния
            document.SetMergeFieldText("Наименование_СМИ", $"{fieldMedia}");
            document.SetMergeFieldText("ИО_Фамилия_предст_СМИ", $"{settings.Партии_ИО_Фамилия_предст_СМИ}");
            document.SetMergeFieldText("Дата", $"{settings.Партии_Дата}");
            document.SetMergeFieldText("ИО_Фамилия_члена_изб_ком", $"{settings.Партии_ИО_Фамилия_члена_изб_ком}");
            //
            try
            {
                document.SetBookmarkText($"Талон", "");
                document.SetBookmarkTable($"Талон", table);
            }
            catch { }
            //
            document.Save(resultPath + $"{fileName}");
            document.Close();
        }

        /// <summary>
        /// Захардкоженная таблица протокола партии
        /// </summary>
        /// <param name="talon"></param>
        /// <returns></returns>
        Table CreateTableParty(Talon talon, string lastRow2CellText = "", string lastRow5CellText = "", List<string> linesCustom = null, string durationCustom = "")
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
            table.Append(tblProp);
            // Заголовки таблицы
            TableRow trHead = new TableRow();
            var tcH4 = new TableCell(CreateParagraph($"Даты и время\r\n" +
                $"выхода в эфир предвыборных\r\n" +
                $"агитационных материалов\r\n" +
                $"(число, месяц, год; время;\r\n" +
                $"количество\r\n" +
                $"минут/секунд"));
            tcH4.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
            trHead.Append(
                new TableCell(CreateParagraph($"№ п/п")),
                new TableCell(CreateParagraph($"Наименование избирательного объединения")),
                new TableCell(CreateParagraph($"Даты и время выхода в эфир\r\n" +
                $"совместных агитационных\r\n" +
                $"мероприятий\r\n" +
                $"(число, месяц, год; время;\r\n" +
                $"количество\r\n" +
                $"минут/секунд")),
                tcH4,
                new TableCell(CreateParagraph($"Фамилия, инициалы\r\n" +
                $"представителя избирательного\r\n" +
                $"объединения, участвовавшего\r\n" +
                $"в жеребьевке (члена\r\n" +
                $"соответствующей\r\n" +
                $"избирательной комиссии с\r\n" +
                $"правом решающего голоса)")),
                new TableCell(CreateParagraph($"Подпись представителя\r\n" +
                $"избирательного объединения,\r\n" +
                $"участвовавшего в жеребьевке\r\n" +
                $"(члена соответствующей\r\n" +
                $"избирательной комиссии с\r\n" +
                $"правом решающего голоса)\r\n" +
                $"и дата подписания"))
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
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
            table.Append(tr);
            // Формируем текст ячейки с талоном
            List<string> lines = new List<string>();
            //
            if (talon != null)
            {
                // Добавляем номер талона
                lines.Add($"Талон № {talon.Id}");
                //
                foreach (var row in talon.TalonRecords)
                {
                    lines.Add($"{row.Date} {row.Time} {row.Duration} {row.Description}");
                }
            }
            else
            {
                lines.Add("");
            }
            // Строка с данными
            tr = new TableRow();
            //
            tc1 = new TableCell(CreateParagraph($""));
            tc2 = new TableCell(CreateParagraph($"{lastRow2CellText}"));
            if (linesCustom != null)
            {
                tc3 = new TableCell(CreateParagraph(linesCustom));
            }
            tc4 = new TableCell(CreateParagraph(lines));
            tc5 = new TableCell(CreateParagraph($"{lastRow5CellText}"));
            tc6 = new TableCell(CreateParagraph($""));
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
            table.Append(tr);
            // Строка "Итого"
            tr = new TableRow();
            tc1 = new TableCell(CreateParagraph($"Итого"));
            tc2 = new TableCell(CreateParagraph($""));
            tc3 = new TableCell(CreateParagraph($"{durationCustom}"));
            if (talon != null && talon.TotalDuration != null && talon.TotalDuration != TimeSpan.Zero)
            {
                tc4 = new TableCell(CreateParagraph($"{talon.TotalDuration}"));
            }
            else
            {
                tc4 = new TableCell(CreateParagraph($""));
            }
            tc5 = new TableCell(CreateParagraph($""));
            tc6 = new TableCell(CreateParagraph($""));
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
            table.Append(tr);
            // Возвращаем
            return table;
        }

        List<string> GetCustomCommonLines_Россия_1()
        {
            List<string> lines = new List<string>();
            //
            lines.Add("17.08.2023 09:34 00:03:00");
            lines.Add("22.08.2023 09:34 00:03:00");
            lines.Add("23.08.2023 09:34 00:03:00");
            lines.Add("24.08.2023 09:34 00:03:00");
            lines.Add("29.08.2023 09:34 00:03:00");
            lines.Add("30.08.2023 09:34 00:03:00");
            lines.Add("31.08.2023 09:34 00:03:00");
            lines.Add("05.09.2023 09:34 00:02:45");
            //
            return lines;
        }

        List<string> GetCustomCommonLines_Маяк()
        {
            List<string> lines = new List<string>();
            //
            lines.Add("17.08.2023 11:00 00:25:00");
            lines.Add("22.08.2023 11:00 00:25:00");
            lines.Add("24.08.2023 11:00 00:25:00");
            lines.Add("29.08.2023 11:00 00:25:00");
            lines.Add("31.08.2023 11:00 00:25:00");
            lines.Add("05.09.2023 11:00 00:25:00");
            //
            return lines;
        }

        List<string> GetCustomCommonLines_Радио_России()
        {
            List<string> lines = new List<string>();
            //
            lines.Add("17.08.2023 20:10 00:03:00");
            lines.Add("22.08.2023 20:10 00:03:00");
            lines.Add("24.08.2023 20:10 00:03:00");
            lines.Add("29.08.2023 20:10 00:03:00");
            lines.Add("31.08.2023 20:10 00:03:00");
            lines.Add("05.09.2023 20:10 00:02:45");
            //
            return lines;
        }

        List<string> GetCustomCommonLines_Россия_24()
        {
            List<string> lines = new List<string>();
            //
            lines.Add("17.08.2023 10:00 00:03:00");
            lines.Add("22.08.2023 10:00 00:03:00");
            lines.Add("24.08.2023 10:00 00:03:00");
            lines.Add("29.08.2023 10:00 00:03:00");
            lines.Add("31.08.2023 10:00 00:03:00");
            lines.Add("05.09.2023 10:00 00:02:45");
            //
            return lines;
        }

        List<string> GetCustomCommonLines_Вести_ФМ()
        {
            List<string> lines = new List<string>();
            //
            lines.Add("17.08.2023 11:30 00:10:00");
            lines.Add("22.08.2023 11:30 00:10:00");
            lines.Add("24.08.2023 11:30 00:10:00");
            lines.Add("29.08.2023 11:30 00:10:00");
            lines.Add("31.08.2023 11:30 00:10:00");
            lines.Add("05.09.2023 11:30 00:10:00");
            //
            return lines;
        }
    }
}
