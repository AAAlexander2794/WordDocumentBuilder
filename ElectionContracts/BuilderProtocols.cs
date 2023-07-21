﻿using DocumentFormat.OpenXml;
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
            // test
            DataTable dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_FilePath, sheetNumber: 0);
            // Получаем список партий
            var parties = BuildParties(talonVariant);
            
            // По каждой партии
            foreach (var party in parties)
            {
                // Если не отмечено на печать, пропускаем
                if (party.Info.На_печать == "") continue;                
                // Формируем путь к документу
                var resultPath = $"{_folderPath}" + $"{party.Info.Партия_Название}\\";
                // Создает путь для документов, если вдруг каких-то папок нет
                Directory.CreateDirectory(resultPath);
                // По каждому СМИ
                CreateProtocol(party, _templatePath, resultPath, "Маяк");
                CreateProtocol(party, _templatePath, resultPath, "Вести ФМ");
                CreateProtocol(party, _templatePath, resultPath, "Радио России");
                CreateProtocol(party, _templatePath, resultPath, "Россия 1");
                CreateProtocol(party, _templatePath, resultPath, "Россия 24");
            }
            //
            return dt;
        }

        private void CreateProtocol(Party party, string templatePath, string resultPath, string mediaresource)
        {
            //
            string fieldMedia = "";
            string fileName = "_";
            switch (mediaresource)
            {
                case "Маяк":
                    fieldMedia = "Радиостанция \"Маяк\"";
                    fileName = "Маяк.docx";
                    break;
                case "Вести ФМ":
                    fieldMedia = "Радиостанция \"Вести ФМ\"";
                    fileName = "Вести ФМ.docx";
                    break;
                case "Радио России":
                    fieldMedia = "Радиостанция \"Радио России\"";
                    fileName = "Радио России.docx";
                    break;
                case "Россия 1":
                    fieldMedia = "Телеканал \"Россия\" (\"Россия-1\")";
                    fileName = "Россия 1.docx";
                    break;
                case "Россия 24":
                    fieldMedia = "Телеканал \"Россия\" (\"Россия-24\")";
                    fileName = "Россия 24.docx";
                    break;

            }
            // Новый протокол
            var document = new WordDocument(templatePath);
            // Заполняем поля слияния
            document.SetMergeFieldText("Медиаресурс", $"{fieldMedia}");
            //
            var partyName = $"{party.Info.Партия_Отделение} {party.Info.Партия_Название}";
            // Фамилия И.О. человека, который подписывает протокол
            var personName = $"{party.Info.Представитель_Фамилия} {party.Info.Представитель_Имя[0]}. {party.Info.Представитель_Отчество[0]}.";
            //
            try
            {
                document.SetBookmarkText($"Талон", "");
                var table = CreateTableParty(party.Талон_Маяк, partyName, personName);
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
        Table CreateTableParty(Talon talon, string lastRow2CellText ="", string lastRow5CellText = "")
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
            // Заголовки таблицы
            TableRow trHead = new TableRow();
            trHead.Append(
                new TableCell(CreateParagraph($"№ п/п")),
                new TableCell(CreateParagraph($"Наименование избирательного объединения")),
                new TableCell(CreateParagraph($"Даты и время выхода в эфир\r\n" +
                $"совместных агитационных\r\n" +
                $"мероприятий\r\n" +
                $"(число, месяц, год; время;\r\n" +
                $"количество\r\n" +
                $"минут/секунд")),
                new TableCell(CreateParagraph($"Даты и время\r\n" +
                $"выхода в эфир предвыборных" +
                $"агитационных материалов\r\n" +
                $"(число, месяц, год; время;\r\n" +
                $"количество\r\n" +
                $"минут/секунд")),
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
            foreach (var row in talon.TalonRecords)
            {
                lines.Add($"{row.Date} {row.Time} {row.Duration} {row.Description}");
            }
            // Строка с данными
            tr = new TableRow();
            tc1 = new TableCell(CreateParagraph($""));
            tc2 = new TableCell(CreateParagraph($"{lastRow2CellText}"));
            tc3 = new TableCell(CreateParagraph($""));
            tc4 = new TableCell(CreateParagraph(lines));
            tc5 = new TableCell(CreateParagraph($"{lastRow5CellText}"));
            tc6 = new TableCell(CreateParagraph($""));
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
            table.Append(tr);
            // Строка "Итого"
            tr = new TableRow();
            tc1 = new TableCell(CreateParagraph($"Итого"));
            tc2 = new TableCell(CreateParagraph($""));
            tc3 = new TableCell(CreateParagraph($""));
            tc4 = new TableCell(CreateParagraph($""));
            tc5 = new TableCell(CreateParagraph($""));
            tc6 = new TableCell(CreateParagraph($""));
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
            table.Append(tr);
            // Возвращаем
            return table;
        }
    }
}
