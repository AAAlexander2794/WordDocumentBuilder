using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Security.Cryptography;
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
        /// По каждому СМИ по каждому округу делаем отдельный протокол.
        /// Разделяем по подпапкам округов, в каждой по 5 протоколов.
        /// </remarks>
        /// <param name="talonVariant"></param>
        /// <returns></returns>
        public DataTable BuildProtocolsCandidates(string talonVariant = "default")
        {
            var _folderPath = $"{Settings.Default.Protocols_FolderPath}{DateTime.Now.ToString().Replace(":", "_")}\\";
            var _templatePath = Settings.Default.Protocols_TemplateFilePath_Candidates;
            var _protocolsFilePath = Settings.Default.Protocols_FilePath;
            // test
            DataTable dt;
            try
            {
                 dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_FilePath, sheetNumber: 0);
            }
            catch
            {
                throw new Exception("Не читает таблицу.");
            }
            // Получаем список Кандидатов
            List<Candidate> candidates;
            try
            {
                candidates = BuildCandidates(talonVariant);
            }
            catch { throw new Exception("Ошибка с талонами."); }
            // Настройки текущей жеребьевки
            ProtocolsInfo settings;
            try
            {
                settings = ReadProtocols(_protocolsFilePath);
            }
            catch { throw new Exception("Не читает настройки протоколов."); }
            // Создаем список протоколов
            List<ProtocolCandidates> protocols;
            try
            {
                protocols = CreateProtocolsCandidates(candidates, settings);
            }
            catch { throw new Exception("Ошибка со списком протоколов."); }
            // По каждому протоколу
            try
            {
                foreach (var protocol in protocols)
                {
                    // Формируем путь к документу
                    var resultPath = $"{_folderPath}" + $"{protocol.Округ}\\";
                    // Создает путь для документов, если вдруг каких-то папок нет
                    Directory.CreateDirectory(resultPath);
                    // По каждому СМИ
                    CreateProtocol(protocol, settings, _templatePath, resultPath, "Маяк");
                    CreateProtocol(protocol, settings, _templatePath, resultPath, "Вести ФМ");
                    CreateProtocol(protocol, settings, _templatePath, resultPath, "Радио России");
                    CreateProtocol(protocol, settings, _templatePath, resultPath, "Россия 1");
                    CreateProtocol(protocol, settings, _templatePath, resultPath, "Россия 24");
                }
            }
            catch { throw new Exception("Ошибка с записью протоколов."); }
            //
            return dt;
        }

        /// <summary>
        /// Из списка кандидатов формирует список протоколов по округам.
        /// </summary>
        /// <remarks>
        /// В каждом протоколе из списка кандидаты лежат с талонами по всем СМИ, т.е. далее надо будет из каждого протокола 
        /// сформировать по 5 непосредственно документов протокола для каждого СМИ.
        /// </remarks>
        /// <param name="candidates"></param>
        /// <returns></returns>
        private List<ProtocolCandidates> CreateProtocolsCandidates(List<Candidate> candidates, ProtocolsInfo settings)
        {
            //
            var protocols = new List<ProtocolCandidates>();
            //
            foreach(var candidate in candidates)
            {
                //
                if (candidate.Info.Фамилия.Trim() == "" || candidate.Info.На_печать != "1") continue;
                //
                bool exist = false;
                // Проход по уже существующим протоколам
                foreach (var protocol in protocols)
                {
                    // Если в протоколах уже есть округ этого кандидата
                    if (protocol.Округ == candidate.Округ_полное_название)
                    {
                        // Добавляем к протоколу кандидата
                        protocol.Candidates.Add(candidate);
                        // Помечаем, что кандидат добавлен
                        exist = true;
                        // Выходим с цикла
                        break;
                    }
                }
                // Если после в существующих протоколах не встретился округ кандидата
                if (exist == false)
                {
                    // Создаем новый протокол
                    var newProtocol = new ProtocolCandidates()
                    {
                        // Округ текущего кандидата
                        Округ = candidate.Округ_полное_название,
                        // Из настроек
                        Изб_ком_Фамилия_ИО = settings.Кандидаты_Фамилия_ИО_члена_изб_ком,
                        СМИ_ИО_Фамилия = settings.Кандидаты_ИО_Фамилия_предст_СМИ
                    };
                    // Добавляем к новому протоколу текущего кандидата
                    newProtocol.Candidates.Add(candidate);
                    // Добавляем новый протокол к списку протоколов
                    protocols.Add(newProtocol);
                }
            }
            //
            return protocols;
        }

        /// <summary>
        /// Читаем файл настроек протоколов и храним это в отдельной сущности.
        /// </summary>
        /// <param name="dataFilePath"></param>
        /// <returns></returns>
        ProtocolsInfo ReadProtocols(string dataFilePath)
        {
            var dt = ExcelProcessor.ReadExcelSheet(dataFilePath, sheetNumber: 0);
            var info = new ProtocolsInfo()
            {
                // Партии
                Партии_Фамилия_ИО_члена_изб_ком = dt.Rows[0].Field<string>(0),
                Партии_ИО_Фамилия_члена_изб_ком = dt.Rows[0].Field<string>(1),
                Партии_ИО_Фамилия_предст_СМИ = dt.Rows[0].Field<string>(2),
                Партии_Дата = dt.Rows[0].Field<string>(3),
                // Кандидаты
                Кандидаты_Фамилия_ИО_члена_изб_ком = dt.Rows[0].Field<string>(4),
                Кандидаты_ИО_Фамилия_члена_изб_ком = dt.Rows[0].Field<string>(5),
                Кандидаты_ИО_Фамилия_предст_СМИ = dt.Rows[0].Field<string>(6),
                Кандидаты_Дата = dt.Rows[0].Field<string>(7),
                // Общее
                Наименование_СМИ_Маяк = dt.Rows[0].Field<string>(8),
                Наименование_СМИ_Вести_ФМ = dt.Rows[0].Field<string>(9),
                Наименование_СМИ_Радио_России = dt.Rows[0].Field<string>(10),
                Наименование_СМИ_Россия_1 = dt.Rows[0].Field<string>(11),
                Наименование_СМИ_Россия_24 = dt.Rows[0].Field<string>(12)                
            };
            return info;
        }

        /// <summary>
        /// Формирует файл протокола кандидатов
        /// </summary>
        private void CreateProtocol(ProtocolCandidates protocol, ProtocolsInfo settings, string templatePath, string resultPath, string mediaresource)
        {
            //
            string fieldMedia = "";
            string fileName = "_";
            switch (mediaresource)
            {
                case "Маяк":
                    fieldMedia = settings.Наименование_СМИ_Маяк;
                    fileName = "Маяк.docx";
                    break;
                case "Вести ФМ":
                    fieldMedia = settings.Наименование_СМИ_Вести_ФМ;
                    fileName = "Вести ФМ.docx";
                    break;
                case "Радио России":
                    fieldMedia = settings.Наименование_СМИ_Радио_России;
                    fileName = "Радио России.docx";
                    break;
                case "Россия 1":
                    fieldMedia = settings.Наименование_СМИ_Россия_1;
                    fileName = "Россия 1.docx";
                    break;
                case "Россия 24":
                    fieldMedia = settings.Наименование_СМИ_Россия_24;
                    fileName = "Россия 24.docx";
                    break;
            }
            // Новый протокол
            var document = new WordDocument(templatePath);
            // Заполняем поля слияния
            document.SetMergeFieldText("Наименование_СМИ", $"{fieldMedia}");
            document.SetMergeFieldText("ИО_Фамилия_предст_СМИ", $"{settings.Кандидаты_ИО_Фамилия_предст_СМИ}");
            document.SetMergeFieldText("Дата", $"{settings.Кандидаты_Дата}");
            document.SetMergeFieldText("ИО_Фамилия_члена_изб_ком", $"{settings.Кандидаты_ИО_Фамилия_члена_изб_ком}");
            //
            try
            {
                document.SetBookmarkText($"Талон", "");
                var table = CreateTableProtocolCandidates(protocol, mediaresource);
                document.SetBookmarkTable($"Талон", table);
            }
            catch { }
            //
            document.Save(resultPath + $"{fileName}");
            document.Close();
        }

        /// <summary>
        /// Захардкоженная таблица протокола кандидатов
        /// </summary>
        /// <param name="talon"></param>
        /// <returns></returns>
        Table CreateTableProtocolCandidates(ProtocolCandidates protocol, string mediaresource)
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
            var tcH4 = new TableCell(CreateParagraph($"Даты и время\r\n" +
                $"выхода в эфир предвыборных\r\n" +
                $"агитационных материалов\r\n" +
                $"(число, месяц, год; время;\r\n" +
                $"количество\r\n" +
                $"минут/секунд"));
            tcH4.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
            trHead.Append(
                new TableCell(CreateParagraph($"№ п/п")),
                new TableCell(CreateParagraph($"Фамилия, инициалы\r\n" +
                $"зарегистрированного кандидата,\r\n" +
                $"№ одномандатного\r\n" +
                $"избирательного округа, по\r\n" +
                $"которому он зарегистрирован")),
                new TableCell(CreateParagraph($"Даты и время выхода в эфир\r\n" +
                $"совместных агитационных\r\n" +
                $"мероприятий\r\n" +
                $"(число, месяц, год; время;\r\n" +
                $"количество\r\n" +
                $"минут/секунд")),
                tcH4,
                new TableCell(CreateParagraph($"Фамилия, инициалы представителя\r\n" +
                $"зарегистрированного кандидата,\r\n" +
                $"участвовавшего\r\n" +
                $"в жеребьевке (члена\r\n" +
                $"соответствующей\r\n" +
                $"избирательной комиссии с\r\n" +
                $"правом решающего голоса)")),
                new TableCell(CreateParagraph($"Подпись зарегистрированного кандидата,\r\n" +
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
            // Добавляем округ одной строкой
            tr = CreateRowMergedCells(protocol.Округ, 6);
            table.Append(tr);
            // Для строки Итого со всех кандидатов длительность берем
            TimeSpan duration = TimeSpan.Zero;
            // По каждому кандидату из протокола
            for (int i = 0; i < protocol.Candidates.Count; i++)
            {
                //
                var c = protocol.Candidates[i];
                //
                string cell5Text = "";
                if (c.Info.Явка_кандидата == "1")
                {
                    // Если кандидат внесен (вообще и так должен быть внесен)
                    if (c.Info.Фамилия.Length > 0 &&
                    c.Info.Имя.Length > 0 &&
                    c.Info.Отчество.Length > 0)
                    {
                        cell5Text = $"{c.Info.Фамилия} {c.Info.Имя[0]}. {c.Info.Отчество[0]}.";
                    }
                }
                else if (c.Info.Явка_представителя == "1")
                {
                    // Если Представитель внесен
                    if (c.Info.Представитель_Фамилия.Length > 0 &&
                    c.Info.Представитель_Имя.Length > 0 &&
                    c.Info.Представитель_Отчество.Length > 0)
                    {
                        cell5Text = $"{c.Info.Представитель_Фамилия} {c.Info.Представитель_Имя[0]}. {c.Info.Представитель_Отчество[0]}.";
                    }
                }
                else
                {
                    cell5Text = $"{protocol.Изб_ком_Фамилия_ИО}";
                }
                //
                Talon talon = null;
                // Определяем, какой из талонов надо использовать
                switch (mediaresource)
                {
                    case "Маяк":
                        talon = c.Талон_Маяк;
                        break;
                    case "Вести ФМ":
                        talon = c.Талон_Вести_ФМ;
                        break;
                    case "Радио России":
                        talon = c.Талон_Радио_России;
                        break;
                    case "Россия 1":
                        talon = c.Талон_Россия_1;
                        break;
                    case "Россия 24":
                        talon = c.Талон_Россия_24;
                        break;
                }
                // Для общей длительности в Итого
                if (talon != null && talon.TotalDuration != null) duration += talon.TotalDuration;
                // Делаем строку кандидата
                tr = CreateRowProtocolCandidates(c, talon, mediaresource, i, cell5Text);
                //
                if (tr == null) continue;
                // Добавляем к таблице
                table.Append(tr);
            }          
            // Строка "Итого"
            tr = new TableRow();
            tc1 = new TableCell(CreateParagraph($"Итого"));
            tc2 = new TableCell(CreateParagraph($""));
            tc3 = new TableCell(CreateParagraph($""));
            if (duration != TimeSpan.Zero)
            {
                tc4 = new TableCell(CreateParagraph($"{duration}"));
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

        private TableRow CreateRowProtocolCandidates(Candidate candidate, Talon talon, string mediaresource, int i, string row5Text)
        {
            var tr = new TableRow();
            
            //
            if (candidate.Info.Фамилия == "") return null;
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
            //// Чтобы не разделялась при переходе на другую страницу
            //var rowProp = new TableRowProperties(new CantSplit());
            //tr.Append(rowProp);
            // 
            var tc1 = new TableCell(CreateParagraph($"{i + 1}"));
            var tc2 = new TableCell(CreateParagraph($"{candidate.Info.Фамилия} {candidate.Info.Имя} {candidate.Info.Отчество}, {candidate.Info.Округ_Номер}"));
            var tc3 = new TableCell(CreateParagraph($""));
            var tc4 = new TableCell(CreateParagraph(lines));
            //tc4.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "60" }));
            var tc5 = new TableCell(CreateParagraph($"{row5Text}"));
            var tc6 = new TableCell(CreateParagraph($""));
            tr.Append(tc1, tc2, tc3, tc4, tc5, tc6);
            //
            return tr;
        }

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
    }
}
