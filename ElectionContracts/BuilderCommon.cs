using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
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
    /// Общие для всех построений методы
    /// </summary>
    /// <remarks>Конечно надо потом их куда-то перенести, чтобы не держать отдельный файл</remarks>
    public partial class Builder
    {
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

        /// <summary>
        /// Создает новый абзац текста
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        Paragraph CreateParagraph(List<string> lines)
        {
            var paragraph = new Paragraph();
            var run = new Run();
            // Добавляем без лишнего переноса на новую строку в конце
            for (int i = 0; i < lines.Count - 1; i++)
            {
                run.AppendChild(new Text(lines[i]));
                run.AppendChild(new Break());
            }
            run.AppendChild(new Text(lines[lines.Count - 1]));
            //
            RunProperties runProperties = new RunProperties();
            FontSize size = new FontSize();
            size.Val = StringValue.FromString("18");
            runProperties.Append(size);
            //
            run.Append(runProperties);
            paragraph.Append(run);
            //
            return paragraph;
        }

    }
    
}
