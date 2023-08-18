using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.EconomicDepartment
{
    public class Builder
    {
        public void BuildTable()
        {
            //
            var _protocolsFilePath = Settings.Default.Protocols_FilePath;
            //
            var builder = new ElectionContracts.Builder();
            // Настройки текущей жеребьевки
            ProtocolsInfo settings;
            try
            {
                settings = builder.ReadProtocols(_protocolsFilePath);
            }
            catch { throw new Exception("Не читает настройки протоколов."); }
            //
            var protocols = builder.BuildProtocolsCandidates(settings, "1");

            //
            DataTable dt = new DataTable();
            // test
            //dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_FilePath, sheetNumber: 0);
            //
            dt.Columns.Add("Канал");
            dt.Columns.Add("Дата");
            dt.Columns.Add("Отрезок");
            dt.Columns.Add("Хрон");
            dt.Columns.Add("Округ");
            dt.Columns.Add("Партия/кандидат");
            dt.Columns.Add("Название партии/ФИО кандидата");
            dt.Columns.Add("Факт время");
            dt.Columns.Add("Форма выступления");
            dt.Columns.Add("Название ролика/тема дебатов");
            //
            foreach (var protocol in protocols)
            {
                foreach (var c in protocol.Candidates)
                {
                    dt.Rows.Add();
                    dt.Rows[dt.Rows.Count - 1][0] = protocol.Наименование_СМИ;
                    dt.Rows[dt.Rows.Count - 1][1] = c.ИО_Фамилия;
                }
            }

            // Запись в файл Excel
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "WorksheetName");
            wb.SaveAs(@".\Настройки\Экономический отдел\excel.xlsx");
        }

        
    }
}
