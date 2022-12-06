using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.ElectionContracts
{
    internal partial class TalonBuilder
    {
        internal static class Variant1
        {
            internal static List<TalonRecord> BuildTalonRecords(DataTable dt, string mediaResource)
            {
                var talonRecordInfos = ParseTalonRecordInfos(dt, mediaResource);
                var talonRecords = BuildTalonRecords(talonRecordInfos);
                return talonRecords;
            }

            static List<TalonRecord> BuildTalonRecords(List<TalonRecordInfo> infos)
            {
                var talonRecords = new List<TalonRecord>();
                //
                foreach (var info in infos)
                {
                    talonRecords.Add(CreateTalonRecord(info));
                }
                return talonRecords;
            }

            static TalonRecord CreateTalonRecord(TalonRecordInfo info)
            {
                var talonRecord = new TalonRecord(
                    int.Parse(info.Id),
                    info.MediaResource,
                    DateOnly.FromDateTime(DateTime.Parse(info.Date)),
                    // Происходит замена точки на запятую (вот такая культура)
                    TimeOnly.FromDateTime(DateTime.Parse(info.Time.Replace('.', ','))),
                    TimeOnly.FromDateTime(DateTime.Parse(info.Duration.Replace('.', ','))).ToTimeSpan(),
                    info.Description
                    );
                return talonRecord;
            }

            /// <summary>
            /// Парсинг таблицы конкретного вида [номер талона] [все строки талона]. Все талоны относятся к одному медиаресурсу.
            /// </summary>
            /// <param name="dt"></param>
            /// <param name="mediaResource"></param>
            /// <returns></returns>
            static List<TalonRecordInfo> ParseTalonRecordInfos(DataTable dt, string mediaResource)
            {
                var result = new List<TalonRecordInfo>();
                // В одной ячейке все строки одного талона
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var talonId = dt.Rows[i].Field<string>(0);
                    // Ячейка со строками одного талона парсится в список строк одного талона
                    var talonRecords = ParseTalonString(talonId, mediaResource, dt.Rows[i].Field<string>(1));
                    // Все записи добавляем к результату
                    foreach (var talonRecord in talonRecords)
                    {
                        result.Add(talonRecord);
                    }
                }
                return result;
            }

            /// <summary>
            /// Парсинг текста конкретного вида в записи талона.
            /// </summary>
            /// <param name="id"></param>
            /// <param name="mediaResource"></param>
            /// <param name="talonString">Текст из ячейки со всеми записями одного талона</param>
            /// <returns>Строки одного талона указанного медиаресурса.</returns>
            private static List<TalonRecordInfo> ParseTalonString(string id, string mediaResource, string talonString)
            {
                var result = new List<TalonRecordInfo>();
                //
                char[] delimitersRow = { '\n', '\r' };
                string[] rows = talonString.Split(delimitersRow);
                //
                char[] delimeterColumn = { ' ' };
                foreach (string row in rows)
                {
                    string[] columns = row.Split(delimeterColumn);
                    //
                    try
                    {
                        result.Add(new TalonRecordInfo(
                        id,
                        mediaResource,
                        columns[0],
                        columns[1],
                        columns[2]));
                    }
                    catch { continue; }
                }
                return result;
            }
        }


        

    }
}
