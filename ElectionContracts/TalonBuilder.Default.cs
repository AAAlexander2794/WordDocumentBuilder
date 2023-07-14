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
        
        /// <summary>
        /// Базовый вариант парсинга талонов. Все талоны вместе в одной таблице Эксель
        /// </summary>
        /// <remarks>
        /// 14.07.2023: Такие талоны никто не делает пока, так что пользуем Variant1.
        /// </remarks>
        internal static class Default
        {
            internal static List<TalonRecord> BuildTalonRecords(DataTable dt)
            {
                //
                var talonRecordInfos = ReadTalonRecords(dt);
                //
                var talonRecords = BuildTalonRecords(talonRecordInfos);
                //
                return talonRecords;
            }

            /// <summary>
            /// Читаем таблицу с строками талонов
            /// </summary>
            /// <remarks>
            /// Там в кучу талоны разных медиаресурсов, то есть могут совпадать ID.
            /// </remarks>
            /// <returns></returns>
            static List<TalonRecordInfo> ReadTalonRecords(DataTable dt)
            {
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

            static List<TalonRecord> BuildTalonRecords(List<TalonRecordInfo> infos)
            {
                var records = new List<TalonRecord>();
                foreach (var inf in infos)
                {
                    records.Add(CreateTalonRecordFromExcelFormat(inf));
                }
                return records;
            }

            /// <summary>
            /// Парсинг строки в формате экселя.
            /// </summary>
            /// <param name="info"></param>
            /// <returns></returns>
            static TalonRecord CreateTalonRecordFromExcelFormat(TalonRecordInfo info)
            {
                var talonRecord = new TalonRecord(
                    int.Parse(info.Id),
                    info.MediaResource,
                    DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(info.Date))),
                    // Происходит замена точки на запятую (вот такая культура)
                    TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(info.Time.Replace('.', ',')))),
                    TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(info.Duration.Replace('.', ',')))).ToTimeSpan(),
                    info.Description
                    );
                return talonRecord;
            }

            
        }
        
    }
}
