using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.Entities
{
    /// <summary>
    /// Одна строка 
    /// </summary>
    internal class TalonRecord
    {
        public int Id { get; set; }

        public string MediaResource { get; set; }

        public DateOnly Date { get; set; }

        public TimeOnly Time { get; set; }

        public TimeSpan Duration { get; set; }

        public TalonRecord(int id, string mediaResource, DateOnly date, TimeOnly time, TimeSpan duration)
        {
            Id = id;
            MediaResource = mediaResource;
            Date = date;
            Time = time;
            Duration = duration;
        }

        public TalonRecord(TalonRecordInfo info)
        {
            Id = int.Parse(info.Id);
            MediaResource = info.MediaResource;
            Date = DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(info.Date)));
            // Происходит замена точки на запятую (вот такая культура)
            Time = TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(info.Time.Replace('.', ','))));
            Duration = TimeOnly.FromDateTime(DateTime.FromOADate(double.Parse(info.Duration.Replace('.', ',')))).ToTimeSpan();
        }
    }
}
