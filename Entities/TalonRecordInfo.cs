using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.Entities
{
    /// <summary>
    /// Одна запись талона в первозданном текстовом виде (из Экселя).
    /// </summary>
    internal class TalonRecordInfo
    {
        /// <summary>
        /// No Талона
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Медиаресурс
        /// </summary>
        public string MediaResource { get; set; }

        /// <summary>
        /// Дата
        /// </summary>
        public string Date { get; set; }

        /// <summary>
        /// Время
        /// </summary>
        public string Time { get; set; }

        /// <summary>
        /// Хронометраж
        /// </summary>
        public string Duration { get; set; }

        public string Description { get;set; }

        public TalonRecordInfo(string id, string mediaResource, string date, string time, string duration, string description = "")
        {
            Id = id;
            MediaResource = mediaResource;
            Date = date;
            Time = time;
            Duration = duration;
            Description = description;
        }

        //public TalonRecord(object? id, object? mediaResource, object? date, object? time, object? duration)
        //{
        //    Id = int.Parse(id.ToString());
        //    MediaResource = (string)mediaResource;

        //    Date = DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(date.ToString())));
        //    //var timeString = time.ToString();
        //    //var timeDouble = double.Parse(timeString);
        //    //var timeDateTime = DateTime.FromOADate(timeDouble);
        //    Time = TimeOnly.FromDateTime(DateTime.Now);
        //    Duration = TimeSpan.FromSeconds(10);
        //}
    }
}
