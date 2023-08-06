using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
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

        public string GetTalonRecordString()
        {
            return $"{Id} {MediaResource} {Date} {Time} {Duration} {Description}";
        }
    }
}
