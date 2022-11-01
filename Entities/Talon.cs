using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.Entities
{
    internal class Talon
    {


        public int Id { get; set; }

        public string MediaResource { get; set; }

        public List<TalonRecord> TalonRecords { get; set; } = new List<TalonRecord>();

        public Talon(int id, string mediaResource, List<TalonRecord> talonRecords)
        {
            Id = id;
            MediaResource = mediaResource;
            // Только записи с совпадающими Медиаресурсом и ID
            TalonRecords = talonRecords.Where(x => x.MediaResource == MediaResource && x.Id == Id).ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>Информация о талоне в текстовом виде.</returns>
        public string GetTalonText()
        {
            var text = "";
            foreach (var rec in TalonRecords)
            {
                text += $"{rec.Id} {rec.MediaResource} {rec.Date} {rec.Time} {rec.Duration}<br/>";
            }
            return text;
        }
    }
}
