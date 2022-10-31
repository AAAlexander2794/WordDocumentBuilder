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

        public List<TalonRecord> TalonRecords { get; set; } = new List<TalonRecord>();

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
