using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
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

        public string Description { get; set; }

        public TalonRecord(int id, string mediaResource, DateOnly date, TimeOnly time, TimeSpan duration, string description)
        {
            Id = id;
            MediaResource = mediaResource;
            Date = date;
            Time = time;
            Duration = duration;
            Description = description;
        }

        
    }
}
