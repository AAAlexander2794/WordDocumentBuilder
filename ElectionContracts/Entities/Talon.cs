﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    public class Talon
    {


        public int Id { get; set; }

        public string MediaResource { get; set; }

        public List<TalonRecord> TalonRecords { get; set; } = new List<TalonRecord>();

        public TimeSpan TotalDuration { get; set; } = TimeSpan.Zero;

        public Talon(int id, string mediaResource, List<TalonRecord> talonRecords)
        {
            Id = id;
            MediaResource = mediaResource;
            // Только записи с совпадающими Медиаресурсом и ID
            TalonRecords = talonRecords.Where(x => x.MediaResource == MediaResource && x.Id == Id).ToList();
            //
            foreach (TalonRecord record in TalonRecords)
            {
                TotalDuration += record.Duration;
            }
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

        //public DataTable GetTalonDataTable()
        //{
        //    DataTable dt = new DataTable();

        //    return dt;
        //}
    }
}
