using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    public class BroadcastRecord
    {
        /// <summary>
        /// Канал
        /// </summary>
        public string MediaResource { get; set; } = "";

        /// <summary>
        /// Дата
        /// </summary>
        public DateOnly Date { get; set; } = DateOnly.MinValue;

        /// <summary>
        /// Отрезок
        /// </summary>
        public TimeOnly Time { get; set; } = TimeOnly.MinValue;

        /// <summary>
        /// Хронометраж вещания номинальный
        /// </summary>
        public TimeSpan DurationNominal { get; set; } = TimeSpan.Zero;

        /// <summary>
        /// Номер округа
        /// </summary>
        public string RegionNumber { get; set; } = "";

        /// <summary>
        /// Партия/кандидат
        /// </summary>
        public string ClientType { get; set; } = "";

        /// <summary>
        /// Название партии/ФИО кандидата
        /// </summary>
        public string ClientName { get; set; } = "";

        /// <summary>
        /// Хронометраж вещания фактический
        /// </summary>
        public TimeSpan DurationActual { get; set; } = TimeSpan.Zero;

        /// <summary>
        /// Форма выступления
        /// </summary>
        public string BroadcastType { get; set; } = "";

        /// <summary>
        /// Название ролика/тема дебатов
        /// </summary>
        public string BroadcastCaption { get; set; } = "";
    }
}
