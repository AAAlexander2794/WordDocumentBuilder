using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.Entities
{
    /// <summary>
    /// Запись кандидата в текстовом виде (из Экселя).
    /// </summary>
    /// <remarks>Да, поля названы по-русски.</remarks>
    internal class CandidateInfo
    {
        /// <summary>
        /// Фамилия
        /// </summary>
        public string Фамилия { get; set; } = "";

        /// <summary>
        /// Имя
        /// </summary>
        public string Имя { get; set; } = "";

        /// <summary>
        /// Отчество
        /// </summary>
        public string Отчество { get; set; } = "";

        /// <summary>
        /// Номер талона
        /// </summary>
        public string Номер_талона { get; set; } = "";

        /// <summary>
        /// Номер договора
        /// </summary>
        public string Номер_договора { get; set; } = "";

        public string Дата_договора { get; set; } = "";

        /// <summary>
        /// Постановление ТИК. В формате "[дата] [номер]"
        /// </summary>
        public string Постановление_ТИК { get; set; } = "";

        public string ИНН { get; set; } = "";

        public string Спец_изб_счет_номер { get; set; } = "";

        public string Фамилия_представителя { get; set; } = "";

        public string Имя_представителя { get; set; } = "";

        public string Отчество_представителя { get; set; } = "";

        /// <summary>
        /// В формате "[номер] от [дата]"
        /// </summary>
        public string Доверенность_на_представителя { get; set; } = "";




    }
}
