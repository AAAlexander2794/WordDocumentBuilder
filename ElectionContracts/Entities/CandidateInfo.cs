using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    /// <summary>
    /// Запись кандидата в текстовом виде (из Экселя).
    /// </summary>
    /// <remarks>Да, поля названы по-русски.</remarks>
    internal class CandidateInfo
    {
        /// <summary>
        /// Поле, где отмечается, создавать договор на этого кандидата или нет
        /// </summary>
        public string На_печать { get; set; } = "";

        // ФИО кандидата

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

        // Округ

        public string Округ_Номер { get; set; } = "";

        public string Округ_Название_падеж_дат { get; set; } = "";

        public string Округ_Название_падеж_им { get; set; } = "";

        // Для договора

        /// <summary>
        /// Постановление ТИК. В формате "[дата] [номер]"
        /// </summary>
        public string Постановление { get; set; } = "";

        /// <summary>
        /// Номер договора
        /// </summary>
        public string Номер_договора { get; set; } = "";

        private string _contractDate = "\"__\" __________ 20__ ";
        public string Дата_договора
        {
            get { return _contractDate; }
            set { if (value != "") _contractDate = value; }
        }

        // Талоны

        public string Талон_Маяк { get; set; } = "";

        public string Талон_Радио_России { get; set; } = "";

        public string Талон_Вести_ФМ { get; set; } = "";

        public string Талон_Россия_1 { get; set; } = "";

        public string Талон_Россия_24 { get; set; } = "";

        // Явка

        public string Явка_кандидата { get; set; } = "";

        public string Явка_представителя { get; set; } = "";

        //

        public string Партия { get; set; } = "";

        public string ИНН { get; set; } = "";

        public string Спец_изб_счет_номер { get; set; } = "";

        // Представитель

        public string Представитель_Фамилия { get; set; } = "";

        public string Представитель_Имя { get; set; } = "";

        public string Представитель_Отчество { get; set; } = "";

        /// <summary>
        /// В формате "[номер] от [дата]"
        /// </summary>
        public string Представитель_Доверенность { get; set; } = "";

    }
}
