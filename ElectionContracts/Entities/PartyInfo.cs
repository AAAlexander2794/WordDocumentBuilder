using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    /// <summary>
    /// Запись партии в текстовом виде (из Экселя)
    /// </summary>
    internal class PartyInfo
    {
        /// <summary>
        /// Поле, где отмечается, создавать договор на эту партию или нет
        /// </summary>
        public string На_печать { get; set; } = "";

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

        /// <summary>
        /// Указание, отделение региональное или федеральное, которое идет перед текстом "Политическая(ой) партия(и)"
        /// </summary>
        public string Партия_Отделение { get; set; } = "";

        public string Партия_Название { get; set; } = "";     

        /// <summary>
        /// Постановление избирательной комиссии. В формате "[дата] [номер]"
        /// </summary>
        public string Постановление { get; set; } = "";

        public string Представитель_Фамилия { get; set; } = "";

        public string Представитель_Имя { get; set; } = "";

        public string Представитель_Отчество { get; set; } = "";

        /// <summary>
        /// В формате "[номер] [дата]"
        /// </summary>
        public string Представитель_Доверенность { get; set; } = "";

        public string Нотариус_Фамилия { get; set; } = "";

        public string Нотариус_Имя { get; set; } = "";

        public string Нотариус_Отчество { get; set; } = "";

        public string Нотариус_Город { get; set; } = "";

        /// <summary>
        /// Номер реестра
        /// </summary>
        public string Нотариус_Реестр { get; set; } = "";

        public string ОГРН { get; set; } = "";

        public string ИНН { get; set; } = "";

        public string КПП { get; set; } = "";

        public string Спец_изб_счет_номер { get; set; } = "";

        public string Место_нахождения { get; set; } = "";

        // Талон Маяк
        public string Талон_Маяк { get; set; } = "";

        public string Талон_Радио_России { get; set; } = "";

        public string Талон_Вести_ФМ { get; set; } = "";

        public string Талон_Россия_1 { get; set; } = "";

        public string Талон_Россия_24 { get; set; } = "";

        public string Явка_представителя { get; set; } = "";
    }
}
