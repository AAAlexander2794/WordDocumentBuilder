using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    public class ProtocolsInfo
    {
        // Партии

        public string Префикс_партии { get; set; } = "";

        public string Партии_Фамилия_ИО_члена_изб_ком { get; set; } = "";

        public string Партии_ИО_Фамилия_члена_изб_ком { get; set; } = "";

        public string Партии_ИО_Фамилия_предст_СМИ { get; set; } = "";

        public string Партии_Дата { get; set; } = "";

        // Кандидаты

        public string Кандидаты_Фамилия_ИО_члена_изб_ком { get; set; } = "";

        public string Кандидаты_ИО_Фамилия_члена_изб_ком { get; set; } = "";

        public string Кандидаты_ИО_Фамилия_предст_СМИ { get; set; } = "";

        public string Кандидаты_Дата { get; set; } = "";

        // Общее

        public string Наименование_СМИ_Маяк { get; set; } = "";

        public string Наименование_СМИ_Вести_ФМ { get; set; } = "";

        public string Наименование_СМИ_Радио_России { get; set; } = "";

        public string Наименование_СМИ_Россия_1 { get; set; } = "";

        public string Наименование_СМИ_Россия_24 { get; set; } = "";
    }
}
