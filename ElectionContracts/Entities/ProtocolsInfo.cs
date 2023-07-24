using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    public class ProtocolsInfo
    {
        public string Префикс_партии { get; set; } = "";

        public string Фамилия_ИО_члена_изб_ком { get; set; } = "";

        public string ИО_Фамилия_члена_изб_ком { get; set; } = "";

        public string ИО_Фамилия_предст_СМИ { get; set; } = "";

        public string Наименование_СМИ_Маяк { get; set; } = "";

        public string Наименование_СМИ_Вести_ФМ { get; set; } = "";

        public string Наименование_СМИ_Радио_России { get; set; } = "";

        public string Наименование_СМИ_Россия_1 { get; set; } = "";

        public string Наименование_СМИ_Россия_24 { get; set; } = "";

        public string Дата { get; set; } = "";
    }
}
