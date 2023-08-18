using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    public class Candidate
    {
        public string Фамилия { get; set; } = "";

        public string Имя { get; set; } = "";

        public string Отчество { get; set; } = "";

        public string ФИО { get; set; } = "";

        

        public string Договор_дата_номер { get; set; } = "";

        //public List<>

        public Candidate(string фамилия, string имя, string отчество)
        {
            Фамилия = фамилия;
            Имя = имя;
            Отчество = отчество;
            ФИО = $"{Фамилия}";
            if (Имя.Length > 0 && Отчество.Length > 0)
            {
                ФИО += $" {Имя[0]} {Отчество[0]}";
            }
        }
    }
}
