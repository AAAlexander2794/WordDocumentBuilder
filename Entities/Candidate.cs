using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.Entities
{
    /// <summary>
    /// Полноценная строчка кандидата для записи в документ.
    /// </summary>
    internal class Candidate
    {
        /// <summary>
        /// Информация о кандидате, формируется на основе записи из Экселя.
        /// </summary>
        public CandidateInfo Info { get; }

        public Talon Talon { get; }

        public string И_О_Фамилия { get; }

        public string И_О_Фамилия_представителя { get; }

        public Candidate(CandidateInfo info, Talon talon)
        {
            Info = info;
            Talon = talon;
            И_О_Фамилия = $"{Info.Имя[0]}.{Info.Отчество[0]} {Info.Фамилия}";
            if (Info.Имя_представителя != "" & Info.Отчество_представителя != "" & Info.Фамилия_представителя != "")
            {
                И_О_Фамилия_представителя = $"{Info.Имя_представителя[0]}.{Info.Отчество_представителя[0]} {Info.Фамилия_представителя}";
            }
        }
    }
}
