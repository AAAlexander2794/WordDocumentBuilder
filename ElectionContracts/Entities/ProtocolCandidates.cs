using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    /// <summary>
    /// Хранит всю инфу, которая должна быть записана в протокол кандидатов (по СМИ и округам).
    /// </summary>
    public class ProtocolCandidates
    {
        /// <summary>
        /// Расширенное название медиаресурса для заголовка протокола.
        /// </summary>
        public string Наименование_СМИ { get; set; } = "";

        public string Округ { get; set; } = "";

        public List<Candidate> Candidates { get; set; } = new List<Candidate>();

        /// <summary>
        /// Фамилия И.О. члена избирательной комиссии.
        /// </summary>
        public string Изб_ком_Фамилия_ИО { get; set; } = "";

        /// <summary>
        /// И.О. Фамилия представителья организации телерадиовещания (директор Дон-ТР).
        /// </summary>
        public string СМИ_ИО_Фамилия { get; set; } = "";
    }
}
