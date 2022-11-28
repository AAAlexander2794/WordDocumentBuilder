using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    /// <summary>
    /// Полноценная строчка кандидата для записи в документ.
    /// </summary>
    /// <remarks>
    /// У каждого кандидата по 5 талонов.
    /// </remarks>
    internal class Candidate
    {
        /// <summary>
        /// Информация о кандидате, формируется на основе записи из Экселя.
        /// </summary>
        public CandidateInfo Info { get; }

        public Talon Талон_Маяк { get; }

        public Talon Талон_Радио_России { get; }

        public Talon Талон_Вести_ФМ { get; }

        public Talon Талон_Россия_1 { get; }

        public Talon Талон_Россия_24 { get; }

        public string ИО_Фамилия { get; }

        public string ИО_Фамилия_представителя { get; }

        public string Округ_для_создания_каталога { get; }

        public Candidate(CandidateInfo info, List<Talon> talons)
        {
            Info = info;
            ИО_Фамилия = $"{Info?.Имя[0]}.{Info?.Отчество[0]}. {Info?.Фамилия}";
            if (Info.Имя_представителя != "" & Info.Отчество_представителя != "" & Info.Фамилия_представителя != "")
            {
                ИО_Фамилия_представителя = $"{Info.Имя_представителя[0]}.{Info.Отчество_представителя[0]}. {Info.Фамилия_представителя}";
            }
            else
            {
                ИО_Фамилия_представителя = "";
            }
            //
            Талон_Маяк = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Маяк && x.MediaResource == "Маяк");
            Талон_Радио_России = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Радио_России && x.MediaResource == "Радио России");
            Талон_Вести_ФМ = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Вести_ФМ && x.MediaResource == "Вести ФМ");
            Талон_Россия_1 = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Россия_1 && x.MediaResource == "Россия 1");
            Талон_Россия_24 = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Россия_24 && x.MediaResource == "Россия 24");
            //
            Regex rgx = new Regex("[^a-zA-Zа-яА-Я0-9 -]");
            Округ_для_создания_каталога = rgx.Replace(Info.Округ_дат_падеж, "");
        }
    }
}
