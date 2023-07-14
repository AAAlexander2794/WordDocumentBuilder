using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordDocumentBuilder.ElectionContracts.Entities
{
    internal class Party
    {
        /// <summary>
        /// Информация о кандидате, формируется на основе записи из Экселя.
        /// </summary>
        public PartyInfo Info { get; }

        public Talon Талон_Маяк { get; }

        public Talon Талон_Радио_России { get; }

        public Talon Талон_Вести_ФМ { get; }

        public Talon Талон_Россия_1 { get; }

        public Talon Талон_Россия_24 { get; }

        public string Представитель_ИО_Фамилия { get; }

        public Party(PartyInfo info, List<Talon> talons)
        {
            Info = info;
            // Формируем ИО_Фамилия представителя для полей подписи
            if (Info.Представитель_Имя != "" & Info.Представитель_Отчество != "" & Info.Представитель_Фамилия != "")
            {
                Представитель_ИО_Фамилия = $"{Info.Представитель_Имя[0]}.{Info.Представитель_Отчество[0]}. {Info.Представитель_Фамилия}";
            }
            else
            {
                Представитель_ИО_Фамилия = "";
            }
            //
            Талон_Маяк = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Маяк && x.MediaResource == "Маяк");
            Талон_Радио_России = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Радио_России && x.MediaResource == "Радио России");
            Талон_Вести_ФМ = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Вести_ФМ && x.MediaResource == "Вести ФМ");
            Талон_Россия_1 = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Россия_1 && x.MediaResource == "Россия 1");
            Талон_Россия_24 = talons.FirstOrDefault(x => x.Id.ToString() == Info.Талон_Россия_24 && x.MediaResource == "Россия 24");
        }
    }
}
