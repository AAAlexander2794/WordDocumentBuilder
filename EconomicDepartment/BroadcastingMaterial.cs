using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    internal class BroadcastingMaterial
    {
        /// <summary>
        /// Форма предвыборной агитации
        /// </summary>
        public string FormName { get; set; } = "";

        /// <summary>
        /// Наименование теле-, радиоматериала
        /// </summary>
        public string Name { get; set; } = "";

        /// <summary>
        /// Дата и время выхода в эфир
        /// </summary>
        public DateTime DateTime { get; set; } = DateTime.MinValue;

        public string Факт_объем_времени { get; set; } = "";

        public BroadcastingMaterial()
        {

        }
    }
}
