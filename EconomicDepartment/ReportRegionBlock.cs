using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    /// <summary>
    /// Блок отчета, содержащий один округ
    /// </summary>
    public class ReportRegionBlock
    {
        public string RegionNumber { get; set; } = "";

        public string RegionCaption { get; set; } = "";

        public List<ReportClientBlock> ClientBlocks { get; set; } = new List<ReportClientBlock>();
    }
}
