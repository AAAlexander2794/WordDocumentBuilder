using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

        public ObservableCollection<ReportClientBlock> ClientBlocks { get; set; } = new ObservableCollection<ReportClientBlock>();

        public TimeSpan TotalDuration { get; set; }

        //public void AddDuration (TimeSpan timeSpan)
        //{
        //    TotalDuration += timeSpan;
        //}

        //public ReportRegionBlock()
        //{
        //    TotalDuration = TimeSpan.Zero;
        //    //
        //    ClientBlocks.CollectionChanged += new System.Collections.Specialized.NotifyCollectionChangedEventHandler(
        //        delegate (object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        //        {
        //            if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
        //            {
        //                foreach (var record in ClientBlocks)
        //                {
        //                    TotalDuration += record.TotalDuration;
        //                }
        //            }
        //        }
        //        );
        //}
    }
}
