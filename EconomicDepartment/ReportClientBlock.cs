using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentBuilder.EconomicDepartment
{
    /// <summary>
    /// Блок отчета, содержащий одного клиента
    /// </summary>
    internal class ReportClientBlock
    {
        /// <summary>
        /// Ф.И.О. клиента (или название партии)
        /// </summary>
        public string ClientName { get; set; } = "";

        /// <summary>
        /// Основание предоставления (дата заключения и номер договора)
        /// </summary>
        public string ClientContract { get; set; } = "";

        /// <summary>
        /// Записи о вещании
        /// </summary>
        ObservableCollection<BroadcastRecord> BroadcastRecords { get; set; } = new ObservableCollection<BroadcastRecord>();

        /// <summary>
        /// Суммарное время, предоставленное клиенту
        /// </summary>
        public TimeSpan TotalDuration { get; set; }

        public ReportClientBlock()
        {
            BroadcastRecords.CollectionChanged += new System.Collections.Specialized.NotifyCollectionChangedEventHandler(
                delegate (object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
                {
                    if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
                    {
                        TotalDuration = TimeSpan.Zero;
                        foreach (var record in BroadcastRecords)
                        {
                            TotalDuration += record.DurationActual;
                        }
                    }
                }
            );
        }
    }

    
}
