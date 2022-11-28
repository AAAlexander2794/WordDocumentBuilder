using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.ElectionContracts
{
    internal partial class TalonBuilder
    {
        const string _dataFilepath = "data.xlsm";
        const string _talonsFilepath = "Талоны.xlsx";

        internal static List<Talon> BuildTalonsVariant1()
        {
            var dt = ExcelProcessor.ReadExcelSheet(_dataFilepath, sheetNumber: 1);
            var talonRecords = Variant1.BuildTalonRecords(dt);
            var talons = BuildTalons(talonRecords);
            return talons;
        }

        internal static List<Talon> BuildTalonsVariant2()
        {
            // todo: добавить обработку всех медиаресурсов
            var dt = ExcelProcessor.ReadExcelSheet(_talonsFilepath, sheetNumber: 1);
            List<TalonRecord> talonRecords = Variant2.BuildTalonRecords(dt, "Маяк");
            //
            var talons = BuildTalons(talonRecords);
            return talons;
        }

        static List<Talon> BuildTalons(List<TalonRecord> talonRecords)
        {
            var talons = new List<Talon>();
            // Берем по уникальным медиаресурсам
            var mediaresources = new List<string>();
            foreach (var record in talonRecords)
            {
                mediaresources.Add(record.MediaResource);
            }
            // Формируем список уникальных медиаресурсов
            var uniqMediaResources = mediaresources.Distinct().ToList();
            // Для каждого медиаресурса
            foreach (var mediaResource in uniqMediaResources)
            {
                // Выбираем все строчки для текущего медиаресурса
                var curMediaTalonRecords = talonRecords.Where(x => x.MediaResource == mediaResource).ToList();
                // Получаем уникальные ID талонов для этих строчек (по сути количество талонов)
                var ids = new List<int>();
                foreach (var rec in curMediaTalonRecords)
                {
                    ids.Add(rec.Id);
                }
                var uniqIds = ids.Distinct().ToList();
                //
                foreach (var id in uniqIds)
                {
                    // Создаем талон с этими записями
                    var talon = new Talon(id, mediaResource, talonRecords);
                    talons.Add(talon);
                }
            }
            return talons;
        }
    }
}
