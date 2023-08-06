using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentBuilder.ElectionContracts.Entities;

namespace WordDocumentBuilder.ElectionContracts
{
    internal partial class TalonBuilder
    {
        

        internal static List<Talon> BuildTalonsCandidates(string variant = "default")
        {
            DataTable dt;
            List<TalonRecord> talonRecords;
            List<Talon> talons = new List<Talon>();
            switch (variant)
            {
                case "1":
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath_Маяк, sheetNumber: 0);
                    talonRecords = Variant1.BuildTalonRecords(dt, "Маяк");
                    talons = BuildTalons(talonRecords, talons);
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath_Вести_ФМ, sheetNumber: 0);
                    talonRecords = Variant1.BuildTalonRecords(dt, "Вести ФМ");
                    talons = BuildTalons(talonRecords, talons);
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath_Радио_России, sheetNumber: 0);
                    talonRecords = Variant1.BuildTalonRecords(dt, "Радио России");
                    talons = BuildTalons(talonRecords, talons);
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath_Россия_1, sheetNumber: 0);
                    talonRecords = Variant1.BuildTalonRecords(dt, "Россия 1");
                    talons = BuildTalons(talonRecords, talons);
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath_Россия_24, sheetNumber: 0);
                    talonRecords = Variant1.BuildTalonRecords(dt, "Россия 24");
                    talons = BuildTalons(talonRecords, talons);
                    break;
                case "test":
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath, sheetNumber: 2);
                    talonRecords = Default.BuildTalonRecords(dt);
                    talons = BuildTalons(talonRecords);
                    break;
                default:
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Candidates_TalonsFilePath, sheetNumber: 0);
                    talonRecords = Default.BuildTalonRecords(dt);
                    talons = BuildTalons(talonRecords);
                    break;
            }
            return talons;
        }

        internal static List<Talon> BuildTalonsParties(string variant = "default")
        {
            DataTable dt;
            List<TalonRecord> talonRecords;
            List<Talon> talons = new List<Talon>();
            switch (variant)
            {
                case "1":
                    try
                    {
                        dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath_Маяк, sheetNumber: 0);
                        talonRecords = Variant1.BuildTalonRecords(dt, "Маяк");
                        talons = BuildTalons(talonRecords, talons);
                        dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath_Вести_ФМ, sheetNumber: 0);
                        talonRecords = Variant1.BuildTalonRecords(dt, "Вести ФМ");
                        talons = BuildTalons(talonRecords, talons);
                        dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath_Радио_России, sheetNumber: 0);
                        talonRecords = Variant1.BuildTalonRecords(dt, "Радио России");
                        talons = BuildTalons(talonRecords, talons);
                        dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath_Россия_1, sheetNumber: 0);
                        talonRecords = Variant1.BuildTalonRecords(dt, "Россия 1");
                        talons = BuildTalons(talonRecords, talons);
                        dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath_Россия_24, sheetNumber: 0);
                        talonRecords = Variant1.BuildTalonRecords(dt, "Россия 24");
                        talons = BuildTalons(talonRecords, talons);
                    }
                    catch(Exception ex)
                    {
                        throw ex;
                    }
                    break;
                case "test":
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath, sheetNumber: 2);
                    talonRecords = Default.BuildTalonRecords(dt);
                    talons = BuildTalons(talonRecords);
                    break;
                default:
                    dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath, sheetNumber: 0);
                    talonRecords = Default.BuildTalonRecords(dt);
                    talons = BuildTalons(talonRecords);
                    break;
            }
            return talons;
        }


        static List<Talon> BuildTalons(List<TalonRecord> talonRecords, List<Talon> talons = null)
        {
            if (talons == null) talons = new List<Talon>();
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
