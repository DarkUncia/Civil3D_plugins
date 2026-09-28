using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;
using ExcelDataReader;

[assembly: CommandClass(typeof(Civil3D_plugins.HatchCoveragePlugin))]

namespace Civil3D_plugins
{
    public class ExcelCoverageRow
    {
        public string CoverageId { get; set; }
        public string CodeClassifier { get; set; }
        public string CodeMaterial { get; set; }
        public string PositionType { get; set; }
        public string Area { get; set; }
        public string Designation { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string PieLayers { get; set; }
        public string Thickness { get; set; }

        public string GetTargetPsdName()
        {
            string t = (Type ?? "").Trim().ToLower();
            string n = (Name ?? "").Trim().ToLower();

            if (t.Contains("плитк") || t.Contains("брусчат") || n.Contains("плитк") || n.Contains("брусчат"))
                return "05_ДП_(плитка)";

            if (t.Contains("тверд") || t.Contains("асфальт") || n.Contains("асфальт") || n.Contains("бетон"))
                return "05_ДП_(твердые)";

            if (t.Contains("озел") || t.Contains("газон") || n.Contains("газон") || n.Contains("экорешет"))
                return "05_ДП_(озеленение)";

            if (t.Contains("мягк") || t.Contains("резин") || n.Contains("резин") || n.Contains("каучук"))
                return "05_ДП_(мягкие)";

            return "05_ДП_(иное)";
        }
    }

    public class HatchCoveragePlugin
    {
        private static readonly string[] AllPsdNames = new string[]
        {
            "05_ДП_(плитка)",
            "05_ДП_(твердые)",
            "05_ДП_(озеленение)",
            "05_ДП_(мягкие)",
            "05_ДП_(иное)"
        };

        [CommandMethod("COVERAGE_FILL_PROPERTIES")]
        public void FillHatchPropertiesFromCurrentDoc()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls",
                Title = "Выберите файл ведомости покрытий"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            string excelPath = openFileDialog.FileName;
            Dictionary<string, ExcelCoverageRow> excelData = LoadCoverageFromExcel(excelPath, ed);

            if (excelData == null || excelData.Count == 0)
            {
                ed.WriteMessage("\n[ОШИБКА] Не удалось загрузить данные из Excel или файл пуст.");
                return;
            }

            int processedCount = 0;
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"--- ЛОГ ОБРАБОТКИ ШТРИХОВОК ---");
            logBuilder.AppendLine($"Дата и время: {DateTime.Now}");
            logBuilder.AppendLine($"Файл Excel: {excelPath}");
            logBuilder.AppendLine($"Загружено строк из Excel: {excelData.Count}\n");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                DictionaryPropertySetDefinitions psdDict = new DictionaryPropertySetDefinitions(db);

                foreach (ObjectId objId in ms)
                {
                    Hatch hatch = tr.GetObject(objId, OpenMode.ForWrite) as Hatch;
                    if (hatch == null) continue;

                    string layerName = hatch.Layer.Trim();

                    if (excelData.TryGetValue(layerName, out ExcelCoverageRow excelRow))
                    {
                        string targetPsdName = excelRow.GetTargetPsdName();

                        if (!psdDict.Has(targetPsdName, tr))
                        {
                            logBuilder.AppendLine($"[ОШИБКА] В чертеже отсутствует определение Набора Характеристик: {targetPsdName}");
                            continue;
                        }

                        ObjectId targetPsdId = psdDict.GetAt(targetPsdName);

                        // Очистка старых НХ
                        foreach (string psdName in AllPsdNames)
                        {
                            if (psdDict.Has(psdName, tr))
                            {
                                try
                                {
                                    PropertyDataServices.RemovePropertySet(hatch, psdDict.GetAt(psdName));
                                }
                                catch { }
                            }
                        }

                        // Назначение актуального НХ
                        PropertyDataServices.AddPropertySet(hatch, targetPsdId);

                        ObjectId psId = ObjectId.Null;
                        foreach (ObjectId id in PropertyDataServices.GetPropertySets(hatch))
                        {
                            var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                            if (testPs != null && testPs.PropertySetDefinition == targetPsdId)
                            {
                                psId = id;
                                break;
                            }
                        }

                        if (!psId.IsNull)
                        {
                            var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                            if (ps != null)
                            {
                                SetProperty(ps, "Название_покрытия", excelRow.CoverageId);
                                SetProperty(ps, "Код_по_классификатору_зданий", excelRow.CodeClassifier);
                                SetProperty(ps, "Код_по_классификатору_материалов", excelRow.CodeMaterial);
                                SetProperty(ps, "Позиция_Тип", excelRow.PositionType);
                                SetProperty(ps, "Обозначение", excelRow.Designation);
                                SetProperty(ps, "Наименование", excelRow.Name);
                                SetProperty(ps, "Пирог_покрытия_со_слоями", excelRow.PieLayers);
                                SetProperty(ps, "Толщина_покрытия", excelRow.Thickness);

                                if (!SetProperty(ps, "Тип_покрытия", excelRow.Type))
                                {
                                    SetProperty(ps, "Тип", excelRow.Type);
                                }

                                processedCount++;
                                logBuilder.AppendLine($"[УСПЕХ] Handle: {hatch.Handle} | Слой: {layerName} -> Назначен PSD: {targetPsdName}");
                            }
                        }
                    }
                }

                tr.Commit();
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string logPath = Path.Combine(desktopPath, "Отработка_Покрытий_Лог.txt");
            File.WriteAllText(logPath, logBuilder.ToString(), Encoding.UTF8);

            ed.WriteMessage($"\n[ГОТОВО] Обработано штриховок: {processedCount}. Подробный лог сохранен на Рабочий стол.");
        }

        [CommandMethod("COVERAGE_MATCH_PROPERTIES")]
        public void MatchCoverageProperties()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nВыберите исходную штриховку с заполненным НХ:");
            peo.AddAllowedClass(typeof(Hatch), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Hatch sourceHatch = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Hatch;
                if (sourceHatch == null) return;

                PropertySet sourcePs = null;
                PropertySetDefinition sourcePsd = null;

                foreach (ObjectId psId in PropertyDataServices.GetPropertySets(sourceHatch))
                {
                    var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForRead);
                    if (ps != null)
                    {
                        var psd = (PropertySetDefinition)tr.GetObject(ps.PropertySetDefinition, OpenMode.ForRead);
                        if (psd != null && Array.IndexOf(AllPsdNames, psd.Name) >= 0)
                        {
                            sourcePs = ps;
                            sourcePsd = psd;
                            break;
                        }
                    }
                }

                if (sourcePs == null || sourcePsd == null)
                {
                    ed.WriteMessage("\n[ОШИБКА] На выбранной штриховке не найден Набор Характеристик серии 05_ДП_.");
                    return;
                }

                PromptSelectionOptions pso = new PromptSelectionOptions
                {
                    MessageForAdding = "\nВыберите целевые штриховки для копирования характеристик:"
                };
                TypedValue[] filterList = new TypedValue[] { new TypedValue((int)DxfCode.Start, "HATCH") };
                SelectionFilter filter = new SelectionFilter(filterList);

                PromptSelectionResult psr = ed.GetSelection(pso, filter);
                if (psr.Status != PromptStatus.OK) return;

                int count = 0;
                DictionaryPropertySetDefinitions psdDict = new DictionaryPropertySetDefinitions(db);

                foreach (SelectedObject so in psr.Value)
                {
                    if (so == null) continue;
                    Hatch targetHatch = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Hatch;
                    if (targetHatch == null) continue;

                    foreach (string psdName in AllPsdNames)
                    {
                        if (psdDict.Has(psdName, tr))
                        {
                            try { PropertyDataServices.RemovePropertySet(targetHatch, psdDict.GetAt(psdName)); } catch { }
                        }
                    }

                    PropertyDataServices.AddPropertySet(targetHatch, sourcePsd.ObjectId);

                    foreach (ObjectId psId in PropertyDataServices.GetPropertySets(targetHatch))
                    {
                        var targetPs = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                        if (targetPs != null && targetPs.PropertySetDefinition == sourcePsd.ObjectId)
                        {
                            foreach (PropertyDefinition pd in sourcePsd.Definitions)
                            {
                                if (!pd.Automatic)
                                {
                                    try
                                    {
                                        int propId = pd.Id;
                                        PropertySetData pData = sourcePs.PropertySetData[propId];
                                        if (pData != null)
                                        {
                                            PropertySetData targetData = targetPs.PropertySetData[propId];
                                            if (targetData != null)
                                            {
                                                targetData.SetData(pData.GetData());
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            count++;
                            break;
                        }
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n[УСПЕХ] Характеристики успешно скопированы на {count} штриховок.");
            }
        }

        private Dictionary<string, ExcelCoverageRow> LoadCoverageFromExcel(string filePath, Editor ed)
        {
            var result = new Dictionary<string, ExcelCoverageRow>(StringComparer.OrdinalIgnoreCase);

            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var resultDataSet = reader.AsDataSet();
                        if (resultDataSet.Tables.Count == 0) return result;

                        var table = resultDataSet.Tables[0];
                        int rowCount = table.Rows.Count;

                        for (int r = 1; r < rowCount; r++)
                        {
                            string coverageId = table.Rows[r][0]?.ToString()?.Trim();
                            if (string.IsNullOrEmpty(coverageId)) continue;

                            var row = new ExcelCoverageRow
                            {
                                CoverageId = coverageId,
                                CodeClassifier = table.Rows[r][1]?.ToString()?.Trim(),
                                CodeMaterial = table.Rows[r][2]?.ToString()?.Trim(),
                                PositionType = table.Rows[r][3]?.ToString()?.Trim(),
                                Area = table.Rows[r][4]?.ToString()?.Trim(),
                                Designation = table.Rows[r][5]?.ToString()?.Trim(),
                                Name = table.Rows[r][6]?.ToString()?.Trim(),
                                Type = table.Rows[r][7]?.ToString()?.Trim(),
                                PieLayers = table.Rows[r][8]?.ToString()?.Trim(),
                                Thickness = table.Rows[r][9]?.ToString()?.Trim()
                            };

                            if (!result.ContainsKey(coverageId))
                            {
                                result.Add(coverageId, row);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[ИСКЛЮЧЕНИЕ] При чтении Excel: {ex.Message}");
            }

            return result;
        }

        private bool SetProperty(PropertySet ps, string propName, object val)
        {
            if (ps == null || val == null) return false;
            try
            {
                Transaction tr = ps.Database.TransactionManager.TopTransaction;
                var psd = (PropertySetDefinition)tr.GetObject(ps.PropertySetDefinition, OpenMode.ForRead);
                if (psd == null) return false;

                foreach (PropertyDefinition pd in psd.Definitions)
                {
                    if (pd.Name.Equals(propName, StringComparison.OrdinalIgnoreCase))
                    {
                        PropertySetData pData = ps.PropertySetData[pd.Id];
                        if (pData != null)
                        {
                            pData.SetData(val.ToString());
                            return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }
    }
}