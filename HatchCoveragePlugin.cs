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
using AcDb = Autodesk.AutoCAD.DatabaseServices;
using CivilDb = Autodesk.Civil.DatabaseServices;

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
        public void FillCoveragePropertiesFromCurrentDoc()
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
            int hatchCount = 0;
            int solidCount = 0;
            int surfaceCount = 0;

            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine("--- ЛОГ ОБРАБОТКИ ОБЪЕКТОВ ПОКРЫТИЙ ---");
            logBuilder.AppendLine($"Дата и время: {DateTime.Now}");
            logBuilder.AppendLine($"Файл Excel: {excelPath}");
            logBuilder.AppendLine($"Загружено строк из Excel (начиная с 3-й строки): {excelData.Count}\n");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                DictionaryPropertySetDefinitions psdDict = new DictionaryPropertySetDefinitions(db);

                foreach (ObjectId objId in ms)
                {
                    AcDb.DBObject dbObj = tr.GetObject(objId, OpenMode.ForRead);
                    if (!TryExtractTargetEntity(dbObj, out AcDb.Entity targetEntity, out string objectTypeLabel))
                        continue;

                    string layerName = (targetEntity.Layer ?? "").Trim();
                    if (!excelData.TryGetValue(layerName, out ExcelCoverageRow excelRow))
                        continue;

                    string targetPsdName = excelRow.GetTargetPsdName();

                    if (!psdDict.Has(targetPsdName, tr))
                    {
                        logBuilder.AppendLine($"[ОШИБКА] В чертеже отсутствует определение Набора Характеристик: {targetPsdName} (объект {objectTypeLabel}, слой: {layerName})");
                        continue;
                    }

                    targetEntity = (AcDb.Entity)tr.GetObject(objId, OpenMode.ForWrite);
                    ObjectId targetPsdId = psdDict.GetAt(targetPsdName);

                    bool applied = ApplyCoveragePropertySet(targetEntity, targetPsdId, psdDict, tr, excelRow, logBuilder);
                    if (applied)
                    {
                        processedCount++;
                        if (targetEntity is Hatch) hatchCount++;
                        else if (targetEntity is Solid3d) solidCount++;
                        else if (targetEntity is CivilDb.Surface) surfaceCount++;

                        logBuilder.AppendLine($"[УСПЕХ] Тип: {objectTypeLabel} | Handle: {targetEntity.Handle} | Слой: {layerName} -> Назначен PSD: {targetPsdName}");
                    }
                }

                tr.Commit();
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string logPath = Path.Combine(desktopPath, "Отработка_Покрытий_Лог.txt");
            File.WriteAllText(logPath, logBuilder.ToString(), Encoding.UTF8);

            ed.WriteMessage(
                $"\n[ГОТОВО] Обработано объектов: {processedCount} " +
                $"(Штриховки: {hatchCount}, 3D-тела: {solidCount}, Поверхности: {surfaceCount}). " +
                $"Подробный лог сохранён на Рабочий стол.");
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

        private bool TryExtractTargetEntity(AcDb.DBObject dbObj, out AcDb.Entity entity, out string typeLabel)
        {
            entity = null;
            typeLabel = null;

            if (dbObj is Hatch hatch)
            {
                entity = hatch;
                typeLabel = "Hatch";
                return true;
            }

            if (dbObj is Solid3d solid)
            {
                entity = solid;
                typeLabel = "Solid3d";
                return true;
            }

            if (dbObj is CivilDb.Surface surface)
            {
                entity = surface as AcDb.Entity;
                typeLabel = surface.GetType().Name;
                return entity != null;
            }

            return false;
        }

        private bool ApplyCoveragePropertySet(
            AcDb.Entity targetEntity,
            ObjectId targetPsdId,
            DictionaryPropertySetDefinitions psdDict,
            Transaction tr,
            ExcelCoverageRow excelRow,
            StringBuilder logBuilder)
        {
            foreach (string psdName in AllPsdNames)
            {
                if (psdDict.Has(psdName, tr))
                {
                    try
                    {
                        PropertyDataServices.RemovePropertySet(targetEntity, psdDict.GetAt(psdName));
                    }
                    catch (System.Exception ex)
                    {
                        logBuilder.AppendLine($"[ПРЕДУПРЕЖДЕНИЕ] Не удалось удалить НХ '{psdName}' с объекта {targetEntity.Handle}: {ex.Message}");
                    }
                }
            }

            PropertyDataServices.AddPropertySet(targetEntity, targetPsdId);

            ObjectId psId = ObjectId.Null;
            foreach (ObjectId id in PropertyDataServices.GetPropertySets(targetEntity))
            {
                var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                if (testPs != null && testPs.PropertySetDefinition == targetPsdId)
                {
                    psId = id;
                    break;
                }
            }

            if (psId.IsNull)
            {
                logBuilder.AppendLine($"[ОШИБКА] Не удалось получить созданный PropertySet для объекта {targetEntity.Handle}.");
                return false;
            }

            var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
            if (ps == null) return false;

            SetProperty(ps, "Название_покрытия", excelRow.CoverageId, logBuilder);
            SetProperty(ps, "Код_по_классификатору_зданий", excelRow.CodeClassifier, logBuilder);
            SetProperty(ps, "Код_по_классификатору_материалов", excelRow.CodeMaterial, logBuilder);
            SetProperty(ps, "Позиция_Тип", excelRow.PositionType, logBuilder);
            SetProperty(ps, "Обозначение", excelRow.Designation, logBuilder);
            SetProperty(ps, "Наименование", excelRow.Name, logBuilder);
            SetProperty(ps, "Пирог_покрытия_со_слоями", excelRow.PieLayers, logBuilder);
            SetProperty(ps, "Толщина_покрытия", excelRow.Thickness, logBuilder);

            if (!SetProperty(ps, "Тип_покрытия", excelRow.Type, logBuilder))
            {
                SetProperty(ps, "Тип", excelRow.Type, logBuilder);
            }

            return true;
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

                        // Начинаем с r = 2 (3-я строка Excel: 1-я — заголовки, 2-я — пример заполнения)
                        for (int r = 2; r < rowCount; r++)
                        {
                            string coverageId = table.Rows[r][0]?.ToString()?.Trim();
                            if (string.IsNullOrEmpty(coverageId)) continue;

                            if (coverageId.StartsWith("Пример заполнения", StringComparison.OrdinalIgnoreCase))
                                continue;

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

        private bool SetProperty(PropertySet ps, string propName, object val, StringBuilder logBuilder = null)
        {
            if (ps == null || val == null) return false;

            try
            {
                Transaction tr = ps.Database.TransactionManager.TopTransaction;
                if (tr == null) return false;

                var psd = tr.GetObject(ps.PropertySetDefinition, OpenMode.ForRead) as PropertySetDefinition;
                if (psd == null) return false;

                foreach (PropertyDefinition pd in psd.Definitions)
                {
                    if (pd.Automatic) continue;

                    if (pd.Name.Equals(propName, StringComparison.OrdinalIgnoreCase))
                    {
                        PropertySetData pData = null;
                        try
                        {
                            pData = ps.PropertySetData[pd.Id];
                        }
                        catch
                        {
                            for (int i = 0; i < ps.PropertySetData.Count; i++)
                            {
                                var item = ps.PropertySetData[i];
                                if (item != null && item.Id == pd.Id)
                                {
                                    pData = item;
                                    break;
                                }
                            }
                        }

                        if (pData == null)
                        {
                            logBuilder?.AppendLine($"  [WARN] Не найдены данные для свойства '{pd.Name}'");
                            return false;
                        }

                        pData.SetData(val.ToString());
                        return true;
                    }
                }

                logBuilder?.AppendLine($"  [WARN] В НХ '{psd.Name}' не найдено свойство '{propName}'");
            }
            catch (System.Exception ex)
            {
                logBuilder?.AppendLine($"  [ERR] Ошибка записи свойства '{propName}' ('{val}'): {ex.Message}");
            }

            return false;
        }
    }
}