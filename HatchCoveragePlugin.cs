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

[assembly: CommandClass(typeof(Civil3D_plugins.HatchCoveragePlugin))]

namespace Civil3D_plugins
{
    // Класс-модель для хранения данных структуры пирога покрытия из Excel
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

        // Метод автоматического определения имени целевого набора характеристик
        public string GetTargetPsdName()
        {
            string t = Type?.ToLower() ?? "";
            string n = Name?.ToLower() ?? "";

            // 1. Проверяем жесткие типы покрытий в первую очередь (защита от ложного срабатывания асфальта)
            if (t.Contains("плитка")) return "05_ДП_(плитка)";
            if (n.Contains("асфальт") || t.Contains("асфальт")) return "05_ДП_(твердые)";
            if (t.Contains("газон") || t.Contains("решетки") || t.Contains("озеленение")) return "05_ДП_(озеленение)";
            if (t.Contains("резин")) return "05_ДП_(мягкие)";

            // 2. И только если не подошли главные типы — проверяем категорию "Иное"
            if (t.Contains("мульча") || n.Contains("мульча") || t.Contains("сыпучее") || n.Contains("сыпучее") ||
                t.Contains("цветники") || n.Contains("цветники") || t.Contains("настил") || n.Contains("настил") ||
                t.Contains("терравей") || n.Contains("терравей") || t.Contains("щепа") || n.Contains("щепа"))
            {
                return "05_ДП_(иное)";
            }

            return "05_ДП_(иное)";
        }
    }
    public class ExcelDataReaderHelper
    {
        public static Dictionary<string, ExcelCoverageRow> ReadExcelData(string filePath, Editor ed)
        {
            var result = new Dictionary<string, ExcelCoverageRow>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    if (!reader.Read()) return null;
                    int idCol = -1, codeClCol = -1, codeMatCol = -1, posCol = -1, areaCol = -1;
                    int desCol = -1, nameCol = -1, typeCol = -1, pieCol = -1, thickCol = -1;

                    for (int col = 0; col < reader.FieldCount; col++)
                    {
                        string h = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? "";
                        h = h.Replace("\r", "").Replace("\n", "");

                        if (h.Contains("покрыт") && (h.Contains("назван") || h.Contains("имя"))) idCol = col;
                        else if (h.Contains("классиф") && h.Contains("здан")) codeClCol = col;
                        else if (h.Contains("классиф") && h.Contains("матер")) codeMatCol = col;
                        else if (h.Contains("позиц") || h.Contains("тип")) { if (posCol == -1) posCol = col; }
                        else if (h.Contains("площад")) areaCol = col;
                        else if (h.Contains("обознач")) desCol = col;
                        else if (h.Contains("наименов")) nameCol = col;
                        else if (h.Contains("тип") && h.Contains("покрыт")) typeCol = col;
                        else if (h.Contains("пирог")) pieCol = col;
                        else if (h.Contains("толщин")) thickCol = col;
                    }

                    if (idCol == -1) idCol = 0;
                    if (codeClCol == -1) codeClCol = 1;
                    if (codeMatCol == -1) codeMatCol = 2;
                    if (posCol == -1) posCol = 3;
                    if (areaCol == -1) areaCol = 4;
                    if (desCol == -1) desCol = 5;
                    if (nameCol == -1) nameCol = 6;
                    if (typeCol == -1) typeCol = 7;
                    if (pieCol == -1) pieCol = 8;
                    if (thickCol == -1) thickCol = 9;

                    reader.Read();

                    while (reader.Read())
                    {
                        string id = reader.GetValue(idCol)?.ToString()?.Trim() ?? "";
                        if (string.IsNullOrEmpty(id) || id.Contains("Пример") || id.Contains("Название")) continue;

                        var row = new ExcelCoverageRow
                        {
                            CoverageId = id,
                            CodeClassifier = codeClCol != -1 ? reader.GetValue(codeClCol)?.ToString() ?? "" : "",
                            CodeMaterial = codeMatCol != -1 ? reader.GetValue(codeMatCol)?.ToString() ?? "" : "",
                            PositionType = posCol != -1 ? reader.GetValue(posCol)?.ToString() ?? "" : "",
                            Area = areaCol != -1 ? reader.GetValue(areaCol)?.ToString() ?? "" : "",
                            Designation = desCol != -1 ? reader.GetValue(desCol)?.ToString() ?? "" : "",
                            Name = nameCol != -1 ? reader.GetValue(nameCol)?.ToString() ?? "" : "",
                            Type = typeCol != -1 ? reader.GetValue(typeCol)?.ToString() ?? "" : "",
                            PieLayers = pieCol != -1 ? reader.GetValue(pieCol)?.ToString() ?? "" : "",
                            Thickness = thickCol != -1 ? reader.GetValue(thickCol)?.ToString() ?? "" : ""
                        };
                        if (!result.ContainsKey(id)) result.Add(id, row);
                    }
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка Excel]: {ex.Message}"); return null; }
            return result;
        }
    }
    public class HatchCoveragePlugin : IExtensionApplication
    {
        private static readonly string[] AllPsdNames = { "05_ДП_(плитка)", "05_ДП_(озеленение)", "05_ДП_(мягкие)", "05_ДП_(твердые)", "05_ДП_(иное)" };

        public void Initialize()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
            if (ed != null) ed.WriteMessage("\n>>> Плагин штриховок ГП загружен! Команды: COVERAGE_FILL_PROPERTIES, COVERAGE_MATCH_PROPERTIES <<<");
        }

        public void Terminate() { }

        [CommandMethod("COVERAGE_FILL_PROPERTIES", CommandFlags.Modal)]
        public void FillHatchPropertiesFromCurrentDoc()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor;
            Database db = activeDoc.Database;

            string excelFilePath = "";
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel (*.xlsx)|*.xlsx";
                ofd.Title = "Выберите Excel-файл покрытий ГП";
                if (ofd.ShowDialog() != DialogResult.OK) return;
                excelFilePath = ofd.FileName;
            }

            var excelData = ExcelDataReaderHelper.ReadExcelData(excelFilePath, ed);
            if (excelData == null || excelData.Count == 0) return;

            string dwgDir = db.Filename != null && File.Exists(db.Filename) ? Path.GetDirectoryName(db.Filename) : Path.GetTempPath();
            string logFilePath = Path.Combine(dwgDir, "Отработка_Покрытий_Лог.txt");
            StringBuilder log = new StringBuilder($"=== ОТЧЕТ {DateTime.Now} ===\nФайл: {Path.GetFileName(db.Filename)}\n\n");

            int processedCount = 0;
            try
            {
                using (DocumentLock docLock = activeDoc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var psdDict = new DictionaryPropertySetDefinitions(db);
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId objId in ms)
                    {
                        if (objId.ObjectClass.Name == "AcDbHatch")
                        {
                            var hatch = (Hatch)tr.GetObject(objId, OpenMode.ForWrite);
                            if (hatch == null) continue;

                            string layerName = hatch.Layer.Trim();
                            if (excelData.TryGetValue(layerName, out ExcelCoverageRow excelRow))
                            {
                                string targetPsdName = excelRow.GetTargetPsdName();
                                if (!psdDict.Has(targetPsdName, tr)) continue;

                                ObjectId targetPsdId = psdDict.GetAt(targetPsdName);

                                // Полная зачистка старых наборов из нашей группы перед записью нового
                                foreach (string psdName in AllPsdNames)
                                {
                                    if (psdDict.Has(psdName, tr))
                                    {
                                        try { PropertyDataServices.RemovePropertySet(hatch, psdDict.GetAt(psdName)); } catch { }
                                    }
                                }

                                // Добавление актуального НХ на чистый объект
                                PropertyDataServices.AddPropertySet(hatch, targetPsdId);

                                ObjectId psId = ObjectId.Null;
                                foreach (ObjectId id in PropertyDataServices.GetPropertySets(hatch))
                                {
                                    var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                    if (testPs != null && testPs.PropertySetDefinition == targetPsdId) { psId = id; break; }
                                }

                                // Заполнение полей свойствами из таблицы Excel
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

                                        hatch.RecordGraphicsModified(true);
                                        processedCount++;
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                ed.WriteMessage($"\n[Успех] Готово! Наполнено штриховок: {processedCount}.");
                log.AppendLine($"[УСПЕХ] Обработано штриховок: {processedCount}");
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка]: {ex.Message}"); log.AppendLine($"[СБОЙ]: {ex.Message}"); }

            try { File.WriteAllText(logFilePath, log.ToString(), Encoding.UTF8); } catch { }
            ed.UpdateScreen();
        }

        // =================================================================================
        // НЕЗАВИСИМАЯ КОМАНДА: BIM-КИСТОЧКА ДЛЯ КОПИРОВАНИЯ НХ МЕЖДУ ШТРИХОВКАМИ
        // =================================================================================
        [CommandMethod("COVERAGE_MATCH_PROPERTIES", CommandFlags.Modal)]
        public void MatchCoverageProperties()
        {
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor;
            Database db = activeDoc.Database;

            // 1. Выбор ИСХОДНОЙ штриховки-эталона
            PromptEntityOptions optSource = new PromptEntityOptions("\nВыберите ИСХОДНУЮ эталонную штриховку (с которой копируем НХ):");
            optSource.AddAllowedClass(typeof(Hatch), true);
            optSource.SetRejectMessage("\nМожно выбирать только объекты Штриховки (Hatch)!");
            PromptEntityResult resSource = ed.GetEntity(optSource);
            if (resSource.Status != PromptStatus.OK) return;

            // 2. Выбор ЦЕЛЕВЫХ штриховок для копирования
            PromptSelectionOptions optDest = new PromptSelectionOptions();
            optDest.MessageForAdding = "\nВыберите ЦЕЛЕВЫЕ штриховки для вставки характеристик:";

            SelectionFilter filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "HATCH") });
            PromptSelectionResult resDest = ed.GetSelection(optDest, filter);
            if (resDest.Status != PromptStatus.OK) return;
            int copiedCount = 0;
            try
            {
                using (DocumentLock docLock = activeDoc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var psdDict = new DictionaryPropertySetDefinitions(db);
                    var sourceHatch = (Hatch)tr.GetObject(resSource.ObjectId, OpenMode.ForRead);
                    if (sourceHatch == null) return;

                    var sourcePropertySets = PropertyDataServices.GetPropertySets(sourceHatch);
                    var propertySetsToCopy = new List<KeyValuePair<ObjectId, Dictionary<string, string>>>();

                    foreach (ObjectId psId in sourcePropertySets)
                    {
                        var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForRead);
                        if (ps != null)
                        {
                            foreach (string psdName in AllPsdNames)
                            {
                                if (psdDict.Has(psdName, tr) && psdDict.GetAt(psdName) == ps.PropertySetDefinition)
                                {
                                    var propertiesValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                    var psd = (PropertySetDefinition)tr.GetObject(ps.PropertySetDefinition, OpenMode.ForRead);

                                    foreach (PropertyDefinition propDef in psd.Definitions)
                                    {
                                        try
                                        {
                                            int propId = ps.PropertyNameToId(propDef.Name);
                                            propertiesValues[propDef.Name] = ps.GetAt(propId)?.ToString() ?? "";
                                        }
                                        catch { }
                                    }
                                    propertySetsToCopy.Add(new KeyValuePair<ObjectId, Dictionary<string, string>>(ps.PropertySetDefinition, propertiesValues));
                                    break;
                                }
                            }
                        }
                    }

                    if (propertySetsToCopy.Count == 0)
                    {
                        ed.WriteMessage("\n[Внимание] На исходной штриховке не найдены наборы характеристик группы 05_ДП_!");
                        return;
                    }

                    // Перенос собранных данных на выбранные целевые штриховки
                    SelectionSet selSet = resDest.Value;
                    foreach (SelectedObject selObj in selSet)
                    {
                        if (selObj.ObjectId == resSource.ObjectId) continue;

                        var destHatch = (Hatch)tr.GetObject(selObj.ObjectId, OpenMode.ForWrite);
                        if (destHatch == null) continue;

                        foreach (string psdName in AllPsdNames)
                        {
                            if (psdDict.Has(psdName, tr))
                            {
                                try { PropertyDataServices.RemovePropertySet(destHatch, psdDict.GetAt(psdName)); } catch { }
                            }
                        }

                        foreach (var pair in propertySetsToCopy)
                        {
                            ObjectId definitionId = pair.Key;
                            var savedValues = pair.Value;

                            PropertyDataServices.AddPropertySet(destHatch, definitionId);

                            foreach (ObjectId targetPsId in PropertyDataServices.GetPropertySets(destHatch))
                            {
                                var destPs = (PropertySet)tr.GetObject(targetPsId, OpenMode.ForWrite);
                                if (destPs != null && destPs.PropertySetDefinition == definitionId)
                                {
                                    foreach (var valPair in savedValues)
                                    {
                                        try
                                        {
                                            int propId = destPs.PropertyNameToId(valPair.Key);
                                            destPs.SetAt(propId, valPair.Value);
                                        }
                                        catch { }
                                    }
                                    break;
                                }
                            }
                        }
                        destHatch.RecordGraphicsModified(true);
                        copiedCount++;
                    }
                    tr.Commit();
                }
                ed.WriteMessage($"\n[Успех] Свойства скопированы! Характеристики перенесены на {copiedCount} штриховок.");
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка кисточки]: {ex.Message}"); }
            ed.UpdateScreen();
        }

        private bool SetProperty(PropertySet ps, string propName, object value)
        {
            try
            {
                int id = ps.PropertyNameToId(propName);
                ps.SetAt(id, value?.ToString() ?? "");
                return true;
            }
            catch { return false; }
        }
    }
}
