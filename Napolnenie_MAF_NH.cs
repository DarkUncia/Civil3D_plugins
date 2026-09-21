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

            if (t.Contains("мульча") || n.Contains("мульча") || t.Contains("сыпучее") || n.Contains("сыпучее") ||
                t.Contains("цветники") || n.Contains("цветники") || t.Contains("настил") || n.Contains("настил") ||
                t.Contains("терравей") || n.Contains("терравей") || t.Contains("щепа") || n.Contains("щепа"))
                return "05_ДП_(иное)";

            if (t.Contains("плитка")) return "05_ДП_(плитка)";
            if (t.Contains("газон") || t.Contains("решетки") || t.Contains("озеленение")) return "05_ДП_(озеленение)";
            if (t.Contains("резин")) return "05_ДП_(мягкие)";
            if (n.Contains("асфальт")) return "05_ДП_(твердые)";

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
            if (ed != null) ed.WriteMessage("\n>>> Плагин штриховок ГП загружен! Команда: COVERAGE_FILL_PROPERTIES <<<");
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

                                // Очистка старых конфликтующих НХ из группы 05_ДП_
                                ObjectIdCollection currentSets = PropertyDataServices.GetPropertySets(hatch);
                                foreach (ObjectId id in currentSets)
                                {
                                    var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                    if (testPs != null)
                                    {
                                        foreach (string psdName in AllPsdNames)
                                        {
                                            if (psdDict.Has(psdName, tr) && psdDict.GetAt(psdName) == testPs.PropertySetDefinition)
                                            {
                                                if (psdName != targetPsdName) PropertyDataServices.RemovePropertySet(hatch, psdDict.GetAt(psdName));
                                                break;
                                            }
                                        }
                                    }
                                }

                                // Поиск или добавление целевого НХ
                                ObjectId psId = ObjectId.Null;
                                currentSets = PropertyDataServices.GetPropertySets(hatch);
                                foreach (ObjectId id in currentSets)
                                {
                                    var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                    if (testPs != null && testPs.PropertySetDefinition == targetPsdId) { psId = id; break; }
                                }

                                if (psId.IsNull)
                                {
                                    PropertyDataServices.AddPropertySet(hatch, targetPsdId);
                                    foreach (ObjectId id in PropertyDataServices.GetPropertySets(hatch))
                                    {
                                        var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                        if (testPs != null && testPs.PropertySetDefinition == targetPsdId) { psId = id; break; }
                                    }
                                }

                                // Запись данных в свойства
                                if (!psId.IsNull)
                                {
                                    var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                                    if (ps != null)
                                    {
                                        // Записываем ID покрытия (строку "ГП_260_Покрытие...") в свойство "Название_покрытия"
                                        SetProperty(ps, "Название_покрытия", excelRow.CoverageId);

                                        SetProperty(ps, "Код_по_классификатору_зданий", excelRow.CodeClassifier);
                                        SetProperty(ps, "Код_по_классификатору_материалов", excelRow.CodeMaterial);
                                        SetProperty(ps, "Позиция_Тип", excelRow.PositionType);
                                        SetProperty(ps, "Обозначение", excelRow.Designation);
                                        SetProperty(ps, "Наименование", excelRow.Name);
                                        SetProperty(ps, "Пирог_покрытия_со_слоями", excelRow.PieLayers);
                                        SetProperty(ps, "Толщина_покрытия", excelRow.Thickness);

                                        // Запись типа покрытия (Плитка, асфальт и т.д.)
                                        if (!SetProperty(ps, "Тип_покрытия", excelRow.Type))
                                        {
                                            SetProperty(ps, "Тип", excelRow.Type);
                                        }

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
