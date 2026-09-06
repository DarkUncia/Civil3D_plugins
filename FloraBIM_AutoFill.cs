using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;
using DocumentFormat.OpenXml.Packaging;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

[assembly: CommandClass(typeof(Civil3D_plugins.FloraBIMSmartFillCommand))]
namespace Civil3D_plugins
{
    public class FloraBIMSmartFillCommand
    {
        private static readonly Dictionary<string, string> SheetToPsetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Деревья", "01_Деревья и кусты (стандарт)" }, { "Кустарники", "01_Деревья и кусты (стандарт)" },
            { "ЖИ", "02_Живая изгородь. Многолетние цветы. Злаки" }, { "Цветники", "03_Цветочные миксы" }
        };

        [CommandMethod("FLORA_SMART_FILL")]
        public void SmartFillFromCatalog()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument; if (doc == null) return;
            Editor ed = doc.Editor; Database db = doc.Database;
            ed.WriteMessage("\n================= [FloraBIM СТАРТ ТРАССИРОВКИ] =================");

            var targetBlockIds = new List<ObjectId>(); var debugBlockNames = new List<string>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId id in ms)
                {
                    var blkRef = tr.GetObject(id, OpenMode.ForRead) as BlockReference; if (blkRef == null) continue;
                    string nameRaw = blkRef.IsDynamicBlock ? ((BlockTableRecord)tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name : blkRef.Name;
                    targetBlockIds.Add(id); if (!debugBlockNames.Contains(nameRaw)) debugBlockNames.Add(nameRaw);
                }
                tr.Commit();
            }
            ed.WriteMessage($"\n[ШАГ 1] Блоков на чертеже: {targetBlockIds.Count}. Уникальных имён: {debugBlockNames.Count}");
            foreach (var bName in debugBlockNames.Take(10)) ed.WriteMessage($"\n  -> Имя на чертеже: '{bName}' (Длина: {bName.Length})");

            var openFileDialog = new Microsoft.Win32.OpenFileDialog { Title = "Выберите мастер-каталог Excel (.xlsx)", Filter = "Excel Workbooks (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() != true) return;
            string excelPath = openFileDialog.FileName;

            var catalogCache = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
            foreach (var psetName in SheetToPsetNameMap.Values) catalogCache[psetName] = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelPath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart; SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
                    foreach (Excel.Sheet sheet in workbookPart.Workbook.Sheets.Cast<Excel.Sheet>())
                    {
                        string sheetName = sheet.Name; if (string.IsNullOrEmpty(sheetName) || !SheetToPsetNameMap.ContainsKey(sheetName)) continue;
                        string targetPsetName = SheetToPsetNameMap[sheetName]; var psetDict = catalogCache[targetPsetName];
                        WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                        Excel.SheetData sheetData = worksheetPart.Worksheet.Elements<Excel.SheetData>().First();
                        var rows = sheetData.Elements<Excel.Row>().ToList(); if (rows.Count <= 1) continue;

                        var colHeaderMap = new Dictionary<int, string>(); var headerCells = rows.First().Descendants<Excel.Cell>().ToList();
                        for (int c = 0; c < headerCells.Count; c++) { string headerText = GetCellValue(headerCells[c], stringTablePart); if (!string.IsNullOrEmpty(headerText)) colHeaderMap[c] = headerText; }

                        int idColIndex = -1;
                        foreach (var kp in colHeaderMap) { if (kp.Value.Equals("id", StringComparison.OrdinalIgnoreCase)) { idColIndex = kp.Key; break; } }
                        if (idColIndex == -1) { ed.WriteMessage($"\n[ОШИБКА] На листе '{sheetName}' нет колонки 'id'!"); continue; }

                        int sheetKeysCount = 0;
                        for (int i = 1; i < rows.Count; i++)
                        {
                            var cells = rows[i].Descendants<Excel.Cell>().ToList(); if (cells.Count <= idColIndex) continue;
                            string rawIdVal = GetCellValue(cells[idColIndex], stringTablePart); if (string.IsNullOrEmpty(rawIdVal)) continue;

                            string cleanExcelIdKey = Convert.ToString(rawIdVal).Replace("-", "").Replace("_", "").Replace(" ", "").Replace("{", "").Replace("}", "").Trim();
                            int dotIdx = cleanExcelIdKey.IndexOf('.'); if (dotIdx != -1) cleanExcelIdKey = cleanExcelIdKey.Substring(0, dotIdx);
                            int commaIdx = cleanExcelIdKey.IndexOf(','); if (commaIdx != -1) cleanExcelIdKey = cleanExcelIdKey.Substring(0, commaIdx);
                            cleanExcelIdKey = cleanExcelIdKey.ToUpper(); if (string.IsNullOrEmpty(cleanExcelIdKey)) continue;

                            if (sheetKeysCount < 5) ed.WriteMessage($"\n  - Лист '{sheetName}', ID из Excel: '{cleanExcelIdKey}' (Длина: {cleanExcelIdKey.Length})");
                            if (!psetDict.ContainsKey(cleanExcelIdKey)) { psetDict[cleanExcelIdKey] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); sheetKeysCount++; }
                            for (int c = 0; c < cells.Count; c++) { if (colHeaderMap.TryGetValue(c, out string propName)) psetDict[cleanExcelIdKey][propName] = GetCellValue(cells[c], stringTablePart); }
                        }
                        ed.WriteMessage($"\n[ШАГ 2] Импортировано из листа '{sheetName}': {sheetKeysCount} уникальных ключей.");
                    }
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка Excel]: {ex.Message}"); return; }
            // ШАГ 3: ТОЧНОЕ СОПОСТАВЛЕНИЕ ПАРАМЕТРОВ С ЛОГИРОВАНИЕМ СБОЕВ СВЕРКИ
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                object rawNod = tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite); var nod = rawNod as DBDictionary; if (nod == null) return;
                DBDictionary propDefs = nod.Contains("Aec_PropertySet_Definitions") ? tr.GetObject(nod.GetAt("Aec_PropertySet_Definitions"), OpenMode.ForWrite) as DBDictionary : null;
                if (propDefs == null) { propDefs = new DBDictionary(); nod.SetAt("Aec_PropertySet_Definitions", propDefs); tr.AddNewlyCreatedDBObject(propDefs, true); }

                var pSetDefIds = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
                var pSetDefObjects = new Dictionary<string, PropertySetDefinition>(StringComparer.OrdinalIgnoreCase);
                foreach (var name in SheetToPsetNameMap.Values)
                {
                    if (pSetDefIds.ContainsKey(name)) continue;
                    if (propDefs.Contains(name))
                    {
                        ObjectId defId = propDefs.GetAt(name); var pSetDef = tr.GetObject(defId, OpenMode.ForWrite) as PropertySetDefinition;
                        if (pSetDef != null) { pSetDefIds[name] = defId; pSetDefObjects[name] = pSetDef; }
                    }
                }

                int updatedCount = 0; int failedMatchesLogCount = 0;
                foreach (ObjectId blockId in targetBlockIds)
                {
                    var blkRef = tr.GetObject(blockId, OpenMode.ForWrite) as BlockReference; if (blkRef == null) continue;
                    string blockNameRaw = blkRef.IsDynamicBlock ? ((BlockTableRecord)tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name : blkRef.Name;

                    string blockNameClean = Convert.ToString(blockNameRaw).Replace("-", "").Replace("_", "").Replace(" ", "").Replace("{", "").Replace("}", "").Trim();
                    int bDotIdx = blockNameClean.IndexOf('.'); if (bDotIdx != -1) blockNameClean = blockNameClean.Substring(0, bDotIdx);
                    int bCommaIdx = blockNameClean.IndexOf(','); if (bCommaIdx != -1) blockNameClean = blockNameClean.Substring(0, bCommaIdx);
                    blockNameClean = blockNameClean.ToUpper();

                    string matchedPsetName = string.Empty; Dictionary<string, string> excelProperties = null;
                    foreach (var kp in catalogCache) { if (kp.Value.TryGetValue(blockNameClean, out excelProperties)) { matchedPsetName = kp.Key; break; } }

                    if (excelProperties == null)
                    {
                        if (failedMatchesLogCount < 10) { ed.WriteMessage($"\n  [ОТКАЗ СВЕРКИ] Блок '{blockNameRaw}' -> Чистый: '{blockNameClean}' -> НЕ НАЙДЕН в Excel!"); failedMatchesLogCount++; }
                        continue;
                    }

                    if (!pSetDefIds.TryGetValue(matchedPsetName, out ObjectId defId)) continue;
                    var pSetDef = pSetDefObjects[matchedPsetName];

                    ObjectId pSetId = ObjectId.Null; ObjectIdCollection attachedSets = PropertyDataServices.GetPropertySets(blkRef);
                    foreach (ObjectId attachedId in attachedSets) { var pSetObj = tr.GetObject(attachedId, OpenMode.ForRead) as PropertySet; if (pSetObj != null && pSetObj.PropertySetDefinition == defId) { pSetId = attachedId; break; } }

                    if (pSetId == ObjectId.Null)
                    {
                        PropertyDataServices.AddPropertySet(blkRef, defId); attachedSets = PropertyDataServices.GetPropertySets(blkRef);
                        foreach (ObjectId attachedId in attachedSets) { var pSetObj = tr.GetObject(attachedId, OpenMode.ForRead) as PropertySet; if (pSetObj != null && pSetObj.PropertySetDefinition == defId) { pSetId = attachedId; break; } }
                    }

                    if (pSetId != ObjectId.Null)
                    {
                        var pSet = tr.GetObject(pSetId, OpenMode.ForWrite) as PropertySet;
                        if (pSet != null)
                        {
                            bool isBlockChanged = false; excelProperties.TryGetValue("Размер", out string targetSizeValue); if (!string.IsNullOrEmpty(targetSizeValue)) targetSizeValue = targetSizeValue.Trim();
                            foreach (PropertyDefinition propDef in pSetDef.Definitions)
                            {
                                if (propDef.Automatic) continue; string propName = propDef.Name;
                                if (propName.Equals("s", StringComparison.OrdinalIgnoreCase) || propName.Equals("m", StringComparison.OrdinalIgnoreCase) || propName.Equals("l", StringComparison.OrdinalIgnoreCase) || propName.Equals("xl", StringComparison.OrdinalIgnoreCase) || propName.Equals("xxl", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (!string.IsNullOrEmpty(targetSizeValue) && propName.Equals(targetSizeValue, StringComparison.OrdinalIgnoreCase)) { excelProperties.TryGetValue(propName, out string specificSizeData); pSet.SetAt(propDef.Id, !string.IsNullOrEmpty(specificSizeData) ? specificSizeData : "1"); isBlockChanged = true; }
                                    continue;
                                }
                                if (excelProperties.TryGetValue(propName, out string catalogValue)) { pSet.SetAt(propDef.Id, catalogValue ?? string.Empty); isBlockChanged = true; }
                            }
                            if (isBlockChanged) updatedCount++;
                        }
                    }
                }
                tr.Commit();
                ed.WriteMessage($"\n================= [FloraBIM ИТОГ ТРАССИРОВКИ] =================");
                ed.WriteMessage($"\n[ИТОГ] Успешно перезаписано блоков: {updatedCount}");

                var emptyArray = new ObjectId[0]; ed.SetImpliedSelection(emptyArray);
            }
        }

        private string GetCellValue(Excel.Cell cell, SharedStringTablePart stringTablePart)
        {
            if (cell == null) return string.Empty;
            string value = cell.CellValue != null ? cell.CellValue.Text : cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == Excel.CellValues.SharedString) return stringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText.Trim();
            return (value ?? string.Empty).Trim();
        }
    }
}
