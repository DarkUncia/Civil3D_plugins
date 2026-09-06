using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;

[assembly: CommandClass(typeof(Civil3D_plugins.FloraBIMImportCommand))]

namespace Civil3D_plugins
{
    public class FloraBIMImportCommand
    {
        // Словарь соответствия вкладок XML-файла и наборов характеристик Civil 3D
        private static readonly Dictionary<string, string> SheetToPsetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Деревья", "01_Деревья и кусты (стандарт)" },
            { "Кустарники", "01_Деревья и кусты (стандарт)" }, // Использует тот же набор
            { "ЖИ", "02_Живая изгородь. Многолетние цветы. Злаки" },
            { "Цветники", "04_Озеленение (индивидуально)" }
        };

        public class FloraExcelRow
        {
            public string HandleStr { get; set; } = string.Empty;
            public string TargetPsetName { get; set; } = string.Empty;
            public Dictionary<string, string> Properties { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        [CommandMethod("FLORA_IMPORT_PSETS")]
        public void ImportFloraFromExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Выберите ведомость FloraBIM для импорта параметров",
                Filter = "Excel XML Spreadsheet (*.xml)|*.xml",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() != true) return;
            ed.WriteMessage($"\n[FloraBIM] Чтение XML файла: {Path.GetFileName(openFileDialog.FileName)}...");

            var parsedRows = new List<FloraExcelRow>();

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(openFileDialog.FileName);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");

                XmlNodeList worksheets = xmlDoc.SelectNodes("//ss:Worksheet", nsmgr);

                foreach (XmlNode worksheet in worksheets)
                {
                    string sheetName = worksheet.Attributes["ss:Name"]?.Value;
                    if (string.IsNullOrEmpty(sheetName) || !SheetToPsetNameMap.ContainsKey(sheetName)) continue;

                    // Определяем целевой набор характеристик для текущего листа
                    string targetPsetName = SheetToPsetNameMap[sheetName];

                    XmlNodeList rows = worksheet.SelectNodes(".//ss:Row", nsmgr);
                    if (rows == null || rows.Count <= 1) continue;

                    // Динамически считываем шапку таблицы (первая строка), чтобы не зависеть от номеров колонок
                    var colHeaderMap = new Dictionary<int, string>();
                    XmlNodeList headerCells = rows[0].SelectNodes(".//ss:Cell", nsmgr);
                    for (int c = 0; c < headerCells.Count; c++)
                    {
                        string headerText = headerCells[c].SelectSingleNode(".//ss:Data", nsmgr)?.InnerText?.Trim();
                        if (!string.IsNullOrEmpty(headerText))
                        {
                            colHeaderMap[c] = headerText;
                        }
                    }

                    // Читаем строки данных
                    for (int i = 1; i < rows.Count; i++)
                    {
                        XmlNodeList cells = rows[i].SelectNodes(".//ss:Cell", nsmgr);
                        if (cells == null || cells.Count == 0) continue;

                        string handle = cells[0].SelectSingleNode(".//ss:Data", nsmgr)?.InnerText?.Trim();
                        if (string.IsNullOrEmpty(handle)) continue;

                        var rowData = new FloraExcelRow { HandleStr = handle, TargetPsetName = targetPsetName };

                        // Считываем ячейки на основе сопоставления заголовков шапки
                        for (int c = 1; c < cells.Count; c++)
                        {
                            if (colHeaderMap.TryGetValue(c, out string propName))
                            {
                                string val = cells[c].SelectSingleNode(".//ss:Data", nsmgr)?.InnerText?.Trim();
                                rowData.Properties[propName] = val ?? string.Empty;
                            }
                        }
                        parsedRows.Add(rowData);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[FloraBIM Ошибка парсинга XML]: {ex.Message}");
                return;
            }

            if (parsedRows.Count == 0)
            {
                ed.WriteMessage("\n[FloraBIM] Данные для импорта не найдены.");
                return;
            }

            ExecuteBimImport(db, ed, parsedRows);
        }
        private void ExecuteBimImport(Database db, Editor ed, List<FloraExcelRow> parsedRows)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                int updatedObjectsCount = 0;

                // Получаем глобальный словарь определений наборов характеристик чертежа
                object rawNod = tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                var nod = rawNod as DBDictionary;
                if (nod == null || !nod.Contains("Aec_PropertySet_Definitions"))
                {
                    ed.WriteMessage("\n[FloraBIM Ошибка] В чертеже полностью отсутствуют Property Set Definitions.");
                    return;
                }

                object rawPropDefs = tr.GetObject(nod.GetAt("Aec_PropertySet_Definitions"), OpenMode.ForRead);
                var propDefs = rawPropDefs as DBDictionary;
                if (propDefs == null) return;

                // Кэшируем идентификаторы определений PropertySet для ускорения работы
                var pSetDefIds = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
                var pSetDefObjects = new Dictionary<string, PropertySetDefinition>(StringComparer.OrdinalIgnoreCase);

                foreach (var name in SheetToPsetNameMap.Values)
                {
                    if (pSetDefIds.ContainsKey(name)) continue;

                    if (propDefs.Contains(name))
                    {
                        ObjectId defId = propDefs.GetAt(name);
                        var pSetDef = tr.GetObject(defId, OpenMode.ForRead) as PropertySetDefinition;
                        if (pSetDef != null)
                        {
                            pSetDefIds[name] = defId;
                            pSetDefObjects[name] = pSetDef;
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\n[FloraBIM Предупреждение] Набор '{name}' отсутствует в Диспетчере стилей чертежа. Объекты этого типа будут пропущены.");
                    }
                }

                // Основной цикл обновления наборов на графических элементах плана
                foreach (var importRow in parsedRows)
                {
                    string targetSetName = importRow.TargetPsetName;
                    if (!pSetDefIds.TryGetValue(targetSetName, out ObjectId defId)) continue;

                    var pSetDef = pSetDefObjects[targetSetName];

                    // Конвертируем Handle из HEX-строки XML обратно в структуру AutoCAD
                    if (!long.TryParse(importRow.HandleStr, System.Globalization.NumberStyles.HexNumber, null, out long handleValue))
                        continue;

                    Handle h = new Handle(handleValue);
                    if (!db.TryGetObjectId(h, out ObjectId objId)) continue;

                    object rawDbObj = tr.GetObject(objId, OpenMode.ForWrite);
                    var dbObj = rawDbObj as DBObject;
                    if (dbObj == null) continue;

                    ObjectId pSetId = ObjectId.Null;
                    ObjectIdCollection attachedSets = PropertyDataServices.GetPropertySets(dbObj);

                    foreach (ObjectId attachedId in attachedSets)
                    {
                        object rawPSetObj = tr.GetObject(attachedId, OpenMode.ForRead);
                        var pSetObj = rawPSetObj as PropertySet;
                        if (pSetObj != null && pSetObj.PropertySetDefinition == defId)
                        {
                            pSetId = attachedId;
                            break;
                        }
                    }

                    // Если набора характеристик ещё нет на блоке, безопасно добавляем его на лету
                    if (pSetId == ObjectId.Null)
                    {
                        PropertyDataServices.AddPropertySet(dbObj, defId);
                        attachedSets = PropertyDataServices.GetPropertySets(dbObj);
                        foreach (ObjectId attachedId in attachedSets)
                        {
                            object rawPSetObj = tr.GetObject(attachedId, OpenMode.ForRead);
                            var pSetObj = rawPSetObj as PropertySet;
                            if (pSetObj != null && pSetObj.PropertySetDefinition == defId)
                            {
                                pSetId = attachedId;
                                break;
                            }
                        }
                    }

                    // Запись характеристик (Property) в экземпляр набора объекта
                    if (pSetId != ObjectId.Null)
                    {
                        object rawPSet = tr.GetObject(pSetId, OpenMode.ForWrite);
                        var pSet = rawPSet as PropertySet;
                        if (pSet != null)
                        {
                            bool objectChanged = false;
                            foreach (PropertyDefinition propDef in pSetDef.Definitions)
                            {
                                if (propDef.Automatic) continue; // Пропускаем автоматические системные параметры

                                if (importRow.Properties.TryGetValue(propDef.Name, out string excelValue))
                                {
                                    pSet.SetAt(propDef.Id, excelValue ?? string.Empty);
                                    objectChanged = true;
                                }
                            }
                            if (objectChanged) updatedObjectsCount++;
                        }
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n[FloraBIM] Импорт параметров успешно завершен. Обновлено объектов: {updatedObjectsCount}");
            }
        }
    }
}
