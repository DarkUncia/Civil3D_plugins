using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDataReader;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;

[assembly: CommandClass(typeof(Civil3D_FloraPlugin.FloraPropertyFiller))]
namespace Civil3D_FloraPlugin
{
    // 1. Модель строки Excel согласно структуре таблицы озеленения
    public class ExcelFloraRow
    {
        public string Id { get; set; } = string.Empty;
        public string CodeClassifierBuilding { get; set; } = string.Empty;
        public string CodeClassifierElement { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string TypePlant { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string ParamXxxs { get; set; } = string.Empty;
        public string ParamXxs { get; set; } = string.Empty;
        public string ParamXs { get; set; } = string.Empty;
        public string ParamS { get; set; } = string.Empty;
        public string ParamM { get; set; } = string.Empty;
        public string ParamL { get; set; } = string.Empty;
        public string ParamXl { get; set; } = string.Empty;
    }
    // 2. Класс для бинарного чтения Excel (Zero-Graphics поток)
    public static class ExcelFloraReader
    {
        public static Dictionary<string, ExcelFloraRow> ReadFloraTable(string filePath, Editor ed)
        {
            var registry = new Dictionary<string, ExcelFloraRow>(StringComparer.OrdinalIgnoreCase);
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Пропуск первых двух служебных строк (шапки таблицы)
                    if (!reader.Read() || !reader.Read()) return registry;

                    while (reader.Read())
                    {
                        string rawId = reader.GetValue(0)?.ToString()?.Trim() ?? string.Empty;
                        if (string.IsNullOrEmpty(rawId)) continue;

                        var row = new ExcelFloraRow
                        {
                            Id = rawId,
                            CodeClassifierBuilding = reader.GetValue(1)?.ToString()?.Trim() ?? string.Empty,
                            CodeClassifierElement = reader.GetValue(2)?.ToString()?.Trim() ?? string.Empty,
                            Name = reader.GetValue(3)?.ToString()?.Trim() ?? string.Empty,
                            TypePlant = reader.GetValue(4)?.ToString()?.Trim() ?? string.Empty,
                            Status = reader.GetValue(5)?.ToString()?.Trim() ?? string.Empty,
                            ParamXxxs = reader.GetValue(6)?.ToString()?.Trim() ?? string.Empty,
                            ParamXxs = reader.GetValue(7)?.ToString()?.Trim() ?? string.Empty,
                            ParamXs = reader.GetValue(8)?.ToString()?.Trim() ?? string.Empty,
                            ParamS = reader.GetValue(9)?.ToString()?.Trim() ?? string.Empty,
                            ParamM = reader.GetValue(10)?.ToString()?.Trim() ?? string.Empty,
                            ParamL = reader.GetValue(11)?.ToString()?.Trim() ?? string.Empty,
                            ParamXl = reader.GetValue(12)?.ToString()?.Trim() ?? string.Empty // Ошибка CS1003 исправлена
                        };

                        if (!registry.ContainsKey(rawId)) registry.Add(rawId, row);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Ошибка FloraExcel]: {ex.Message}");
            }
            return registry;
        }
    }
    // 3. Основной класс плагина Озеленения и точка входа для Civil 3D
    public class FloraPropertyFiller : IExtensionApplication
    {
        // Имя набора характеристик зафиксировано строго со скриншота диспетчера стилей
        public const string PropertySetName = "01_Деревья и кусты (стандарт)";

        public void Initialize()
        {
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
            if (ed != null)
            {
                ed.WriteMessage($"\n>>> Модуль Озеленения загружен успешно! Команда: FLORA_FILL_PROPERTIES <<<");
            }
        }

        public void Terminate() { }

        [CommandMethod("FLORA_FILL_PROPERTIES", CommandFlags.Modal)]
        public void RunFloraSynchronization()
        {
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor;
            Database db = activeDoc.Database;

            // Шаг 1: Запрос файла ведомости у пользователя
            string targetPath = string.Empty;
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Книга Excel (*.xlsx)|*.xlsx";
                dialog.Title = "Укажите файл базы данных Озеленения";
                if (dialog.ShowDialog() != DialogResult.OK) return;
                targetPath = dialog.FileName;
            }

            var databaseMap = ExcelFloraReader.ReadFloraTable(targetPath, ed);
            if (databaseMap == null || databaseMap.Count == 0)
            {
                ed.WriteMessage("\n[Отмена] Выбранная таблица пуста или занята другим процессом.");
                return;
            }

            // Шаг 2: Выбор элементов на чертеже
            ed.WriteMessage($"\nВыберите блоки для привязки характеристик '{PropertySetName}'...");
            PromptSelectionResult selResult = ed.GetSelection();
            if (selResult.Status != PromptStatus.OK || selResult.Value == null) return;

            int alteredObjectsCount = 0;

            using (DocumentLock docLock = activeDoc.LockDocument())
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    var psdDictionary = new DictionaryPropertySetDefinitions(db);
                    if (!psdDictionary.Has(PropertySetName, trans))
                    {
                        ed.WriteMessage($"\n[Ошибка] Набор '{PropertySetName}' не обнаружен в данном DWG-шаблоне!");
                        trans.Commit();
                        return;
                    }
                    ObjectId globalPsdId = psdDictionary.GetAt(PropertySetName);

                    foreach (SelectedObject selected in selResult.Value)
                    {
                        if (selected == null) continue;
                        DBObject currentDbObj = trans.GetObject(selected.ObjectId, OpenMode.ForWrite);

                        if (currentDbObj is BlockReference blockRef)
                        {
                            string rawBlockName = blockRef.IsDynamicBlock
                                ? ((BlockTableRecord)trans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                                : blockRef.Name;

                            string cleanedKey = rawBlockName.ToUpper().Trim();
                            if (cleanedKey.EndsWith("_2D")) cleanedKey = cleanedKey.Substring(0, cleanedKey.Length - 3).Trim();
                            if (cleanedKey.EndsWith("_3D")) cleanedKey = cleanedKey.Substring(0, cleanedKey.Length - 3).Trim();

                            if (databaseMap.TryGetValue(cleanedKey, out ExcelFloraRow dataRow))
                            {
                                ObjectId activePsId = ObjectId.Null;
                                ObjectIdCollection assignedSets = PropertyDataServices.GetPropertySets(blockRef);

                                foreach (ObjectId setInterfaceId in assignedSets)
                                {
                                    var currentSet = (PropertySet)trans.GetObject(setInterfaceId, OpenMode.ForRead);
                                    if (currentSet != null && currentSet.PropertySetDefinition == globalPsdId)
                                    {
                                        activePsId = setInterfaceId;
                                        break;
                                    }
                                }

                                if (activePsId.IsNull)
                                {
                                    PropertyDataServices.AddPropertySet(blockRef, globalPsdId);
                                    ObjectIdCollection freshSets = PropertyDataServices.GetPropertySets(blockRef);
                                    foreach (ObjectId setInterfaceId in freshSets)
                                    {
                                        var currentSet = (PropertySet)trans.GetObject(setInterfaceId, OpenMode.ForRead);
                                        if (currentSet != null && currentSet.PropertySetDefinition == globalPsdId)
                                        {
                                            activePsId = setInterfaceId;
                                            break;
                                        }
                                    }
                                }

                                if (!activePsId.IsNull)
                                {
                                    var writeableSet = (PropertySet)trans.GetObject(activePsId, OpenMode.ForWrite);
                                    if (writeableSet != null)
                                    {
                                        InjectProperty(writeableSet, "ID", dataRow.Id, trans);
                                        InjectProperty(writeableSet, "Код по классификатору зданий", dataRow.CodeClassifierBuilding, trans);
                                        InjectProperty(writeableSet, "Код по классификатору элементов", dataRow.CodeClassifierElement, trans);
                                        InjectProperty(writeableSet, "Наименование", dataRow.Name, trans);
                                        InjectProperty(writeableSet, "Тип", dataRow.TypePlant, trans);
                                        InjectProperty(writeableSet, "Статус", dataRow.Status, trans);

                                        InjectProperty(writeableSet, "Характеристики-xxxs", dataRow.ParamXxxs, trans);
                                        InjectProperty(writeableSet, "Характеристики-xxs", dataRow.ParamXxs, trans);
                                        InjectProperty(writeableSet, "Характеристики-xs", dataRow.ParamXs, trans);
                                        InjectProperty(writeableSet, "Характеристики-s", dataRow.ParamS, trans);
                                        InjectProperty(writeableSet, "Характеристики-m", dataRow.ParamM, trans);
                                        InjectProperty(writeableSet, "Характеристики-l", dataRow.ParamL, trans);
                                        InjectProperty(writeableSet, "Характеристики-xl", dataRow.ParamXl, trans);

                                        alteredObjectsCount++;
                                    }
                                }
                            }
                        }
                    }
                    trans.Commit();
                }
            }

            ed.WriteMessage($"\n[Успех] Заполнение завершено. Обработано блоков озеленения: {alteredObjectsCount}");
            ed.UpdateScreen();
        }

        private void InjectProperty(PropertySet ps, string propName, string value, Transaction tr)
        {
            try
            {
                int internalId = ps.PropertyNameToId(propName);
                ps.SetAt(internalId, value ?? string.Empty);
            }
            catch
            {
                try
                {
                    ObjectId metadataId = ps.PropertySetDefinition;
                    DBObject metadataObj = tr.GetObject(metadataId, OpenMode.ForRead);
                    var propertyCollectionInfo = metadataObj.GetType().GetProperty("Properties");
                    var collection = propertyCollectionInfo?.GetValue(metadataObj, null) as System.Collections.ICollection;

                    if (collection != null)
                    {
                        foreach (object singleDefinition in collection)
                        {
                            string targetName = singleDefinition.GetType().GetProperty("Name")?.GetValue(singleDefinition, null) as string ?? string.Empty;
                            if (!string.IsNullOrEmpty(targetName) && targetName.Equals(propName, StringComparison.OrdinalIgnoreCase))
                            {
                                int dynamicInternalId = ps.PropertyNameToId(targetName);
                                ps.SetAt(dynamicInternalId, value ?? string.Empty);
                                break;
                            }
                        }
                    }
                }
                catch { }
            }
        }
    }
}
