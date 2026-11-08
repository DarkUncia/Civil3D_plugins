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

[assembly: CommandClass(typeof(Civil3D_plugins.LightingPropertyFiller))]

namespace Civil3D_plugins
{
    // 1. Модель данных для освещения (на основе скриншота Excel)
    public class ExcelLightingRowFiller
    {
        public string Id { get; set; }                      // Колонка A (0): Артикул
        public string NameElement { get; set; }             // Колонка B (1): Наименование элемента
        public string CodeClassifierBuilding { get; set; }  // Колонка C (2): Код по классификатору зданий
        public string CodeClassifierElement { get; set; }   // Колонка D (3): Код по классификатору элементов
        public string CodeClassifierMaterial { get; set; }  // Колонка E (4): Код по классификатору материалов
        public string SupportType { get; set; }             // Колонка F (5): Тип опоры
        public string SupportHeight { get; set; }           // Колонка G (6): Высота опоры
        public string Dimensions { get; set; }              // Колонка H (7): Габаритные размеры
        public string Position { get; set; }                // Колонка I (8): Положение
        public string Construction { get; set; }            // Колонка J (9): Конструкция
    }

    // 2. Парсер Excel для освещения
    public static class LightingExcelReader
    {
        public static Dictionary<string, ExcelLightingRowFiller> ReadExcelData(string filePath, Editor ed)
        {
            var result = new Dictionary<string, ExcelLightingRowFiller>(StringComparer.OrdinalIgnoreCase);
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    // Шаг 1 и 2: Пропускаем две строки шапки
                    if (!reader.Read()) return null;
                    if (!reader.Read()) return null;

                    // Шаг 3: Читаем данные
                    while (reader.Read())
                    {
                        string rowId = reader.GetValue(0)?.ToString()?.Trim() ?? string.Empty;
                        if (string.IsNullOrEmpty(rowId)) continue;

                        var rowData = new ExcelLightingRowFiller
                        {
                            Id = rowId,
                            NameElement = reader.GetValue(1)?.ToString()?.Trim() ?? "",
                            CodeClassifierBuilding = reader.GetValue(2)?.ToString()?.Trim() ?? "",
                            CodeClassifierElement = reader.GetValue(3)?.ToString()?.Trim() ?? "",
                            CodeClassifierMaterial = reader.GetValue(4)?.ToString()?.Trim() ?? "",
                            SupportType = reader.GetValue(5)?.ToString()?.Trim() ?? "",
                            SupportHeight = reader.GetValue(6)?.ToString()?.Trim() ?? "",
                            Dimensions = reader.GetValue(7)?.ToString()?.Trim() ?? "",
                            Position = reader.GetValue(8)?.ToString()?.Trim() ?? "",
                            Construction = reader.GetValue(9)?.ToString()?.Trim() ?? ""
                        };

                        if (!result.ContainsKey(rowId))
                        {
                            result.Add(rowId, rowData);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Ошибка чтения Excel (Освещение)]: {ex.Message}");
                return null;
            }
            return result;
        }
    }

    // 3. Основной класс команды (аналог MafOnlyPropertyFiller)
    public class LightingPropertyFiller
    {
        public const string TargetPsdName = "06_Освещение";

        [CommandMethod("LIGHTING_FILL_ATTRIBUTES", CommandFlags.Modal)]
        public void FillPropertiesProcess()
        {
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor;

            // 1. Выбор файла Excel
            string excelFilePath = string.Empty;
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel (*.xlsx)|*.xlsx";
                ofd.Title = "Выберите БД Excel элементов освещения";
                if (ofd.ShowDialog() != DialogResult.OK) return;
                excelFilePath = ofd.FileName;
            }

            var excelData = LightingExcelReader.ReadExcelData(excelFilePath, ed);
            if (excelData == null || excelData.Count == 0)
            {
                ed.WriteMessage("\n[Ошибка] Не удалось загрузить базу данных Excel или она пуста.");
                return;
            }

            // 2. Выбор блоков
            ed.WriteMessage($"\nВыберите блоки освещения для записи параметров в НХ '{TargetPsdName}'...");
            PromptSelectionResult psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK || psr.Value == null || psr.Value.Count == 0) return;

            int successCount = 0;
            Database db = activeDoc.Database;

            using (DocumentLock docLock = activeDoc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var psdDict = new DictionaryPropertySetDefinitions(db);
                    if (!psdDict.Has(TargetPsdName, tr))
                    {
                        ed.WriteMessage($"\n[Критическая ошибка] Набор характеристик '{TargetPsdName}' отсутствует в текущем чертеже!");
                        tr.Commit();
                        return;
                    }
                    ObjectId psdId = psdDict.GetAt(TargetPsdName);

                    foreach (SelectedObject selObj in psr.Value)
                    {
                        if (selObj == null) continue;
                        DBObject obj = tr.GetObject(selObj.ObjectId, OpenMode.ForWrite);

                        if (obj is BlockReference br)
                        {
                            string blockName = br.IsDynamicBlock
                                ? ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                                : br.Name;

                            // Очистка от _2D и _3D
                            string cleanArticul = blockName.ToUpper().Trim();
                            if (cleanArticul.EndsWith("_2D")) cleanArticul = cleanArticul.Substring(0, cleanArticul.Length - 3).Trim();
                            if (cleanArticul.EndsWith("_3D")) cleanArticul = cleanArticul.Substring(0, cleanArticul.Length - 3).Trim();

                            if (excelData.TryGetValue(cleanArticul, out ExcelLightingRowFiller excelRow))
                            {
                                ObjectId psId = ObjectId.Null;
                                ObjectIdCollection currentPropertySets = PropertyDataServices.GetPropertySets(br);

                                foreach (ObjectId id in currentPropertySets)
                                {
                                    var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                    if (testPs != null && testPs.PropertySetDefinition == psdId)
                                    {
                                        psId = id;
                                        break;
                                    }
                                }

                                if (psId.IsNull)
                                {
                                    PropertyDataServices.AddPropertySet(br, psdId);
                                    ObjectIdCollection updatedPropertySets = PropertyDataServices.GetPropertySets(br);
                                    foreach (ObjectId id in updatedPropertySets)
                                    {
                                        var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                        if (testPs != null && testPs.PropertySetDefinition == psdId)
                                        {
                                            psId = id;
                                            break;
                                        }
                                    }
                                }

                                if (!psId.IsNull)
                                {
                                    var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                                    if (ps != null)
                                    {
                                        // Записываем данные в НХ (имена строго по вашему скриншоту)
                                        SafeSetProperty(ps, "ID", excelRow.Id, tr);
                                        SafeSetProperty(ps, "Высота опоры", excelRow.SupportHeight, tr);
                                        SafeSetProperty(ps, "Габаритные размеры", excelRow.Dimensions, tr);
                                        SafeSetProperty(ps, "Код по классификатору зданий", excelRow.CodeClassifierBuilding, tr);
                                        SafeSetProperty(ps, "Код по классификатору материалов", excelRow.CodeClassifierMaterial, tr);
                                        SafeSetProperty(ps, "Код по классификатору элементов", excelRow.CodeClassifierElement, tr);
                                        SafeSetProperty(ps, "Конструкция", excelRow.Construction, tr);
                                        SafeSetProperty(ps, "Наименование элемента", excelRow.NameElement, tr);
                                        SafeSetProperty(ps, "Положение", excelRow.Position, tr);
                                        SafeSetProperty(ps, "Тип опоры", excelRow.SupportType, tr);

                                        successCount++;
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }

            ed.WriteMessage($"\n[Успех] Обработка завершена. Заполнено параметров у блоков освещения: {successCount}");
            ed.UpdateScreen();
        }

        // Точная копия метода из МАФ для безопасной записи
        private void SafeSetProperty(PropertySet ps, string exactPropName, string value, Transaction tr)
        {
            try
            {
                int propId = ps.PropertyNameToId(exactPropName);
                ps.SetAt(propId, value ?? string.Empty);
            }
            catch
            {
                try
                {
                    ObjectId defId = ps.PropertySetDefinition;
                    DBObject defObj = tr.GetObject(defId, OpenMode.ForRead);

                    var propsProp = defObj.GetType().GetProperty("Properties");
                    var propsCollection = propsProp?.GetValue(defObj, null) as System.Collections.ICollection;

                    if (propsCollection != null)
                    {
                        foreach (object prop in propsCollection)
                        {
                            string propName = prop.GetType().GetProperty("Name")?.GetValue(prop, null) as string;
                            if (!string.IsNullOrEmpty(propName) && propName.Equals(exactPropName, StringComparison.OrdinalIgnoreCase))
                            {
                                int dynamicId = ps.PropertyNameToId(propName);
                                ps.SetAt(dynamicId, value ?? string.Empty);
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