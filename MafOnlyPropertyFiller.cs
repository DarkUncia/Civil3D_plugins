using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;

// Обязательные пространства имен AutoCAD и Civil 3D
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;

[assembly: CommandClass(typeof(Civil3D_plugins.MafOnlyPropertyFiller))]

namespace Civil3D_plugins
{
    // Уникальный класс модели данных, исключающий любые дублирования в проекте
    public class ExcelMafRowFiller
    {
        public string MafId { get; set; }                    // Колонка А (индекс 0): ID (Артикул)
        public string CodeClassifierBuilding { get; set; }   // Колонка B (индекс 1): Код по классификатору зданий
        public string CodeClassifierElement { get; set; }    // Колонка C (индекс 2): Код по классификатору элементов
        public string Position { get; set; }                 // Колонка D (индекс 3): Позиция
        public string Name { get; set; }                     // Колонка E (индекс 4): Наименование элемента
        public string AgeGroup { get; set; }                 // Колонка F (индекс 5): Возрастная группа
        public string Dimensions { get; set; }               // Колонка G (индекс 6): Габаритные размеры
        public string TypeName { get; set; }                 // Колонка H (индекс 7): Формируем имя блока (Тип МАФ)
    }

    public static class ExcelDataReaderHelperForOnlyFiller
    {
        public static Dictionary<string, ExcelMafRowFiller> ReadExcelData(string filePath, Editor ed)
        {
            var result = new Dictionary<string, ExcelMafRowFiller>(StringComparer.OrdinalIgnoreCase);
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    // Шаг 1: Пропускаем самую первую строку (Строка заголовков: ID, Код по кла...)
                    if (!reader.Read()) return null;

                    // Шаг 2: ЖЕСТКО ИГНОРИРУЕМ И ПРОПУСКАЕМ ВТОРУЮ СТРОКУ (Информационный текст: X.X.X.X и т.д.)
                    if (!reader.Read()) return null;

                    // Шаг 3: Начиная со строки №3, считываем чистые параметры объектов по индексам колонок
                    while (reader.Read())
                    {
                        string rowId = reader.GetValue(0)?.ToString()?.Trim() ?? string.Empty;

                        // Если ячейка артикула пустая — пропускаем строчку
                        if (string.IsNullOrEmpty(rowId)) continue;

                        var rowData = new ExcelMafRowFiller
                        {
                            MafId = rowId,
                            CodeClassifierBuilding = reader.GetValue(1)?.ToString()?.Trim() ?? "",
                            CodeClassifierElement = reader.GetValue(2)?.ToString()?.Trim() ?? "",
                            Position = reader.GetValue(3)?.ToString()?.Trim() ?? "",
                            Name = reader.GetValue(4)?.ToString()?.Trim() ?? "",
                            AgeGroup = reader.GetValue(5)?.ToString()?.Trim() ?? "",
                            Dimensions = reader.GetValue(6)?.ToString()?.Trim() ?? "",
                            TypeName = reader.GetValue(7)?.ToString()?.Trim() ?? ""
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
                ed.WriteMessage($"\n[Ошибка чтения Excel]: {ex.Message}");
                return null;
            }
            return result;
        }
    }
    public class MafOnlyPropertyFiller
    {
        public const string TargetPsdName = "07_МАФ";

        [CommandMethod("MAF_FILL_ATTRIBUTES", CommandFlags.Modal)]
        public void FillPropertiesProcess()
        {
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor;

            // 1. ВЫБОР ФАЙЛА EXCEL
            string excelFilePath = string.Empty;
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel (*.xlsx)|*.xlsx";
                ofd.Title = "Выберите БД Excel элементов МАФ";
                if (ofd.ShowDialog() != DialogResult.OK) return;
                excelFilePath = ofd.FileName;
            }

            var excelData = ExcelDataReaderHelperForOnlyFiller.ReadExcelData(excelFilePath, ed);
            if (excelData == null || excelData.Count == 0)
            {
                ed.WriteMessage("\n[Ошибка] Не удалось загрузить базу данных Excel или она пуста.");
                return;
            }

            // 2. ВЫБОР БЛОКОВ ПОЛЬЗОВАТЕЛЕМ НА ЭКРАНЕ
            ed.WriteMessage("\nВыберите блоки на экране для записи параметров в НХ 07_МАФ...");
            PromptSelectionResult psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK || psr.Value == null || psr.Value.Count == 0) return;

            int successCount = 0;
            Database db = activeDoc.Database;

            using (DocumentLock docLock = activeDoc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Получаем ID определения стиля "07_МАФ" в текущем чертеже
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
                            // Получаем имя дефиниции блока с защитой от динамических имен
                            string blockName = br.IsDynamicBlock
                                ? ((BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                                : br.Name;

                            // Очищаем имя блока от суффиксов _2D и _3D для точного сопоставления с Excel (Adanat_10020_2D -> Adanat_10020)
                            string cleanArticul = blockName.ToUpper().Trim();
                            if (cleanArticul.EndsWith("_2D")) cleanArticul = cleanArticul.Substring(0, cleanArticul.Length - 3).Trim();
                            if (cleanArticul.EndsWith("_3D")) cleanArticul = cleanArticul.Substring(0, cleanArticul.Length - 3).Trim();

                            // Поиск строки в считанном словаре Excel
                            if (excelData.TryGetValue(cleanArticul, out ExcelMafRowFiller excelRow))
                            {
                                ObjectId psId = ObjectId.Null;
                                ObjectIdCollection currentPropertySets = PropertyDataServices.GetPropertySets(br);

                                // Проверяем, насажен ли уже набор 07_МАФ на этот блок
                                foreach (ObjectId id in currentPropertySets)
                                {
                                    var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                    if (testPs != null && testPs.PropertySetDefinition == psdId)
                                    {
                                        psId = id;
                                        break;
                                    }
                                }

                                // Если набора нет, принудительно насаживаем его из чертежа
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

                                // Наполняем свойства Набора Характеристик точными данными из Excel
                                if (!psId.IsNull)
                                {
                                    var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                                    if (ps != null)
                                    {
                                        // Записываем данные с автоматическим обходом регистра букв свойств (ID, Type...)
                                        SafeSetProperty(ps, "ID", excelRow.MafId, tr);
                                        SafeSetProperty(ps, "Type", excelRow.TypeName, tr);
                                        SafeSetProperty(ps, "Возрастная группа", excelRow.AgeGroup, tr);
                                        SafeSetProperty(ps, "Габаритные размеры", excelRow.Dimensions, tr);
                                        SafeSetProperty(ps, "Код по классификатору зданий", excelRow.CodeClassifierBuilding, tr);
                                        SafeSetProperty(ps, "Код по классификатору элементов", excelRow.CodeClassifierElement, tr);
                                        SafeSetProperty(ps, "Наименование элемента", excelRow.Name, tr);
                                        SafeSetProperty(ps, "Позиция", excelRow.Position, tr);

                                        successCount++;
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }

            ed.WriteMessage($"\n[Успех] Обработка завершена. Заполнено параметров у блоков: {successCount}");
            ed.UpdateScreen();
        }

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
                    // Открываем дефиницию набора характеристик
                    ObjectId defId = ps.PropertySetDefinition;
                    DBObject defObj = tr.GetObject(defId, OpenMode.ForRead);

                    // Безопасное извлечение коллекции Properties через рефлексию (устраняет ошибку CS1061)
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
