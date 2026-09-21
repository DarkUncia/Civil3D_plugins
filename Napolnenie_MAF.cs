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

[assembly: CommandClass(typeof(Civil3D_plugins.MafDataPlugin))]

namespace Civil3D_plugins
{
    // Модель данных для малых архитектурных форм (МАФ) из Excel
    public class ExcelMafRow
    {
        public string MafId { get; set; }                // Имя блока (ID элемента)
        public string CodeClassifier { get; set; }       // Код по классификатору
        public string Designation { get; set; }          // Обозначение (ГОСТ/ТУ)
        public string Name { get; set; }                 // Наименование элемента
        public string Quantity { get; set; }             // Количество

        // Для МАФ всегда используется один фиксированный НХ
        public string GetTargetPsdName()
        {
            return "05_ДП_МАФ";
        }
    }
    public class ExcelMafReaderHelper
    {
        public static Dictionary<string, ExcelMafRow> ReadExcelData(string filePath, Editor ed)
        {
            var result = new Dictionary<string, ExcelMafRow>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    if (!reader.Read()) return null;
                    int idCol = -1, codeCol = -1, desCol = -1, nameCol = -1, qtyCol = -1;

                    // Поиск колонок по ключевым словам шапки МАФ
                    for (int col = 0; col < reader.FieldCount; col++)
                    {
                        string h = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? "";
                        h = h.Replace("\r", "").Replace("\n", "");

                        if (h.Contains("имя блока") || h.Contains("id") || h.Contains("марка покрытия")) idCol = col;
                        else if (h.Contains("классиф") && h.Contains("здан")) codeCol = col;
                        else if (h.Contains("обознач") || h.Contains("гост")) desCol = col;
                        else if (h.Contains("наименов") || h.Contains("элемент")) nameCol = col;
                        else if (h.Contains("кол-во") || h.Contains("количест")) qtyCol = col;
                    }

                    // Жесткие индексы по умолчанию (A, B, C, D, E), если ячейки объединены
                    if (idCol == -1) idCol = 0;
                    if (codeCol == -1) codeCol = 1;
                    if (desCol == -1) desCol = 2;
                    if (nameCol == -1) nameCol = 3;
                    if (qtyCol == -1) qtyCol = 4;

                    reader.Read(); // Пропускаем строку-пример

                    while (reader.Read())
                    {
                        string id = reader.GetValue(idCol)?.ToString()?.Trim() ?? "";
                        if (string.IsNullOrEmpty(id) || id.Contains("Пример") || id.Contains("Имя")) continue;

                        var row = new ExcelMafRow
                        {
                            MafId = id,
                            CodeClassifier = codeCol != -1 ? reader.GetValue(codeCol)?.ToString() ?? "" : "",
                            Designation = desCol != -1 ? reader.GetValue(desCol)?.ToString() ?? "" : "",
                            Name = nameCol != -1 ? reader.GetValue(nameCol)?.ToString() ?? "" : "",
                            Quantity = qtyCol != -1 ? reader.GetValue(qtyCol)?.ToString() ?? "" : ""
                        };
                        if (!result.ContainsKey(id)) result.Add(id, row);
                    }
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка Excel МАФ]: {ex.Message}"); return null; }
            return result;
        }
    }
    public class MafDataPlugin : IExtensionApplication
    {
        public void Initialize()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
            if (ed != null) ed.WriteMessage("\n>>> Плагин данных МАФ загружен! Команда: MAF_FILL_PROPERTIES <<<");
        }

        public void Terminate() { }

        [CommandMethod("MAF_FILL_PROPERTIES", CommandFlags.Modal)]
        public void FillMafPropertiesFromCurrentDoc()
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
                ofd.Title = "Выберите Excel-файл элементов МАФ";
                if (ofd.ShowDialog() != DialogResult.OK) return;
                excelFilePath = ofd.FileName;
            }

            var excelData = ExcelMafReaderHelper.ReadExcelData(excelFilePath, ed);
            if (excelData == null || excelData.Count == 0) return;

            string dwgDir = db.Filename != null && File.Exists(db.Filename) ? Path.GetDirectoryName(db.Filename) : Path.GetTempPath();
            string logFilePath = Path.Combine(dwgDir, "Отработка_МАФ_Лог.txt");
            StringBuilder log = new StringBuilder($"=== ОТЧЕТ МАФ {DateTime.Now} ===\nФайл: {Path.GetFileName(db.Filename)}\n\n");

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
                        // Ищем строго Вставки Блоков (AcDbBlockReference)
                        if (objId.ObjectClass.Name == "AcDbBlockReference")
                        {
                            var blockRef = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);
                            if (blockRef == null) continue;

                            // Получаем имя блока (с учетом динамических блоков)
                            string blockName = blockRef.IsDynamicBlock
                                ? ((BlockTableRecord)tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name.Trim()
                                : blockRef.Name.Trim();

                            if (excelData.TryGetValue(blockName, out ExcelMafRow excelRow))
                            {
                                string targetPsdName = excelRow.GetTargetPsdName(); // "05_ДП_МАФ"
                                if (!psdDict.Has(targetPsdName, tr)) continue;

                                ObjectId targetPsdId = psdDict.GetAt(targetPsdName);

                                // Зачистка старого набора "05_ДП_МАФ" на объекте перед записью нового
                                try { PropertyDataServices.RemovePropertySet(blockRef, targetPsdId); } catch { }

                                // Привязка чистого набора характеристик
                                PropertyDataServices.AddPropertySet(blockRef, targetPsdId);

                                ObjectId psId = ObjectId.Null;
                                foreach (ObjectId id in PropertyDataServices.GetPropertySets(blockRef))
                                {
                                    var testPs = (PropertySet)tr.GetObject(id, OpenMode.ForRead);
                                    if (testPs != null && testPs.PropertySetDefinition == targetPsdId) { psId = id; break; }
                                }

                                // Запись данных из Excel в НХ блока
                                if (!psId.IsNull)
                                {
                                    var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                                    if (ps != null)
                                    {
                                        SetProperty(ps, "Наименование_по_классификатору", excelRow.Name);
                                        SetProperty(ps, "Код_по_классификатору_зданий", excelRow.CodeClassifier);
                                        SetProperty(ps, "Обозначение", excelRow.Designation);
                                        SetProperty(ps, "Количество", excelRow.Quantity);

                                        blockRef.RecordGraphicsModified(true);
                                        processedCount++;
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                ed.WriteMessage($"\n[Успех] Готово! Наполнено блоков МАФ: {processedCount}.");
                log.AppendLine($"[УСПЕХ] Обработано блоков МАФ: {processedCount}");
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка МАФ]: {ex.Message}"); log.AppendLine($"[СБОЙ]: {ex.Message}"); }

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
