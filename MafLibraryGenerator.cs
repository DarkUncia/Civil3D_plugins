using Autodesk.Aec.PropertyData.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Civil3D_plugins;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Civil3D_plugins
{
    // Модель данных для малых архитектурных форм (МАФ) из Excel
    public class ExcelMafRow
    {
        public string MafId { get; set; }                    // ID элемента (Имя блока)
        public string CodeClassifierBuilding { get; set; }   // Код по классификатору зданий
        public string CodeClassifierElement { get; set; }    // Код по классификатору элементов
        public string Position { get; set; }                 // Позиция
        public string Name { get; set; }                     // Наименование элемента
        public string AgeGroup { get; set; }                 // Возрастная группа
        public string Dimensions { get; set; }               // Габаритные размеры
        public string TypeName { get; set; }                 // Type (последняя колонка)

        // Имя целевого набора характеристик
        public string GetTargetPsdName() => "07_МАФ";
    }
}
public class ExcelReaderHelper
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
                int idCol = -1, codeBldCol = -1, codeElCol = -1, posCol = -1, nameCol = -1, ageCol = -1, dimCol = -1, typeCol = -1;

                // Поиск индексов колонок по именам из новой шапки
                for (int col = 0; col < reader.FieldCount; col++)
                {
                    string h = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? "";
                    h = h.Replace("\r", "").Replace("\n", "");
                    if (h == "id") idCol = col;
                    else if (h.Contains("классификатору зданий")) codeBldCol = col;
                    else if (h.Contains("классификатору элементов")) codeElCol = col;
                    else if (h == "позиция") posCol = col;
                    else if (h.Contains("наименование элемента")) nameCol = col;
                    else if (h.Contains("возрастная группа")) ageCol = col;
                    else if (h.Contains("габаритные размеры")) dimCol = col;
                    else if (h == "type") typeCol = col;
                }

                // Индексы по умолчанию на случай, если ячейки объединены
                if (idCol == -1) idCol = 0; if (codeBldCol == -1) codeBldCol = 1; if (codeElCol == -1) codeElCol = 2; if (posCol == -1) posCol = 3;
                if (nameCol == -1) nameCol = 4; if (ageCol == -1) ageCol = 5; if (dimCol == -1) dimCol = 6; if (typeCol == -1) typeCol = 7;

                reader.Read(); // ГАРАНТИРОВАННО ИГНОРИРУЕМ строку-пояснение под шапкой ("Артикул", "X.X.X.X"...)

                // Построчное чтение параметров МАФ
                while (reader.Read())
                {
                    string id = reader.GetValue(idCol)?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(id) || id.Equals("Артикул", StringComparison.OrdinalIgnoreCase)) continue;

                    var row = new ExcelMafRow
                    {
                        MafId = id,
                        CodeClassifierBuilding = codeBldCol != -1 ? reader.GetValue(codeBldCol)?.ToString()?.Trim() ?? "" : "",
                        CodeClassifierElement = codeElCol != -1 ? reader.GetValue(codeElCol)?.ToString()?.Trim() ?? "" : "",
                        Position = posCol != -1 ? reader.GetValue(posCol)?.ToString()?.Trim() ?? "" : "",
                        Name = nameCol != -1 ? reader.GetValue(nameCol)?.ToString()?.Trim() ?? "" : "",
                        AgeGroup = ageCol != -1 ? reader.GetValue(ageCol)?.ToString()?.Trim() ?? "" : "",
                        Dimensions = dimCol != -1 ? reader.GetValue(dimCol)?.ToString()?.Trim() ?? "" : "",
                        TypeName = typeCol != -1 ? reader.GetValue(typeCol)?.ToString()?.Trim() ?? "" : ""
                    };
                    if (!result.ContainsKey(id)) result.Add(id, row);
                }
            }
        }
        catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка Excel]: {ex.Message}"); return null; }
        return result;
    }
}
    public class MafLibraryGenerator
{
    [CommandMethod("MAF_BUILD_LIBRARY", CommandFlags.Modal)]
    public void BuildLibraryProcess()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (activeDoc == null) return;
        Editor ed = activeDoc.Editor; Database db = activeDoc.Database;

        string excelFilePath = "";
        using (OpenFileDialog ofd = new OpenFileDialog())
        {
            ofd.Filter = "Excel (*.xlsx)|*.xlsx"; ofd.Title = "Выберите новую БД Excel элементов МАФ";
            if (ofd.ShowDialog() != DialogResult.OK) return;
            excelFilePath = ofd.FileName;
        }
        var excelData = ExcelReaderHelper.ReadExcelData(excelFilePath, ed);
        if (excelData == null || excelData.Count == 0) { ed.WriteMessage("\n[Ошибка]: БД Excel пуста."); return; }

        string sourceFolder = "";
        using (FolderBrowserDialog fbd = new FolderBrowserDialog())
        {
            fbd.Description = "Выберите папку с исходными файлами блоков 2D и 3D";
            if (fbd.ShowDialog() != DialogResult.OK) return;
            sourceFolder = fbd.SelectedPath;
        }
        string baseOutputDir = Path.Combine(Path.GetDirectoryName(excelFilePath), "Экспорт_МАФ");
        Directory.CreateDirectory(baseOutputDir);
        int successCount = 0; int errorCount = 0;

        using (DocumentLock docLock = activeDoc.LockDocument())
        {
            foreach (var pair in excelData)
            {
                string id = pair.Key; ExcelMafRow row = pair.Value;
                string file2D = Path.Combine(sourceFolder, $"{id}_2D.dwg");
                string file3D = Path.Combine(sourceFolder, $"{id}_3D.dwg");
                if (!File.Exists(file2D) || !File.Exists(file3D)) continue;
                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        string subBlockName2D = $"{id}_INTERNAL_2D";
                        string subBlockName3D = $"{id}_INTERNAL_3D";
                        string matreshkaBlockName = $"{id}#1";
                        ObjectId id2D = ObjectId.Null; ObjectId id3D = ObjectId.Null;

                        using (Database sourceDb = new Database(false, true))
                        { sourceDb.ReadDwgFile(file2D, FileShare.ReadWrite, true, ""); id2D = db.Insert(subBlockName2D, sourceDb, false); }
                        using (Database sourceDb = new Database(false, true))
                        { sourceDb.ReadDwgFile(file3D, FileShare.ReadWrite, true, ""); id3D = db.Insert(subBlockName3D, sourceDb, false); }
                        var psdDict = new DictionaryPropertySetDefinitions(db);
                        string targetPsdName = row.GetTargetPsdName();
                        if (psdDict.Has(targetPsdName, tr))
                        {
                            ObjectId targetPsdId = psdDict.GetAt(targetPsdName);
                            ObjectId[] internalBlockIds = new ObjectId[] { id2D, id3D };
                            foreach (ObjectId internalBtrId in internalBlockIds)
                            {
                                var internalBtr = (BlockTableRecord)tr.GetObject(internalBtrId, OpenMode.ForWrite);

                                // Удаление всех старых наборов характеристик
                                foreach (ObjectId oldPsId in PropertyDataServices.GetPropertySets(internalBtr))
                                { try { PropertyDataServices.RemovePropertySet(internalBtr, oldPsId); } catch { } }

                                // Накатываем чистый НХ 07_МАФ
                                PropertyDataServices.AddPropertySet(internalBtr, targetPsdId);
                                ObjectId psId = ObjectId.Null;
                                foreach (ObjectId pId in PropertyDataServices.GetPropertySets(internalBtr))
                                {
                                    var testPs = (PropertySet)tr.GetObject(pId, OpenMode.ForRead);
                                    if (testPs != null && testPs.PropertySetDefinition == targetPsdId) { psId = pId; break; }
                                }

                                // Заполнение свойств точными данными из базы Excel (соответствует Диспетчеру стилей)
                                if (!psId.IsNull)
                                {
                                    var ps = (PropertySet)tr.GetObject(psId, OpenMode.ForWrite);
                                    if (ps != null)
                                    {
                                        SetProperty(ps, "ID", row.MafId);
                                        SetProperty(ps, "Type", row.TypeName);
                                        SetProperty(ps, "Возрастная группа", row.AgeGroup);
                                        SetProperty(ps, "Габаритные размеры", row.Dimensions);
                                        SetProperty(ps, "Код по классификатору зданий", row.CodeClassifierBuilding);
                                        SetProperty(ps, "Код по классификатору элементов", row.CodeClassifierElement);
                                        SetProperty(ps, "Наименование элемента", row.Name);
                                        SetProperty(ps, "Позиция", row.Position);
                                    }
                                }
                            }
                        }
                        // ---- ШАГ 3: СБОРКА РОДИТЕЛЬСКОЙ МАТРЕШКИ С СУФФИКСОМ #1 ----
                        var matreshkaBtr = new BlockTableRecord { Name = matreshkaBlockName };
                        ObjectId matreshkaBtrId = bt.Add(matreshkaBtr);
                        tr.AddNewlyCreatedDBObject(matreshkaBtr, true);

                        var ref2D = new BlockReference(Autodesk.AutoCAD.Geometry.Point3d.Origin, id2D);
                        matreshkaBtr.AppendEntity(ref2D); tr.AddNewlyCreatedDBObject(ref2D, true);
                        var ref3D = new BlockReference(Autodesk.AutoCAD.Geometry.Point3d.Origin, id3D);
                        matreshkaBtr.AppendEntity(ref3D); tr.AddNewlyCreatedDBObject(ref3D, true);
                        tr.Commit();

                        // ---- ШАГ 4: ЭКСПОРТ В НОВЫЙ DWG С ОПРЕДЕЛЕННЫМ ИМЕНЕМ ----
                        string provider = id.Contains("_") ? id.Split('_')[0] : "Разное";
                        string providerDir = Path.Combine(baseOutputDir, provider);
                        Directory.CreateDirectory(providerDir);
                        string cleanTypeName = row.TypeName.Replace("\\", "_").Replace("/", "_");
                        string newFileName = $"{id}~{cleanTypeName}~0~7.dwg";
                        string finalDwgPath = Path.Combine(providerDir, newFileName);

                        using (Database outputDb = db.Wblock(matreshkaBtrId))
                        { outputDb.SaveAs(finalDwgPath, DwgVersion.Current); }
                        successCount++;
                    }
                }
                catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка {id}]: {ex.Message}"); errorCount++; }
            }
        }
        ed.WriteMessage($"\n\n=== ЗАВЕРШЕНО ===\nУспешно создано безопасных матрешек: {successCount}, Ошибок: {errorCount}");
    }

    private bool SetProperty(PropertySet ps, string propName, object value)
    {
        try
        {
            int id = ps.PropertyNameToId(propName);
            ps.SetAt(id, value?.ToString() ?? ""); return true;
        }
        catch { return false; }
    }
}
