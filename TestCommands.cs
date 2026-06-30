using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;
using ExcelDataReader; // Библиотека чтения Excel

[assembly: CommandClass(typeof(Civil3D_plugins.TestMafCommands))]

namespace Civil3D_plugins
{
    public class TestMafCommands : IExtensionApplication
    {
        public const string PsdName = "07_МАФ";

        public void Initialize()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor
                .WriteMessage("\n>>> Плагин связи МАФ с Excel успешно загружен! <<<");
        }

        public void Terminate()
        {
        }

        [CommandMethod("TEST_FILL_MAF", CommandFlags.Modal)]
        public void TestFillMultipleBlocks()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string excelFilePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                openFileDialog.Title = "Выберите Excel-файл с ведомостью МАФ";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                }
            }

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                ed.WriteMessage("\n[Инфо] Выбор файла отменен или файл не найден.");
                return;
            }

            var excelData = ReadExcelData(excelFilePath, ed);
            if (excelData == null || excelData.Count == 0)
            {
                ed.WriteMessage("\n[Ошибка] Не удалось прочитать данные из Excel или таблица пуста.");
                return;
            }

            using (doc.LockDocument())
            {
                ObjectId psdId = ObjectId.Null;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var psdDict = new DictionaryPropertySetDefinitions(db);
                    if (!psdDict.Has(PsdName, tr))
                    {
                        ed.WriteMessage($"\n[Ошибка] Набор характеристик '{PsdName}' не найден в этом чертеже!");
                        return;
                    }
                    psdId = psdDict.GetAt(PsdName);
                    tr.Commit();
                }

                int processedCount = 0;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId objId in modelSpace)
                    {
                        if (objId.ObjectClass.Name == "AcDbBlockReference")
                        {
                            var blockRef = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);
                            if (blockRef == null) continue;

                            string blockName = blockRef.IsDynamicBlock
                                ? ((BlockTableRecord)tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                                : blockRef.Name;

                            if (excelData.TryGetValue(blockName.ToUpper().Trim(), out ExcelMafRow excelRow))
                            {
                                ObjectId psId = ObjectId.Null;
                                ObjectIdCollection currentPropertySets = PropertyDataServices.GetPropertySets(blockRef);
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
                                    PropertyDataServices.AddPropertySet(blockRef, psdId);
                                    ObjectIdCollection updatedPropertySets = PropertyDataServices.GetPropertySets(blockRef);
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
                                        SetProperty(ps, "id", excelRow.Id);
                                        SetProperty(ps, "manufacturer", excelRow.Manufacturer);
                                        SetProperty(ps, "name", excelRow.Name);
                                        SetProperty(ps, "dimensions", excelRow.Dimensions);
                                        SetProperty(ps, "note", excelRow.Note);
                                        SetProperty(ps, "guid", blockName);
                                        SetProperty(ps, "type", excelRow.Type);
                                        SetProperty(ps, "type2", excelRow.Type2);
                                        SetProperty(ps, "weight", excelRow.Weight);
                                        SetProperty(ps, "group_id", excelRow.GroupId);
                                        SetProperty(ps, "Version", excelRow.Version);
                                        SetProperty(ps, "ADSK_Name", excelRow.AdskName);
                                        SetProperty(ps, "ADSK_Set_of_drawings", excelRow.AdskSetOfDrawings);
                                        SetProperty(ps, "ACER_CodeCollision", excelRow.AcerCodeCollision);
                                        SetProperty(ps, "KRTRS_Code_by_classifier", excelRow.KrtrsCodeByClassifier);

                                        processedCount++;
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                ed.WriteMessage($"\n>>> [Успех] Синхронизация завершена! Успешно обновлено блоков: {processedCount}. <<<");
            }
            ed.UpdateScreen();
        }
        private Dictionary<string, ExcelMafRow> ReadExcelData(string filePath, Editor ed)
        {
            var result = new Dictionary<string, ExcelMafRow>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        if (!reader.Read()) return null;

                        int guidCol = -1, idCol = -1, manCol = -1, nameCol = -1, dimCol = -1, noteCol = -1;
                        int typeCol = -1, type2Col = -1, weightCol = -1, groupIdCol = -1, versionCol = -1;
                        int adskNameCol = -1, adskSetCol = -1, acerCol = -1, krtrsCol = -1;

                        for (int col = 0; col < reader.FieldCount; col++)
                        {
                            string headerText = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? string.Empty;

                            if (headerText == "guid") guidCol = col;
                            else if (headerText == "id") idCol = col;
                            else if (headerText == "manufacturer") manCol = col;
                            else if (headerText == "name") nameCol = col;
                            else if (headerText == "dimensions") dimCol = col;
                            else if (headerText == "note") noteCol = col;
                            else if (headerText == "type") typeCol = col;
                            else if (headerText == "type2") type2Col = col;
                            else if (headerText == "weight") weightCol = col;
                            else if (headerText == "group_id") groupIdCol = col;
                            else if (headerText == "version") versionCol = col;
                            else if (headerText == "adsk_name") adskNameCol = col;
                            else if (headerText == "adsk_set_of_drawings") adskSetCol = col;
                            else if (headerText == "acer_codecollision") acerCol = col;
                            else if (headerText == "krtrs_code_by_classifier") krtrsCol = col;
                        }

                        if (guidCol == -1)
                        {
                            ed.WriteMessage("\n[Ошибка Excel] В первой строчке таблицы не найден обязательный латинский заголовок 'guid'!");
                            return null;
                        }

                        while (reader.Read())
                        {
                            string rowGuid = reader.GetValue(guidCol)?.ToString()?.ToUpper()?.Trim() ?? string.Empty;
                            if (string.IsNullOrEmpty(rowGuid)) continue;

                            var rowData = new ExcelMafRow
                            {
                                Id = idCol != -1 ? reader.GetValue(idCol)?.ToString() ?? "" : "",
                                Manufacturer = manCol != -1 ? reader.GetValue(manCol)?.ToString() ?? "" : "",
                                Name = nameCol != -1 ? reader.GetValue(nameCol)?.ToString() ?? "" : "",
                                Dimensions = dimCol != -1 ? reader.GetValue(dimCol)?.ToString() ?? "" : "",
                                Note = noteCol != -1 ? reader.GetValue(noteCol)?.ToString() ?? "" : "",
                                Type = typeCol != -1 ? reader.GetValue(typeCol)?.ToString() ?? "" : "",
                                Type2 = type2Col != -1 ? reader.GetValue(type2Col)?.ToString() ?? "" : "",
                                Weight = weightCol != -1 ? reader.GetValue(weightCol)?.ToString() ?? "" : "",
                                GroupId = groupIdCol != -1 ? reader.GetValue(groupIdCol)?.ToString() ?? "" : "",
                                Version = versionCol != -1 ? reader.GetValue(versionCol)?.ToString() ?? "" : "",
                                AdskName = adskNameCol != -1 ? reader.GetValue(adskNameCol)?.ToString() ?? "" : "",
                                AdskSetOfDrawings = adskSetCol != -1 ? reader.GetValue(adskSetCol)?.ToString() ?? "" : "",
                                AcerCodeCollision = acerCol != -1 ? reader.GetValue(acerCol)?.ToString() ?? "" : "",
                                KrtrsCodeByClassifier = krtrsCol != -1 ? reader.GetValue(krtrsCol)?.ToString() ?? "" : ""
                            };

                            if (!result.ContainsKey(rowGuid))
                            {
                                result.Add(rowGuid, rowData);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n>>> [Ошибка чтения Excel]: {ex.Message} <<<");
                return null;
            }

            return result;
        }
        private void SetProperty(PropertySet ps, string propName, object value)
        {
            try
            {
                int propId = ps.PropertyNameToId(propName);
                string strValue = value?.ToString() ?? string.Empty;
                ps.SetAt(propId, strValue);
            }
            catch
            {
                // Игнорируем ошибки отсутствия свойств в наборе
            }
        }
    }

    public class ExcelMafRow
    {
        public string Id { get; set; }
        public string Manufacturer { get; set; }
        public string Name { get; set; }
        public string Dimensions { get; set; }
        public string Note { get; set; }
        public string Type { get; set; }
        public string Type2 { get; set; }
        public string Weight { get; set; }
        public string GroupId { get; set; }
        public string Version { get; set; }
        public string AdskName { get; set; }
        public string AdskSetOfDrawings { get; set; }
        public string AcerCodeCollision { get; set; }
        public string KrtrsCodeByClassifier { get; set; }
    }
}
