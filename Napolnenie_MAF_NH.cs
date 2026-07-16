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

[assembly: CommandClass(typeof(Civil3D_plugins.TestMafCommands))]

namespace Civil3D_plugins
{
    // Класс-модель для хранения данных одной строки из Excel
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

    public class ExcelDataReaderHelper
    {
        public static Dictionary<string, ExcelMafRow> ReadExcelData(string filePath, Editor ed)
        {
            // Регистр-независимый словарь, где КЛЮЧОМ является ID (Артикул) элемента из Excel
            var result = new Dictionary<string, ExcelMafRow>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                    {
                        if (!reader.Read()) return null;

                        int idCol = -1, manCol = -1, nameCol = -1, dimCol = -1, noteCol = -1;
                        int typeCol = -1, type2Col = -1, weightCol = -1, groupIdCol = -1, versionCol = -1;
                        int adskNameCol = -1, adskSetCol = -1, acerCol = -1, krtrsCol = -1;

                        for (int col = 0; col < reader.FieldCount; col++)
                        {
                            string headerText = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? string.Empty;

                            if (headerText == "id") idCol = col;
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

                        if (idCol == -1)
                        {
                            ed.WriteMessage("\n[Ошибка Excel] В первой строчке таблицы не найден обязательный латинский заголовок 'id'!");
                            return null;
                        }

                        while (reader.Read())
                        {
                            string rowId = reader.GetValue(idCol)?.ToString()?.Trim() ?? string.Empty;
                            if (string.IsNullOrEmpty(rowId)) continue;

                            var rowData = new ExcelMafRow
                            {
                                Id = rowId,
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

                            if (!result.ContainsKey(rowId))
                            {
                                result.Add(rowId, rowData);
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
    }
    public class TestMafCommands : IExtensionApplication
    {
        public const string PsdName = "07_МАФ";

        public void Initialize()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
            if (ed != null)
            {
                ed.WriteMessage("\n>>> Плагин пакетного наполнения МАФ характеристиками успешно загружен! <<<");
                ed.WriteMessage("\n>>> Используйте команду: MAF_FILL_PROPERTIES <<<");
            }
        }

        public void Terminate() { }

        [CommandMethod("MAF_FILL_PROPERTIES", CommandFlags.Modal)]
        public void TestFillBatchFromFolder()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor;

            // 1. ВЫБОР ОДНОГО ФАЙЛА EXCEL
            string excelFilePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                openFileDialog.Title = "Выберите ОДИН Excel-файл с общей ведомостью МАФ";
                openFileDialog.Multiselect = false;
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    ed.WriteMessage("\n[Инфо] Выбор Excel-файла отменен.");
                    return;
                }
                excelFilePath = openFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                ed.WriteMessage("\n[Ошибка] Файл Excel не найден.");
                return;
            }

            // 2. ВЫБОР ПАПКИ С ЧЕРТЕЖАМИ DWG
            string folderPath = string.Empty;
            using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
            {
                folderBrowser.Description = "Выберите папку, содержащую чертежи DWG для обработки";
                folderBrowser.ShowNewFolderButton = false;
                if (folderBrowser.ShowDialog() != DialogResult.OK)
                {
                    ed.WriteMessage("\n[Инфо] Выбор папки отменен.");
                    return;
                }
                folderPath = folderBrowser.SelectedPath;
            }

            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                ed.WriteMessage("\n[Ошибка] Указанная папка не найдена.");
                return;
            }

            string[] dwgFiles = Directory.GetFiles(folderPath, "*.dwg");
            if (dwgFiles.Length == 0)
            {
                ed.WriteMessage("\n[Ошибка] В выбранной папке не найдены чертежи DWG.");
                return;
            }

            // Читаем общие данные из Excel
            var excelData = ExcelDataReaderHelper.ReadExcelData(excelFilePath, ed);
            if (excelData == null || excelData.Count == 0)
            {
                ed.WriteMessage("\n[Ошибка] Не удалось прочитать данные из Excel или таблица пуста.");
                return;
            }

            // ИНИЦИАЛИЗАЦИЯ ТЕКСТОВОГО ЛОГА В ВЫБРАННОЙ ПАПКЕ
            string logFilePath = Path.Combine(folderPath, "Отработка_МАФ_Лог.txt");
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"=== ОТЧЕТ ОБРАБОТКИ БЛОКОВ МАФ от {DateTime.Now} ===");
            logBuilder.AppendLine($"Файл Excel: {excelFilePath}");
            logBuilder.AppendLine($"Обрабатываемая папка: {folderPath}");
            logBuilder.AppendLine(new string('-', 80));

            Database currentDb = HostApplicationServices.WorkingDatabase;
            int totalUpdatedDrawings = 0;

            // 3. ПАКЕТНАЯ ОБРАБОТКА КАЖДОГО DWG ФАЙЛА В ПАПКЕ
            foreach (string dwgFilePath in dwgFiles)
            {
                string currentFileName = Path.GetFileName(dwgFilePath);
                ed.WriteMessage($"\n\n[Процесс] Обработка файла: {currentFileName}...");
                logBuilder.AppendLine($"\nФАЙЛ: {currentFileName}");

                using (Database db = new Database(false, true))
                {
                    try
                    {
                        db.ReadDwgFile(dwgFilePath, FileOpenMode.OpenForReadAndAllShare, false, null);
                        HostApplicationServices.WorkingDatabase = db;
                        int processedCount = 0;
                        ObjectId psdId = ObjectId.Null;

                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var psdDict = new DictionaryPropertySetDefinitions(db);
                            if (!psdDict.Has(PsdName, tr))
                            {
                                string noPsdMsg = $"  [ПРОПУЩЕН] Набор характеристик '{PsdName}' не найден в этом чертеже.";
                                ed.WriteMessage($"\n{noPsdMsg}");
                                logBuilder.AppendLine(noPsdMsg);
                                tr.Commit();
                                continue;
                            }
                            psdId = psdDict.GetAt(PsdName);

                            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                            foreach (ObjectId objId in modelSpace)
                            {
                                if (objId.ObjectClass.Name == "AcDbBlockReference")
                                {
                                    var blockRef = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);
                                    if (blockRef == null) continue;

                                    // Получаем настоящее имя главного блока в пространстве модели 
                                    // (с защитой от анонимных/динамических имен вроде *U...)
                                    string blockName = blockRef.IsDynamicBlock
                                        ? ((BlockTableRecord)tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                                        : blockRef.Name;

                                    ExcelMafRow excelRow = null;
                                    string matchedArticul = string.Empty;

                                    // УНИВЕРСАЛЬНАЯ ПРОВЕРКА: Ищем вхождение артикула из Excel в оригинальном имени блока
                                    foreach (var pair in excelData)
                                    {
                                        // pair.Key — это артикул из Excel (цифры, латиница или составной с дефисами)
                                        // Проверяем, содержится ли он целиком внутри имени блока (с игнорированием регистра букв)
                                        if (blockName.IndexOf(pair.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            excelRow = pair.Value;
                                            matchedArticul = pair.Key;
                                            break; // Совпадение найдено, прекращаем поиск для этого блока
                                        }
                                    }

                                    // Если артикул успешно сопоставлен со строкой из Excel
                                    if (excelRow != null)
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
                                                bool hasPropertyErrors = false;

                                                // Наполняем характеристики и собираем лог внутренних ошибок записи полей
                                                hasPropertyErrors |= !SetProperty(ps, "id", excelRow.Id, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "manufacturer", excelRow.Manufacturer, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "name", excelRow.Name, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "dimensions", excelRow.Dimensions, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "note", excelRow.Note, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "guid", blockName, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "type", excelRow.Type, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "type2", excelRow.Type2, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "weight", excelRow.Weight, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "group_id", excelRow.GroupId, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "Version", excelRow.Version, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "ADSK_Name", excelRow.AdskName, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "ADSK_Set_of_drawings", excelRow.AdskSetOfDrawings, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "ACER_CodeCollision", excelRow.AcerCodeCollision, logBuilder);
                                                hasPropertyErrors |= !SetProperty(ps, "KRTRS_Code_by_classifier", excelRow.KrtrsCodeByClassifier, logBuilder);

                                                if (!hasPropertyErrors)
                                                {
                                                    logBuilder.AppendLine($"  [УСПЕХ] Блок '{blockName}' успешно связан с артикулом '{matchedArticul}'");
                                                }
                                                else
                                                {
                                                    logBuilder.AppendLine($"  [ВНИМАНИЕ] Блок '{blockName}' (Артикул '{matchedArticul}') записан с частичными ошибками полей (см. ошибки свойств выше).");
                                                }

                                                processedCount++;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // КОСЯК: Блок верхнего уровня есть, но его имя не содержит ни один артикул из текущего Excel
                                        logBuilder.AppendLine($"  [НЕ НАЙДЕН АРТИКУЛ] Блок с именем '{blockName}' пропущен. Ни один ID из Excel не содержится в этом имени.");
                                    }
                                }
                            }
                            tr.Commit();
                        }

                        // Сохранение фоновой базы данных чертежа поверх исходного файла
                        db.SaveAs(dwgFilePath, false, DwgVersion.Current, db.SecurityParameters);
                        ed.WriteMessage($"\n[Успех] Файл {currentFileName} обработан. Найдено и заполнено блоков: {processedCount}.");
                        totalUpdatedDrawings++;
                    }
                    catch (System.Exception ex)
                    {
                        string critError = $"  [КРИТИЧЕСКАЯ ОШИБКА ФАЙЛА] Сбой при обработке файла: {ex.Message}";
                        ed.WriteMessage($"\n{critError}");
                        logBuilder.AppendLine(critError);
                    }
                    finally
                    {
                        HostApplicationServices.WorkingDatabase = currentDb;
                    }
                }
            }

            // СОХРАНЕНИЕ ТЕКСТОВОГО ЛОГА НА ДИСК
            try
            {
                logBuilder.AppendLine("\n" + new string('-', 80));
                logBuilder.AppendLine($"Пакетная обработка завершена. Успешно обновлено файлов: {totalUpdatedDrawings} из {dwgFiles.Length}");
                File.WriteAllText(logFilePath, logBuilder.ToString(), Encoding.UTF8);
                ed.WriteMessage($"\n\n>>> Пакетная обработка завершена! Создан подробный лог-файл: {logFilePath} <<<");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Ошибка записи лога] Не удалось сохранить файл отчета: {ex.Message}");
            }

            ed.UpdateScreen();
        }

        private bool SetProperty(PropertySet ps, string propName, object value, StringBuilder logBuilder)
        {
            try
            {
                int propId = ps.PropertyNameToId(propName);
                string strValue = value?.ToString() ?? string.Empty;
                ps.SetAt(propId, strValue);
                return true;
            }
            catch (System.Exception ex)
            {
                // Фиксируем в логе несовпадение имен полей или типов данных в диспетчере стилей
                logBuilder.AppendLine($"    └─ [ОШИБКА СВОЙСТВА] Не удалось записать значение в поле '{propName}'. Причина: {ex.Message}");
                return false;
            }
        } // Закрывает метод SetProperty
    } // Закрывает класс TestMafCommands
} // Закрывает namespace Civil3D_plugins