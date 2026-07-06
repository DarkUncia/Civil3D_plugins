using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using ExcelDataReader;

[assembly: CommandClass(typeof(Civil3D_plugins.MafLibraryGenerator))]

namespace Civil3D_plugins
{
    // Класс для хранения считанных данных из одной строки Excel
    public class ExcelRowData
    {
        public string Guid { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
    }

    // Первая часть основного класса (Команда запуска конвейера)
    public partial class MafLibraryGenerator
    {
        [CommandMethod("C3D_GENERATE_MAF_LIBRARY", CommandFlags.Modal)]
        public void RunLibraryGeneration()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            IntPtr windowHandle = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Handle;
            if (windowHandle == IntPtr.Zero && doc.Window != null)
            {
                windowHandle = doc.Window.Handle;
            }
            IWin32Window ownerWindow = windowHandle != IntPtr.Zero ? NativeWindow.FromHandle(windowHandle) : null;

            ed.WriteMessage("\n[Инфо] Старт выбора исходных данных...");

            // 1. ВЫБОР EXCEL
            string excelPath = string.Empty;
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
                ofd.Title = "Шаг 1: Выберите Excel-файл с ведомостью МАФ";
                if (ofd.ShowDialog(ownerWindow) != DialogResult.OK) return;
                excelPath = ofd.FileName;
            }

            var excelRows = ReadExcelData(excelPath, ed);
            if (excelRows == null || excelRows.Count == 0) return;

            // 2. ВЫБОР ИСХОДНОЙ ПАПКИ
            string sourceFolder = string.Empty;
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Шаг 2: Выберите папку с исходными чертежами (*_2D.dwg и *_3D.dwg)";
                if (fbd.ShowDialog(ownerWindow) != DialogResult.OK) return;
                sourceFolder = fbd.SelectedPath;
            }

            // 3. ВЫБОР ЦЕЛЕВОЙ ПАПКИ
            string targetFolder = string.Empty;
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Шаг 3: Выберите целевую папку для сохранения библиотеки";
                if (fbd.ShowDialog(ownerWindow) != DialogResult.OK) return;
                targetFolder = fbd.SelectedPath;
            }

            string[] dwgFiles = Directory.GetFiles(sourceFolder, "*.dwg", SearchOption.TopDirectoryOnly);
            if (dwgFiles.Length == 0)
            {
                ed.WriteMessage("\n[Инфо] В папке не найдено файлов .dwg.");
                return;
            }

            int successCount = 0;
            ed.WriteMessage($"\n>>> Старт конвейера Civil 3D. Файлов к обработке: {dwgFiles.Length} <<<");

            foreach (string file in dwgFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file).Trim();
                string dimensionPrefix = "2D";
                string cleanIdForMatch = fileName;

                // Определение типа блока (2D или 3D) по окончанию имени файла
                if (fileName.EndsWith("_2D", StringComparison.OrdinalIgnoreCase))
                {
                    dimensionPrefix = "2D";
                    cleanIdForMatch = fileName.Substring(0, fileName.Length - 3).Trim();
                }
                else if (fileName.EndsWith("_3D", StringComparison.OrdinalIgnoreCase))
                {
                    dimensionPrefix = "3D";
                    cleanIdForMatch = fileName.Substring(0, fileName.Length - 3).Trim();
                }

                // Ищем строку в Excel по очищенному артикулу (Id)
                ExcelRowData matchedRow = FindExcelMatchById(cleanIdForMatch, excelRows);
                if (matchedRow == null)
                {
                    ed.WriteMessage($"\n[Пропуск] Артикул '{cleanIdForMatch}' (из файла {fileName}) не найден в столбце ID таблицы Excel.");
                    continue;
                }

                if (ProcessSideDatabase(file, targetFolder, matchedRow, dimensionPrefix, ed))
                {
                    successCount++;
                }
            }

            ed.WriteMessage($"\n>>> [Успех] Обработка завершена! Успешно создано файлов: {successCount} из {dwgFiles.Length} <<<");
        }
    }
}
namespace Civil3D_plugins
{
    public partial class MafLibraryGenerator
    {
        /// <summary>
        /// Умный поиск строки в списке данных Excel по артикулу (Id).
        /// </summary>
        private ExcelRowData FindExcelMatchById(string idToFind, List<ExcelRowData> excelRows)
        {
            if (string.IsNullOrEmpty(idToFind)) return null;

            // 1. Очищаем имя файла от возможных технических суффиксов, ломающих логику
            string cleanFileId = idToFind.ToUpper().Trim();
            if (cleanFileId.EndsWith("_#1")) cleanFileId = cleanFileId.Substring(0, cleanFileId.Length - 3).Trim();
            if (cleanFileId.EndsWith("-#1")) cleanFileId = cleanFileId.Substring(0, cleanFileId.Length - 3).Trim();

            // Стандартизируем разделители в имени файла для поиска
            string fileIdWithDash = cleanFileId.Replace('_', '-');
            string fileIdWithUnderscore = cleanFileId.Replace('-', '_');

            // 2. Первый проход: Ищем строгое совпадение артикула (для ваших первых 20 файлов)
            foreach (var row in excelRows)
            {
                if (string.IsNullOrEmpty(row.Id)) continue;
                string excelId = row.Id.ToUpper().Trim();

                if (excelId == cleanFileId || excelId == fileIdWithDash || excelId == fileIdWithUnderscore)
                {
                    return row;
                }
            }

            // 3. Второй проход: Ищем частичное совпадение (для файлов с длинными описаниями типа качелей)
            foreach (var row in excelRows)
            {
                if (string.IsNullOrEmpty(row.Id)) continue;
                string excelId = row.Id.ToUpper().Trim();

                if (cleanFileId.Contains(excelId) ||
                    fileIdWithDash.Contains(excelId.Replace('_', '-')) ||
                    excelId.Contains(cleanFileId))
                {
                    return row;
                }
            }

            return null;
        }

        /// <summary>
        /// Чтение данных из файла Excel с динамическим поиском индексов колонок.
        /// </summary>
        private List<ExcelRowData> ReadExcelData(string filePath, Autodesk.AutoCAD.EditorInput.Editor ed)
        {
            var result = new List<ExcelRowData>();
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                    {
                        if (!reader.Read()) return null;

                        int guidCol = -1, idCol = -1, nameCol = -1, typeCol = -1;
                        for (int col = 0; col < reader.FieldCount; col++)
                        {
                            string headerText = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? string.Empty;
                            if (headerText == "guid") guidCol = col;
                            else if (headerText == "id") idCol = col;
                            else if (headerText == "name") nameCol = col;
                            else if (headerText == "type") typeCol = col;
                        }

                        if (guidCol == -1 || idCol == -1 || nameCol == -1 || typeCol == -1)
                        {
                            ed.WriteMessage("\n[Ошибка Excel] В шапке таблицы не найдены обязательные столбцы 'guid', 'id', 'name' или 'type'!");
                            return null;
                        }

                        while (reader.Read())
                        {
                            string gVal = reader.GetValue(guidCol)?.ToString() ?? "";
                            string idVal = reader.GetValue(idCol)?.ToString() ?? "";
                            string nameVal = reader.GetValue(nameCol)?.ToString() ?? "";
                            string typeVal = reader.GetValue(typeCol)?.ToString() ?? "";

                            if (string.IsNullOrEmpty(gVal) || string.IsNullOrEmpty(idVal)) continue;

                            result.Add(new ExcelRowData
                            {
                                Guid = gVal.Trim(),
                                Id = idVal.Trim(),
                                Name = nameVal.Trim(),
                                Type = typeVal.Trim()
                            });
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
}
namespace Civil3D_plugins
{
    public partial class MafLibraryGenerator
    {
        /// <summary>
        /// Фоновый процесс модификации базы данных .dwg чертежа.
        /// </summary>
        private bool ProcessSideDatabase(string sourceFile, string targetFolder, ExcelRowData excelData, string dimensionPrefix, Autodesk.AutoCAD.EditorInput.Editor ed)
        {
            // Формируем финальное имя файла по маске: 2D_Id~Type~0~7.dwg или 3D_Id~Type~0~7.dwg
            string newFileName = $"{dimensionPrefix}_{excelData.Id}~{excelData.Type}~0~7.dwg";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                newFileName = newFileName.Replace(c, '_');
            }
            string targetPath = System.IO.Path.Combine(targetFolder, newFileName);

            using (Autodesk.AutoCAD.DatabaseServices.Database sideDb = new Autodesk.AutoCAD.DatabaseServices.Database(false, true))
            {
                try
                {
                    sideDb.ReadDwgFile(sourceFile, System.IO.FileShare.ReadWrite, true, "");
                    sideDb.CloseInput(true);

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction tr = sideDb.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable bt = (Autodesk.AutoCAD.DatabaseServices.BlockTable)tr.GetObject(sideDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        Autodesk.AutoCAD.DatabaseServices.ObjectId oldBlockId = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null;

                        // Шаг 1: Ищем блок, имя которого изначально совпадает с артикулом (Id из Excel)
                        foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId bId in bt)
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(bId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            if (!btr.IsLayout && !btr.IsAnonymous)
                            {
                                string cleanBlockName = btr.Name.Replace('_', '-').ToUpper().Trim();
                                string cleanExcelId = excelData.Id.Replace('_', '-').ToUpper().Trim();

                                if (cleanBlockName == cleanExcelId || cleanBlockName.Contains(cleanExcelId))
                                {
                                    oldBlockId = bId;
                                    break;
                                }
                            }
                        }

                        // Запасной вариант на случай, если точного совпадения по имени блока нет
                        if (oldBlockId.IsNull)
                        {
                            foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId bId in bt)
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(bId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                if (!btr.IsLayout && !btr.IsAnonymous)
                                {
                                    oldBlockId = bId;
                                    break;
                                }
                            }
                        }

                        if (oldBlockId.IsNull)
                        {
                            ed.WriteMessage($"\n[Ошибка] В файле '{System.IO.Path.GetFileName(sourceFile)}' не найдено блоков для переименования.");
                            return false;
                        }

                        // Шаг 2: Переименовываем найденный блок в GUID из ячейки Excel
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord targetBtr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(oldBlockId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (targetBtr.Name.ToUpper() != excelData.Guid.ToUpper())
                        {
                            targetBtr.Name = excelData.Guid;
                        }

                        // Шаг 3: Обертываем переименованный блок в новый контейнер с суффиксом #1 (GUID#1)
                        string wrapperBlockName = $"{excelData.Guid}#1";
                        if (!bt.Has(wrapperBlockName))
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord wrapperBtr = new Autodesk.AutoCAD.DatabaseServices.BlockTableRecord())
                            {
                                wrapperBtr.Name = wrapperBlockName;
                                wrapperBtr.Origin = new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0);
                                bt.Add(wrapperBtr);
                                tr.AddNewlyCreatedDBObject(wrapperBtr, true);

                                using (Autodesk.AutoCAD.DatabaseServices.BlockReference br = new Autodesk.AutoCAD.DatabaseServices.BlockReference(new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0), oldBlockId))
                                {
                                    wrapperBtr.AppendEntity(br);
                                    tr.AddNewlyCreatedDBObject(br, true);
                                }
                            }
                        }

                        // Шаг 4: Полностью очищаем ModelSpace чертежа от старой геометрии
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord modelSpace = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(bt[Autodesk.AutoCAD.DatabaseServices.BlockTableRecord.ModelSpace], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId entId in modelSpace)
                        {
                            Autodesk.AutoCAD.DatabaseServices.DBObject obj = tr.GetObject(entId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            if (!obj.IsErased)
                            {
                                obj.Erase();
                            }
                        }

                        // Шаг 5: Вставляем только что созданный блок-контейнер (GUID#1) в очищенный ModelSpace
                        Autodesk.AutoCAD.DatabaseServices.ObjectId wrapperBlockId = bt[wrapperBlockName];
                        using (Autodesk.AutoCAD.DatabaseServices.BlockReference modelBr = new Autodesk.AutoCAD.DatabaseServices.BlockReference(new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0), wrapperBlockId))
                        {
                            modelSpace.AppendEntity(modelBr);
                            tr.AddNewlyCreatedDBObject(modelBr, true);
                        }

                        tr.Commit();
                    }

                    // Шаг 6: Безопасно выгружаем измененную фоновую базу данных в новый файл через Wblock
                    using (Autodesk.AutoCAD.DatabaseServices.Database targetDb = sideDb.Wblock())
                    {
                        targetDb.SaveAs(targetPath, Autodesk.AutoCAD.DatabaseServices.DwgVersion.Current);
                    }
                    return true;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[Критическая ошибка файла {System.IO.Path.GetFileName(sourceFile)}]: {ex.Message}");
                    return false;
                }
            }
        }
    }
}
