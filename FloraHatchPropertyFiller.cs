using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDataReader;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;
using Autodesk.AutoCAD.Colors;

[assembly: CommandClass(typeof(Civil3D_plugins.FloraHatchPropertyFiller))]
namespace Civil3D_plugins
{
    // 1. Модель строки Excel для штриховок озеленения FloraBIM
    public class FloraBIM_HatchRow
    {
        public string ColorRGB { get; set; } = string.Empty;       // Оригинальный "color" (для поиска)
        public string ColorAgo { get; set; } = string.Empty;       // Специальный "color_AGO"
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
        public string HatchPattern { get; set; } = string.Empty;
    }
    // 2. Умный бинарный парсер Excel с раздельным сбором оригинального цвета и цвета по АГО
    public static class FloraHatchExcelReader
    {
        // Метод нормализации RGB строки в строгий формат "R,G,B" без пробелов
        public static string NormalizeRgbString(string rawColor)
        {
            if (string.IsNullOrEmpty(rawColor)) return string.Empty;
            string clean = rawColor.Replace(" ", "").Trim();

            if (!clean.Contains(",") && clean.Length >= 7 && clean.Length <= 9)
            {
                try
                {
                    return $"{int.Parse(clean.Substring(0, 3))},{int.Parse(clean.Substring(3, 3))},{int.Parse(clean.Substring(6))}";
                }
                catch { }
            }

            string[] parts = clean.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 3 && int.TryParse(parts[0], out int r) && int.TryParse(parts[1], out int g) && int.TryParse(parts[2], out int b))
            {
                return $"{r},{g},{b}";
            }
            return clean;
        }

        public static Dictionary<string, FloraBIM_HatchRow> ReadGreeneryTable(string filePath, Editor ed)
        {
            var registry = new Dictionary<string, FloraBIM_HatchRow>(StringComparer.OrdinalIgnoreCase);
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int sheetIndex = 0;
                    do
                    {
                        ed.WriteMessage($"\n[Excel] Сканирую лист №{sheetIndex + 1}: {reader.Name}");

                        int colorCol = -1, colorAgoCol = -1, idCol = -1, classBldCol = -1, classElCol = -1, nameCol = -1, typeCol = -1, statusCol = -1, hatchCol = -1;
                        int xxxsCol = -1, xxsCol = -1, xsCol = -1, sCol = -1, mCol = -1, lCol = -1, xlCol = -1;

                        int localRowCount = 0;

                        while (reader.Read())
                        {
                            localRowCount++;

                            // ЖЕСТКАЯ БЛОКИРОВКА ВТОРОЙ СТРОКИ: Полностью пропускаем её, это примеры для человека!
                            if (localRowCount == 2)
                            {
                                continue;
                            }

                            // Если это первая строка (строка заголовков), считываем индексы колонок
                            if (localRowCount == 1)
                            {
                                for (int col = 0; col < reader.FieldCount; col++)
                                {
                                    string header = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? "";
                                    header = header.Replace("\r", "").Replace("\n", "");

                                    // РАЗДЕЛЬНЫЙ ПОИСК КОЛОНОК ЦВЕТА
                                    if (header == "color_ago" || header == "цвет_аго") colorAgoCol = col;
                                    else if (header == "color" || header == "цвет") colorCol = col;

                                    else if (header.Contains("id") || header.Contains("артикул")) idCol = col;
                                    else if (header.Contains("классиф") && header.Contains("здан")) classBldCol = col;
                                    else if (header.Contains("классиф") && header.Contains("элемен")) classElCol = col;
                                    else if (header.Contains("наименов") || header.Contains("назван")) nameCol = col;
                                    else if (header.Contains("тип") || header.Contains("форма")) typeCol = col;
                                    else if (header.Contains("статус")) statusCol = col;
                                    else if (header.Contains("штриховка") || header.Contains("образец")) hatchCol = col;
                                    else if (header.Contains("xxxs")) xxxsCol = col;
                                    else if (header.Contains("xxs")) xxsCol = col;
                                    else if (header.Contains("xs")) xsCol = col;
                                    else if (header.Equals("s") || header.Contains("-s")) sCol = col;
                                    else if (header.Equals("m") || header.Contains("-m")) mCol = col;
                                    else if (header.Equals("l") || header.Contains("-l")) lCol = col;
                                    else if (header.Contains("xl")) xlCol = col;
                                }
                                continue;
                            }

                            // Поиск по ключевому цвету
                            int actualColorCol = colorCol != -1 ? colorCol : 13;
                            object rawValue = reader.GetValue(actualColorCol);
                            if (rawValue == null) continue;

                            string rawColor = rawValue.ToString().Trim();
                            if (string.IsNullOrEmpty(rawColor)) continue;

                            string normalizedKey = NormalizeRgbString(rawColor);
                            if (string.IsNullOrEmpty(normalizedKey)) continue;

                            // Сбор значения color_AGO из его личной колонки (или дефолт 14 столбец, если не нашли имя)
                            int actualAgoCol = colorAgoCol != -1 ? colorAgoCol : 14;
                            string rawAgoValue = reader.GetValue(actualAgoCol)?.ToString()?.Trim() ?? string.Empty;

                            var row = new FloraBIM_HatchRow
                            {
                                ColorRGB = normalizedKey,
                                ColorAgo = NormalizeRgbString(rawAgoValue), // Корректно сохраняем раздельный цвет АГО
                                Id = idCol != -1 ? reader.GetValue(idCol)?.ToString()?.Trim() ?? "" : "",
                                CodeClassifierBuilding = classBldCol != -1 ? reader.GetValue(classBldCol)?.ToString()?.Trim() ?? "" : "",
                                CodeClassifierElement = classElCol != -1 ? reader.GetValue(classElCol)?.ToString()?.Trim() ?? "" : "",
                                Name = nameCol != -1 ? reader.GetValue(nameCol)?.ToString()?.Trim() ?? "" : "",
                                TypePlant = typeCol != -1 ? reader.GetValue(typeCol)?.ToString()?.Trim() ?? "" : "",
                                Status = statusCol != -1 ? reader.GetValue(statusCol)?.ToString()?.Trim() ?? "" : "",
                                ParamXxxs = xxxsCol != -1 ? reader.GetValue(xxxsCol)?.ToString()?.Trim() ?? "" : "",
                                ParamXxs = xxsCol != -1 ? reader.GetValue(xxsCol)?.ToString()?.Trim() ?? "" : "",
                                ParamXs = xsCol != -1 ? reader.GetValue(xsCol)?.ToString()?.Trim() ?? "" : "",
                                ParamS = sCol != -1 ? reader.GetValue(sCol)?.ToString()?.Trim() ?? "" : "",
                                ParamM = mCol != -1 ? reader.GetValue(mCol)?.ToString()?.Trim() ?? "" : "",
                                ParamL = lCol != -1 ? reader.GetValue(lCol)?.ToString()?.Trim() ?? "" : "",
                                ParamXl = xlCol != -1 ? reader.GetValue(xlCol)?.ToString()?.Trim() ?? "" : "",
                                HatchPattern = hatchCol != -1 ? reader.GetValue(hatchCol)?.ToString()?.Trim() ?? "" : ""
                            };

                            if (!registry.ContainsKey(normalizedKey))
                            {
                                registry.Add(normalizedKey, row);
                            }
                        }
                        sheetIndex++;
                    } while (reader.NextResult());
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка FloraHatch Excel]: {ex.Message}"); }
            return registry;
        }
    }
    // 3. Главный исполняемый класс для синхронизации штриховок (Изгороди, Цветники, Злаки, Миксы)
    public class FloraHatchPropertyFiller : IExtensionApplication
    {
        public const string TargetPsdName = "02_Живая изгородь. Многолетние цветы. Злаки. Миксы";

        public void Initialize()
        {
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
            if (ed != null) ed.WriteMessage($"\n>>> Модуль FloraHatchPropertyFiller загружен! Команда: FLORA_HATCH_FILL_PROPERTIES <<<");
        }

        public void Terminate() { }

        [CommandMethod("FLORA_HATCH_FILL_PROPERTIES", CommandFlags.Modal)]
        public void FillHatchPropertiesFromExcel()
        {
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor; Database db = activeDoc.Database;

            string targetPath = string.Empty;
            using (OpenFileDialog dialog = new OpenFileDialog { Filter = "Книга Excel (*.xlsx)|*.xlsx", Title = "Выберите БД штриховок" })
            {
                if (dialog.ShowDialog() != DialogResult.OK) return;
                targetPath = dialog.FileName;
            }

            var databaseMap = FloraHatchExcelReader.ReadGreeneryTable(targetPath, ed);
            if (databaseMap == null || databaseMap.Count == 0)
            {
                ed.WriteMessage("\n[ОШИБКА]: Из Excel загружено 0 ключей памяти!");
                return;
            }

            ed.WriteMessage($"\n[УСПЕХ]: Всего оригинальных RGB-ключей 'color' загружено в память: {databaseMap.Count}");
            ed.WriteMessage("\n--- ПЕРВЫЕ 10 КЛЮЧЕЙ В ПАМЯТИ ПЛАГИНА ---");
            int debugCounter = 0;
            foreach (var excelKey in databaseMap.Keys)
            {
                ed.WriteMessage($"\n  Ключ №{debugCounter + 1}: '{excelKey}' -> АГО: '{databaseMap[excelKey].ColorAgo}' | Куст: '{databaseMap[excelKey].Name}'");
                debugCounter++;
                if (debugCounter >= 10) break;
            }
            ed.WriteMessage("\n-----------------------------------------");

            ed.WriteMessage("\nВыберите контуры штриховок (Hatch) на чертеже...");
            PromptSelectionOptions selOpts = new PromptSelectionOptions();

            TypedValue[] typedValueArray = new TypedValue[] { new TypedValue((int)DxfCode.Start, "HATCH") };
            PromptSelectionResult selResult = ed.GetSelection(selOpts, new SelectionFilter(typedValueArray));
            if (selResult.Status != PromptStatus.OK || selResult.Value == null) return;

            int alteredObjectsCount = 0;
            using (DocumentLock docLock = activeDoc.LockDocument())
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                var psdDictionary = new DictionaryPropertySetDefinitions(db);
                if (!psdDictionary.Has(TargetPsdName, trans))
                {
                    ed.WriteMessage($"\n[Критическая ошибка] Набор '{TargetPsdName}' отсутствует в текущем чертеже!");
                    trans.Commit(); return;
                }
                ObjectId globalPsdId = psdDictionary.GetAt(TargetPsdName);

                foreach (SelectedObject selected in selResult.Value)
                {
                    if (selected == null) continue;
                    var hatch = trans.GetObject(selected.ObjectId, OpenMode.ForWrite) as Hatch;
                    if (hatch == null) continue;

                    string rawColorStr = GetHatchColorString(hatch);
                    string hatchColorKey = FloraHatchExcelReader.NormalizeRgbString(rawColorStr);
                    ed.WriteMessage($"\n[Анализ штриховки {hatch.Handle}]: Цвет CAD = '{rawColorStr}' | Поиск = '{hatchColorKey}'");

                    if (databaseMap.TryGetValue(hatchColorKey, out FloraBIM_HatchRow dataRow))
                    {
                        ed.WriteMessage(" -> [Матч! Данные найдены в Excel]");
                        ObjectId activePsId = ObjectId.Null;
                        foreach (ObjectId setInterfaceId in PropertyDataServices.GetPropertySets(hatch))
                        {
                            var currentSet = (PropertySet)trans.GetObject(setInterfaceId, OpenMode.ForRead);
                            if (currentSet != null && currentSet.PropertySetDefinition == globalPsdId) { activePsId = setInterfaceId; break; }
                        }

                        if (activePsId.IsNull)
                        {
                            PropertyDataServices.AddPropertySet(hatch, globalPsdId);
                            foreach (ObjectId setInterfaceId in PropertyDataServices.GetPropertySets(hatch))
                            {
                                var currentSet = (PropertySet)trans.GetObject(setInterfaceId, OpenMode.ForRead);
                                if (currentSet != null && currentSet.PropertySetDefinition == globalPsdId) { activePsId = setInterfaceId; break; }
                            }
                        }

                        if (!activePsId.IsNull)
                        {
                            var writeableSet = (PropertySet)trans.GetObject(activePsId, OpenMode.ForWrite);
                            if (writeableSet != null)
                            {
                                var props = new List<KeyValuePair<string, string>>
                                {
                                    // РАЗДЕЛЬНЫЙ ИНЖЕКТ ЦВЕТОВ: Каждому свойству своё честное значение из таблицы!
                                    new KeyValuePair<string, string>("color", dataRow.ColorRGB),
                                    new KeyValuePair<string, string>("color_AGO", dataRow.ColorAgo),

                                    new KeyValuePair<string, string>("ID", dataRow.Id),
                                    new KeyValuePair<string, string>("Код по классификатору зданий", dataRow.CodeClassifierBuilding),
                                    new KeyValuePair<string, string>("Код по классификатору элементов", dataRow.CodeClassifierElement),
                                    new KeyValuePair<string, string>("Наименование", dataRow.Name),
                                    new KeyValuePair<string, string>("Тип", dataRow.TypePlant),
                                    new KeyValuePair<string, string>("Статус", dataRow.Status),
                                    new KeyValuePair<string, string>("Характеристики-xxxs", dataRow.ParamXxxs),
                                    new KeyValuePair<string, string>("Характеристики-xxs", dataRow.ParamXxs),
                                    new KeyValuePair<string, string>("Характеристики-xs", dataRow.ParamXs),
                                    new KeyValuePair<string, string>("Характеристики-s", dataRow.ParamS),
                                    new KeyValuePair<string, string>("Характеристики-m", dataRow.ParamM),
                                    new KeyValuePair<string, string>("Характеристики-l", dataRow.ParamL),
                                    new KeyValuePair<string, string>("Характеристики-xl", dataRow.ParamXl),
                                    new KeyValuePair<string, string>("Штриховка", dataRow.HatchPattern)
                                };

                                foreach (var p in props) InjectProperty(writeableSet, p.Key, p.Value);

                                hatch.RecordGraphicsModified(true); alteredObjectsCount++;
                            }
                        }
                    }
                    else ed.WriteMessage(" -> [Пропуск] В Excel нет строки с таким оригинальным цветом.");
                }
                trans.Commit();
            }
            ed.WriteMessage($"\n\n[Итог] Синхронизация завершена! Успешно заполнено штриховок: {alteredObjectsCount}");
            ed.UpdateScreen();
        }

        private string GetHatchColorString(Hatch hatch)
        {
            if (hatch.Color.IsByLayer)
            {
                using (var tr = hatch.Database.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(hatch.Database.LayerTableId, OpenMode.ForRead);
                    if (lt.Has(hatch.Layer))
                    {
                        var layer = (LayerTableRecord)tr.GetObject(lt[hatch.Layer], OpenMode.ForRead);
                        return GetColorName(layer.Color);
                    }
                }
            }
            return GetColorName(hatch.Color);
        }

        private string GetColorName(Autodesk.AutoCAD.Colors.Color acColor)
        {
            if (acColor.ColorMethod == ColorMethod.ByColor) return $"{acColor.Red},{acColor.Green},{acColor.Blue}";
            return acColor.ColorIndex.ToString();
        }

        private void InjectProperty(PropertySet ps, string propName, string value)
        {
            try
            {
                int internalId = ps.PropertyNameToId(propName);
                ps.SetAt(internalId, value ?? string.Empty);
            }
            catch { }
        }
    }
}
