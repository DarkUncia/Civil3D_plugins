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

// ИСПРАВЛЕНО: Пространство имен заменено на Civil3D_plugins
[assembly: CommandClass(typeof(Civil3D_plugins.FloraHatchPropertyFiller))]
namespace Civil3D_plugins
{
    public class FloraBIM_HatchRow
    {
        public string ColorRGB { get; set; } = string.Empty;
        public string ColorAgo { get; set; } = string.Empty;
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

    public static class FloraHatchExcelReader
    {
        public static string NormalizeRgbString(string rawColor)
        {
            if (string.IsNullOrEmpty(rawColor)) return string.Empty;
            string clean = rawColor.Replace(" ", "").Trim();
            if (!clean.Contains(",") && clean.Length >= 7 && clean.Length <= 9)
            {
                try { return $"{int.Parse(clean.Substring(0, 3))},{int.Parse(clean.Substring(3, 3))},{int.Parse(clean.Substring(6))}"; } catch { }
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
                        int colorCol = -1, colorAgoCol = -1, idCol = -1, classBldCol = -1, classElCol = -1, nameCol = -1, typeCol = -1, statusCol = -1, hatchCol = -1;
                        int xxxsCol = -1, xxsCol = -1, xsCol = -1, sCol = -1, mCol = -1, lCol = -1, xlCol = -1;
                        int localRowCount = 0;

                        while (reader.Read())
                        {
                            localRowCount++;
                            if (localRowCount == 2) continue; // Пропуск примера

                            if (localRowCount == 1)
                            {
                                for (int col = 0; col < reader.FieldCount; col++)
                                {
                                    string header = reader.GetValue(col)?.ToString()?.Trim()?.ToLower() ?? "";
                                    header = header.Replace("\r", "").Replace("\n", "");

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

                            int actualColorCol = colorCol != -1 ? colorCol : 13;
                            object rawValue = reader.GetValue(actualColorCol);
                            if (rawValue == null) continue;

                            string rawColor = rawValue.ToString().Trim();
                            if (string.IsNullOrEmpty(rawColor)) continue;

                            string normalizedKey = NormalizeRgbString(rawColor);
                            if (string.IsNullOrEmpty(normalizedKey)) continue;

                            int actualAgoCol = colorAgoCol != -1 ? colorAgoCol : 14;
                            string rawAgoValue = reader.GetValue(actualAgoCol)?.ToString()?.Trim() ?? string.Empty;

                            var row = new FloraBIM_HatchRow
                            {
                                ColorRGB = normalizedKey,
                                ColorAgo = NormalizeRgbString(rawAgoValue),
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
            catch (System.Exception ex) { ed.WriteMessage($"\n[Ошибка Excel]: {ex.Message}"); }
            return registry;
        }
    }

    public class FloraHatchPropertyFiller : IExtensionApplication
    {
        public const string TargetPsdName = "02_Живая изгородь. Многолетние цветы. Злаки. Миксы";

        public void Initialize()
        {
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor;
            if (ed != null) ed.WriteMessage($"\n>>> Модуль Площадных объектов (3D/Штриховки) загружен! Команда: FLORA_HATCH_FILL_PROPERTIES <<<");
        }

        public void Terminate() { }

        [CommandMethod("FLORA_HATCH_FILL_PROPERTIES", CommandFlags.Modal)]
        public void FillHatchPropertiesFromExcel()
        {
            Document activeDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;
            Editor ed = activeDoc.Editor; Database db = activeDoc.Database;

            string targetPath = string.Empty;
            using (OpenFileDialog dialog = new OpenFileDialog { Filter = "Книга Excel (*.xlsx)|*.xlsx", Title = "Выберите БД озеленения" })
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

            ed.WriteMessage($"\n[УСПЕХ]: Загружено ключей: {databaseMap.Count}");
            ed.WriteMessage("\nВыберите 3D-тела, Поверхности или Штриховки на чертеже...");
            PromptSelectionResult selResult = ed.GetSelection();
            if (selResult.Status != PromptStatus.OK || selResult.Value == null) return;

            int alteredObjectsCount = 0;
            using (DocumentLock docLock = activeDoc.LockDocument())
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                var psdDictionary = new DictionaryPropertySetDefinitions(db);
                if (!psdDictionary.Has(TargetPsdName, trans))
                {
                    ed.WriteMessage($"\n[Критическая ошибка] Набор '{TargetPsdName}' отсутствует в текущем чертеже!");
                    return;
                }
                ObjectId globalPsdId = psdDictionary.GetAt(TargetPsdName);

                foreach (SelectedObject selected in selResult.Value)
                {
                    if (selected == null) continue;
                    DBObject currentDbObj = trans.GetObject(selected.ObjectId, OpenMode.ForWrite);

                    if (!(currentDbObj is Hatch || currentDbObj is Solid3d || currentDbObj is Autodesk.Civil.DatabaseServices.Surface)) continue;

                    string colorKey = GetObjectColorKey(currentDbObj, trans);
                    ed.WriteMessage($"\n[Анализ {currentDbObj.GetType().Name} {currentDbObj.Handle}]: Цвет = '{colorKey}'");

                    if (!string.IsNullOrEmpty(colorKey) && databaseMap.TryGetValue(colorKey, out FloraBIM_HatchRow dataRow))
                    {
                        ed.WriteMessage(" -> [Матч! Данные найдены в Excel]");
                        ObjectId activePsId = ObjectId.Null;
                        foreach (ObjectId setInterfaceId in PropertyDataServices.GetPropertySets(currentDbObj))
                        {
                            var currentSet = (PropertySet)trans.GetObject(setInterfaceId, OpenMode.ForRead);
                            if (currentSet != null && currentSet.PropertySetDefinition == globalPsdId) { activePsId = setInterfaceId; break; }
                        }

                        if (activePsId.IsNull)
                        {
                            PropertyDataServices.AddPropertySet(currentDbObj, globalPsdId);
                            foreach (ObjectId setInterfaceId in PropertyDataServices.GetPropertySets(currentDbObj))
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

                                if (currentDbObj is Hatch h) h.RecordGraphicsModified(true);
                                else if (currentDbObj is Solid3d s) s.RecordGraphicsModified(true);

                                alteredObjectsCount++;
                                ed.WriteMessage($" -> [УСПЕХ: Данные записаны!]");
                            }
                        }
                    }
                    else ed.WriteMessage(" -> [Пропуск] В Excel нет строки с таким цветом.");
                }
                trans.Commit();
            }
            ed.WriteMessage($"\n\n[Итог] Заполнение завершено! Обработано объектов: {alteredObjectsCount}");
            ed.UpdateScreen();
        }

        private string GetObjectColorKey(DBObject obj, Transaction trans)
        {
            if (obj is Entity ent)
            {
                if (ent is Autodesk.Civil.DatabaseServices.Surface surface)
                {
                    var col = ent.Color;
                    if (col.ColorMethod == Autodesk.AutoCAD.Colors.ColorMethod.ByLayer || col.ColorIndex == 256)
                    {
                        var layer = (LayerTableRecord)trans.GetObject(surface.LayerId, OpenMode.ForRead);
                        var layerColor = layer.Color;
                        return FloraHatchExcelReader.NormalizeRgbString($"{layerColor.Red},{layerColor.Green},{layerColor.Blue}");
                    }
                    return FloraHatchExcelReader.NormalizeRgbString($"{col.Red},{col.Green},{col.Blue}");
                }

                if (ent is Solid3d || ent is Hatch)
                {
                    var col = ent.Color;
                    if (col.ColorMethod == Autodesk.AutoCAD.Colors.ColorMethod.ByLayer || col.ColorIndex == 256)
                    {
                        var layer = (LayerTableRecord)trans.GetObject(ent.LayerId, OpenMode.ForRead);
                        var layerColor = layer.Color;
                        return FloraHatchExcelReader.NormalizeRgbString($"{layerColor.Red},{layerColor.Green},{layerColor.Blue}");
                    }
                    return FloraHatchExcelReader.NormalizeRgbString($"{col.Red},{col.Green},{col.Blue}");
                }
            }
            return null;
        }

        private void InjectProperty(PropertySet ps, string propName, string value)
        {
            try
            {
                int internalId = ps.PropertyNameToId(propName);
                if (internalId >= 0) ps.SetAt(internalId, value ?? string.Empty);
            }
            catch { }
        }
    }
}