using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Aec.PropertyData.DatabaseServices;

[assembly: CommandClass(typeof(Civil3D_plugins.FloraBIMExportCommand))]

namespace Civil3D_plugins
{
    public class FloraBIMExportCommand
    {
        // Список допустимых наборов характеристик, которые плагин будет считывать
        private static readonly HashSet<string> TargetPsetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "01_Деревья и кусты (стандарт)",
            "02_Живая изгородь. Многолетние цветы. Злаки",
            "04_Озеленение (индивидуально)"
        };

        private static readonly string[] ColumnHeaders = new string[]
        {
            "Handle", "guid", "id", "name", "namelat", "description", "note", "type",
            "treetrunk", "version", "ADSK_name", "ADSK_Set_of_drawings",
            "ACER_Collision_Code", "KRTRS_Classifier_code", "s", "m", "l", "xl", "xxl", "xxxl", "Размер"
        };

        private static readonly string[] SheetNames = new string[] { "Деревья", "Кустарники", "ЖИ", "Цветники" };

        [CommandMethod("FLORA_EXPORT_PSETS")]
        public void ExportFloraToExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptSelectionOptions pso = new PromptSelectionOptions { MessageForAdding = "\nВыберите блоки озеленения для выгрузки: " };
            SelectionFilter filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") });
            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK || psr.Value == null) return;

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Сохранить ведомость FloraBIM",
                Filter = "Excel XML Spreadsheet (*.xml)|*.xml",
                FileName = $"Ведомость_Озеленения_{DateTime.Now:dd_MM_yyyy}.xml"
            };

            if (saveFileDialog.ShowDialog() != true) return;
            string filePath = saveFileDialog.FileName;

            var dataStorage = new Dictionary<string, List<Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in SheetNames) dataStorage[name] = new List<Dictionary<string, string>>();

            try
            {
                ObjectId[] ids = psr.Value.GetObjectIds();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    for (int i = 0; i < ids.Length; i++)
                    {
                        var blkRef = tr.GetObject(ids[i], OpenMode.ForRead) as BlockReference;
                        if (blkRef == null) continue;

                        string handleStr = blkRef.Handle.ToString();
                        string blockName = blkRef.IsDynamicBlock
                            ? ((BlockTableRecord)tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead)).Name
                            : blkRef.Name;

                        ObjectIdCollection attachedSets = PropertyDataServices.GetPropertySets(blkRef);
                        PropertySet targetPset = null;
                        string matchedPsetName = string.Empty;

                        // Ищем любой из разрешенных наборов характеристик на объекте
                        foreach (ObjectId pSetId in attachedSets)
                        {
                            object rawObj = tr.GetObject(pSetId, OpenMode.ForRead);
                            var pSet = rawObj as PropertySet;
                            if (pSet != null)
                            {
                                object rawDef = tr.GetObject(pSet.PropertySetDefinition, OpenMode.ForRead);
                                var pSetDef = rawDef as PropertySetDefinition;
                                if (pSetDef != null && (TargetPsetNames.Contains(pSetDef.Name) ||
                                    pSetDef.Name.IndexOf("Озеленен", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    pSetDef.Name.IndexOf("Дерев", StringComparison.OrdinalIgnoreCase) >= 0))
                                {
                                    targetPset = pSet;
                                    matchedPsetName = pSetDef.Name;
                                    break;
                                }
                            }
                        }

                        // Если на объекте нет подходящего набора, пропускаем его
                        if (targetPset == null) continue;

                        var pSetDefinition = tr.GetObject(targetPset.PropertySetDefinition, OpenMode.ForRead) as PropertySetDefinition;
                        if (pSetDefinition == null) continue;

                        // Собираем все свойства, которые физически заполнены в чертеже
                        var propVals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (PropertyDefinition propDef in pSetDefinition.Definitions)
                        {
                            object val = targetPset.GetAt(propDef.Id);
                            propVals[propDef.Name] = val?.ToString() ?? string.Empty;
                        }

                        var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        row["Handle"] = handleStr;

                        // Если в наборе есть реальный guid, берем его, иначе используем имя блока
                        row["guid"] = propVals.ContainsKey("guid") && !string.IsNullOrEmpty(propVals["guid"])
                            ? propVals["guid"]
                            : blockName;

                        // Заполняем остальные колонки из набора характеристик
                        foreach (string header in ColumnHeaders)
                        {
                            if (header == "Handle" || header == "guid") continue;
                            row[header] = propVals.ContainsKey(header) ? propVals[header] : string.Empty;
                        }

                        // Логика сортировки по вкладкам Excel на основе свойства "type" или имени набора
                        row.TryGetValue("type", out string typeVal);
                        string targetSheet = "Цветники"; // По умолчанию

                        if (!string.IsNullOrEmpty(typeVal))
                        {
                            if (typeVal.IndexOf("дерев", StringComparison.OrdinalIgnoreCase) >= 0) targetSheet = "Деревья";
                            else if (typeVal.IndexOf("куст", StringComparison.OrdinalIgnoreCase) >= 0) targetSheet = "Кустарники";
                            else if (typeVal.IndexOf("жи", StringComparison.OrdinalIgnoreCase) >= 0 || typeVal.IndexOf("изгород", StringComparison.OrdinalIgnoreCase) >= 0) targetSheet = "ЖИ";
                        }
                        else
                        {
                            // Подстраховка: если тип не указан, сортируем по имени самого набора характеристик
                            if (matchedPsetName.IndexOf("дерев", StringComparison.OrdinalIgnoreCase) >= 0) targetSheet = "Деревья";
                            else if (matchedPsetName.IndexOf("изгород", StringComparison.OrdinalIgnoreCase) >= 0) targetSheet = "ЖИ";
                            else if (matchedPsetName.IndexOf("цвет", StringComparison.OrdinalIgnoreCase) >= 0) targetSheet = "Цветники";
                        }

                        dataStorage[targetSheet].Add(row);
                    }
                    tr.Commit();
                }
                using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    WriteHeader(sw);
                    foreach (string sheetName in SheetNames) WriteWorksheet(sw, sheetName, dataStorage[sheetName]);
                    WriteFooter(sw);
                }
                ed.WriteMessage($"\n[FloraBIM] Успешно экспортировано в: {filePath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[FloraBIM] Ошибка экспорта: {ex.Message}");
            }
        }

        private void WriteHeader(StreamWriter sw)
        {
            sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sw.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
            sw.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            sw.WriteLine(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            sw.WriteLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            sw.WriteLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            sw.WriteLine(" xmlns:html=\"http://w3.org\">");
            sw.WriteLine("  <Styles>");
            sw.WriteLine("    <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
            sw.WriteLine("      <Alignment ss:Vertical=\"Bottom\"/>");
            sw.WriteLine("      <Font ss:FontName=\"Arial\" x:CharSet=\"204\" x:Family=\"Swiss\" ss:Size=\"10\"/>");
            sw.WriteLine("    </Style>");
            sw.WriteLine("    <Style ss:ID=\"HeaderStyle\">");
            sw.WriteLine("      <Font ss:FontName=\"Arial\" x:CharSet=\"204\" x:Family=\"Swiss\" ss:Size=\"10\" ss:Bold=\"1\"/>");
            sw.WriteLine("      <Interior ss:Color=\"#D3D3D3\" ss:Pattern=\"Solid\"/>");
            sw.WriteLine("    </Style>");
            sw.WriteLine("  </Styles>");
        }

        private void WriteWorksheet(StreamWriter sw, string sheetName, List<Dictionary<string, string>> rows)
        {
            sw.WriteLine($"  <Worksheet ss:Name=\"{sheetName}\">");
            sw.WriteLine("    <Table>");

            sw.WriteLine("      <Row ss:StyleID=\"HeaderStyle\">");
            foreach (string header in ColumnHeaders)
            {
                sw.WriteLine($"        <Cell><Data ss:Type=\"String\">{EscapeXml(header)}</Data></Cell>");
            }
            sw.WriteLine("      </Row>");

            foreach (var row in rows)
            {
                sw.WriteLine("      <Row>");
                foreach (string header in ColumnHeaders)
                {
                    string val = row.ContainsKey(header) ? row[header] : string.Empty;
                    sw.WriteLine($"        <Cell><Data ss:Type=\"String\">{EscapeXml(val)}</Data></Cell>");
                }
                sw.WriteLine("      </Row>");
            }

            sw.WriteLine("    </Table>");
            sw.WriteLine("  </Worksheet>");
        }

        private void WriteFooter(StreamWriter sw)
        {
            sw.WriteLine("</Workbook>");
        }

        private string EscapeXml(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value.Replace("&", "&amp;")
                        .Replace("<", "&lt;")
                        .Replace(">", "&gt;")
                        .Replace("\"", "&quot;")
                        .Replace("'", "&apos;");
        }
    }
}
