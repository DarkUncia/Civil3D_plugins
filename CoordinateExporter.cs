using System;
using System.IO;
using System.Xml.Linq;
using System.Globalization;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Civil3D_plugins
{
    public class CoordinateExporter
    {
        [CommandMethod("ExportProjectCoordinates")]
        public void ExportProjectCoordinates()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Настройки поиска
            const string targetBlockName = "КЛЕН_Базовая точка проекта";
            const string attributeTag = "НОМЕР_ДОМА";

            // 1. Выбор папки через стандартный диалог Windows
            string exportFolder = string.Empty;
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Выберите папку для сохранения XML";
                if (fbd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    ed.WriteMessage("\nВыгрузка отменена.");
                    return;
                }
                exportFolder = fbd.SelectedPath;
            }

            // 2. Фильтр выбора: только нужные блоки
            TypedValue[] filterList = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "INSERT"),
                new TypedValue((int)DxfCode.BlockName, targetBlockName)
            };
            SelectionFilter filter = new SelectionFilter(filterList);

            PromptSelectionOptions selOpts = new PromptSelectionOptions();
            selOpts.MessageForAdding = $"\nВыберите блоки '{targetBlockName}' на чертеже: ";

            PromptSelectionResult selRes = ed.GetSelection(selOpts, filter);
            if (selRes.Status != PromptStatus.OK) return;

            SelectionSet ss = selRes.Value;
            int exportedCount = 0;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in ss)
                    {
                        if (so == null) continue;

                        BlockReference br = (BlockReference)tr.GetObject(so.ObjectId, OpenMode.ForRead);

                        // Получение данных трансформации
                        Point3d pos = br.Position;
                        double rotationRad = br.Rotation;
                        string houseNumber = GetAttributeValue(br, attributeTag, tr);

                        // Формирование имени файла
                        string safeHouseNum = MakeValidFileName(houseNumber);
                        string fileName = $"Код проекта_Очередь_{safeHouseNum}.xml";
                        string fullPath = Path.Combine(exportFolder, fileName);

                        // Устранение конфликтов имен (если блоки имеют одинаковый номер дома)
                        int copyIndex = 1;
                        while (File.Exists(fullPath))
                        {
                            fullPath = Path.Combine(exportFolder, $"Код проекта_Очередь_{safeHouseNum}_{copyIndex}.xml");
                            copyIndex++;
                        }

                        // ФОРМИРОВАНИЕ XML
                        // X, Y и Rotation — максимальная точность (как в Dynamo)
                        // Z — принудительно 2 знака (например, 0.00)
                        XElement xmlRoot = new XElement("CoordSysZup",
                            new XAttribute("Units", "M"),
                            new XElement("OriginX", pos.X.ToString("G17", CultureInfo.InvariantCulture)),
                            new XElement("OriginY", pos.Y.ToString("G17", CultureInfo.InvariantCulture)),
                            new XElement("OriginZ", pos.Z.ToString("F2", CultureInfo.InvariantCulture)),
                            new XElement("RotationInXYPlane", rotationRad.ToString("G17", CultureInfo.InvariantCulture))
                        );

                        xmlRoot.Save(fullPath);
                        exportedCount++;
                    }
                    tr.Commit();
                }
                ed.WriteMessage($"\n[Успех]: Сформировано файлов: {exportedCount}. Путь: {exportFolder}");

                // Опционально: открыть папку после завершения
                System.Diagnostics.Process.Start("explorer.exe", exportFolder);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Критическая ошибка]: {ex.Message}");
            }
        }

        private string GetAttributeValue(BlockReference br, string tag, Transaction tr)
        {
            foreach (ObjectId attId in br.AttributeCollection)
            {
                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                if (attRef.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                {
                    return string.IsNullOrWhiteSpace(attRef.TextString) ? "БЕЗ_НОМЕРА" : attRef.TextString;
                }
            }
            return "АТРИБУТ_НЕ_НАЙДЕН";
        }

        private string MakeValidFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
