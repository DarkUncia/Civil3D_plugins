using System;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

[assembly: CommandClass(typeof(Civil3D_plugins.C3D_ExportClasses))]

namespace Civil3D_plugins
{
    public class C3D_ExportClasses
    {
        // КОМАНДА 1: Простой экспорт данных
        [CommandMethod("C3D_ExportData")]
        public void ExportToTxt()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            if (acDoc == null) return;

            Editor ed = acDoc.Editor;
            Database db = acDoc.Database;

            try
            {
                using (DocumentLock docLock = acDoc.LockDocument())
                {
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string path = Path.Combine(desktop, "C3D_Report.txt");

                    using (StreamWriter sw = new StreamWriter(path))
                    {
                        sw.WriteLine($"ОТЧЕТ ПО ОБЪЕКТУ: {acDoc.Name}");
                        sw.WriteLine(new string('=', 40));
                        sw.WriteLine($"Дата экспорта: {DateTime.Now}");
                        sw.WriteLine("Плагин работает корректно.");
                    }
                    ed.WriteMessage($"\nУспех! Файл создан на рабочем столе: {path}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nОшибка: {ex.Message}");
            }
        }

        // КОМАНДА 2: Глубокий анализ слоев и владельцев (контейнеров)
        [CommandMethod("C3D_AnalyzeLayers")]
        public void AnalyzeLayers()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            if (acDoc == null) return;

            Editor ed = acDoc.Editor;
            Database db = acDoc.Database;

            PromptSelectionResult selRes = ed.GetSelection();
            if (selRes.Status != PromptStatus.OK) return;

            try
            {
                using (DocumentLock docLock = acDoc.LockDocument())
                {
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string path = Path.Combine(desktop, "C3D_Deep_Analysis.txt");

                    using (StreamWriter sw = new StreamWriter(path))
                    {
                        sw.WriteLine("ГЛУБОКИЙ АНАЛИЗ: СЛОИ + КОНТЕЙНЕРЫ (ВЛАДЕЛЬЦЫ)");
                        sw.WriteLine(new string('=', 100));
                        sw.WriteLine($"{"Тип объекта",-20} | {"Handle",-8} | {"Слой",-25} | {"Контейнер (Где лежит?)",-25}");
                        sw.WriteLine(new string('-', 100));

                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            foreach (SelectedObject so in selRes.Value)
                            {
                                Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                                if (ent != null)
                                {
                                    // Определяем владельца (блок или пространство модели)
                                    SymbolTableRecord owner = tr.GetObject(ent.OwnerId, OpenMode.ForRead) as SymbolTableRecord;
                                    string containerName = owner != null ? owner.Name : "Unknown";
                                    string typeName = ent.GetType().Name;

                                    sw.WriteLine($"{typeName,-20} | {ent.Handle,-8} | {ent.Layer,-25} | {containerName,-25}");
                                }
                            }
                            tr.Commit();
                        }
                    }
                    ed.WriteMessage($"\n[Анализ]: Готово! Проверь колонку 'Контейнер' в файле: {path}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nОшибка: {ex.Message}");
            }
        }
    }
}