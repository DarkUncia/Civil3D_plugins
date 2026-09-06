using System;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;

[assembly: ExtensionApplication(typeof(Civil3D_plugins.Loader))]
[assembly: CommandClass(typeof(Civil3D_plugins.CoordinateExporter))]
[assembly: CommandClass(typeof(Civil3D_plugins.C3D_ExportClasses))]
// Теперь здесь четкое, понятное и уникальное имя класса
[assembly: CommandClass(typeof(Civil3D_plugins.CogoFromGeometryCommands))]

namespace Civil3D_plugins
{
    public class Loader : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                // Регистрируем кодировки для корректной работы с файлами
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                // Добавляем автоматический поиск зависимых DLL в папке плагина
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            }
            catch { }

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\n>>> Плагин Civil3D_plugins успешно загружен! <<<");
            }
        }

        public void Terminate() { }

        // Событие, которое принудительно находит ClosedXML.Parser.dll рядом с основным плагином
        private System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                string asmPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (string.IsNullOrEmpty(asmPath)) return null;

                string name = new System.Reflection.AssemblyName(args.Name).Name;
                string file = Path.Combine(asmPath, name + ".dll");

                if (File.Exists(file))
                {
                    return System.Reflection.Assembly.LoadFrom(file);
                }
            }
            catch { }
            return null;
        }
    }
}
