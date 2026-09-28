using System;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;

[assembly: ExtensionApplication(typeof(Civil3D_plugins.Loader))]

// Ваши существующие рабочие команды
[assembly: CommandClass(typeof(Civil3D_plugins.CoordinateExporter))]
[assembly: CommandClass(typeof(Civil3D_plugins.C3D_ExportClasses))]
[assembly: CommandClass(typeof(Civil3D_plugins.CogoFromGeometryCommands))]

// Новая изолированная команда заполнения НХ
[assembly: CommandClass(typeof(Civil3D_plugins.MafOnlyPropertyFiller))]

namespace Civil3D_plugins
{
    public class Loader : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            }
            catch { }

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\n>>> Плагин Civil3D_plugins успешно загружен! <<<");
                doc.Editor.WriteMessage("\n>>> Используйте команду: MAF_FILL_ATTRIBUTES <<<");
            }
        }

        public void Terminate() { }

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
