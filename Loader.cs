using System;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;

[assembly: ExtensionApplication(typeof(Civil3D_plugins.Loader))]

// Ваши существующие рабочие команды
[assembly: CommandClass(typeof(Civil3D_plugins.CoordinateExporter))]
[assembly: CommandClass(typeof(Civil3D_plugins.C3D_ExportClasses))]
[assembly: CommandClass(typeof(Civil3D_plugins.CogoFromGeometryCommands))]

// Команды для работы с НХ (Наборами Характеристик)
[assembly: CommandClass(typeof(Civil3D_plugins.MafOnlyPropertyFiller))]       // МАФ (MAF_FILL_ATTRIBUTES)
[assembly: CommandClass(typeof(Civil3D_plugins.FloraPropertyFiller))]         // Озеленение: Блоки (FLORA_FILL_PROPERTIES)
[assembly: CommandClass(typeof(Civil3D_plugins.FloraHatchPropertyFiller))]    // Озеленение: Штриховки/3D (FLORA_HATCH_FILL_PROPERTIES)
[assembly: CommandClass(typeof(Civil3D_plugins.LightingPropertyFiller))]      // Освещение (LIGHTING_FILL_ATTRIBUTES) <--- ДОБАВЛЕНО

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
                doc.Editor.WriteMessage("\n>>> Доступные команды: FLORA_FILL_PROPERTIES, FLORA_HATCH_FILL_PROPERTIES, MAF_FILL_ATTRIBUTES, LIGHTING_FILL_ATTRIBUTES <<<");
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