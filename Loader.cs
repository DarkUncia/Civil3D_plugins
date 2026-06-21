using System;
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
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\n>>> Плагин Civil3D_plugins успешно загружен! <<<");
            }
        }

        public void Terminate() { }
    }
}