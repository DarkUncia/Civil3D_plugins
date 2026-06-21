using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Entity = Autodesk.AutoCAD.DatabaseServices.Entity;

[assembly: CommandClass(typeof(Civil3D_plugins.CogoFromGeometryCommands))]
namespace Civil3D_plugins
{
    public class CogoFromGeometryCommands
    {
        [CommandMethod("CreateCogoSmartVerify")]
        public void CreateCogoSmartVerify()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            Database db = doc.Database;
            CivilDocument civilDoc = CivilApplication.ActiveDocument;

            // --- НАСТРОЙКИ ---
            double searchRadius = 1.0; // Радиус поиска текста (только для штриховок)
            string layerError = "_АНАЛИЗ_ОШИБОК";

            TypedValue[] filter = new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "HATCH,CIRCLE")
            };

            PromptSelectionResult sel = ed.GetSelection(new SelectionFilter(filter));
            if (sel.Status != PromptStatus.OK) return;

            using (var lockDoc = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Проверяем/создаем слой для ошибок
                EnsureLayerExists(db, tr, layerError);

                // Кэш существующих COGO-точек для удаления дубликатов
                var cogoMap = new Dictionary<string, ObjectId>();
                foreach (ObjectId pId in civilDoc.CogoPoints)
                {
                    var p = tr.GetObject(pId, OpenMode.ForRead) as CogoPoint;
                    if (p == null) continue;
                    string key = $"{p.Easting:F3}_{p.Northing:F3}";
                    if (!cogoMap.ContainsKey(key))
                        cogoMap.Add(key, pId);
                }

                int createdCount = 0;

                foreach (SelectedObject sObj in sel.Value)
                {
                    Entity ent = tr.GetObject(sObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    Point3d finalPos = Point3d.Origin;
                    bool hasValidZ = false;
                    string description = "Verified_Z";

                    // ВЕРНУЛИ КЛАССНЫЙ ВАРИАНТ ДЛЯ ОКРУЖНОСТЕЙ
                    if (ent is Circle circle)
                    {
                        if (circle.Center.Z > 0.0001)
                        {
                            finalPos = circle.Center;
                            hasValidZ = true;
                            description = "Circle_GeometricZ";
                        }
                        else
                        {
                            MarkError(circle, 1, layerError); // На слой ошибок и в красный цвет
                            continue;
                        }
                    }
                    // ЛОГИКА ДЛЯ ШТРИХОВОК
                    else if (ent is Hatch hatch)
                    {
                        Extents3d ex = hatch.GeometricExtents;
                        Point3d center2D = new Point3d((ex.MinPoint.X + ex.MaxPoint.X) / 2.0, (ex.MinPoint.Y + ex.MaxPoint.Y) / 2.0, 0);

                        double? parsedZ = GetZFromTextContent(ed, tr, center2D, searchRadius);
                        if (parsedZ.HasValue)
                        {
                            if (parsedZ.Value <= 0.0001)
                            {
                                MarkError(hatch, 1, layerError);
                                continue;
                            }

                            double finalZ = Math.Round(parsedZ.Value, 2);

                            hatch.UpgradeOpen();
                            hatch.Elevation = finalZ;

                            finalPos = new Point3d(center2D.X, center2D.Y, finalZ);
                            hasValidZ = true;
                            description = "Hatch_TextZ";
                        }
                        else
                        {
                            MarkError(hatch, 1, layerError);
                            continue;
                        }
                    }

                    // Создание COGO-точки с затиранием старых дублей
                    if (hasValidZ)
                    {
                        string key = $"{finalPos.X:F3}_{finalPos.Y:F3}";
                        if (cogoMap.TryGetValue(key, out ObjectId oldId))
                        {
                            var oldP = tr.GetObject(oldId, OpenMode.ForRead) as CogoPoint;
                            if (oldP != null)
                            {
                                oldP.UpgradeOpen();
                                oldP.Erase();
                            }
                            cogoMap.Remove(key);
                        }

                        ObjectId newPointId = civilDoc.CogoPoints.Add(finalPos, false);
                        var newPoint = tr.GetObject(newPointId, OpenMode.ForWrite) as CogoPoint;
                        if (newPoint != null)
                        {
                            newPoint.RawDescription = description;
                        }
                        createdCount++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n[Готово] Обработано объектов. Создано/обновлено COGO-точек: {createdCount}.");
            }
        }

        private double? GetZFromTextContent(Editor ed, Transaction tr, Point3d center, double rad)
        {
            var res = ed.SelectCrossingWindow(center.Add(new Vector3d(-rad, -rad, 0)),
                                              center.Add(new Vector3d(rad, rad, 0)),
                                              new SelectionFilter(new[] { new TypedValue(0, "TEXT,MTEXT") }));

            if (res.Status != PromptStatus.OK) return null;

            Entity bestText = null;
            double minDist = double.MaxValue;
            double foundVal = 0;

            foreach (SelectedObject sObj in res.Value)
            {
                Entity ent = tr.GetObject(sObj.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                string raw = string.Empty;
                Point3d tPos = Point3d.Origin;

                if (ent is DBText t)
                {
                    raw = t.TextString;
                    tPos = t.Position;
                }
                else if (ent is MText mt)
                {
                    raw = mt.Contents;
                    tPos = mt.Location;
                }
                else
                {
                    continue;
                }

                string clean = Regex.Replace(raw, @"(\\S+;)|(\\{)|(\\})|(\\[AcLopPhntT])", "");

                if (double.TryParse(clean.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                {
                    double dist = tPos.DistanceTo(new Point3d(center.X, center.Y, tPos.Z));

                    if (dist < minDist)
                    {
                        minDist = dist;
                        bestText = ent;
                        foundVal = val;
                    }
                }
            }
            return bestText != null ? foundVal : (double?)null;
        }

        private void MarkError(Entity e, short colorIndex, string layerName)
        {
            if (!e.IsWriteEnabled) e.UpgradeOpen();
            e.Layer = layerName;
            e.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
        }

        private void EnsureLayerExists(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                using (LayerTableRecord ltr = new LayerTableRecord())
                {
                    ltr.Name = layerName;
                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 1); // 1 = Красный
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
            }
        }
    }
}