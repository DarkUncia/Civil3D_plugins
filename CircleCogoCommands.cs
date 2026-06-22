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
            double searchRadius = 1.5; // Радиус поиска текста вокруг объектов
            string layerError = "_АНАЛИЗ_ОШИБОК";
            string layerWater = "ВОДОЕМ"; // Имя целевого слоя для водных объектов

            // ФИЛЬТР ВЫСОТ: Все числа вне этого диапазона (диаметры труб, параметры берез) игнорируются
            double minValidElevation = 30.0;  // Минимальная правдоподобная отметка земли
            double maxValidElevation = 350.0; // Максимальная правдоподобная отметка земли

            // Выбираем штриховки, окружности и весь текст для анализа
            TypedValue[] filter = new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "HATCH,CIRCLE,TEXT,MTEXT")
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

                // Списки для разделения логики обработки
                List<Entity> geometryEntities = new List<Entity>();
                List<Entity> standaloneTexts = new List<Entity>();
                HashSet<ObjectId> usedTextIds = new HashSet<ObjectId>();

                // Сортируем выборку пользователя
                foreach (SelectedObject sObj in sel.Value)
                {
                    Entity ent = tr.GetObject(sObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    // Все объекты со слоя ВОДОЕМ защищаем от превращения в одиночные точки суши
                    if (ent.Layer.Equals(layerWater, StringComparison.OrdinalIgnoreCase))
                    {
                        if (ent is Hatch || ent is Circle)
                        {
                            geometryEntities.Add(ent);
                        }
                        else if (ent is DBText || ent is MText)
                        {
                            usedTextIds.Add(ent.ObjectId);
                        }
                    }
                    else
                    {
                        if (ent is Circle || ent is Hatch)
                        {
                            geometryEntities.Add(ent);
                        }
                        else if (ent is DBText || ent is MText)
                        {
                            standaloneTexts.Add(ent);
                        }
                    }
                }

                int createdCount = 0;

                // ЭТАП 1: Обработка геометрии (Окружности и Штриховки)
                foreach (Entity ent in geometryEntities)
                {
                    Point3d finalPos = Point3d.Origin;
                    bool hasValidZ = false;
                    string description = "Verified_Z";

                    if (ent is Circle circle)
                    {
                        Point3d center2D = new Point3d(circle.Center.X, circle.Center.Y, 0);
                        ObjectId foundTextId;
                        double? parsedZ = GetZFromTextContent(ed, tr, center2D, searchRadius, minValidElevation, maxValidElevation, out foundTextId);

                        // Если рядом с кругом найден корректный текст, блокируем его от дублирования
                        if (foundTextId != ObjectId.Null)
                        {
                            usedTextIds.Add(foundTextId);
                        }

                        // Если у круга уже есть нормальный Z в диапазоне
                        if (circle.Center.Z >= minValidElevation && circle.Center.Z <= maxValidElevation)
                        {
                            finalPos = circle.Center;
                            hasValidZ = true;
                            description = "Circle_GeometricZ";
                        }
                        // Если круг плоский, но текст рядом успешно прошел фильтрацию по высоте
                        else if (parsedZ.HasValue)
                        {
                            double finalZ = Math.Round(parsedZ.Value, 2);
                            circle.UpgradeOpen();
                            circle.Center = new Point3d(circle.Center.X, circle.Center.Y, finalZ);

                            finalPos = new Point3d(center2D.X, center2D.Y, finalZ);
                            hasValidZ = true;
                            description = "Circle_TextZ";
                        }
                        else
                        {
                            MarkError(circle, 1, layerError);
                            continue;
                        }
                    }
                    else if (ent is Hatch hatch)
                    {
                        Extents3d ex = hatch.GeometricExtents;
                        Point3d center2D = new Point3d((ex.MinPoint.X + ex.MaxPoint.X) / 2.0, (ex.MinPoint.Y + ex.MaxPoint.Y) / 2.0, 0);

                        ObjectId foundTextId;
                        double? parsedZ = GetZFromTextContent(ed, tr, center2D, searchRadius, minValidElevation, maxValidElevation, out foundTextId);
                        if (parsedZ.HasValue)
                        {
                            double finalZ = Math.Round(parsedZ.Value, 2);

                            hatch.UpgradeOpen();
                            hatch.Elevation = finalZ;

                            finalPos = new Point3d(center2D.X, center2D.Y, finalZ);
                            hasValidZ = true;
                            description = "Hatch_TextZ";

                            if (foundTextId != ObjectId.Null) usedTextIds.Add(foundTextId);
                        }
                        else
                        {
                            MarkError(hatch, 1, layerError);
                            continue;
                        }
                    }

                    if (hasValidZ)
                    {
                        CreateOrReplaceCogoPoint(civilDoc, tr, cogoMap, finalPos, description);
                        createdCount++;
                    }
                }

                // ЭТАП 2: Обработка ОДИНОЧНОГО текста суши (Исключая характеристики деревьев и слой ВОДОЕМ)
                foreach (Entity textEnt in standaloneTexts)
                {
                    if (usedTextIds.Contains(textEnt.ObjectId)) continue;

                    string raw = string.Empty;
                    Point3d tPos = Point3d.Origin;

                    if (textEnt is DBText t)
                    {
                        raw = t.TextString;
                        tPos = t.Position;
                    }
                    else if (textEnt is MText mt)
                    {
                        raw = mt.Contents;
                        tPos = mt.Location;
                    }

                    string clean = Regex.Replace(raw, @"(\\S+;)|(\\{)|(\\})|(\\[AcLopPhntT])", "");
                    if (string.IsNullOrWhiteSpace(clean)) continue;

                    if (double.TryParse(clean.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                    {
                        // Применяем фильтр диапазона высот для одиночного текста
                        if (val < minValidElevation || val > maxValidElevation) continue;

                        double finalZ = Math.Round(val, 2);
                        Point3d finalPos = new Point3d(tPos.X, tPos.Y, finalZ);

                        CreateOrReplaceCogoPoint(civilDoc, tr, cogoMap, finalPos, "Text_DirectZ");
                        createdCount++;
                    }
                }

                tr.Commit();
                ed.WriteMessage($"\n[Готово] Обработка завершена. Всего создано/обновлено COGO-точек: {createdCount}.");
            }
        }

        private void CreateOrReplaceCogoPoint(CivilDocument civilDoc, Transaction tr, Dictionary<string, ObjectId> cogoMap, Point3d pos, string desc)
        {
            string key = $"{pos.X:F3}_{pos.Y:F3}";
            if (cogoMap.TryGetValue(key, out ObjectId oldId))
            {
                var oldP = tr.GetObject(oldId, OpenMode.ForWrite) as CogoPoint;
                if (oldP != null) oldP.Erase();

                cogoMap.Remove(key);
            }
            ObjectId newPointId = civilDoc.CogoPoints.Add(pos, false);
            var newPoint = tr.GetObject(newPointId, OpenMode.ForWrite) as CogoPoint;
            if (newPoint != null)
            {
                newPoint.RawDescription = desc;
            }
        }
        private double? GetZFromTextContent(Editor ed, Transaction tr, Point3d center, double rad, double minZ, double maxZ, out ObjectId foundTextId)
        {
            foundTextId = ObjectId.Null;
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
                string clean = Regex.Replace(raw, @"(\S+;)|(\{)|(\})|(\[AcLopPhntT])", "");
                if (string.IsNullOrWhiteSpace(clean)) continue;
                if (double.TryParse(clean.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                {
                    if (val < minZ || val > maxZ) continue;
                    double dist = tPos.DistanceTo(new Point3d(center.X, center.Y, tPos.Z));
                    if (dist < minDist)
                    {
                        minDist = dist;
                        bestText = ent;
                        foundVal = val;
                    }
                }
            }
            if (bestText != null)
            {
                foundTextId = bestText.ObjectId;
                return foundVal;
            }
            return null;
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
            if (lt == null) return;
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                using (LayerTableRecord ltr = new LayerTableRecord())
                {
                    ltr.Name = layerName;
                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 1);
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
            }
        }
    }
}

