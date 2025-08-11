using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace PanelDivider
{
    [Transaction(TransactionMode.Manual)]
    //————————————————————————————————————————————————————————————————
    // 1. Entry point: pick panels, run splits, handle Revit transaction
    //————————————————————————————————————————————————————————————————
    // Main Revit command. Prompts the user for panels to split and cutting panels, then processes each panel inside a single Transaction.
    public class DivideAnalyticalPanels : IExternalCommand
    {
        private const double Tolerance = 0.001;
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "RevitScriptLog.txt");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!File.Exists(LogPath))
                    File.WriteAllText(LogPath, "Начало лога\n");
            }
            catch { }

            Log("======== СКРИПТ ЗАПУЩЕН ========");

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Log("Начало основной логики скрипта");

                // Выбор панелей для разделения
                Log("Запрос выбора панелей для разделения...");
                IList<Reference> panelRefs;
                try
                {
                    panelRefs = uidoc.Selection.PickObjects(ObjectType.Element,
                        new AnalyticalPanelFilter(), "Выберите панели для разделения");
                    Log($"Выбрано объектов: {panelRefs?.Count ?? 0}");
                }
                catch (Exception ex)
                {
                    Log($"Ошибка выбора панелей: {ex.Message}");
                    return Result.Cancelled;
                }

                var panelsToSplit = new List<AnalyticalPanel>();
                foreach (var reference in panelRefs)
                {
                    var element = doc.GetElement(reference);
                    if (element is AnalyticalPanel panel)
                    {
                        panelsToSplit.Add(panel);
                        Log($"Добавлена панель: {panel.Id}");
                    }
                }
                if (panelsToSplit.Count == 0)
                {
                    Log("Не выбрано ни одной панели");
                    TaskDialog.Show("Ошибка", "Не выбрано ни одной панели");
                    return Result.Failed;
                }

                // Выбор режущих панелей
                Log("Запрос выбора режущих панелей...");
                IList<Reference> cuttingRefs;
                try
                {
                    cuttingRefs = uidoc.Selection.PickObjects(ObjectType.Element,
                        new AnalyticalPanelFilter(), "Выберите режущие панели");
                    Log($"Выбрано режущих панелей: {cuttingRefs?.Count ?? 0}");
                }
                catch (Exception ex)
                {
                    Log($"Ошибка выбора режущих панелей: {ex.Message}");
                    return Result.Cancelled;
                }

                var cuttingPanels = new List<AnalyticalPanel>();
                foreach (var reference in cuttingRefs)
                {
                    var element = doc.GetElement(reference);
                    if (element is AnalyticalPanel panel)
                    {
                        cuttingPanels.Add(panel);
                        Log($"Добавлена режущая панель: {panel.Id}");
                    }
                }

                Log($"Выбрано панелей для разделения: {panelsToSplit.Count}");
                Log($"Выбрано режущих панелей: {cuttingPanels.Count}");

                using (Transaction tx = new Transaction(doc, "Разделение панелей"))
                {
                    Log("Начало основной транзакции");
                    tx.Start();

                    foreach (var panel in panelsToSplit)
                    {
                        Log($"Обработка панели: {panel.Id}");
                        // Передаём cuttingPanels в качестве параметра; имя параметра теперь "cutters" 
                        ProcessPanel(doc, uidoc, panel, cuttingPanels);
                    }

                    tx.Commit();
                    Log("Основная транзакция завершена");
                }
            }
            catch (Exception ex)
            {
                Log($"КРИТИЧЕСКАЯ ОШИБКА: {ex}");
                TaskDialog.Show("Ошибка", ex.Message);
                return Result.Failed;
            }

            Log("======== СКРИПТ УСПЕШНО ЗАВЕРШЕН ========");
            return Result.Succeeded;
        }

        //————————————————————————————————————————————————————————————————
        // 2. High-level per-panel logic
        //————————————————————————————————————————————————————————————————
        // For one AnalyticalPanel, computes the outer contour, calls SplitPanel, then deletes the original if at least two pieces were created.
        private void ProcessPanel(Document doc, UIDocument uidoc,
            AnalyticalPanel panel, List<AnalyticalPanel> cutters)
        {
            Log($"Обработка панели {panel.Id}");
            var validCutters = cutters.Where(c => c.Id != panel.Id).ToList();

            var outerContour = panel.GetOuterContour();
            if (outerContour == null || !outerContour.Any())
            {
                Log($"Не удалось получить контур панели {panel.Id}");
                return;
            }

            var newPanels = SplitPanel(doc, panel, outerContour, validCutters);
            Log($"SplitPanel вернул {newPanels.Count} новых панелей.");

            if (newPanels.Count >= 2)
            {
                Log($"Удаляю исходную панель {panel.Id} — создано {newPanels.Count} частей");
                doc.Delete(panel.Id);
            }
            else
            {
                Log($"Не удаляю исходную панель {panel.Id}. Новых частей = {newPanels.Count} (<2)");
            }
        }

        //————————————————————————————————————————————————————————————————
        // 3. Core splitting: from one panel → many
        //————————————————————————————————————————————————————————————————
        // Given an original panel and its outer CurveLoop, intersects it with each “cutter” plane, splits the loop iteratively, then builds new panels.
        private List<AnalyticalPanel> SplitPanel(
            Document doc,
            AnalyticalPanel original,
            CurveLoop baseContour,
            List<AnalyticalPanel> cuttingPanels)
        {
            Log($"SplitPanel START: исходная панель {original.Id}, режущих панелей = {cuttingPanels.Count}");

            var newPanels = new List<AnalyticalPanel>();

            // 1. Плоскость исходной панели
            Plane targetPlane = GetPanelPlane(baseContour);
            if (targetPlane == null)
            {
                Log("SplitPanel: не удалось определить плоскость исходной панели — выходим");
                return newPanels;
            }

            // 2. Формируем список линий разреза
            var splitLines = new List<Line>();
            foreach (var cutter in cuttingPanels)
            {
                Plane cutPlane = CreateCuttingPlane(doc, cutter);
                if (cutPlane == null)
                {
                    Log($"SplitPanel: не удалось создать режущую плоскость для панели {cutter.Id}");
                    continue;
                }

                Line L = GetIntersectionLine(baseContour, targetPlane, cutter, cutPlane);
                if (L != null)
                    splitLines.Add(L);
                else
                    Log($"SplitPanel: линии пересечения с панелью {cutter.Id} не найдены");
            }

            // 3. Фильтрация и дедупликация
            double tol = doc.Application.ShortCurveTolerance;
            splitLines = splitLines
                .Where(l => l.Length > tol)
                .Distinct(new LineEndpointsComparer(tol))
                .ToList();

            Log($"SplitPanel: после фильтрации осталось {splitLines.Count} валидных линий разреза");
            if (!splitLines.Any())
            {
                Log("SplitPanel: нет линий разреза — возвращаем пустой список");
                return newPanels;
            }

            // 4. Итеративное разбиение контура
            var loops = new List<CurveLoop> { baseContour };
            Log($"SplitPanel: стартовое кол-во контуров = {loops.Count}");

            for (int i = 0; i < splitLines.Count; i++)
            {
                var L = splitLines[i];
                Log($"SplitPanel: разрез #{i}: {FormatXYZ(L.GetEndPoint(0))} → {FormatXYZ(L.GetEndPoint(1))}");
                Log($"SplitPanel: контур(ов) ДО разреза #{i} = {loops.Count}");

                var nextLoops = new List<CurveLoop>();
                foreach (var loop in loops)
                {
                    var parts = SplitContourByLine2D(doc, loop, L, targetPlane);

                    if (parts != null && parts.Count > 0)
                        nextLoops.AddRange(parts);
                    else
                        nextLoops.Add(loop);
                }

                loops = nextLoops;
                Log($"SplitPanel: контур(ов) ПОСЛЕ разреза #{i} = {loops.Count}");
            }

            // 5. Создание новых панелей из полученных контуров
            Log($"SplitPanel: начинаю создание панелей из {loops.Count} контуров");
            doc.Regenerate();  // гарантируем актуальность модели

            for (int idx = 0; idx < loops.Count; idx++)
            {
                var rawLoop = loops[idx];
                Log($"SplitPanel: rawLoop[{idx}] — сегментов = {rawLoop.Count()}");

                // Очищаем и проектируем
                var cleanLoop = BuildCleanLoopOrderedSafe(rawLoop, targetPlane, tol);
                if (cleanLoop == null)
                {
                    Log($"SplitPanel: cleanLoop[{idx}] = null — пропускаем");
                    continue;
                }
                Log($"SplitPanel: cleanLoop[{idx}] сегментов = {cleanLoop.Count()}");

                try
                {
                    var panel = AnalyticalPanel.Create(doc, cleanLoop);
                    if (panel == null)
                    {
                        Log($"SplitPanel: Create вернул null для cleanLoop[{idx}]");
                        continue;
                    }

                    doc.Regenerate();  // сразу фиксируем створенную панель в модели
                    Log($"SplitPanel: успешно создана панель {panel.Id}");

                    // копируем тип и параметры
                    panel.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
                         .Set(original.GetTypeId());
                    CopyPanelParameters(original, panel);
                    CopyAnalyticalOpenings(doc, original, panel, cleanLoop);

                    newPanels.Add(panel);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException argEx)
                {
                    Log($"SplitPanel: ArgumentException [{idx}]: {argEx.Message}");
                }
                catch (Exception ex)
                {
                    Log($"SplitPanel: Ошибка создания панели [{idx}]: {ex.Message}");
                }
            }

            Log($"SplitPanel END: создано новых панелей = {newPanels.Count}");
            return newPanels;
        }

        //————————————————————————————————————————————————————————————————
        // 4. Split one 2D loop into two by a line (robust at vertices-on-line)
        //————————————————————————————————————————————————————————————————
        // Walks the loop’s vertices, classifies each point by signed distance to the splitting line, inserts intersections, and returns two new loops.
        private List<CurveLoop> SplitContourByLine2D(Document doc, CurveLoop contour, Line splittingLine, Plane plane)
        {
            double tol = Math.Max(doc.Application.ShortCurveTolerance, 1e-6);

            // Build an orthonormal frame on the plane (u, v, n)
            XYZ n = plane.Normal.Normalize();

            XYZ u = null;
            foreach (Curve c in contour)
            {
                var dir = (c.GetEndPoint(1) - c.GetEndPoint(0));
                if (dir.GetLength() > tol)
                {
                    var cand = dir - n.DotProduct(dir) * n; // project to plane
                    if (cand.GetLength() > tol) { u = cand.Normalize(); break; }
                }
            }
            if (u == null) return null;
            XYZ v = n.CrossProduct(u).Normalize();

            // Local 2D mapping
            Func<XYZ, (double x, double y)> to2D = p => {
                var d = p - plane.Origin;
                return (d.DotProduct(u), d.DotProduct(v));
            };

            // 2D splitting line (project its end points onto the plane)
            var s0 = to2D(splittingLine.GetEndPoint(0));
            var s1 = to2D(splittingLine.GetEndPoint(1));
            var dx = s1.x - s0.x;
            var dy = s1.y - s0.y;
            double len = Math.Sqrt(dx * dx + dy * dy);
            if (len < tol) return null;

            (double x, double y) d2 = (dx / len, dy / len);
            (double x, double y) n2 = (-d2.y, d2.x); // in-plane normal (perpendicular)

            // Signed distance to the splitting line
            Func<(double x, double y), double> sd = q => (q.x - s0.x) * n2.x + (q.y - s0.y) * n2.y;

            // Collect ordered vertices of contour in 3D and close
            var verts3D = new List<XYZ>();
            foreach (Curve c in contour)
            {
                var a = c.GetEndPoint(0);
                if (verts3D.Count == 0 || !verts3D.Last().IsAlmostEqualTo(a, tol))
                    verts3D.Add(a);
                var b = c.GetEndPoint(1);
                if (!a.IsAlmostEqualTo(b, tol))
                    verts3D.Add(b);
            }
            if (!verts3D.First().IsAlmostEqualTo(verts3D.Last(), tol))
                verts3D.Add(verts3D.First());

            var verts2D = verts3D.Select(to2D).ToList();
            var polyNeg = new List<XYZ>(); // side where signed distance < 0
            var polyPos = new List<XYZ>(); // side where signed distance > 0

            for (int i = 0; i < verts2D.Count - 1; i++)
            {
                var A3 = verts3D[i];
                var B3 = verts3D[i + 1];
                var A2 = verts2D[i];
                var B2 = verts2D[i + 1];

                double sa = sd(A2);
                double sb = sd(B2);
                int ia = SignWithTol(sa, tol);
                int ib = SignWithTol(sb, tol);

                // 1) Add A to the appropriate side(s)
                if (ia == 0)
                {
                    AddPointUnique(polyNeg, A3, tol);
                    AddPointUnique(polyPos, A3, tol);
                }
                else if (ia < 0)
                {
                    AddPointUnique(polyNeg, A3, tol);
                }
                else // ia > 0
                {
                    AddPointUnique(polyPos, A3, tol);
                }

                // 2) If the segment crosses the line, insert the intersection on both
                if (ia * ib < 0)
                {
                    double t = sa / (sa - sb); // in [0,1] ideally
                    XYZ I3 = A3 + t * (B3 - A3);
                    AddPointUnique(polyNeg, I3, tol);
                    AddPointUnique(polyPos, I3, tol);
                }
                else if (ib == 0 && ia != 0)
                {
                    // Edge ends exactly on the line: add B to both to keep the corner
                    AddPointUnique(polyNeg, B3, tol);
                    AddPointUnique(polyPos, B3, tol);
                }
                // Else: no crossing; B will be handled as next A in the next iteration
            }

            // Close, clean, and reorder each polygon
            List<CurveLoop> result = new List<CurveLoop>();
            foreach (var poly in new[] { polyNeg, polyPos })
            {
                var closed = CloseAndClean(poly, plane, tol);
                if (closed != null)
                {
                    ReorderLoopByPerpendicular(closed, plane, splittingLine);
                    result.Add(closed);
                }
            }

            return result;
        }

        //————————————————————————————————————————————————————————————————
        // 5. Clean & close each raw list → a valid CurveLoop
        //————————————————————————————————————————————————————————————————
        /// Projects a raw vertex list onto the given plane, removes consecutive duplicates, collapses tiny edges and colinear triples, ensures closure and CCW winding,  and constructs a valid CurveLoop (returns null if the loop is degenerate).
        private CurveLoop CloseAndClean(List<XYZ> verts, Plane plane, double tol)
        {
            // Project, dedup consecutive, enforce closure
            var v = new List<XYZ>();
            foreach (var p in verts)
            {
                var q = ProjectPointToPlane(p, plane);
                AddPointUnique(v, q, tol);
            }
            if (v.Count < 3) return null;
            if (!v.First().IsAlmostEqualTo(v.Last(), tol)) v.Add(v.First());

            // Merge only truly colinear steps (same direction, almost zero angle)
            v = CollapseTinyAndColinear(v, tol);
            if (v.Count < 4) return null; // need at least 3 unique + closure

            // Ensure a stable, CCW orientation relative to plane.Normal
            // Build 2D axes
            XYZ n = plane.Normal.Normalize();
            XYZ u = null;
            for (int i = 0; i < v.Count - 1 && u == null; i++)
            {
                var dir = v[i + 1] - v[i];
                if (dir.GetLength() > tol)
                {
                    var cand = dir - n.DotProduct(dir) * n;
                    if (cand.GetLength() > tol) u = cand.Normalize();
                }
            }
            if (u == null) u = XYZ.BasisX; // fallback
            XYZ w = n.CrossProduct(u).Normalize();
            Func<XYZ, (double x, double y)> to2D = p => {
                var d = p - plane.Origin;
                return (d.DotProduct(u), d.DotProduct(w));
            };
            // Signed area
            double area2 = 0.0;
            for (int i = 0; i < v.Count - 1; i++)
            {
                var a = to2D(v[i]);
                var b = to2D(v[i + 1]);
                area2 += (a.x * b.y - b.x * a.y);
            }
            if (area2 < 0) v.Reverse(); // make CCW

            // Build loop
            var loop = new CurveLoop();
            for (int i = 0; i < v.Count - 1; i++)
            {
                var a = v[i];
                var b = v[i + 1];
                if (a.DistanceTo(b) > tol)
                    loop.Append(Line.CreateBound(a, b));
            }

            // Sanity closure
            var start = loop.First().GetEndPoint(0);
            var end = loop.Last().GetEndPoint(1);
            if (!start.IsAlmostEqualTo(end, tol))
                loop.Append(Line.CreateBound(end, start));

            return loop;
        }


        //————————————————————————————————————————————————————————————————
        // 6. Re-order a closed loop so it starts at the “min” point perpendicular to cut
        //————————————————————————————————————————————————————————————————   
        // Returns a new CurveLoop whose first vertex is the one with smallest projection along the axis perpendicular to splittingLine.
        private CurveLoop ReorderLoopByPerpendicular(
            CurveLoop loop,
            Plane plane,
            Line splittingLine)
        {
            // 1. Extract ordered vertices (closing the polygon)
            var verts = new List<XYZ>();
            foreach (var c in loop)
            {
                var pt = c.GetEndPoint(0);
                if (verts.Count == 0 || !verts.Last().IsAlmostEqualTo(pt, 1e-9))
                    verts.Add(pt);
            }
            var endPt = loop.Last().GetEndPoint(1);
            if (!verts.Last().IsAlmostEqualTo(endPt, 1e-9))
                verts.Add(endPt);
            if (!verts.First().IsAlmostEqualTo(verts.Last(), 1e-9))
                verts.Add(verts.First());

            if (verts.Count < 4)
                return loop; // nothing to reorder

            // 2. Build 2D frame on plane
            var n = plane.Normal.Normalize();
            XYZ u = null;
            for (int i = 0; i < verts.Count - 1 && u == null; i++)
            {
                var d = verts[i + 1] - verts[i];
                var proj = d - n.DotProduct(d) * n;
                if (proj.GetLength() > 1e-9)
                    u = proj.Normalize();
            }
            if (u == null) u = XYZ.BasisX;
            var w = n.CrossProduct(u).Normalize();

            Func<XYZ, (double x, double y)> to2D = p =>
            {
                var d = p - plane.Origin;
                return (d.DotProduct(u), d.DotProduct(w));
            };

            // 3. Compute splitting line in 2D
            var s0 = to2D(splittingLine.GetEndPoint(0));
            var s1 = to2D(splittingLine.GetEndPoint(1));
            var lineDir2D = (x: s1.x - s0.x, y: s1.y - s0.y);
            var len2D = Math.Sqrt(lineDir2D.x * lineDir2D.x + lineDir2D.y * lineDir2D.y);
            if (len2D < 1e-9)
                return loop;

            var perpAxis = (x: -lineDir2D.y / len2D, y: lineDir2D.x / len2D);

            // 4. Find vertex with minimal projection on perpAxis
            int startIdx = 0;
            double minProj = double.MaxValue;
            for (int i = 0; i < verts.Count - 1; i++)
            {
                var v2 = to2D(verts[i]);
                double proj = (v2.x - s0.x) * perpAxis.x + (v2.y - s0.y) * perpAxis.y;
                if (proj < minProj)
                {
                    minProj = proj;
                    startIdx = i;
                }
            }

            // 5. Rotate vertices so startIdx is first, then close loop
            var rotated = new List<XYZ>();
            int count = verts.Count - 1;
            for (int k = 0; k < count; k++)
                rotated.Add(verts[(startIdx + k) % count]);
            rotated.Add(rotated[0]);

            // 6. Build new CurveLoop
            var newLoop = new CurveLoop();
            for (int i = 0; i < rotated.Count - 1; i++)
                newLoop.Append(Line.CreateBound(rotated[i], rotated[i + 1]));

            return newLoop;
        }

        //————————————————————————————————————————————————————————————————
        // 7. Build the final clean, closed loop from raw curves
        //————————————————————————————————————————————————————————————————
        // Projects each vertex to the panel plane, removes tiny/duplicate steps, merges true-colinear triples, and enforces CCW orientation.
        private CurveLoop BuildCleanLoopOrderedSafe(CurveLoop rawLoop, Plane plane, double tol)
        {
            // Collect ordered projected vertices
            var verts = new List<XYZ>();
            foreach (Curve c in rawLoop)
            {
                var p0 = ProjectPointToPlane(c.GetEndPoint(0), plane);
                var p1 = ProjectPointToPlane(c.GetEndPoint(1), plane);

                if (verts.Count == 0) verts.Add(p0);
                else if (!verts.Last().IsAlmostEqualTo(p0, tol)) verts.Add(p0);

                if (!p0.IsAlmostEqualTo(p1, tol)) verts.Add(p1);
            }

            if (verts.Count < 3) return null;
            if (!verts.First().IsAlmostEqualTo(verts.Last(), tol)) verts.Add(verts.First());

            verts = CollapseTinyAndColinear(verts, tol);
            if (verts.Count < 4) return null;

            // Build loop
            var loop = new CurveLoop();
            for (int i = 0; i < verts.Count - 1; i++)
            {
                var a = verts[i];
                var b = verts[i + 1];
                if (a.DistanceTo(b) > tol)
                    loop.Append(Line.CreateBound(a, b));
            }

            // Sanity closure
            var start = loop.First().GetEndPoint(0);
            var end = loop.Last().GetEndPoint(1);
            if (!start.IsAlmostEqualTo(end, tol))
                loop.Append(Line.CreateBound(end, start));

            return loop;
        }

        //————————————————————————————————————————————————————————————————
        // 8. Remove tiny edges & colinear triples from a point list
        //————————————————————————————————————————————————————————————————
        // Drops edges shorter than tol and merges only truly colinear triples (angle≈0°, same direction).
        private List<XYZ> CollapseTinyAndColinear(List<XYZ> v, double tol)
        {
            if (v.Count < 3) return v;

            // Remove tiny consecutive edges
            var outV = new List<XYZ>();
            for (int i = 0; i < v.Count; i++)
            {
                if (outV.Count == 0 || outV.Last().DistanceTo(v[i]) > tol)
                    outV.Add(v[i]);
            }

            if (outV.Count < 3) return outV;
            if (!outV.First().IsAlmostEqualTo(outV.Last(), tol)) outV.Add(outV.First());

            // Merge truly colinear triples (same direction). Be conservative.
            int iIdx = 0;
            while (outV.Count > 3 && iIdx < outV.Count - 2)
            {
                XYZ a = outV[iIdx], b = outV[iIdx + 1], c = outV[iIdx + 2];
                var abVec = b - a;
                var bcVec = c - b;
                double abLen = abVec.GetLength();
                double bcLen = bcVec.GetLength();

                if (abLen <= tol || bcLen <= tol)
                {
                    // Let tiny edges be removed by the next pass
                    iIdx++;
                    continue;
                }

                var ab = abVec / abLen;
                var bc = bcVec / bcLen;

                // Colinear if angle ~ 0 (same direction), not 180
                double dot = ab.DotProduct(bc); // 1 for same dir
                var cross = ab.CrossProduct(bc);
                double sinMag = cross.GetLength(); // ~0 if colinear

                if (sinMag <= 1e-6 && dot > 0.9999)
                {
                    // Remove middle vertex
                    outV.RemoveAt(iIdx + 1);
                    // do not advance iIdx to catch chains
                }
                else
                {
                    iIdx++;
                }
            }

            // Ensure closure again
            if (!outV.First().IsAlmostEqualTo(outV.Last(), tol))
                outV.Add(outV.First());

            return outV;
        }

        //————————————————————————————————————————————————————————————————
        // 9. Tiny sign test with tolerance
        //————————————————————————————————————————————————————————————————
        // Returns –1, 0 or +1 if s is negative, near zero, or positive (|s| ≤ tol).
        private static int SignWithTol(double s, double tol)
        {
            if (s > tol) return 1;
            if (s < -tol) return -1;
            return 0;
        }

        //————————————————————————————————————————————————————————————————
        // 10. Safe push: only add point if not within tol of last
        //————————————————————————————————————————————————————————————————
        // Appends p to dst only if dst is empty or last point is > tol away.
        private static void AddPointUnique(List<XYZ> dst, XYZ p, double tol)
        {
            if (dst.Count == 0 || !dst[dst.Count - 1].IsAlmostEqualTo(p, tol))
                dst.Add(p);
        }

        //————————————————————————————————————————————————————————————————
        // 11. Fit a plane through the panel contour
        //————————————————————————————————————————————————————————————————
        // Takes the first three non-colinear contour points and builds a Plane.
        private Plane GetPanelPlane(CurveLoop contour)
        {
            List<XYZ> pts = new List<XYZ>();
            foreach (Curve c in contour)
            {
                pts.Add(c.GetEndPoint(0));
            }
            if (pts.Count < 3)
                return null;

            for (int i = 0; i < pts.Count - 2; i++)
            {
                XYZ v1 = pts[i + 1] - pts[i];
                for (int j = i + 2; j < pts.Count; j++)
                {
                    XYZ v2 = pts[j] - pts[i];
                    XYZ normal = v1.CrossProduct(v2);
                    if (normal.GetLength() > Tolerance)
                        return Plane.CreateByNormalAndOrigin(normal.Normalize(), pts[i]);
                }
            }
            return null;
        }


        //————————————————————————————————————————————————————————————————
        // 12. Build an infinite cutting plane from a cutter’s contour
        //————————————————————————————————————————————————————————————————
        // Finds three non-colinear points on the cutter’s outer contour and returns the Plane. Optionally creates a SketchPlane for debug.
        private Plane CreateCuttingPlane(Document doc, AnalyticalPanel cutter)
        {
            var loop = cutter.GetOuterContour();
            if (loop == null || !loop.Any()) return null;

            var pts = loop.Select(c => c.GetEndPoint(0)).ToList();
            Plane plane = null;

            for (int i = 0; i < pts.Count - 2 && plane == null; i++)
            {
                var v1 = pts[i + 1] - pts[i];
                for (int j = i + 2; j < pts.Count; j++)
                {
                    var v2 = pts[j] - pts[i];
                    var n = v1.CrossProduct(v2);
                    if (n.GetLength() > Tolerance)
                    {
                        plane = Plane.CreateByNormalAndOrigin(n.Normalize(), pts[i]);
                        break;
                    }
                }
            }
            if (plane == null) return null;

            // SketchPlane for debug
            SketchPlane.Create(doc, plane);
            return plane;
        }

        //————————————————————————————————————————————————————————————————
        // 13. Intersection line between two planes, clipped to target contour
        //————————————————————————————————————————————————————————————————
        // Computes the 3D line of intersection between targetPlane and cuttingPlane, then bounds it by the edges of targetContour.
        private Line GetIntersectionLine(
            CurveLoop targetContour,
            Plane targetPlane,
            AnalyticalPanel cutter,
            Plane cuttingPlane)
        {
            // параллельность?
            if (Math.Abs(targetPlane.Normal.DotProduct(cuttingPlane.Normal)) > 0.999)
                return null;

            var dir = targetPlane.Normal.CrossProduct(cuttingPlane.Normal).Normalize();
            double d1 = targetPlane.Normal.DotProduct(targetPlane.Origin);
            double d2 = cuttingPlane.Normal.DotProduct(cuttingPlane.Origin);
            var nCross = targetPlane.Normal.CrossProduct(cuttingPlane.Normal);
            double denom = nCross.DotProduct(nCross);
            if (denom < Tolerance) return null;

            // точка на прямой
            var origin = ((d1 * cuttingPlane.Normal - d2 * targetPlane.Normal)
                          .CrossProduct(nCross)) / denom;

            // ищем реальные пересечения с гранями контура
            var hits = new List<XYZ>();
            foreach (var c in targetContour.OfType<Line>())
            {
                var unb = Line.CreateUnbound(origin, dir);
                if (c.Intersect(unb, out IntersectionResultArray arr) == SetComparisonResult.Overlap
                    && arr != null && arr.Size > 0)
                {
                    hits.Add(arr.get_Item(0).XYZPoint);
                }
            }
            if (hits.Count < 2) return null;

            // сортируем по проекции на dir
            hits.Sort((a, b) =>
                (a - origin).DotProduct(dir).CompareTo((b - origin).DotProduct(dir)));

            return Line.CreateBound(hits.First(), hits.Last());
        }

        //————————————————————————————————————————————————————————————————
        // 14. Project an XYZ point onto a Plane
        //————————————————————————————————————————————————————————————————
        // Returns the perpendicular projection of point onto plane.
        private XYZ ProjectPointToPlane(XYZ point, Plane plane)
        {
            double distance = (point - plane.Origin).DotProduct(plane.Normal);
            return point - distance * plane.Normal;
        }

        //————————————————————————————————————————————————————————————————
        // 15. Copy all writable parameters from one panel to another
        //————————————————————————————————————————————————————————————————
        // Iterates source.Parameters and, for each non-read-only match in target, assigns its value.
        private void CopyPanelParameters(AnalyticalPanel source, AnalyticalPanel target)
        {
            try
            {
                Log($"Начало копирования параметров с панели {source.Id} на панель {target.Id}");
                foreach (Parameter sourceParam in source.Parameters)
                {
                    if (!sourceParam.IsReadOnly)
                    {
                        // Используем имя параметра для поиска соответствующего в target
                        string paramName = sourceParam.Definition.Name;
                        Parameter targetParam = target.LookupParameter(paramName);
                        if (targetParam != null && !targetParam.IsReadOnly)
                        {
                            try
                            {
                                switch (sourceParam.StorageType)
                                {
                                    case StorageType.Double:
                                        targetParam.Set(sourceParam.AsDouble());
                                        break;
                                    case StorageType.Integer:
                                        targetParam.Set(sourceParam.AsInteger());
                                        break;
                                    case StorageType.String:
                                        targetParam.Set(sourceParam.AsString());
                                        break;
                                    case StorageType.ElementId:
                                        targetParam.Set(sourceParam.AsElementId());
                                        break;
                                }
                                Log($"Скопирован параметр '{paramName}' со значением {targetParam.AsValueString()}.");
                            }
                            catch (Exception exParam)
                            {
                                Log($"Ошибка копирования параметра '{paramName}': {exParam.Message}");
                            }
                        }
                        else
                        {
                            Log($"Параметр '{paramName}' отсутствует или недоступен для записи на целевой панели.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Общая ошибка при копировании параметров: {ex.Message}");
            }
        }

        //————————————————————————————————————————————————————————————————
        // 16. Copy openings (door/window holes) into new panels
        //————————————————————————————————————————————————————————————————
        // For each opening on source, projects and clips its contour and, if its centroid lies inside the new panel, creates it.
        private void CopyAnalyticalOpenings(Document doc, AnalyticalPanel source, AnalyticalPanel target, CurveLoop targetContour)
        {
            try
            {
                Log($"Копирование проемов с панели {source.Id} на {target.Id}");
                var openingIds = source.GetAnalyticalOpeningsIds() ?? new HashSet<ElementId>();
                Log($"Найдено проемов: {openingIds.Count}");

                foreach (ElementId openingId in openingIds)
                {
                    var openingElement = doc.GetElement(openingId);
                    if (!(openingElement is AnalyticalOpening opening))
                        continue;

                    Log($"Копирование проема {openingId}");
                    CurveLoop openingContour = opening.GetOuterContour();
                    if (openingContour == null || !openingContour.Any())
                    {
                        Log($"Не удалось получить контур для проема {openingId} или он пуст.");
                        continue;
                    }

                    // Вычисляем центроид проёма
                    XYZ centroid = GetCentroid(openingContour);

                    // Проверяем, попадает ли центроид в новый контур панели
                    if (!IsPointInsideCurveLoop(targetContour, centroid))
                    {
                        Log($"Проем {openingId}: центроид {FormatXYZ(centroid)} не лежит внутри нового контура панели {target.Id}. Пропускаем копирование.");
                        continue;
                    }

                    // Проецируем контур проёма на плоскость панели (если требуется)
                    CurveLoop adjustedOpeningContour = ProjectCurveLoopToPlane(openingContour, targetContour);
                    if (adjustedOpeningContour == null || !adjustedOpeningContour.Any())
                    {
                        Log($"Проем {openingId}: не удалось получить корректный проецированный контур для панели {target.Id}.");
                        continue;
                    }

                    // Альтернативно можно вычислить пересечение adjustedOpeningContour с targetContour,
                    // чтобы гарантировать, что он полностью лежит внутри новой панели.
                    CurveLoop finalOpeningContour = GetIntersectionCurveLoop(adjustedOpeningContour, targetContour);
                    if (finalOpeningContour == null || !finalOpeningContour.Any())
                    {
                        Log($"Проем {openingId}: пересечение с новым контуром не дало корректного результата. Пропускаем создание проема.");
                        continue;
                    }

                    var newOpening = AnalyticalOpening.Create(doc, finalOpeningContour, target.Id);
                    if (newOpening != null)
                    {
                        Log($"Создан новый проем {newOpening.Id} для панели {target.Id}");
                    }
                    else
                    {
                        Log($"Не удалось создать новый проем для панели {target.Id} из проема {openingId}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка копирования проемов: {ex.Message}");
            }
        }

        //————————————————————————————————————————————————————————————————
        // 17. Compute centroid of a closed CurveLoop
        //————————————————————————————————————————————————————————————————
        // Averages the loop’s first endpoint of each curve to find the centroid.
        private XYZ GetCentroid(CurveLoop loop)
        {
            List<XYZ> points = new List<XYZ>();
            foreach (Curve curve in loop)
            {
                points.Add(curve.GetEndPoint(0));
            }
            double sumX = 0, sumY = 0, sumZ = 0;
            int count = points.Count;
            foreach (XYZ pt in points)
            {
                sumX += pt.X;
                sumY += pt.Y;
                sumZ += pt.Z;
            }
            return new XYZ(sumX / count, sumY / count, sumZ / count);
        }

        //————————————————————————————————————————————————————————————————
        // 18. 2D point-in-polygon test by ray-casting
        //————————————————————————————————————————————————————————————————
        // Projects both polygon and testPoint to the loop’s plane then runs a standard ray-casting test.
        private bool IsPointInsideCurveLoop(CurveLoop loop, XYZ testPoint)
        {
            List<XYZ> pts = new List<XYZ>();
            foreach (Curve c in loop)
            {
                pts.Add(c.GetEndPoint(0));
            }
            if (pts.Count < 3)
                return false;

            XYZ v1 = pts[1] - pts[0];
            XYZ v2 = pts[2] - pts[0];
            XYZ normal = v1.CrossProduct(v2).Normalize();
            Plane plane = Plane.CreateByNormalAndOrigin(normal, pts[0]);

            XYZ uDir = (pts[1] - pts[0]).Normalize();
            XYZ vDir = normal.CrossProduct(uDir);

            // Функция для проекции точки в 2D с использованием собственного Vector2
            Func<XYZ, Vector2> projectTo2D = (XYZ p) =>
            {
                double u = (p - pts[0]).DotProduct(uDir);
                double v = (p - pts[0]).DotProduct(vDir);
                return new Vector2((float)u, (float)v);
            };

            List<Vector2> poly = new List<Vector2>();
            foreach (XYZ p in pts)
            {
                poly.Add(projectTo2D(p));
            }
            Vector2 ptTest = projectTo2D(testPoint);

            bool inside = false;
            for (int i = 0, j = poly.Count - 1; i < poly.Count; j = i++)
            {
                // Реализуем алгоритм Ray-casting для собственного Vector2.
                if (((poly[i].Y > ptTest.Y) != (poly[j].Y > ptTest.Y)) &&
                    (ptTest.X < (poly[j].X - poly[i].X) * (ptTest.Y - poly[i].Y) / (poly[j].Y - poly[i].Y) + poly[i].X))
                {
                    inside = !inside;
                }
            }
            return inside;
        }

        //————————————————————————————————————————————————————————————————
        // 19. Project one CurveLoop onto another loop’s plane
        //————————————————————————————————————————————————————————————————
        // Takes each segment of loop, projects its endpoints perpendicularly onto the plane of targetContour, and reconstructs a new loop.
        private CurveLoop ProjectCurveLoopToPlane(CurveLoop loop, CurveLoop targetContour)
        {
            // Получим плоскость для targetContour.
            List<XYZ> pts = new List<XYZ>();
            foreach (Curve c in targetContour)
            {
                pts.Add(c.GetEndPoint(0));
            }
            if (pts.Count < 3)
                return null;
            XYZ v1 = pts[1] - pts[0];
            XYZ v2 = pts[2] - pts[0];
            XYZ normal = v1.CrossProduct(v2).Normalize();
            Plane plane = Plane.CreateByNormalAndOrigin(normal, pts[0]);

            CurveLoop newLoop = new CurveLoop();
            foreach (Curve curve in loop)
            {
                // Проецируем начальную и конечную точки
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                XYZ projStart = start - (start - plane.Origin).DotProduct(normal) * normal;
                XYZ projEnd = end - (end - plane.Origin).DotProduct(normal) * normal;
                // Создаем линию между проецированными точками.
                Line projLine = Line.CreateBound(projStart, projEnd);
                newLoop.Append(projLine);
            }
            return newLoop;
        }

        //————————————————————————————————————————————————————————————————
        // 20. Clip one loop by another—naïve centroid-inside check
        //————————————————————————————————————————————————————————————————
        // If loop1’s centroid lies inside loop2, returns loop1; otherwise null.
        private CurveLoop GetIntersectionCurveLoop(CurveLoop loop1, CurveLoop loop2)
        {
            XYZ centroid = GetCentroid(loop1);
            if (IsPointInsideCurveLoop(loop2, centroid))
            {
                // Дополнительно можно вернуть уже обработанный (например, проецированный) контур.
                return loop1;
            }
            return null;
        }

        //————————————————————————————————————————————————————————————————
        // 21. Debug & trace logging
        //————————————————————————————————————————————————————————————————
        private void Log(string message)
        {
            try
            {
                string entry = $"{DateTime.Now:HH:mm:ss} - {message}";
                File.AppendAllText(LogPath, entry + Environment.NewLine);
                Trace.WriteLine(entry);        // вывод в окно Debug/Trace
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка записи лога", ex.Message);
            }
        }

        //————————————————————————————————————————————————————————————————
        // 22. Format an XYZ for logs
        //————————————————————————————————————————————————————————————————
        private string FormatXYZ(XYZ point)
        {
            return $"X={point.X:F4}, Y={point.Y:F4}, Z={point.Z:F4}";
        }

        //————————————————————————————————————————————————————————————————
        // 23. Nested support types
        //————————————————————————————————————————————————————————————————
        public struct Vector2
        {
            public float X;
            public float Y;

            public Vector2(float x, float y)
            {
                X = x;
                Y = y;
            }
        }

        private class LineEndpointsComparer : IEqualityComparer<Line>
        {
            private readonly double _tol;
            public LineEndpointsComparer(double tolerance) => _tol = tolerance;

            public bool Equals(Line a, Line b)
            {
                bool sameDir = a.GetEndPoint(0).IsAlmostEqualTo(b.GetEndPoint(0), _tol)
                             && a.GetEndPoint(1).IsAlmostEqualTo(b.GetEndPoint(1), _tol);
                bool oppoDir = a.GetEndPoint(0).IsAlmostEqualTo(b.GetEndPoint(1), _tol)
                             && a.GetEndPoint(1).IsAlmostEqualTo(b.GetEndPoint(0), _tol);
                return sameDir || oppoDir;
            }
            public int GetHashCode(Line l) => 0;
        }

        #region Фильтры
        private class AnalyticalPanelFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is AnalyticalPanel;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        private class XYZComparer : IEqualityComparer<XYZ>
        {
            private readonly double _tolerance;
            public XYZComparer(double tolerance) => _tolerance = tolerance;
            public bool Equals(XYZ a, XYZ b) => a.DistanceTo(b) < _tolerance;
            public int GetHashCode(XYZ obj) => obj.X.GetHashCode() ^ obj.Y.GetHashCode() ^ obj.Z.GetHashCode();
        }
        #endregion
    }
}
