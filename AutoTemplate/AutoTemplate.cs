﻿using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using static AutoTemplate.AutoTemplate;
using static System.Windows.Forms.LinkLabel;

namespace AutoTemplate
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AutoTemplate : IExternalCommand
    {
        // 需要載入的磁磚族群
        public class Tiles
        {
            public static string tiles = "VC73001A 純白霧面7.5*30丁掛磚 對縫";
        }
        // 計算圖型的長度與高度
        public class LengthOrHeight
        {
            public int count { get; set; }
            public double heightOrHeight { get; set; }
        }
        // 儲存座標點
        public class PointToMatrix
        {
            public int cols { get; set; }
            public int rows { get; set; }
            public XYZ xyz { get; set; }
            public int isRectangle { get; set; }
        }
        // 最大矩形資訊結構
        public class Rectangle
        {
            public int MaxArea { get; set; }
            public (int Row, int Col) TopLeft { get; set; }
            public (int Row, int Col) BottomRight { get; set; }
        }
        // 干涉的元件
        public class IntersectionElem
        {
            public Element hostElem = null;
            public List<Element> elemList = new List<Element>();
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            DateTime timeStart = DateTime.Now; // 計時開始 取得目前時間

            // 儲存所有柱
            List<Element> columns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Columns).WhereElementIsNotElementType().ToList();
            // 取得所有的牆
            List<Wall> walls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().Cast<Wall>().ToList();
            using (Transaction trans = new Transaction(doc, "自動放置模板"))
            {
                // 關閉警示視窗
                FailureHandlingOptions options = trans.GetFailureHandlingOptions();
                CloseWarnings closeWarnings = new CloseWarnings();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(closeWarnings);
                trans.SetFailureHandlingOptions(options);
                trans.Start();

                // 取得磁磚族群的FamilySymbol, 並啟動
                List<FamilySymbol> familySymbolList = new List<FamilySymbol>();
                familySymbolList = GetFamilySymbols(doc);
                foreach (FamilySymbol familySymbol in familySymbolList)
                {
                    if (!familySymbol.IsActive) { familySymbol.Activate(); }
                }
                FamilySymbol fs = familySymbolList.FirstOrDefault();

                // 找到邊界外圍框
                List<Element> elems = ColumnsAndWallsElems(doc); // 查詢所有柱牆的Element                
                List<IntersectionElem> list = IntersectGroup(doc, elems); // ElementId第一個有接觸到的所有元件
                List<Solid> solids = new List<Solid>(); // 儲存所有柱牆的Solid                
                foreach (Element elem in elems) { foreach (Solid solid in GetSolids(doc, elem)) { solids.Add(solid); } }
                Solid hostSolid = UnionSolids(solids, solids[0]); // 將所有Solid聯集

                if (null != hostSolid)
                {
                    List<Face> sideFaces = GetFaces(new List<Solid> { hostSolid }, "side");
                    Face maxFace = sideFaces.OrderByDescending(x => x.Area).FirstOrDefault();
                    sideFaces = new List<Face> { maxFace };
                    foreach (Face face in sideFaces)
                    {
                        (List<PointToMatrix>, List<List<Curve>>, int, int, List<Curve>) pointToMatrix = GenerateUniformPoints(face); // 將Face網格化, 每100cm佈一個點
                        List<PointToMatrix> pointToMatrixs = pointToMatrix.Item1;
                        List<List<Curve>> curveLoopList = pointToMatrix.Item2; // 開口的封閉曲線
                        int rows = pointToMatrix.Item3;
                        int cols = pointToMatrix.Item4;
                        List<Curve> drawCurves = pointToMatrix.Item5;

                        // 區塊分割計算
                        List<Curve> leftRightCurves = new List<Curve>(); // 左右線段
                        List<Curve> leftCurves = new List<Curve>(); // 左線段
                        List<Curve> rightCurves = new List<Curve>(); // 右線段
                        List<Curve> upDownCurves = new List<Curve>(); // 上下線段
                        List<Curve> upCurves = new List<Curve>(); // 上線段
                        List<Curve> downCurves = new List<Curve>(); // 下線段

                        // 儲存所有線段
                        foreach (List<Curve> curveLoop in curveLoopList)
                        {
                            foreach(Curve curve in curveLoop)
                            {
                                Line line = curve as Line;
                                double directionX = ToZeroIfCloseToZero(line.Direction.X);
                                double directionY = ToZeroIfCloseToZero(line.Direction.Y);
                                double directionZ = ToZeroIfCloseToZero(line.Direction.Z);
                                if (directionX.Equals(0) && directionY.Equals(0))
                                {
                                    if (directionZ.Equals(1)) { leftCurves.Add(curve); leftRightCurves.Add(curve); }
                                    else if (directionZ.Equals(-1)) { rightCurves.Add(curve); leftRightCurves.Add(curve); }
                                }
                                if (directionZ.Equals(0) && directionY.Equals(0))
                                {
                                    if (directionX.Equals(-1)) { upCurves.Add(curve); upDownCurves.Add(curve); }
                                    else if (directionX.Equals(1)) { downCurves.Add(curve); upDownCurves.Add(curve); }
                                }
                                //DrawLine(doc, curve); doc.Regenerate(); uidoc.RefreshActiveView();
                            }
                        }
                        // 進行線段左右排序
                        if (leftRightCurves.Count > 2) // 有開口
                        {
                            BoundingBoxUV bboxUV = face.GetBoundingBox();
                            XYZ minXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
                            XYZ maxXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
                            leftRightCurves = leftRightCurves.OrderBy(x => x.Project(minXYZ).Distance).ToList();
                            leftCurves = leftCurves.OrderBy(x => x.Project(minXYZ).Distance).ToList();
                            rightCurves = rightCurves.OrderBy(x => x.Project(maxXYZ).Distance).ToList();
                            Curve leftestCurve = leftCurves.OrderBy(x => x.Project(minXYZ).Distance).FirstOrDefault(); // 最左邊的邊
                            Curve rightestCurve = rightCurves.OrderBy(x => x.Project(maxXYZ).Distance).FirstOrDefault(); // 最右邊的邊
                            leftRightCurves = new List<Curve>();
                            // 將邊界與最旁邊的邊連結成矩形, 左線段連結左邊、右線段連結右邊
                            PlanarFace planarFace = face as PlanarFace;
                            UseCurveLoopToCreateSolid(uidoc, doc, planarFace.FaceNormal, leftCurves, leftestCurve, upDownCurves);
                            //foreach (Curve drawCurve in leftRightCurves) { DrawLine(doc, drawCurve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                        }
                        else // 沒有開口
                        {

                        }


                        //List<XYZ> normalizes = new List<XYZ>();
                        //foreach (Curve curve in drawCurves)
                        //{
                        //    for(int i = 0; i < curve.Tessellate().Count - 1; i++)
                        //    {
                        //        XYZ startPoint = curve.GetEndPoint(i);
                        //        XYZ endPoint = curve.GetEndPoint(i + 1);
                        //        XYZ direction = endPoint - startPoint;
                        //        double directionX = ToZeroIfCloseToZero(direction.Normalize().X);
                        //        double directionY = ToZeroIfCloseToZero(direction.Normalize().Y);
                        //        double directionZ = ToZeroIfCloseToZero(direction.Normalize().Z);
                        //        XYZ sameDirection = normalizes.Where(x => x.X.Equals(directionX) && x.Y.Equals(directionY) && x.Z.Equals(directionZ)).FirstOrDefault();
                        //        if (sameDirection == null) { normalizes.Add(new XYZ(directionX, directionY, directionZ)); }
                        //    }
                        //}
                        //normalizes = normalizes.Distinct().ToList();
                        //XYZ normalize = normalizes.LastOrDefault();
                        //foreach (Curve curve in drawCurves)
                        //{
                        //    for (int i = 0; i < curve.Tessellate().Count - 1; i++)
                        //    {
                        //        XYZ startPoint = curve.GetEndPoint(i);
                        //        XYZ endPoint = curve.GetEndPoint(i + 1);
                        //        XYZ direction = endPoint - startPoint;

                        //        double directionX = ToZeroIfCloseToZero(direction.Normalize().X);
                        //        double directionY = ToZeroIfCloseToZero(direction.Normalize().Y);
                        //        double directionZ = ToZeroIfCloseToZero(direction.Normalize().Z);
                        //        double normalizeX = ToZeroIfCloseToZero(normalize.X);
                        //        double normalizeY = ToZeroIfCloseToZero(normalize.Y);
                        //        double normalizeZ = ToZeroIfCloseToZero(normalize.Z);
                        //        if ((directionX == normalizeX) && (directionY == normalizeY) && (directionZ == normalizeZ))
                        //        {
                        //            //DrawLine(doc, curve); doc.Regenerate(); uidoc.RefreshActiveView();
                        //        }
                        //        if (directionX == XYZ.BasisX.X && directionY == XYZ.BasisX.Y && directionZ == XYZ.BasisX.Z) { upDownCurves.Add(curve); }
                        //        else if (directionX == -XYZ.BasisX.X && directionY == -XYZ.BasisX.Y && directionZ == -XYZ.BasisX.Z) { upDownCurves.Add(curve); }
                        //        //else if (directionX == XYZ.BasisZ.X && directionY == XYZ.BasisZ.Y && directionZ == XYZ.BasisZ.Z) { leftRightCurves.Add(curve); }
                        //        //else if (directionX == -XYZ.BasisZ.X && directionY == -XYZ.BasisZ.Y && directionZ == -XYZ.BasisZ.Z) { leftRightCurves.Add(curve); }
                        //    }
                        //}

                        //drawCurves = new List<Curve>();
                        //foreach (Curve upDownCurve in upDownCurves) 
                        //{ 
                        //    for(int i = 0; i < upDownCurve.Tessellate().Count - 1; i++)
                        //    {
                        //        XYZ startXYZ = upDownCurve.Tessellate()[i];
                        //        XYZ endXYZ = upDownCurve.Tessellate()[i + 1];
                        //        try
                        //        {
                        //            XYZ startToCurveXYZ = leftRightCurves.OrderByDescending(x => x.Distance(startXYZ)).FirstOrDefault().Project(startXYZ).XYZPoint;
                        //            drawCurves.Add(Line.CreateBound(startXYZ, startToCurveXYZ));
                        //        }
                        //        catch (Exception) { }
                        //        try
                        //        {
                        //            XYZ endToCurveXYZ = leftRightCurves.OrderByDescending(x => x.Distance(endXYZ)).FirstOrDefault().Project(endXYZ).XYZPoint;
                        //            drawCurves.Add(Line.CreateBound(endXYZ, endToCurveXYZ));
                        //        }
                        //        catch (Exception) { }
                        //    }
                        //}
                        //foreach (Curve drawCurve in drawCurves) { DrawLine(doc, drawCurve); doc.Regenerate(); uidoc.RefreshActiveView(); }

                        //(List<Line>, List<Line>, List<Line>) leftRightLines = LeftRightLines(face, curveLoopList);
                        //List<Line> leftLines = leftRightLines.Item1; // 左邊的線段
                        //List<Line> rightLines = leftRightLines.Item2; // 右邊的線段
                        //List<Line> allLines = leftRightLines.Item3; // 全部的線段
                        ////foreach (Line line in allLines) { DrawLine(doc, line); doc.Regenerate(); uidoc.RefreshActiveView(); }

                        //CurveLoop maxCurveLoops = face.GetEdgesAsCurveLoops().ToList().OrderByDescending(x => x.GetExactLength()).ToList().FirstOrDefault(); // Face最大的範圍邊界
                        //List<Line> maxCurveList = new List<Line>();
                        //foreach (Curve maxCurve in maxCurveLoops) { maxCurveList.Add(Line.CreateBound(maxCurve.Tessellate()[0], maxCurve.Tessellate()[maxCurve.Tessellate().Count - 1])); }
                        //maxCurveList = maxCurveList.Where(x => x.Direction.Z > 0 || x.Direction.Z < 0).ToList();
                        //if (maxCurveList.Count > 1)
                        //{
                        //    BoundingBoxUV bboxUV = face.GetBoundingBox();
                        //    Line minXYZLine = maxCurveList.OrderBy(x => VectorDistance(x.Origin, face.Evaluate(bboxUV.Min))).FirstOrDefault();
                        //    Line line = allLines[0];
                        //    XYZ point1 = minXYZLine.Project(line.GetEndPoint(0)).XYZPoint;
                        //    XYZ point2 = minXYZLine.Project(line.GetEndPoint(1)).XYZPoint;

                        //    // 左邊框
                        //    List<Curve> curves = new List<Curve>();
                        //    try
                        //    {
                        //        curves.Add(Line.CreateBound(point1, point2));
                        //        curves.Add(Line.CreateBound(point2, line.GetEndPoint(1)));
                        //        curves.Add(line);
                        //        curves.Add(Line.CreateBound(line.GetEndPoint(0), point1));
                        //        //foreach (Line curve in curves) { DrawLine(doc, curve); }
                        //        //allLines.Remove(line);
                        //    }
                        //    catch { }
                        //    // 右邊框
                        //    try
                        //    {
                        //        Line maxXYZLine = maxCurveList.OrderByDescending(x => VectorDistance(x.Origin, face.Evaluate(bboxUV.Min))).FirstOrDefault();
                        //        line = allLines[allLines.Count - 1];
                        //        point1 = maxXYZLine.Project(line.GetEndPoint(0)).XYZPoint;
                        //        point2 = maxXYZLine.Project(line.GetEndPoint(1)).XYZPoint;
                        //        curves = new List<Curve>();
                        //        curves.Add(Line.CreateBound(point1, point2));
                        //        curves.Add(Line.CreateBound(point2, line.GetEndPoint(1)));
                        //        curves.Add(line);
                        //        curves.Add(Line.CreateBound(line.GetEndPoint(0), point1));
                        //        //foreach (Line curve in curves) { DrawLine(doc, curve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                        //        //allLines.Remove(line);
                        //    }
                        //    catch {}
                        //    // 中間框
                        //    try
                        //    {
                        //        for (int i = 0; i < allLines.Count - 1; i += 2)
                        //        {
                        //            try
                        //            {
                        //                Line line1 = allLines[i];
                        //                Line line2 = allLines[i + 1];
                        //                if (line1.Length < line2.Length) { line1 = allLines[i + 1]; line2 = allLines[i]; }
                        //                point1 = line1.Project(line2.GetEndPoint(0)).XYZPoint;
                        //                point2 = line1.Project(line2.GetEndPoint(1)).XYZPoint;
                        //                curves = new List<Curve>();
                        //                curves.Add(Line.CreateBound(point1, point2));
                        //                curves.Add(Line.CreateBound(point2, line2.GetEndPoint(1)));
                        //                curves.Add(line2);
                        //                curves.Add(Line.CreateBound(line2.GetEndPoint(0), point1));
                        //                //foreach (Line curve in curves) { DrawLine(doc, curve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                        //            }
                        //            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                        //        }
                        //    }
                        //    catch { }
                        //}

                        //if (leftLines.Count > 1)
                        //{
                        //    List<Curve> curves = new List<Curve>();
                        //    Line line = leftLines[0];
                        //    XYZ point1 = line.Project(leftLines[1].GetEndPoint(0)).XYZPoint;
                        //    XYZ point2 = line.Project(leftLines[1].GetEndPoint(1)).XYZPoint;
                        //    curves.Add(Line.CreateBound(point1, point2));
                        //    curves.Add(Line.CreateBound(point2, leftLines[1].GetEndPoint(1)));
                        //    curves.Add(leftLines[1]);
                        //    curves.Add(Line.CreateBound(leftLines[1].GetEndPoint(0), point1));
                        //    foreach(Curve curve in curves) { DrawLine(doc, curve); }

                        //    //leftLines.Remove(leftLines[0]); leftLines.Remove(leftLines[1]);
                        //}
                        //if (rightLines.Count > 1)
                        //{
                        //    List<Curve> curves = new List<Curve>();
                        //    Line line = rightLines.Last();
                        //    XYZ point1 = line.Project(rightLines[rightLines.Count - 2].GetEndPoint(0)).XYZPoint;
                        //    XYZ point2 = line.Project(rightLines[rightLines.Count - 2].GetEndPoint(1)).XYZPoint;
                        //    curves.Add(Line.CreateBound(point1, point2));
                        //    curves.Add(Line.CreateBound(point2, rightLines[rightLines.Count - 2].GetEndPoint(1)));
                        //    curves.Add(rightLines[rightLines.Count - 2]);
                        //    curves.Add(Line.CreateBound(rightLines[rightLines.Count - 2].GetEndPoint(0), point1));
                        //    foreach (Curve curve in curves) { DrawLine(doc, curve); }

                        //    //rightLines.Remove(rightLines[rightLines.Count - 2]); rightLines.Remove(rightLines.Last());
                        //}


                        //Form1 form1 = new Form1(pointToMatrixs, cols, rows);
                        //form1.ShowDialog();

                        //// 分割幾何圖形成n個矩形
                        //List<Rectangle> rectangles = new List<Rectangle>();
                        //SaveRectangle(uidoc, doc, rows, cols, pointToMatrixs, rectangles);
                        //for (int i = 0; i < rows; i++)
                        //{
                        //    for (int j = 0; j < cols - 1; j++)
                        //    {
                        //        PointToMatrix start = pointToMatrixs.Where(x => x.rows.Equals(i) && x.cols.Equals(j)).FirstOrDefault();
                        //        PointToMatrix end = pointToMatrixs.Where(x => x.rows.Equals(i) && x.cols.Equals(j + 1)).FirstOrDefault();
                        //        if (start.isRectangle == 1 && end.isRectangle == 1) { DrawLine(doc, Line.CreateBound(start.xyz, end.xyz)); }
                        //    }
                        //}
                    }
                }

                //PutExterInteriorTiles(uidoc, doc, walls, fs); // 放置內外牆的磁磚

                trans.Commit();
                DateTime timeEnd = DateTime.Now; // 計時結束 取得目前時間
                TimeSpan totalTime = timeEnd - timeStart;
                //TaskDialog.Show("Revit", "完成，耗時：" + totalTime.Minutes + " 分 " + totalTime.Seconds + " 秒。\n\n");
            }

            return Result.Succeeded;
        }
        /// <summary>
        /// 將邊界與最旁邊的邊連結成矩形, 左線段連結左邊、右線段連結右邊, 並建立為Solid
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="faceNormal"></param>
        /// <param name="curves"></param>
        /// <param name="leftOrRightestCurve"></param>
        private void UseCurveLoopToCreateSolid(UIDocument uidoc, Document doc, XYZ faceNormal, List<Curve> curves, Curve leftOrRightestCurve, List<Curve> upDownCurves)
        {
            SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
            foreach (Curve curve in curves)
            {
                try
                {
                    XYZ point1 = curve.Tessellate()[0];
                    XYZ point2 = curve.Tessellate()[curve.Tessellate().Count - 1];
                    XYZ point3 = leftOrRightestCurve.Project(point2).XYZPoint;
                    XYZ point4 = leftOrRightestCurve.Project(point1).XYZPoint;
                    Curve curve1 = Line.CreateBound(point1, point2);
                    Curve curve2 = Line.CreateBound(point2, point3);
                    Curve curve3 = Line.CreateBound(point3, point4);
                    Curve curve4 = Line.CreateBound(point4, point1);
                    CurveLoop curveLoop = CurveLoop.Create(new List<Curve>() { curve1, curve2, curve3, curve4 });
                    IList<CurveLoop> curveLoops = new List<CurveLoop>() { curveLoop };
                    Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, faceNormal, 0.1, options);
                    DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                    ds.SetShape(new List<GeometryObject>() { solid });
                    doc.Regenerate();

                    List<XYZ> xyzs = new List<XYZ>();
                    List<Face> solidFaces = new List<Face>();
                    foreach (Face face in solid.Faces)
                    {
                        PlanarFace pf = face as PlanarFace;
                        double faceNormalX = ToZeroIfCloseToZero(pf.FaceNormal.X);
                        double faceNormalY = ToZeroIfCloseToZero(pf.FaceNormal.Y);
                        double faceNormalZ = ToZeroIfCloseToZero(pf.FaceNormal.Z);
                        if (faceNormal.X.Equals(-faceNormalX) && faceNormal.Y.Equals(-faceNormalY) && faceNormal.Z.Equals(-faceNormalZ))
                        {
                            solidFaces.Add(face);
                        }
                    }
                    if (solidFaces.Count > 0) 
                    { 
                        Face soildFace = solidFaces.FirstOrDefault();
                        List<Curve> containCurves = new List<Curve>();
                        List<Curve> upCurves = new List<Curve>();
                        List<Curve> downCurves = new List<Curve>();
                        foreach (Curve upDownCurve in upDownCurves)
                        {
                            bool trueOrFalse = false;
                            foreach (XYZ upDownXYZ in upDownCurve.Tessellate())
                            {
                                if (IsPointInsideSolid(faceNormal, soildFace, upDownXYZ)) { trueOrFalse = true; }
                                else { trueOrFalse = false; break; }
                            }
                            if (trueOrFalse) 
                            {
                                containCurves.Add(upDownCurve);

                                Line line = upDownCurve as Line;
                                double directionX = ToZeroIfCloseToZero(line.Direction.X);
                                double directionY = ToZeroIfCloseToZero(line.Direction.Y);
                                double directionZ = ToZeroIfCloseToZero(line.Direction.Z);
                                if (directionZ.Equals(0) && directionY.Equals(0))
                                {
                                    if (directionX.Equals(-1)) { upCurves.Add(upDownCurve); }
                                    else if (directionX.Equals(1)) { downCurves.Add(upDownCurve); }
                                }
                            }
                        }
                        if (containCurves.Count > 0)
                        {
                            // 只包含上方的線
                            if(upCurves.Count > 0 && downCurves.Count == 0)
                            {
                                Curve heightestCurve = upCurves.OrderBy(x => x.Tessellate()[0].Z).FirstOrDefault();
                                DrawLine(doc, heightestCurve); doc.Regenerate(); uidoc.RefreshActiveView();
                            }
                            // 只包含下方的線
                            if (upCurves.Count == 0 && downCurves.Count > 0)
                            {
                                Curve lowestCurve = downCurves.OrderByDescending(x => x.Tessellate()[0].Z).FirstOrDefault();
                                DrawLine(doc, lowestCurve); doc.Regenerate(); uidoc.RefreshActiveView();
                            }
                            // 包含上方與下方的線
                            if (upCurves.Count > 0 && downCurves.Count > 0)
                            {
                                Curve lowestCurve = downCurves.OrderByDescending(x => x.Tessellate()[0].Z).FirstOrDefault();
                                DrawLine(doc, lowestCurve); doc.Regenerate(); uidoc.RefreshActiveView();
                            }
                            //foreach (Curve drawCurve in containCurves)
                            //{
                            //    DrawLine(doc, drawCurve); doc.Regenerate(); uidoc.RefreshActiveView();
                            //}
                        }
                        else
                        {
                            foreach(Curve solidCurve in curveLoop)
                            {
                                DrawLine(doc, solidCurve); doc.Regenerate(); uidoc.RefreshActiveView();
                            }
                        }
                    }
                    doc.Delete(ds.Id);
                    doc.Regenerate();

                    //List<Element> elems = new List<Element>() { ds };
                    //List<IntersectionElem> list = IntersectGroup(doc, elems);
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        public bool IsPointInsideSolid(XYZ faceNormal, Face soildFace, XYZ point)
        {
            int intersections = 0;
            IntersectionResultArray results;
            Line ray = Line.CreateBound(point, point + faceNormal * 1000);
            if (soildFace.Intersect(ray, out results) == SetComparisonResult.Overlap && results != null) { intersections += results.Size; }

            // 若交點數為奇數，則點在Solid內
            bool trueOrFalse = false;
            if((intersections % 2) == 1) { trueOrFalse = true; }
            return trueOrFalse;
        }
        private double ToZeroIfCloseToZero(double value)
        {
            double threshold = 1e-14;
            return Math.Abs(value) < threshold ? 0 : value;
        }
        /// <summary>
        /// 查詢所有柱牆的Element
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private List<Element> ColumnsAndWallsElems(Document doc)
        {
            List<Element> elems = new List<Element>();

            ElementCategoryFilter columnsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Columns);
            ElementCategoryFilter wallsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            LogicalOrFilter filter = new LogicalOrFilter(columnsFilter, wallsFilter);
            elems = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements().ToList();

            return elems;
        }
        /// <summary>
        /// ElementId第一個有接觸到的所有元件
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="elems"></param>
        /// <returns></returns>
        private List<IntersectionElem> IntersectGroup(Document doc, List<Element> elems)
        {
            List<IntersectionElem> list = new List<IntersectionElem>();
            IntersectionElem intersectionElem = new IntersectionElem();

            foreach (Element elem in elems)
            {
                List<Element> elemList = new List<Element>();
                intersectionElem = new IntersectionElem();

                // 找到選取元件的輪廓線
                View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().Cast<View3D>().Where(x => x.Name.Equals("{3D}")).FirstOrDefault();
                // 創建BoundingBoxIntersectsFilter找到其他與之交接的元件
                BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(new Outline(elem.get_BoundingBox(view3D).Min, elem.get_BoundingBox(view3D).Max));
                // 排除點選元件本身
                ICollection<ElementId> idsExclude = new List<ElementId>() { elem.Id };
                // 存放到容器內, 兩個都是快篩, 所以順序不重要  
                elemList = new FilteredElementCollector(doc/*, doc.ActiveView.Id*/).Excluding(idsExclude).WherePasses(bbFilter).WhereElementIsNotElementType().ToElements().ToList();
                intersectionElem.hostElem = elem;
                intersectionElem.elemList = elemList;
                list.Add(intersectionElem);
            }

            return list;
        }
        /// <summary>
        /// 將所有Solid聯集
        /// </summary>
        /// <param name="solids"></param>
        /// <param name="hostSolid"></param>
        /// <returns></returns>
        private Solid UnionSolids(IList<Solid> solids, Solid hostSolid)
        {
            Solid unionSolid = null;
            foreach (Solid subSolid in solids)
            {
                if (subSolid.Volume > 0)
                {
                    try
                    {
                        unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(hostSolid, subSolid, BooleanOperationsType.Union);
                        hostSolid = unionSolid;
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
            }

            return hostSolid;
        }
        /// <summary>
        /// 儲存所有Solid
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="elem"></param>
        /// <returns></returns>
        private List<Solid> GetSolids(Document doc, Element elem)
        {
            List<Solid> solidList = new List<Solid>();

            // 1.讀取Geometry Option
            Options options = new Options();
            //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            // 得到幾何元素
            GeometryElement geomElem = elem.get_Geometry(options);
            List<Solid> solids = GeometrySolids(geomElem);
            foreach (Solid solid in solids)
            {
                solidList.Add(solid);
            }

            return solidList;
        }
        /// <summary>
        /// 取得Solid
        /// </summary>
        /// <param name="geoObj"></param>
        /// <returns></returns>
        private List<Solid> GeometrySolids(GeometryObject geoObj)
        {
            List<Solid> solids = new List<Solid>();
            if (geoObj is Solid)
            {
                Solid solid = (Solid)geoObj;
                if (solid.Faces.Size > 0)
                {
                    solids.Add(solid);
                }
            }
            if (geoObj is GeometryInstance)
            {
                GeometryInstance geoIns = geoObj as GeometryInstance;
                GeometryElement geometryElement = (geoObj as GeometryInstance).GetSymbolGeometry(geoIns.Transform); // 座標轉換
                foreach (GeometryObject o in geometryElement)
                {
                    solids.AddRange(GeometrySolids(o));
                }
            }
            else if (geoObj is GeometryElement)
            {
                GeometryElement geometryElement2 = (GeometryElement)geoObj;
                foreach (GeometryObject o in geometryElement2)
                {
                    solids.AddRange(GeometrySolids(o));
                }
            }
            return solids;
        }
        /// <summary>
        /// 取得幾何圖形的面
        /// </summary>
        /// <param name="solidList"></param>
        /// <returns></returns>
        private List<Face> GetFaces(List<Solid> solidList, string direction)
        {
            List<Face> faces = new List<Face>();
            foreach (Solid solid in solidList)
            {
                foreach (Face face in solid.Faces)
                {
                    //if (face.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate())) { faces.Add(face); }
                    double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                    if (direction.Equals("top")) { if (faceTZ > 0.0) { faces.Add(face); } } // 頂面
                    else if (direction.Equals("bottom")) { if (faceTZ < 0.0) { faces.Add(face); } } // 底面
                    else if (direction.Equals("bevel")) { if (faceTZ < 0 && faceTZ != -1) { faces.Add(face); } } // 斜邊(突沿)
                    else if (direction.Equals("side")) { if (faceTZ == 0.0) { faces.Add(face); } } // 側邊
                }
            }
            return faces;
        }
        /// <summary>
        /// 取得磁磚族群
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private List<FamilySymbol> GetFamilySymbols(Document doc)
        {
            List<FamilySymbol> familySymbols = new List<FamilySymbol>();
            List<string> tiles = new List<string>() { Tiles.tiles };
            List<FamilySymbol> familySymbolList = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            // 篩選出逃生標誌需要使用的族群
            List<FamilySymbol> saveFamilySymbols = new List<FamilySymbol>();
            foreach (string tile in tiles)
            {
                FamilySymbol saveFamilySymbol = familySymbolList.Where(x => x.Name.Equals(tile)).FirstOrDefault();
                if (saveFamilySymbol != null) { saveFamilySymbols.Add(saveFamilySymbol); }
            }
            saveFamilySymbols = saveFamilySymbols.OrderBy(x => x.Family.Name).ToList(); // 排序
            return saveFamilySymbols;
        }
        public static bool AreCurvesPerpendicular(Curve curve1, Curve curve2, double tolerance = 1e-9)
        {
            // 獲取兩條曲線的方向向量
            XYZ direction1 = GetCurveDirection(curve1);
            XYZ direction2 = GetCurveDirection(curve2);

            if (direction1 == null || direction2 == null)
            {
                throw new ArgumentException("曲線類型不支持計算方向向量");
            }

            // 計算點積
            double dotProduct = direction1.DotProduct(direction2);

            // 檢查點積是否接近 0
            return Math.Abs(dotProduct) < tolerance;
        }
        private static XYZ GetCurveDirection(Curve curve)
        {
            if (curve is Line line)
            {
                // 對於直線，使用起點和終點計算方向
                return (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
            }
            else if (curve is Arc arc)
            {
                // 對於弧線，取起點和弧線上的中點計算方向
                XYZ startPoint = arc.GetEndPoint(0);
                XYZ midPoint = arc.Evaluate(0.5, true); // 弧線的中間點
                return (midPoint - startPoint).Normalize();
            }

            // 其他類型的曲線不支持
            return null;
        }
        /// <summary>
        /// 放置內外牆的磁磚
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="doc"></param>
        /// <param name="walls"></param>
        /// <param name="fs"></param>
        private void PutExterInteriorTiles(UIDocument uidoc, Document doc, List<Wall> walls, FamilySymbol fs)
        {
            foreach (Wall wall in walls)
            {
                List<Solid> wallSolids = GetSolids(doc, wall);
                List<Face> wallFaces = GetFaces(wallSolids, "side");
                Level level = doc.GetElement(wall.LevelId) as Level;
                try
                {
                    if (level != null)
                    {
                        // 在這個位置創建長方形
                        IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior); // 外牆
                        CalculateGeom(uidoc, doc, exteriorFaces, fs, wall.LevelId); // 計算放置各尺寸的數量
                        //IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior); // 內牆
                        //CalculateGeom(uidoc, doc, interiorFaces, fs, wall.LevelId); // 計算放置各尺寸的數量
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        /// <summary>
        /// 計算放置各尺寸的數量
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="doc"></param>
        /// <param name="referenceFaces"></param>
        /// <param name="fs"></param>
        /// <param name="elemId"></param>
        private void CalculateGeom(UIDocument uidoc, Document doc, IList<Reference> referenceFaces, FamilySymbol fs, ElementId elemId)
        {
            foreach (Reference referenceFace in referenceFaces)
            {
                try
                {
                    Face face = uidoc.Document.GetElement(referenceFace).GetGeometryObjectFromReference(referenceFace) as Face;
                    (List<PointToMatrix>, List<List<Curve>>, int, int, List<Curve>) pointToMatrix = GenerateUniformPoints(face); // 將Face網格化, 每100cm佈一個點
                    List<PointToMatrix> pointToMatrixs = pointToMatrix.Item1;
                    List<List<Curve>> curveLoopList = pointToMatrix.Item2; // 開口的封閉曲線
                    int rows = pointToMatrix.Item3;
                    int cols = pointToMatrix.Item4;
                    List<Curve> drawCurves = pointToMatrix.Item5;

                    // 分割幾何圖形成n個矩形
                    //List<Rectangle> rectangles = new List<Rectangle>();
                    //SaveRectangle(doc, rows, cols, pointToMatrixs, rectangles);

                    // 放置磁磚
                    List<Line> rowLines = new List<Line>(); // 橫向的線
                    List<Line> colLines = new List<Line>(); // 縱向的線
                    foreach (CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
                    {
                        foreach (Curve curve in curveLoop)
                        {
                            Line line = curve as Line;
                            if (line.Direction.Z == 1.0 || line.Direction.Z == -1.0) { rowLines.Add(line); }
                            else { colLines.Add(line); }
                        }
                    }
                    Line bottomLine = rowLines.OrderBy(x => x.Origin.Z).FirstOrDefault(); // 最底的邊
                    Line topLine = rowLines.OrderByDescending(x => x.Origin.Z).FirstOrDefault(); // 最高的邊
                    List<double> heights = rowLines.Select(x => x.Origin.Z - bottomLine.Origin.Z).Distinct().OrderBy(x => x).ToList();

                    for (int i = 0; i < heights.Count; i++)
                    {
                        BoundingBoxUV bboxUV = face.GetBoundingBox();
                        XYZ startXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
                        XYZ endXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
                        Vector vector = new Vector(endXYZ.X - startXYZ.X, endXYZ.Y - startXYZ.Y); // 方向向量
                        vector = GetVectorOffset(vector, RevitAPI.ConvertToInternalUnits(100, "millimeters"));
                        XYZ minXYZ = new XYZ(startXYZ.X + vector.X, startXYZ.Y + vector.Y, startXYZ.Z + heights[i]); // 預留起始的空間
                        XYZ maxXYZ = new XYZ(endXYZ.X - vector.X, endXYZ.Y - vector.Y, startXYZ.Z + heights[i + 1]); // 預留結束的空間

                        double wallHeight = minXYZ.DistanceTo(new XYZ(minXYZ.X, minXYZ.Y, maxXYZ.Z)); // 牆高度
                        double wallLength = minXYZ.DistanceTo(new XYZ(maxXYZ.X, maxXYZ.Y, minXYZ.Z)); // 牆長度

                        List<int> sizes = new List<int>() { 300, 200, 100 };
                        List<LengthOrHeight> heightList = LengthAndHeightTiles(wallHeight, sizes); // 面積長度的磁磚尺寸與數量
                        List<LengthOrHeight> lengthList = LengthAndHeightTiles(wallLength, sizes); // 面積高度的磁磚尺寸與數量

                        XYZ location = minXYZ;
                        foreach (LengthOrHeight heightItem in heightList)
                        {
                            for (int row = 0; row < heightItem.count; row++)
                            {
                                foreach (LengthOrHeight lengthItem in lengthList)
                                {
                                    for (int count = 0; count < lengthItem.count; count++)
                                    {
                                        try
                                        {
                                            FamilyInstance tiles = doc.Create.NewFamilyInstance(referenceFace, location, XYZ.Zero, fs); // 放置磁磚
                                            tiles.LookupParameter("長").Set(lengthItem.heightOrHeight);
                                            tiles.LookupParameter("寬").Set(heightItem.heightOrHeight);
                                            tiles.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(elemId);

                                            vector = new Vector(maxXYZ.X - minXYZ.X, maxXYZ.Y - minXYZ.Y); // 方向向量
                                            vector = GetVectorOffset(vector, lengthItem.heightOrHeight);
                                            location = new XYZ(location.X + vector.X, location.Y + vector.Y, location.Z);
                                        }
                                        catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                                    }
                                }
                                location = new XYZ(minXYZ.X, minXYZ.Y, location.Z + heightItem.heightOrHeight);
                            }
                        }
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        /// <summary>
        /// 將Face網格化, 每100cm佈一個點
        /// </summary>
        /// <param name="face"></param>
        /// <returns></returns>
        private (List<PointToMatrix>, List<List<Curve>>, int, int, List<Curve>) GenerateUniformPoints(Face face)
        {
            List<PointToMatrix> pointToMatrixs = new List<PointToMatrix>();
            List<List<Curve>> curveLoopList = new List<List<Curve>>(); // 儲存所有開口的封閉曲線
            List<Curve> drawCurves = new List<Curve>(); // 測試要繪出的線段

            BoundingBoxUV bboxUV = face.GetBoundingBox();
            XYZ startXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
            XYZ endXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
            double minZ = startXYZ.Z;
            double maxZ = endXYZ.Z;
            if(minZ < maxZ) { startXYZ = new XYZ(startXYZ.X, startXYZ.Y, maxZ); endXYZ = new XYZ(endXYZ.X, endXYZ.Y, minZ); } // 確認座標是由左上到右下
            double offset = RevitAPI.ConvertToInternalUnits(100, "millimeters");
            double height = startXYZ.DistanceTo(new XYZ(startXYZ.X, startXYZ.Y, endXYZ.Z));
            int rows = (int)Math.Ceiling(height / offset);
            double length = startXYZ.DistanceTo(new XYZ(endXYZ.X, endXYZ.Y, startXYZ.Z));
            int cols = (int)Math.Ceiling(length / offset);
            Vector vector = new Vector(endXYZ.X - startXYZ.X, endXYZ.Y - startXYZ.Y); // 方向向量
            vector = GetVectorOffset(vector, offset);
            XYZ newXYZ = startXYZ;

            // 計算最外圍邊界有的開口處
            List<XYZ> points = new List<XYZ>() { startXYZ, new XYZ(endXYZ.X, endXYZ.Y, startXYZ.Z), endXYZ, new XYZ(startXYZ.X, startXYZ.Y, endXYZ.Z) };
            List<Curve> boundingBoxCurves = new List<Curve>();
            for(int i = 0; i < points.Count - 1; i++) { boundingBoxCurves.Add(Line.CreateBound(points[i], points[i+1])); }
            boundingBoxCurves.Add(Line.CreateBound(new XYZ(startXYZ.X, startXYZ.Y, endXYZ.Z), startXYZ));
            CurveLoop boundingBoxCurveLoop = CurveLoop.Create(boundingBoxCurves);

            // 將CurveLoop轉換為Solid
            PlanarFace planarFace = face as PlanarFace;
            Solid solid1 = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { boundingBoxCurveLoop }, planarFace.FaceNormal, 1);
            Solid solid2 = GeometryCreationUtilities.CreateExtrusionGeometry(face.GetEdgesAsCurveLoops(), planarFace.FaceNormal, 1);
            // 計算兩個Solid的差異
            Solid differenceSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Difference);            
            if (differenceSolid != null) // 分析結果
            {
                List<Solid> differenceSolids = new List<Solid>() { differenceSolid };
                List<Face> differenceSolidFaces = GetFaces(differenceSolids, "side");
                Face differenceSolidFace = differenceSolidFaces.Where(x => SameFaceNormal(x, -planarFace.FaceNormal)).FirstOrDefault();
                if (differenceSolidFace != null)
                {
                    foreach (CurveLoop curveLoop in differenceSolidFace.GetEdgesAsCurveLoops())
                    {
                        curveLoopList.Add(curveLoop.ToList()); // 牆面所有開口處
                        foreach (Curve curve in curveLoop) { drawCurves.Add(curve); }
                    }
                    foreach (Curve bbCurve in boundingBoxCurveLoop) { drawCurves.Add(bbCurve); }
                    curveLoopList.Add(boundingBoxCurveLoop.ToList());
                }
            }

            for (int i = 0; i < rows; i++)
            {
                if (i != 0) { newXYZ = new XYZ(startXYZ.X, startXYZ.Y, newXYZ.Z - offset); }
                for (int j = 0; j < cols; j++)
                {
                    PointToMatrix pointToMatrix = new PointToMatrix();
                    pointToMatrix.rows = i;
                    pointToMatrix.cols = j;
                    pointToMatrix.xyz = newXYZ;
                    IntersectionResult result = face.Project(newXYZ);
                    if (result != null) { pointToMatrix.isRectangle = 1; }
                    else { pointToMatrix.isRectangle = 0; }
                    //bool isInside = true;
                    //foreach(List<Line> lines in linesList)
                    //{
                    //    isInside = IsInsideOutline(newXYZ, vector, lines); // 檢查座標點是否在開口內
                    //    if (isInside) { pointToMatrix.isRectangle = 0; break; } // 在開口內的話則退出, 不儲存座標點
                    //}
                    pointToMatrixs.Add(pointToMatrix);
                    newXYZ = new XYZ(newXYZ.X + vector.X, newXYZ.Y + vector.Y, newXYZ.Z);
                }
            }

            //UseEdgeGetDoorOpening(points, face, drawCurves); // 使用邊界計算門開口處

            return (pointToMatrixs, curveLoopList, rows, cols, drawCurves);
        }
        private bool SameFaceNormal(Face face, XYZ faceNormal)
        {
            bool trueOrFalse = false;
            double threshold = 1e-14;
            double faceNormalX = Math.Abs(faceNormal.X) < threshold ? 0 : faceNormal.X;
            double faceNormalY = Math.Abs(faceNormal.Y) < threshold ? 0 : faceNormal.Y;
            double faceNormalZ = Math.Abs(faceNormal.Z) < threshold ? 0 : faceNormal.Z;
            try
            {
                PlanarFace planarFace = face as PlanarFace;
                double planarFaceX = Math.Abs(planarFace.FaceNormal.X) < threshold ? 0 : planarFace.FaceNormal.X;
                double planarFaceY = Math.Abs(planarFace.FaceNormal.Y) < threshold ? 0 : planarFace.FaceNormal.Y;
                double planarFaceZ = Math.Abs(planarFace.FaceNormal.Z) < threshold ? 0 : planarFace.FaceNormal.Z;
                if (planarFaceX == faceNormalX && planarFaceY == faceNormalY && planarFaceZ == faceNormalZ) { trueOrFalse = true; }
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            return trueOrFalse;
        }
        /// <summary>
        /// 使用邊界計算門開口處
        /// </summary>
        /// <param name="points"></param>
        /// <param name="maxFaceCurveLoop"></param>
        /// <param name="drawCurves"></param>
        private void UseEdgeGetDoorOpening(List<XYZ> points, Face face, List<Curve> drawCurves)
        {
            // 取得最大面封閉曲線中的所有座標點
            CurveLoop maxFaceCurveLoop = face.GetEdgesAsCurveLoops().ToList().OrderByDescending(x => x.GetExactLength()).FirstOrDefault();
            List<XYZ> facePoints = new List<XYZ>();
            foreach (Curve curve in maxFaceCurveLoop)
            {
                for (int i = 0; i < curve.Tessellate().Count - 1; i++) { facePoints.Add(curve.Tessellate()[i]); }
            }
            // 最大面的四個端點座標
            List<XYZ> maxFacePoints = new List<XYZ>();
            foreach (XYZ point in points) { maxFacePoints.Add(facePoints.OrderBy(x => x.DistanceTo(point)).FirstOrDefault()); }
            // 排序
            List<XYZ> sortMaxFacePoints = new List<XYZ>();
            foreach (XYZ facePoint in facePoints)
            {
                XYZ sortPoint = maxFacePoints.Where(x => x.Equals(facePoint)).FirstOrDefault();
                if (sortPoint != null) { sortMaxFacePoints.Add(sortPoint); }
            }
            // 取得最大面的線段中, 開口的座標點
            for (int i = 0; i < sortMaxFacePoints.Count - 1; i++)
            {
                List<Curve> maxFaceOpeningLines = new List<Curve>();
                int startPoint = facePoints.IndexOf(facePoints.Where(x => x.Equals(sortMaxFacePoints[i])).FirstOrDefault());
                int endPoint = facePoints.IndexOf(facePoints.Where(x => x.Equals(sortMaxFacePoints[i + 1])).FirstOrDefault());
                if ((endPoint - startPoint) > 1)
                {
                    for (int j = startPoint + 1; j < endPoint - 1; j++)
                    {
                        Curve curve = Line.CreateBound(facePoints[j], facePoints[j + 1]);
                        maxFaceOpeningLines.Add(curve);
                        drawCurves.Add(curve);
                    }
                    maxFaceOpeningLines.Add(Line.CreateBound(facePoints[endPoint - 1], facePoints[startPoint + 1]));
                    drawCurves.Add(Line.CreateBound(facePoints[endPoint - 1], facePoints[startPoint + 1]));
                }
            }
        }
        /// <summary>
        /// 檢查座標點是否在開口內
        /// </summary>
        /// <param name="xyz"></param>
        /// <param name="lines"></param>
        /// <returns></returns>
        private bool IsInsideOutline(XYZ xyz, Vector vector, List<Line> lines)
        {
            bool result = true;
            int insertCount = 0;
            vector = GetVectorOffset(vector, 1000); // 偏移長度
            XYZ rayXYZ = new XYZ(xyz.X + vector.X, xyz.Y + vector.Y, xyz.Z);
            Line rayLine = Line.CreateBound(xyz, rayXYZ);
            List<Line> overlap = lines.Where(x => x.Distance(xyz) < 0.0001).ToList();
            if (overlap.Count == 0) 
            {
                foreach (Line line in lines)
                {
                    //if (line.Distance(xyz) < 0.0001) { break; }
                    SetComparisonResult interResult = line.Intersect(rayLine, out IntersectionResultArray resultArray);
                    IntersectionResult insPoint = resultArray?.get_Item(0);
                    if (insPoint != null) { insertCount++; }
                }
                // 如果次數為偶數就在外面, 奇數就在裡面
                if (insertCount % 2 == 0) { return result = false; }
            }
            else { return result = false; }
            

            return result;
        }
        /// <summary>
        /// 辨識在左右邊的線段
        /// </summary>
        /// <param name="face"></param>
        /// <param name="curveLoopList"></param>
        /// <returns></returns>
        private (List<Line>, List<Line>, List<Line>) LeftRightLines(Face face, List<List<Curve>> curveLoopList)
        {
            List<Line> allLines = new List<Line>();
            List<Line> leftLines = new List<Line>();
            List<Line> rightLines = new List<Line>();
            CurveLoop maxCurveLoops = face.GetEdgesAsCurveLoops().ToList().OrderByDescending(x => x.GetExactLength()).ToList().FirstOrDefault(); // Face最大的範圍邊界
            List<Line> maxCurveList = new List<Line>();
            foreach (Curve maxCurve in maxCurveLoops) { maxCurveList.Add(Line.CreateBound(maxCurve.Tessellate()[0], maxCurve.Tessellate()[maxCurve.Tessellate().Count - 1])); }

            // 針對邊界進行排列
            //linesList.Add(maxCurveList);
            foreach (List<Curve> curves in curveLoopList)
            {
                //// 先找到上下的邊
                //List<Curve> orderByLines = curves.Where(x => x.Direction.Z > 0 || x.Direction.Z < 0).ToList();
                //// 辨識線段在左右邊
                //if (orderByLines.Count > 1)
                //{
                //    XYZ xyz1 = orderByLines[0].Origin;
                //    XYZ xyz2 = orderByLines[orderByLines.Count - 1].Origin;
                //    Vector vector1 = new Vector(xyz2.X - xyz1.X, xyz2.Y - xyz1.Y);
                //    if (vector1.X > 0)
                //    {
                //        leftLines.Add(orderByLines[0]);
                //        rightLines.Add(orderByLines[orderByLines.Count - 1]);
                //    }
                //    else if (vector1.X < 0)
                //    {
                //        leftLines.Add(orderByLines[orderByLines.Count - 1]);
                //        rightLines.Add(orderByLines[0]);
                //    }
                //    else
                //    {
                //        if (vector1.Y > 0)
                //        {
                //            leftLines.Add(orderByLines[0]);
                //            rightLines.Add(orderByLines[orderByLines.Count - 1]);
                //        }
                //        else if (vector1.Y < 0)
                //        {
                //            leftLines.Add(orderByLines[orderByLines.Count - 1]);
                //            rightLines.Add(orderByLines[0]);
                //        }
                //    }
                //    allLines.Add(orderByLines[0]);
                //    allLines.Add(orderByLines[orderByLines.Count - 1]);
                //}
            }
            BoundingBoxUV bboxUV = face.GetBoundingBox();
            leftLines = leftLines.OrderBy(x => VectorDistance(x.Origin, face.Evaluate(bboxUV.Min))).ToList();
            rightLines = rightLines.OrderBy(x => VectorDistance(x.Origin, face.Evaluate(bboxUV.Min))).ToList();
            allLines = allLines.OrderBy(x => VectorDistance(x.Origin, face.Evaluate(bboxUV.Min))).ToList();

            return (leftLines, rightLines, allLines);
        }
        private void SaveRectangle(UIDocument uidoc, Document doc, int rows, int cols, List<PointToMatrix> pointToMatrixs, List<Rectangle> rectangles)
        {
            // 假設 1 表示矩形的一部分，0 表示空白
            int[,] grid = new int[rows, cols];
            int count = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    grid[i, j] = pointToMatrixs[count].isRectangle;
                    count++;
                }
            }
            Rectangle rectangle = MaximalRectangle(grid);
            rectangles.Add(rectangle);
            for(int i = rectangle.TopLeft.Row; i <= rectangle.BottomRight.Row; i++)
            {
                for(int j = rectangle.TopLeft.Col; j <= rectangle.BottomRight.Col; j++)
                {
                    PointToMatrix pointToMatrix = pointToMatrixs.Where(x => x.rows == i && x.cols == j).FirstOrDefault();
                    pointToMatrix.isRectangle = 0;
                }
            }

            XYZ startPoint = pointToMatrixs.Where(x => x.rows == rectangle.TopLeft.Row && x.cols == rectangle.TopLeft.Col).Select(x => x.xyz).FirstOrDefault();
            XYZ endPoint = pointToMatrixs.Where(x => x.rows == rectangle.BottomRight.Row && x.cols == rectangle.BottomRight.Col).Select(x => x.xyz).FirstOrDefault();
            List<Curve> curves = new List<Curve>();
            try
            {
                curves.Add(Line.CreateBound(startPoint, new XYZ(endPoint.X, endPoint.Y, startPoint.Z)));
                curves.Add(Line.CreateBound(new XYZ(endPoint.X, endPoint.Y, startPoint.Z), endPoint));
                curves.Add(Line.CreateBound(endPoint, new XYZ(startPoint.X, startPoint.Y, endPoint.Z)));
                curves.Add(Line.CreateBound(new XYZ(startPoint.X, startPoint.Y, endPoint.Z), startPoint));
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            foreach (Curve curve in curves) { DrawLine(doc, curve); }

            //doc.Regenerate();
            //uidoc.RefreshActiveView();

            int isRectangle = pointToMatrixs.Where(x => x.isRectangle == 1).ToList().Count;
            if(isRectangle > 0) { SaveRectangle(uidoc, doc, rows, cols, pointToMatrixs, rectangles); }
        }
        /// <summary>
        /// 計算最大矩形面積並返回位置
        /// </summary>
        /// <param name="matrix"></param>
        /// <returns></returns>
        public static Rectangle MaximalRectangle(int[,] matrix)
        {
            if (matrix == null || matrix.Length == 0) return null;

            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            int[] heights = new int[cols];

            int maxArea = 0;
            (int Row, int Col) topLeft = (0, 0);
            (int Row, int Col) bottomRight = (0, 0);

            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    heights[col] = (matrix[row, col] == 1) ? heights[col] + 1 : 0;
                }

                var (area, startCol, endCol, height) = LargestRectangleAreaWithPosition(heights);
                if (area > maxArea)
                {
                    maxArea = area;
                    topLeft = (row - height + 1, startCol);
                    bottomRight = (row, endCol);
                }
            }

            return new Rectangle
            {
                MaxArea = maxArea,
                TopLeft = topLeft,
                BottomRight = bottomRight
            };
        }
        /// <summary>
        /// 使用單調棧計算柱狀圖的最大矩形面積和位置
        /// </summary>
        /// <param name="heights"></param>
        /// <returns></returns>
        private static (int Area, int StartCol, int EndCol, int Height) LargestRectangleAreaWithPosition(int[] heights)
        {
            Stack<int> stack = new Stack<int>();
            int maxArea = 0;
            int startCol = 0;
            int endCol = 0;
            int height = 0;

            for (int i = 0; i <= heights.Length; i++)
            {
                int h = (i == heights.Length) ? 0 : heights[i];
                while (stack.Count > 0 && h < heights[stack.Peek()])
                {
                    height = heights[stack.Pop()];
                    int width = (stack.Count == 0) ? i : i - stack.Peek() - 1;
                    int area = height * width;

                    if (area > maxArea)
                    {
                        maxArea = area;
                        startCol = (stack.Count == 0) ? 0 : stack.Peek() + 1;
                        endCol = i - 1;
                    }
                }
                stack.Push(i);
            }

            return (maxArea, startCol, endCol, height);
        }
        /// <summary>
        /// 計算與Face原點的距離後排序
        /// </summary>
        /// <param name="xyz1"></param>
        /// <param name="minXYZ"></param>
        /// <returns></returns>
        private double VectorDistance(XYZ xyz1, XYZ minXYZ)
        {
            Vector vector = new Vector(minXYZ.X - xyz1.X, minXYZ.Y - xyz1.Y);
            return vector.Length;
        }













        /// <summary>
        /// 面積長度與高度的磁磚尺寸與數量
        /// </summary>
        /// <param name="length"></param>
        /// <param name="sizes"></param>
        /// <returns></returns>
        private List<LengthOrHeight> LengthAndHeightTiles(double length, List<int> sizes)
        {
            List<LengthOrHeight> lengthList = new List<LengthOrHeight>();
            foreach (int size in sizes)
            {
                LengthOrHeight lengthOrHeight = new LengthOrHeight();
                lengthOrHeight.heightOrHeight = RevitAPI.ConvertToInternalUnits(size, "millimeters"); // 磁磚的長度
                lengthOrHeight.count = Convert.ToInt32(Math.Floor(length / lengthOrHeight.heightOrHeight)); // 長度的數量
                length = length % lengthOrHeight.heightOrHeight; // 餘數繼續計算
                lengthList.Add(lengthOrHeight);
            }
            return lengthList;
        }
        /// <summary>
        /// 取得向量偏移的距離
        /// </summary>
        /// <param name="vector"></param>
        /// <param name="newLength"></param>
        /// <returns></returns>
        private Vector GetVectorOffset(Vector vector, double newLength)
        {
            double length = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y); // 計算向量長度
            if (length != 0) { vector = new Vector(vector.X / length * newLength, vector.Y / length * newLength); }
            return vector;
        }
        /// <summary>
        /// 3D視圖中畫模型線
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="curve"></param>
        private void DrawLine(Document doc, Curve curve)
        {
            try
            {
                Line line = Line.CreateBound(curve.Tessellate()[0], curve.Tessellate()[curve.Tessellate().Count - 1]);
                XYZ normal = new XYZ(line.Direction.Z - line.Direction.Y, line.Direction.X - line.Direction.Z, line.Direction.Y - line.Direction.X); // 使用與線不平行的任意向量
                Plane plane = Plane.CreateByNormalAndOrigin(normal, curve.Tessellate()[0]);
                SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
                ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
        }
        /// <summary>
        /// 關閉警示視窗
        /// </summary>
        public class CloseWarnings : IFailuresPreprocessor
        {
            FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                String transactionName = failuresAccessor.GetTransactionName();
                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
                if (fmas.Count == 0) { return FailureProcessingResult.Continue; }
                if (transactionName.Equals("EXEMPLE"))
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        if (fma.GetSeverity() == FailureSeverity.Error)
                        {
                            failuresAccessor.DeleteAllWarnings();
                            return FailureProcessingResult.ProceedWithRollBack;
                        }
                        else { failuresAccessor.DeleteWarning(fma); }
                    }
                }
                else
                {
                    foreach (FailureMessageAccessor fma in fmas) { failuresAccessor.DeleteAllWarnings(); }
                }
                return FailureProcessingResult.Continue;
            }
        }
    }
}