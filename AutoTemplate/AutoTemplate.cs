using Autodesk.Revit.ApplicationServices;
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
                    //sideFaces = new List<Face> { maxFace, sideFaces.OrderByDescending(x => x.Area).ToList()[1] };
                    foreach (Face face in sideFaces)
                    {
                        double areas = face.Area;
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
                            List<Curve> upAndDownCurves = new List<Curve>();
                            foreach (Curve curve in curveLoop)
                            {
                                Line line = curve as Line;
                                XYZ direction = ToZeroIfCloseToZero(line.Direction);
                                if (direction.X.Equals(0) && direction.Y.Equals(0))
                                {
                                    if (direction.Z.Equals(1)) { leftCurves.Add(curve); leftRightCurves.Add(curve); }
                                    else if (direction.Z.Equals(-1)) { rightCurves.Add(curve); leftRightCurves.Add(curve); }
                                }
                                if (direction.Z.Equals(0) && direction.Y.Equals(0))
                                {
                                    if (direction.X.Equals(1) || direction.X.Equals(-1)) { upDownCurves.Add(curve); upAndDownCurves.Add(curve); }
                                }
                            }
                            if (upAndDownCurves.Count >= 2)
                            {
                                upAndDownCurves = upAndDownCurves.OrderBy(x => x.Tessellate()[0].Z).ToList();
                                upCurves.Add(upAndDownCurves.LastOrDefault());
                                downCurves.Add(upAndDownCurves.FirstOrDefault());
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
                            // 將邊界與最旁邊的邊連結成矩形, 左線段連結左邊、右線段連結右邊
                            PlanarFace planarFace = face as PlanarFace;
                            UseCurveLoopToCreateSolid(uidoc, doc, planarFace.FaceNormal, leftCurves, leftestCurve, upDownCurves, upCurves, downCurves, leftRightCurves);
                        }
                        else // 沒有開口
                        {
                            BoundingBoxUV bboxUV = face.GetBoundingBox();
                            XYZ startXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
                            XYZ endXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
                            double minZ = startXYZ.Z;
                            double maxZ = endXYZ.Z;
                            if (minZ < maxZ) { startXYZ = new XYZ(startXYZ.X, startXYZ.Y, maxZ); endXYZ = new XYZ(endXYZ.X, endXYZ.Y, minZ); } // 確認座標是由左上到右下
                            // 計算最外圍邊界有的開口處
                            List<XYZ> points = new List<XYZ>() { startXYZ, new XYZ(endXYZ.X, endXYZ.Y, startXYZ.Z), endXYZ, new XYZ(startXYZ.X, startXYZ.Y, endXYZ.Z) };
                            List<Curve> boundingBoxCurves = new List<Curve>();
                            for (int i = 0; i < points.Count - 1; i++) { boundingBoxCurves.Add(Line.CreateBound(points[i], points[i + 1])); }
                            boundingBoxCurves.Add(Line.CreateBound(new XYZ(startXYZ.X, startXYZ.Y, endXYZ.Z), startXYZ));
                            CurveLoop boundingBoxCurveLoop = CurveLoop.Create(boundingBoxCurves);
                            //foreach (Curve boundingBoxCurve in boundingBoxCurves) { DrawLine(doc, boundingBoxCurve); };
                        }

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
        private void UseCurveLoopToCreateSolid(UIDocument uidoc, Document doc, XYZ faceNormal, List<Curve> curves, Curve leftOrRightestCurve, List<Curve> upDownCurves, List<Curve> upCurves, List<Curve> downCurves, List<Curve> leftRightCurves)
        {
            SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
            foreach (Curve curve in curves)
            {
                try
                {
                    // 建立要干涉的矩形
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
                    doc.Regenerate(); uidoc.RefreshActiveView();

                    List<Face> solidFaces = new List<Face>();
                    foreach (Face face in solid.Faces)
                    {
                        PlanarFace pf = face as PlanarFace;
                        faceNormal = ToZeroIfCloseToZero(faceNormal);
                        XYZ pfFaceNormal = ToZeroIfCloseToZero(pf.FaceNormal);
                        if (faceNormal.X.Equals(-pfFaceNormal.X) && faceNormal.Y.Equals(-pfFaceNormal.Y) && faceNormal.Z.Equals(-pfFaceNormal.Z)) { solidFaces.Add(face); }
                    }
                    if (solidFaces.Count > 0)
                    {
                        // 找到此面所有干涉的線段
                        Face solidFace = solidFaces.OrderBy(x => x.Area).FirstOrDefault();
                        List<Curve> containCurves = new List<Curve>();
                        List<Curve> containUpCurves = new List<Curve>();
                        List<Curve> containDownCurves = new List<Curve>();

                        // solidFace是否有涵蓋垂直的線段
                        List<Curve> verticalCurves = new List<Curve>();
                        foreach(Curve leftRightCurve in leftRightCurves.Where(x => x != curve && x != leftOrRightestCurve).ToList())
                        {
                            SolidCurveIntersection result = solid.IntersectWithCurve(leftRightCurve, new SolidCurveIntersectionOptions());
                            if(result.SegmentCount > 0) { verticalCurves.Add(leftRightCurve); }
                        }
                        foreach (Curve upCurve in upCurves)
                        {
                            bool trueOrFalse = false;
                            foreach (XYZ upDownXYZ in upCurve.Tessellate())
                            {
                                if (solidFace.Project(upDownXYZ) != null) { trueOrFalse = true; }
                                else { trueOrFalse = false; break; }
                            }
                            if (trueOrFalse)
                            {
                                containCurves.Add(upCurve); // solidFace包含整個線段
                                containUpCurves.Add(upCurve);
                            }
                        }
                        foreach (Curve downCurve in downCurves)
                        {
                            bool trueOrFalse = false;
                            foreach (XYZ upDownXYZ in downCurve.Tessellate())
                            {
                                if (solidFace.Project(upDownXYZ) != null) { trueOrFalse = true; }
                                else { trueOrFalse = false; break; }
                            }
                            if (trueOrFalse)
                            {
                                containCurves.Add(downCurve); // solidFace包含整個線段
                                containDownCurves.Add(downCurve);
                            }
                        }

                        // 線段的高點與低點
                        XYZ topXYZ = curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault();
                        XYZ bottomXYZ = curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault();
                        if (containCurves.Count > 0)
                        {
                            if (verticalCurves.Count > 0)
                            {
                                Curve verticalCurve = verticalCurves.OrderBy(x => ClosestDistance(x, curve)).FirstOrDefault(); // 找到離開口最近的垂直線
                                // 包含上方的線
                                if (containUpCurves.Count > 0)
                                {
                                    Curve parallelCurve = containUpCurves.OrderBy(x => ClosestDistance(x, curve)).FirstOrDefault(); // 找到離開口最近的平行線
                                    double vDistance = verticalCurve.Tessellate().Select(x => curve.Distance(x)).OrderBy(x => x).FirstOrDefault();
                                    double pDistance = parallelCurve.Tessellate().Select(x => curve.Distance(x)).OrderBy(x => x).FirstOrDefault();
                                    if (pDistance <= vDistance)
                                    {
                                        if (topXYZ.Z > parallelCurve.Tessellate()[0].Z)
                                        {
                                            // 搜尋是否到最左最右邊界前, 有其他的垂直線段
                                            List<Curve> secondVerticalCurves = verticalCurves.Where(x => !SamePoint(x, parallelCurve)).OrderBy(x => curve.Distance(x.Tessellate()[0])).ToList();
                                            XYZ p1 = new XYZ();
                                            XYZ p2 = new XYZ();
                                            XYZ p3 = new XYZ();
                                            XYZ p4 = new XYZ();
                                            if (secondVerticalCurves.Count > 0)
                                            {
                                                Curve secondVerticalCurve = secondVerticalCurves[0];
                                                double z = secondVerticalCurve.Tessellate().Select(x => x.Z).OrderByDescending(x => x).FirstOrDefault();
                                                if (z > parallelCurve.Tessellate()[0].Z)
                                                {
                                                    p1 = curve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    p2 = secondVerticalCurve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    if(secondVerticalCurve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault().Z > topXYZ.Z)
                                                    {
                                                        p3 = secondVerticalCurve.Project(curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault()).XYZPoint;
                                                        p4 = curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault();
                                                    }
                                                    else
                                                    {
                                                        p3 = secondVerticalCurve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault();
                                                        p4 = curve.Project(secondVerticalCurve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault()).XYZPoint;
                                                    }
                                                }
                                                else
                                                {
                                                    p1 = curve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    p2 = leftOrRightestCurve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    p3 = leftOrRightestCurve.Project(curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault()).XYZPoint;
                                                    p4 = curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault();
                                                }
                                            }
                                            else
                                            {
                                                p1 = curve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                p2 = leftOrRightestCurve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                p3 = leftOrRightestCurve.Project(curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault()).XYZPoint;
                                                p4 = curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault();
                                            }
                                            try
                                            {
                                                List<Curve> rectangleCurves = new List<Curve>() { Line.CreateBound(p1, p2), Line.CreateBound(p2, p3), Line.CreateBound(p3, p4), Line.CreateBound(p4, p1) };
                                                foreach (Curve rectangleCurve in rectangleCurves) { DrawLine(doc, rectangleCurve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                                            }
                                            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                                        }
                                    }
                                }
                                // 包含下方的線
                                if (containDownCurves.Count > 0)
                                {
                                    Curve parallelCurve = containDownCurves.OrderBy(x => ClosestDistance(x, curve)).FirstOrDefault(); // 找到離開口最近的平行線
                                    double vDistance = verticalCurve.Tessellate().Select(x => curve.Distance(x)).OrderBy(x => x).FirstOrDefault();
                                    double pDistance = parallelCurve.Tessellate().Select(x => curve.Distance(x)).OrderBy(x => x).FirstOrDefault();
                                    if (pDistance <= vDistance)
                                    {
                                        if (bottomXYZ.Z < parallelCurve.Tessellate()[0].Z)
                                        {
                                            // 搜尋是否到最左最右邊界前, 有其他的垂直線段
                                            List<Curve> secondVerticalCurves = verticalCurves.Where(x => !SamePoint(x, parallelCurve)).OrderBy(x => ClosestDistance(x, curve)).ToList();
                                            XYZ p1 = new XYZ();
                                            XYZ p2 = new XYZ();
                                            XYZ p3 = new XYZ();
                                            XYZ p4 = new XYZ();
                                            if (secondVerticalCurves.Count > 0)
                                            {
                                                Curve secondVerticalCurve = secondVerticalCurves[0];
                                                List<double> zs = secondVerticalCurve.Tessellate().Select(x => x.Z).OrderByDescending(x => x).ToList();
                                                double z = secondVerticalCurve.Tessellate().Select(x => x.Z).OrderBy(x => x).FirstOrDefault();
                                                if (z < parallelCurve.Tessellate()[0].Z)
                                                {
                                                    p1 = curve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    p2 = secondVerticalCurve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    if (secondVerticalCurve.Tessellate().OrderBy(x => x.Z).FirstOrDefault().Z < bottomXYZ.Z)
                                                    {
                                                        p3 = secondVerticalCurve.Project(curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault()).XYZPoint;
                                                        p4 = curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault();
                                                    }
                                                    else
                                                    {
                                                        p3 = secondVerticalCurve.Tessellate().OrderBy(x => x.Z).FirstOrDefault();
                                                        p4 = curve.Project(secondVerticalCurve.Tessellate().OrderBy(x => x.Z).FirstOrDefault()).XYZPoint;
                                                    }
                                                }
                                                else
                                                {
                                                    p1 = curve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    p2 = leftOrRightestCurve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                    p3 = leftOrRightestCurve.Project(curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault()).XYZPoint;
                                                    p4 = curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault();
                                                }
                                            }
                                            else
                                            {
                                                p1 = curve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                p2 = leftOrRightestCurve.Project(parallelCurve.Tessellate()[0]).XYZPoint;
                                                p3 = leftOrRightestCurve.Project(curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault()).XYZPoint;
                                                p4 = curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault();
                                            }
                                            try
                                            {
                                                List<Curve> rectangleCurves = new List<Curve>() { Line.CreateBound(p1, p2), Line.CreateBound(p2, p3), Line.CreateBound(p3, p4), Line.CreateBound(p4, p1) };
                                                foreach (Curve rectangleCurve in rectangleCurves) { DrawLine(doc, rectangleCurve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                                            }
                                            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                                        }
                                    }
                                }
                                if(verticalCurves.Count > 0)
                                {
                                    List<Curve> twoCurves = new List<Curve>() { curve, verticalCurve };
                                    // 比對垂直跟平行線誰的距離較近
                                    List<XYZ> sortXYZs = new List<XYZ>() { curve.Tessellate()[0], curve.Tessellate()[curve.Tessellate().Count - 1], verticalCurve.Tessellate()[0], verticalCurve.Tessellate()[verticalCurve.Tessellate().Count - 1] };
                                    sortXYZs = sortXYZs.OrderBy(x => x.Z).ToList();
                                    // 找到第2、3個座標點的Curve
                                    XYZ p1 = ToZeroIfCloseToZero(sortXYZs[1]);
                                    XYZ p2 = ToZeroIfCloseToZero(sortXYZs[2]);
                                    Curve secondXYZCurve = null;
                                    Curve thirdXYZCurve = null;
                                    foreach(Curve twoCurve in twoCurves)
                                    {
                                        foreach(XYZ twoXYZ in twoCurve.Tessellate())
                                        {
                                            XYZ xyz = ToZeroIfCloseToZero(twoXYZ);
                                            if (xyz.X.Equals(p1.X) && xyz.Y.Equals(p1.Y) && xyz.Z.Equals(p1.Z)) { secondXYZCurve = twoCurve;  }
                                            if (xyz.X.Equals(p2.X) && xyz.Y.Equals(p2.Y) && xyz.Z.Equals(p2.Z)) { thirdXYZCurve = twoCurve;  }
                                        }
                                    }
                                    Curve otherCurve1 = twoCurves.Where(x => x != secondXYZCurve).FirstOrDefault();
                                    XYZ p3 = ToZeroIfCloseToZero(otherCurve1.Project(p1).XYZPoint);
                                    Curve otherCurve2 = twoCurves.Where(x => x != thirdXYZCurve).FirstOrDefault();
                                    XYZ p4 = ToZeroIfCloseToZero(otherCurve2.Project(p2).XYZPoint);
                                    List<Curve> rectangleCurves = new List<Curve>();
                                    if (p1.DistanceTo(p2) <= p1.DistanceTo(p4))
                                    {
                                        rectangleCurves.Add(Line.CreateBound(p1, p2));
                                        rectangleCurves.Add(Line.CreateBound(p2, p4));
                                        rectangleCurves.Add(Line.CreateBound(p4, p3));
                                        rectangleCurves.Add(Line.CreateBound(p3, p1));
                                    }
                                    else
                                    {
                                        rectangleCurves.Add(Line.CreateBound(p1, p4));
                                        rectangleCurves.Add(Line.CreateBound(p4, p2));
                                        rectangleCurves.Add(Line.CreateBound(p2, p3));
                                        rectangleCurves.Add(Line.CreateBound(p3, p1));
                                    }
                                    foreach (Curve rectangleCurve in rectangleCurves) { DrawLine(doc, rectangleCurve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                                }
                            }
                        }
                        else
                        {
                            if (verticalCurves.Count > 0)
                            {
                                Curve verticalCurve = verticalCurves.OrderBy(x => curve.Distance(x.Tessellate()[0])).FirstOrDefault(); // 找到離開口最近的垂直線
                                List<Curve> secondVerticalCurves = verticalCurves.OrderBy(x => curve.Distance(x.Tessellate()[0])).ToList();
                                XYZ p1 = curve.Tessellate().OrderByDescending(x => x.Z).FirstOrDefault();
                                XYZ p2 = curve.Tessellate().OrderBy(x => x.Z).FirstOrDefault();
                                XYZ p3 = verticalCurve.Project(p2).XYZPoint;
                                XYZ p4 = verticalCurve.Project(p1).XYZPoint;
                                try
                                {
                                    List<Curve> rectangleCurves = new List<Curve>() { Line.CreateBound(p1, p2), Line.CreateBound(p2, p3), Line.CreateBound(p3, p4), Line.CreateBound(p4, p1) };
                                    foreach (Curve rectangleCurve in rectangleCurves) { DrawLine(doc, rectangleCurve); doc.Regenerate(); uidoc.RefreshActiveView(); }
                                }
                                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                            }
                            else
                            {

                                foreach (Curve solidCurve in curveLoop)
                                {
                                    DrawLine(doc, solidCurve); doc.Regenerate(); uidoc.RefreshActiveView();
                                }
                            }
                        }
                    }
                    doc.Delete(ds.Id);
                    doc.Regenerate();
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        /// <summary>
        /// 找到兩條Curve最近的距離
        /// </summary>
        /// <param name="vCurve"></param>
        /// <param name="curve"></param>
        /// <returns></returns>
        private double ClosestDistance(Curve vCurve, Curve curve)
        {
            double minDistance = curve.Distance(vCurve.Tessellate()[0]);
            foreach (XYZ vCurveXYZ in vCurve.Tessellate())
            {
                double distance = curve.Distance(vCurveXYZ);
            }
            return minDistance;
        }
        /// <summary>
        /// 檢查垂直線的座標點是否包含在最高、最低平行線
        /// </summary>
        /// <param name="curve"></param>
        /// <param name="parallelCurve"></param>
        /// <returns></returns>
        private bool SamePoint(Curve curve, Curve parallelCurve)
        {
            bool trueOrFalse = false;
            foreach(XYZ parallelCurveXYZ in parallelCurve.Tessellate())
            {
                XYZ point = ToZeroIfCloseToZero(parallelCurveXYZ);
                foreach (XYZ xyz in curve.Tessellate())
                {
                    XYZ curveXYZ = ToZeroIfCloseToZero(xyz);
                    if (curveXYZ.X.Equals(point.X) && curveXYZ.Y.Equals(point.Y) && curveXYZ.Z.Equals(point.Z)) { trueOrFalse = true; break; }
                }
            }
            return trueOrFalse;
        }
        /// <summary>
        /// 轉換很接近0的數值為0
        /// </summary>
        /// <param name="xyz"></param>
        /// <returns></returns>
        private XYZ ToZeroIfCloseToZero(XYZ xyz)
        {
            double threshold = 1e-14;
            double x = Math.Abs(xyz.X) < threshold ? 0 : xyz.X;
            double y = Math.Abs(xyz.Y) < threshold ? 0 : xyz.Y;
            double z = Math.Abs(xyz.Z) < threshold ? 0 : xyz.Z;
            return new XYZ(x, y, z);
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