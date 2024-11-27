using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Windows;

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
                // 查詢所有柱牆的Element
                List<Element> elems = ColumnsAndWallsElem(doc);
                // ElementId第一個有接觸到的所有元件
                List<IntersectionElem> list = IntersectGroup(doc, elems);
                List<Solid> solids = new List<Solid>(); // 儲存所有柱牆的Solid                
                foreach (Element elem in elems) { foreach(Solid solid in GetSolids(doc, elem)) {solids.Add(solid); } }                
                Solid hostSolid = UnionSolids(solids, solids[0]); // 將所有Solid聯集

                List<Curve> curveList = new List<Curve>();
                if (null != hostSolid)
                {
                    foreach (Face face in hostSolid.Faces)
                    {
                        PlanarFace pf = face as PlanarFace;
                        //if (null != pf && pf.Normal.IsAlmostEqualTo(XYZ.BasisZ))
                        if (null != pf && pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                        {
                            EdgeArrayArray loops = pf.EdgeLoops;
                            foreach (EdgeArray loop in loops)
                            {
                                CurveArray curves = app.Create.NewCurveArray();
                                foreach (Edge edge in loop)
                                {
                                    IList<XYZ> edgePoints = edge.Tessellate();
                                    Arc arc = null;
                                    if (edgePoints.Count > 2)
                                    {
                                        arc = Arc.Create(edgePoints[0], edgePoints[edgePoints.Count - 1], edgePoints[1]);
                                        if (arc != null)
                                        {
                                            curves.Append(arc);
                                        }
                                    }
                                    else
                                    {
                                        Line line = Line.CreateBound(edgePoints[0], edgePoints[1]);
                                        curves.Append(line);
                                        Curve curve = Line.CreateBound(edgePoints[0], edgePoints[1]);
                                        curveList.Add(curve);
                                    }
                                }
                            }
                        }
                    }
                }

                foreach(Curve curve in curveList) { DrawLine(doc, curve); }
                

                //foreach (Wall wall in walls)
                //{
                //    List<Solid> wallSolids = GetSolids(doc, wall);
                //    List<Face> wallFaces = GetFaces(wallSolids, "side");
                //    Level level = doc.GetElement(wall.LevelId) as Level;
                //    try
                //    {
                //        if (level != null)
                //        {
                //            // 在這個位置創建長方形
                //            IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior); // 外牆
                //            CalculateGeom(uidoc, doc, exteriorFaces, fs, wall.LevelId); // 計算放置各尺寸的數量
                //            //IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior); // 內牆
                //            //CalculateGeom(uidoc, doc, interiorFaces, fs, wall.LevelId); // 計算放置各尺寸的數量
                //        }
                //    }
                //    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                //}

                trans.Commit();
                DateTime timeEnd = DateTime.Now; // 計時結束 取得目前時間
                TimeSpan totalTime = timeEnd - timeStart;
                //TaskDialog.Show("Revit", "完成，耗時：" + totalTime.Minutes + " 分 " + totalTime.Seconds + " 秒。\n\n");
            }

            return Result.Succeeded;
        }
        // 查詢所有柱牆的Element
        private List<Element> ColumnsAndWallsElem(Document doc)
        {
            List<Element> elems = new List<Element>();

            ElementCategoryFilter columnsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Columns);
            ElementCategoryFilter wallsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            LogicalOrFilter filter = new LogicalOrFilter(columnsFilter, wallsFilter);
            elems = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements().ToList();

            return elems;
        }
        //  ElementId第一個有接觸到的所有元件
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

            string info = string.Empty;
            foreach (IntersectionElem abc in list)
            {
                info += "\nHostId : " + abc.hostElem.Id + "\nIntersect : ";
                foreach (Element e in abc.elemList)
                {
                    info += "\n" + e.Id;
                }
                info += "\n";
            }

            return list;
        }
        // 將所有Solid聯集
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
        /// 取得面
        /// </summary>
        /// <param name="solidList"></param>
        /// <param name="where"></param>
        /// <returns></returns>
        private List<Face> GetFaces(List<Solid> solidList, string where)
        {
            List<Face> faces = new List<Face>();
            foreach (Solid solid in solidList)
            {
                foreach (Face face in solid.Faces)
                {
                    if (where.Equals("top")) // 頂面
                    {
                        double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                        if (faceTZ > 0.0) { faces.Add(face); }
                    }
                    else if (where.Equals("bottom")) // 底面
                    {
                        //if (face.ComputeNormal(new UV(0.5, 0.5)).IsAlmostEqualTo(XYZ.BasisZ.Negate())) { topOrBottomFaces.Add(face); }
                        double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                        if (faceTZ < 0.0) { faces.Add(face); }
                    }
                    else if (where.Equals("side")) // 側面
                    {
                        double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                        if (faceTZ != 1.0 && faceTZ != -1.0) { faces.Add(face); }
                    }
                    else
                    {
                        faces.Add(face);
                    }
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
                    (List<XYZ>, List<PointToMatrix>, int, int) pointToMatrix = GenerateUniformPoints(face); // 將Face網格化, 每100cm佈一個點
                    List<XYZ> xyzs = pointToMatrix.Item1;
                    List<PointToMatrix> pointToMatrixs = pointToMatrix.Item2;
                    int rows = pointToMatrix.Item3;
                    int cols = pointToMatrix.Item4;

                    List<Rectangle> rectangles = new List<Rectangle>();
                    SaveRectangle(doc, rows, cols, pointToMatrixs, rectangles);










                    //List<Line> rowLines = new List<Line>(); // 橫向的線
                    //List<Line> colLines = new List<Line>(); // 縱向的線
                    //foreach(CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
                    //{
                    //    foreach(Curve curve in curveLoop)
                    //    {
                    //        Line line = curve as Line;
                    //        if (line.Direction.Z == 1.0 || line.Direction.Z == -1.0) { rowLines.Add(line); }
                    //        else { colLines.Add(line); }
                    //    }
                    //}
                    //Line bottomLine = rowLines.OrderBy(x => x.Origin.Z).FirstOrDefault(); // 最底的邊
                    //Line topLine = rowLines.OrderByDescending(x => x.Origin.Z).FirstOrDefault(); // 最高的邊
                    //List<double> heights = rowLines.Select(x => x.Origin.Z - bottomLine.Origin.Z).Distinct().OrderBy(x => x).ToList();

                    //for (int i = 0; i < heights.Count; i++)
                    //{
                    //    BoundingBoxUV bboxUV = face.GetBoundingBox();
                    //    XYZ startXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
                    //    XYZ endXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
                    //    Vector vector = new Vector(endXYZ.X - startXYZ.X, endXYZ.Y - startXYZ.Y); // 方向向量
                    //    vector = GetVectorOffset(vector, RevitAPI.ConvertToInternalUnits(100, "millimeters"));
                    //    XYZ minXYZ = new XYZ(startXYZ.X + vector.X, startXYZ.Y + vector.Y, startXYZ.Z + heights[i]); // 預留起始的空間
                    //    XYZ maxXYZ = new XYZ(endXYZ.X - vector.X, endXYZ.Y - vector.Y, startXYZ.Z + heights[i + 1]); // 預留結束的空間

                    //    double wallHeight = minXYZ.DistanceTo(new XYZ(minXYZ.X, minXYZ.Y, maxXYZ.Z)); // 牆高度
                    //    double wallLength = minXYZ.DistanceTo(new XYZ(maxXYZ.X, maxXYZ.Y, minXYZ.Z)); // 牆長度

                    //    List<int> sizes = new List<int>() { 300, 200, 100 };
                    //    List<LengthOrHeight> heightList = LengthAndHeightTiles(wallHeight, sizes); // 面積長度的磁磚尺寸與數量
                    //    List<LengthOrHeight> lengthList = LengthAndHeightTiles(wallLength, sizes); // 面積高度的磁磚尺寸與數量

                    //    XYZ location = minXYZ;
                    //    foreach (LengthOrHeight heightItem in heightList)
                    //    {
                    //        for (int row = 0; row < heightItem.count; row++)
                    //        {
                    //            foreach (LengthOrHeight lengthItem in lengthList)
                    //            {
                    //                for (int count = 0; count < lengthItem.count; count++)
                    //                {
                    //                    try
                    //                    {
                    //                        FamilyInstance tiles = doc.Create.NewFamilyInstance(referenceFace, location, XYZ.Zero, fs); // 放置磁磚
                    //                        tiles.LookupParameter("長").Set(lengthItem.heightOrHeight);
                    //                        tiles.LookupParameter("寬").Set(heightItem.heightOrHeight);
                    //                        tiles.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(elemId);

                    //                        vector = new Vector(maxXYZ.X - minXYZ.X, maxXYZ.Y - minXYZ.Y); // 方向向量
                    //                        vector = GetVectorOffset(vector, lengthItem.heightOrHeight);
                    //                        location = new XYZ(location.X + vector.X, location.Y + vector.Y, location.Z);
                    //                    }
                    //                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    //                }
                    //            }
                    //            location = new XYZ(minXYZ.X, minXYZ.Y, location.Z + heightItem.heightOrHeight);
                    //        }
                    //    }
                    //}
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        /// <summary>
        /// 將Face網格化, 每100cm佈一個點
        /// </summary>
        /// <param name="face"></param>
        /// <returns></returns>
        public (List<XYZ>, List<PointToMatrix>, int, int) GenerateUniformPoints(Face face)
        {
            List<PointToMatrix> pointToMatrixs = new List<PointToMatrix>();
            List<XYZ> xyzs = new List<XYZ>();

            List<CurveLoop> curveLoops = face.GetEdgesAsCurveLoops().ToList().OrderByDescending(x => x.GetExactLength()).ToList();
            curveLoops.RemoveAt(0); // 移除最大的邊界
            // 儲存所有開口的封閉曲線
            List<List<Line>> linesList = new List<List<Line>>();
            foreach(CurveLoop curveLoop in curveLoops)
            {
                List<Line> lines = new List<Line>();
                foreach (Curve curve in curveLoop)
                {
                    Line line = Line.CreateBound(curve.Tessellate()[0], curve.Tessellate()[curve.Tessellate().Count() - 1]);
                    lines.Add(line);
                }
                linesList.Add(lines);
            }

            BoundingBoxUV bboxUV = face.GetBoundingBox();
            XYZ startXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
            XYZ endXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
            double offset = RevitAPI.ConvertToInternalUnits(100, "millimeters");
            double height = startXYZ.DistanceTo(new XYZ(startXYZ.X, startXYZ.Y, endXYZ.Z));
            int rows = (int)Math.Ceiling(height / offset);
            double length = startXYZ.DistanceTo(new XYZ(endXYZ.X, endXYZ.Y, startXYZ.Z));
            int cols = (int)Math.Ceiling(length / offset);
            Vector vector = new Vector(endXYZ.X - startXYZ.X, endXYZ.Y - startXYZ.Y); // 方向向量
            vector = GetVectorOffset(vector, offset);
            XYZ newXYZ = startXYZ;
            for (int i = 0; i < rows; i++)
            {
                if (i != 0) { newXYZ = new XYZ(startXYZ.X, startXYZ.Y, newXYZ.Z + offset); }
                for (int j = 0; j < cols; j++)
                {
                    PointToMatrix pointToMatrix = new PointToMatrix();
                    pointToMatrix.rows = i;
                    pointToMatrix.cols = j;
                    pointToMatrix.xyz = newXYZ;
                    pointToMatrix.isRectangle = 1;
                    bool isInside = true;
                    foreach(List<Line> lines in linesList)
                    {
                        isInside = IsInsideOutline(newXYZ, lines); // 檢查座標點是否在開口內
                        if (isInside) { pointToMatrix.isRectangle = 0; break; } // 在開口內的話則退出, 不儲存座標點
                    }
                    //if (!isInside) 
                    //{ 
                        xyzs.Add(newXYZ); 
                    //}
                    pointToMatrixs.Add(pointToMatrix);
                    newXYZ = new XYZ(newXYZ.X + vector.X, newXYZ.Y + vector.Y, newXYZ.Z);
                }
            }

            return (xyzs, pointToMatrixs, rows, cols);
        }
        /// <summary>
        /// 檢查座標點是否在開口內
        /// </summary>
        /// <param name="xyz"></param>
        /// <param name="lines"></param>
        /// <returns></returns>
        public bool IsInsideOutline(XYZ xyz, List<Line> lines)
        {
            bool result = true;
            int insertCount = 0;
            Line rayLine = Line.CreateBound(xyz, xyz.Add(XYZ.BasisX * 1000));
            foreach (var arealine in lines)
            {
                SetComparisonResult interResult = arealine.Intersect(rayLine, out IntersectionResultArray resultArray);
                IntersectionResult insPoint = resultArray?.get_Item(0);
                if (insPoint != null)
                {
                    insertCount++;
                }
            }
            // 如果次數為偶數就在外面, 奇數就在裡面
            if(insertCount % 2 == 0)
            {
                return result = false;
            }

            return result;
        }
        private void SaveRectangle(Document doc, int rows, int cols, List<PointToMatrix> pointToMatrixs, List<Rectangle> rectangles)
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
            for(int i = rectangle.TopLeft.Row; i < rectangle.BottomRight.Row; i++)
            {
                for(int j = rectangle.TopLeft.Col;  j < rectangle.BottomRight.Col; j++)
                {
                    PointToMatrix pointToMatrix = pointToMatrixs.Where(x => x.rows == i && x.cols == j).FirstOrDefault();
                    pointToMatrix.isRectangle = 0;
                }
            }

            XYZ startPoint = pointToMatrixs.Where(x => x.rows == rectangle.TopLeft.Row && x.cols == rectangle.TopLeft.Col).Select(x => x.xyz).FirstOrDefault();
            XYZ endPoint = pointToMatrixs.Where(x => x.rows == rectangle.BottomRight.Row && x.cols == rectangle.BottomRight.Col).Select(x => x.xyz).FirstOrDefault();
            List<Curve> curves = new List<Curve>();
            curves.Add(Line.CreateBound(startPoint, new XYZ(endPoint.X, endPoint.Y, startPoint.Z)));
            curves.Add(Line.CreateBound(new XYZ(endPoint.X, endPoint.Y, startPoint.Z), endPoint));
            curves.Add(Line.CreateBound(endPoint, new XYZ(startPoint.X, startPoint.Y, endPoint.Z)));
            curves.Add(Line.CreateBound(new XYZ(startPoint.X, startPoint.Y, endPoint.Z), startPoint));
            foreach (Curve curve in curves) { DrawLine(doc, curve); }

            int isRectangle = pointToMatrixs.Where(x => x.isRectangle == 1).ToList().Count;
            if(isRectangle > 0) { SaveRectangle(doc, rows, cols, pointToMatrixs, rectangles); }
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