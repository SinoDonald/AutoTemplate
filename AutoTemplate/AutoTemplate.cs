using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;

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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            DateTime timeStart = DateTime.Now; // 計時開始 取得目前時間

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

                trans.Commit();
                DateTime timeEnd = DateTime.Now; // 計時結束 取得目前時間
                TimeSpan totalTime = timeEnd - timeStart;
                TaskDialog.Show("Revit", "完成，耗時：" + totalTime.Minutes + " 分 " + totalTime.Seconds + " 秒。\n\n");
            }

            return Result.Succeeded;
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

                    List<Line> rowLines = new List<Line>(); // 橫向的線
                    List<Line> colLines = new List<Line>(); // 縱向的線
                    foreach(CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
                    {
                        foreach(Curve curve in curveLoop)
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