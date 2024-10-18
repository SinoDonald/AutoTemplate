using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
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
        // 牆面與挖空的面
        public class FaceAndHoles
        {
            public Face face { get; set; }
            public List<Surface> holes = new List<Surface>();
        }
        // 牆與面
        public class WallAndFace
        {
            public Element wall { get; set; }
            public List<FaceAndHoles> faces = new List<FaceAndHoles>();
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // 取得所有的牆的面
            List<WallAndFace> wallsAndFaces = new List<WallAndFace>();
            List<Element> walls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToList();
            foreach (Element wall in walls)
            {
                WallAndFace wallAndFaces = new WallAndFace();
                List<Solid> wallSolids = GetSolids(doc, wall);
                List<Face> wallFaces = GetFaces(wallSolids, "side");
                foreach (Face wallFace in wallFaces)
                {
                    FaceAndHoles faceAndHoles = new FaceAndHoles();
                    faceAndHoles.face = wallFace;
                    List<CurveLoop> curveLoops = wallFace.GetEdgesAsCurveLoops().OrderByDescending(x => x.GetExactLength()).ToList();
                    for (int i = 0; i < curveLoops.Count(); i++)
                    {
                        if (i != 0) { faceAndHoles.holes.Add(curveLoops[i].GetPlane()); }
                    }
                    wallAndFaces.faces.Add(faceAndHoles);
                }
                //Face wallFace = wallFaces.Where(x => x.Area.Equals(wallFaces.Max(y => y.Area))).FirstOrDefault(); // 取得面積最大的面
                wallAndFaces.wall = wall;
                //wallAndFaces.faces.Add(wallFace);
                wallsAndFaces.Add(wallAndFaces);
            }

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

                // 放置磁磚
                foreach (WallAndFace wallAndFace in wallsAndFaces)
                {
                    Wall wall = wallAndFace.wall as Wall;
                    Level level = doc.GetElement(wall.LevelId) as Level;
                    try
                    {
                        if (level != null)
                        {
                            foreach (FaceAndHoles faceAndHoles in wallAndFace.faces)
                            {
                                Face face = faceAndHoles.face;

                                // 取得牆的中心點
                                BoundingBoxUV bboxUV = face.GetBoundingBox();
                                UV center = (bboxUV.Max + bboxUV.Min) / 2.0;
                                XYZ minXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
                                XYZ maxXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
                                //XYZ location = face.Evaluate(center);
                                bool inInside = face.IsInside(center);

                                double length = RevitAPI.ConvertToInternalUnits(200, "millimeters");  // 長方形的長度
                                double width = RevitAPI.ConvertToInternalUnits(200, "millimeters"); // 長方形的寬度
                                //Vector vector = new Vector(doorFaceOrigin.X - otherDoorFaceOrigin.X, doorFaceOrigin.Y - otherDoorFaceOrigin.Y); // 方向向量
                                //vector = GetVectorOffset(vector, frontOffset);
                                double availableArea = face.Area;
                                int rows = (int)(availableArea / (length * width));
                                int cols = (int)(availableArea / (length * width));

                                // 開始在面上佈滿長方形，跳過挖空區域
                                for (int i = 0; i < rows; i++)
                                {
                                    for (int j = 0; j < cols; j++)
                                    {
                                        XYZ location = new XYZ(i * length, j * width, 0);
                                        //// 在這個位置創建長方形
                                        //IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                                        //FamilyInstance tiles = doc.Create.NewFamilyInstance(exteriorFaces.FirstOrDefault(), location, XYZ.Zero, familySymbolList.FirstOrDefault());
                                        //tiles.LookupParameter("長").Set(length);
                                        //tiles.LookupParameter("寬").Set(width);
                                        //tiles.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(wall.LevelId);
                                    }
                                }
                                //// 取得牆的中心點
                                //BoundingBoxUV bboxUV = face.GetBoundingBox();
                                //UV center = (bboxUV.Max + bboxUV.Min) / 2.0;
                                //XYZ location = face.Evaluate(center);
                                //bool inInside = face.IsInside(center);
                                //if (inInside)
                                //{
                                //    double length = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                                //    double width = RevitAPI.ConvertToInternalUnits(200, "millimeters");

                                //    // 取得外牆
                                //    IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                                //    foreach (Reference exteriorFace in exteriorFaces)
                                //    {
                                //        Element elem = doc.GetElement(exteriorFace);
                                //        // 使用參考面和放置點創建磚塊族群
                                //        FamilyInstance tiles = doc.Create.NewFamilyInstance(exteriorFace, location, XYZ.Zero, familySymbolList.FirstOrDefault());
                                //        tiles.LookupParameter("長").Set(length);
                                //        tiles.LookupParameter("寬").Set(width);
                                //        tiles.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(wall.LevelId);
                                //    }
                                //    // 取得內牆
                                //    IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);
                                //    foreach (Reference interiorFace in interiorFaces)
                                //    {
                                //        // 使用參考面和放置點創建磚塊族群
                                //        FamilyInstance tiles = doc.Create.NewFamilyInstance(interiorFace, location, XYZ.Zero, familySymbolList.FirstOrDefault());
                                //        tiles.LookupParameter("長").Set(length);
                                //        tiles.LookupParameter("寬").Set(width);
                                //        tiles.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(wall.LevelId);
                                //    }
                                //}
                            }
                        }
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }

                trans.Commit();
                TaskDialog.Show("Revit", "完成");
            }

            //// 協助冠鋐測試畫線
            //Element elem = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).FirstOrDefault();
            //Tuple<List<Curve>, List<PolyLine>> curvesPolyLines = GetCurvesPolyLines(doc, elem);
            //List<Curve> curves = curvesPolyLines.Item1;
            //List<PolyLine> polyLines = curvesPolyLines.Item2;

            //using (Transaction trans = new Transaction(doc, "畫線"))
            //{
            //    trans.Start();
            //    foreach (Curve curve in curves)
            //    {
            //        try
            //        {
            //            Line line = Line.CreateBound(curve.Tessellate()[0], curve.Tessellate()[curve.Tessellate().Count - 1]);
            //            XYZ normal = new XYZ(line.Direction.Z - line.Direction.Y, line.Direction.X - line.Direction.Z, line.Direction.Y - line.Direction.X); // 使用與線不平行的任意向量
            //            Plane plane = Plane.CreateByNormalAndOrigin(normal, curve.Tessellate()[0]);
            //            SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
            //            ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);
            //        }
            //        catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            //    }
            //    foreach (PolyLine polyLine in polyLines)
            //    {
            //        for (int i = 0; i < polyLine.GetCoordinates().Count; i++)
            //        {
            //            try
            //            {
            //                Line line = Line.CreateBound(polyLine.GetCoordinate(i), polyLine.GetCoordinate(i + 1));
            //                XYZ normal = new XYZ(line.Direction.Z - line.Direction.Y, line.Direction.X - line.Direction.Z, line.Direction.Y - line.Direction.X); // 使用與線不平行的任意向量
            //                Plane plane = Plane.CreateByNormalAndOrigin(normal, polyLine.GetCoordinate(i));
            //                SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
            //                ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);
            //            }
            //            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            //        }
            //    }
            //    trans.Commit();
            //}

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
        ///// <summary>
        ///// 協助冠鋐測試畫線
        ///// </summary>
        ///// <param name="doc"></param>
        ///// <param name="elem"></param>
        ///// <returns></returns>
        //private Tuple<List<Curve>, List<PolyLine>> GetCurvesPolyLines(Document doc, Element elem)
        //{
        //    // 1.讀取Geometry Option
        //    Options options = new Options();
        //    //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
        //    options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
        //    options.ComputeReferences = true;
        //    options.IncludeNonVisibleObjects = true;
        //    // 得到幾何元素
        //    GeometryElement geomElem = elem.get_Geometry(options);
        //    List<Curve> curves = GeometryCurves(geomElem); // 取得Curve
        //    List<PolyLine> polyLines = GeometryPolyLines(geomElem); // 取得PolyLine

        //    return Tuple.Create(curves, polyLines);
        //}
        ///// <summary>
        ///// 取得Curve
        ///// </summary>
        ///// <param name="geoObj"></param>
        ///// <returns></returns>
        //private List<Curve> GeometryCurves(GeometryObject geoObj)
        //{
        //    List<Curve> curves = new List<Curve>();
        //    if (geoObj is Arc)
        //    {
        //        Arc curve = (Arc)geoObj;
        //        curves.Add(curve);
        //    }
        //    if (geoObj is Line)
        //    {
        //        Line curve = (Line)geoObj;
        //        curves.Add(curve);
        //    }
        //    if (geoObj is GeometryInstance)
        //    {
        //        GeometryInstance geoIns = geoObj as GeometryInstance;
        //        GeometryElement geometryElement = (geoObj as GeometryInstance).GetSymbolGeometry(geoIns.Transform); // 座標轉換
        //        foreach (GeometryObject o in geometryElement)
        //        {
        //            curves.AddRange(GeometryCurves(o));
        //        }
        //    }
        //    else if (geoObj is GeometryElement)
        //    {
        //        GeometryElement geometryElement2 = (GeometryElement)geoObj;
        //        foreach (GeometryObject o in geometryElement2)
        //        {
        //            curves.AddRange(GeometryCurves(o));
        //        }
        //    }
        //    return curves;
        //}
        ///// <summary>
        ///// 取得PolyLine
        ///// </summary>
        ///// <param name="geoObj"></param>
        ///// <returns></returns>
        //private List<PolyLine> GeometryPolyLines(GeometryObject geoObj)
        //{
        //    List<PolyLine> polyLines = new List<PolyLine>();
        //    if (geoObj is PolyLine)
        //    {
        //        PolyLine curve = (PolyLine)geoObj;
        //        polyLines.Add(curve);
        //    }
        //    if (geoObj is GeometryInstance)
        //    {
        //        GeometryInstance geoIns = geoObj as GeometryInstance;
        //        GeometryElement geometryElement = (geoObj as GeometryInstance).GetSymbolGeometry(geoIns.Transform); // 座標轉換
        //        foreach (GeometryObject o in geometryElement)
        //        {
        //            polyLines.AddRange(GeometryPolyLines(o));
        //        }
        //    }
        //    else if (geoObj is GeometryElement)
        //    {
        //        GeometryElement geometryElement2 = (GeometryElement)geoObj;
        //        foreach (GeometryObject o in geometryElement2)
        //        {
        //            polyLines.AddRange(GeometryPolyLines(o));
        //        }
        //    }
        //    return polyLines;
        //}
    }
}