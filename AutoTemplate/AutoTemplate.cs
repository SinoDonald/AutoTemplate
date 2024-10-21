using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Media3D;

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
        public class TileType
        {
            public int A { get; set; }
            public int B { get; set; }
            public int C { get; set; }
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
                            CalculateGeom(uidoc, doc, exteriorFaces, fs, wall.LevelId); // 放置磁磚
                            //IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior); // 內牆
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
                    BoundingBoxUV bboxUV = face.GetBoundingBox();
                    XYZ minXYZ = face.Evaluate(bboxUV.Min); // 最小座標點
                    XYZ maxXYZ = face.Evaluate(bboxUV.Max); // 最大座標點
                    Vector vector = new Vector(maxXYZ.X - minXYZ.X, maxXYZ.Y - minXYZ.Y); // 方向向量
                    vector = GetVectorOffset(vector, RevitAPI.ConvertToInternalUnits(100, "millimeters"));
                    minXYZ = new XYZ(minXYZ.X + vector.X, minXYZ.Y + vector.Y, minXYZ.Z); // 預留起始的空間
                    maxXYZ = new XYZ(maxXYZ.X - vector.X, maxXYZ.Y - vector.Y, maxXYZ.Z); // 預留結束的空間

                    double wallLength = minXYZ.DistanceTo(new XYZ(maxXYZ.X, maxXYZ.Y, minXYZ.Z)); // 牆長度
                    double wallWidth = minXYZ.DistanceTo(new XYZ(minXYZ.X, minXYZ.Y, maxXYZ.Z)); // 牆寬度

                    // 磁磚的寬度
                    double dWidth = RevitAPI.ConvertToInternalUnits(300, "millimeters");
                    int d = Convert.ToInt32(Math.Floor(wallWidth / dWidth));
                    double d1 = wallWidth % dWidth;
                    double eWidth = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                    int e = Convert.ToInt32(Math.Floor(d1 / eWidth));
                    double e1 = d1 % eWidth;
                    double fWidth = RevitAPI.ConvertToInternalUnits(100, "millimeters");
                    int f = Convert.ToInt32(Math.Floor(e1 / fWidth));

                    // 磁磚的長度
                    double aLength = RevitAPI.ConvertToInternalUnits(300, "millimeters");
                    int a = Convert.ToInt32(Math.Floor(wallLength / aLength));
                    double a1 = wallLength % aLength;
                    double bLength = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                    int b = Convert.ToInt32(Math.Floor(a1 / bLength));
                    double b1 = a1 % bLength;
                    double cLength = RevitAPI.ConvertToInternalUnits(100, "millimeters");
                    int c = Convert.ToInt32(Math.Floor(b1 / cLength));

                    XYZ location = minXYZ;
                    for(int j = 0; j < 3; j++)
                    {
                        int frequency = d;
                        double width = RevitAPI.ConvertToInternalUnits(300, "millimeters");
                        if (j == 1)
                        {
                            frequency = e;
                            width = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                        }
                        else if (j == 2)
                        {
                            frequency = f;
                            width = RevitAPI.ConvertToInternalUnits(100, "millimeters");
                        }
                        for (int k = 0; k < frequency; k++)
                        {
                            for (int i = 0; i < 3; i++)
                            {
                                int counts = a;
                                double length = RevitAPI.ConvertToInternalUnits(300, "millimeters");
                                if (i == 1)
                                {
                                    counts = b;
                                    length = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                                }
                                else if (i == 2)
                                {
                                    counts = c;
                                    length = RevitAPI.ConvertToInternalUnits(100, "millimeters");
                                }
                                location = PutTiles(doc, counts, referenceFace, location, fs, elemId, vector, length, width, k - 1, maxXYZ, minXYZ);
                            }
                            location = new XYZ(minXYZ.X, minXYZ.Y, width * k+1);
                        }
                    }
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }
        }
        /// <summary>
        /// 放置磁磚
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="counts"></param>
        /// <param name="referenceFace"></param>
        /// <param name="location"></param>
        /// <param name="fs"></param>
        /// <param name="elemId"></param>
        /// <param name="vector"></param>
        /// <param name="length"></param>
        /// <param name="width"></param>
        /// <param name="maxXYZ"></param>
        /// <param name="minXYZ"></param>
        /// <returns></returns>
        private XYZ PutTiles(Document doc, int counts, Reference referenceFace, XYZ location, FamilySymbol fs, ElementId elemId, Vector vector, double length, double width, int k, XYZ maxXYZ, XYZ minXYZ)
        {
            for (int count = 0; count < counts; count++)
            {
                try
                {
                    // 放置外牆
                    FamilyInstance tiles = doc.Create.NewFamilyInstance(referenceFace, location, XYZ.Zero, fs);
                    tiles.LookupParameter("長").Set(length);
                    tiles.LookupParameter("寬").Set(width);
                    tiles.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(elemId);

                    vector = new Vector(maxXYZ.X - minXYZ.X, maxXYZ.Y - minXYZ.Y); // 方向向量
                    vector = GetVectorOffset(vector, length);
                    location = new XYZ(location.X + vector.X, location.Y + vector.Y, width * k + 1);

                    //doc.Regenerate();
                    //uidoc.RefreshActiveView();
                }
                catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
            }

            return location;
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