using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

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
        // 牆與面
        public class WallAndFace
        {
            public Element wall { get; set; }
            public List<Face> faces = new List<Face>();
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
                wallAndFaces.wall = wall;
                wallAndFaces.faces = wallFaces;
                wallsAndFaces.Add(wallAndFaces);
            }

            using(Transaction trans = new Transaction(doc, "自動放置模板"))
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
                            foreach (Face face in wallAndFace.faces)
                            {
                                //// 獲取平面的BoundingBox, 找到左上角的點
                                //BoundingBoxUV boundingBox = face.GetBoundingBox();
                                //UV min = boundingBox.Min;
                                //UV max = boundingBox.Max;
                                //XYZ location = face.Evaluate(new UV(max.U, max.V));
                                PlanarFace planarFace = face as PlanarFace;
                                if (planarFace != null)
                                {
                                    // Get the center of the wall
                                    BoundingBoxUV bboxUV = face.GetBoundingBox();
                                    UV center = (bboxUV.Max + bboxUV.Min) / 2.0;
                                    XYZ location = face.Evaluate(center);
                                    XYZ normal = face.ComputeNormal(center);
                                    XYZ refDir = planarFace.FaceNormal;


                                    // 获取墙的侧面
                                    IList<Autodesk.Revit.DB.Reference> sideFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                                    // 定位点（族实例的放置点，可以根据需要调整）
                                    XYZ insertionPoint = new XYZ(0, 0, 0);
                                    // 使用参考面和插入点创建族实例
                                    FamilyInstance tiles = doc.Create.NewFamilyInstance(sideFaces.First(), location, XYZ.Zero, familySymbolList.FirstOrDefault());

                                    //FamilyInstance tiles = doc.Create.NewFamilyInstance(location, familySymbolList.FirstOrDefault(), wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    double length = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                                    double width = RevitAPI.ConvertToInternalUnits(200, "millimeters");
                                    double heigth = RevitAPI.ConvertToInternalUnits(3800, "millimeters");
                                    tiles.LookupParameter("長").Set(length);
                                    tiles.LookupParameter("寬").Set(width);
                                    tiles.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(heigth);
                                }
                            }
                        }
                    }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }

                trans.Commit();
                TaskDialog.Show("Revit", "完成");
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