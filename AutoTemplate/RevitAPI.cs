using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace AutoTemplate
{
    public class RevitAPI : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        public static string picPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "Pic"); // 圖片路徑
        //public static string picPath = Path.Combine(@"C:\ProgramData\Autodesk\Revit\Addins\2020\AutoSign", "Pic"); // 圖片路徑
        public Result OnStartup(UIControlledApplication a)
        {
            string autoSignPath = Path.Combine(Directory.GetParent(addinAssmeblyPath).FullName, "AutoTemplate.dll");
            //string autoSignPath = Path.Combine(@"C:\ProgramData\Autodesk\Revit\Addins\2020\AutoSign", "AutoSign.dll");
            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("自動放置模板"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("自動放置模板", "模板建置"); }
            catch
            {
                List<RibbonPanel> panel_list = new List<RibbonPanel>();
                panel_list = a.GetRibbonPanels("自動放置模板");
                foreach (RibbonPanel rp in panel_list)
                {
                    if (rp.Name == "模板建置")
                    {
                        ribbonPanel = rp;
                    }
                }
            }
            //PushButton autoPutBtn = ribbonPanel.AddItem(new PushButtonData("AutoPut", "自動佈設", autoSignPath, "AutoSign.AutoPut")) as PushButton;
            //autoPutBtn.LargeImage = convertFromBitmap(Properties.Resources.自動佈設);

            return Result.Succeeded;
        }

        BitmapSource convertFromBitmap(System.Drawing.Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
        /// <summary>
        /// ElementId轉換為數字
        /// </summary>
        /// <param name="elemId"></param>
        /// <returns></returns>
        public static int GetValue(ElementId elemId)
        {
            return elemId.IntegerValue; // 2020
            //return ((int)elemId.Value); // 2024
        }
        ///// <summary>
        ///// 取得TaggedLocalElementId
        ///// </summary>
        ///// <param name="independentTags"></param>
        ///// <returns></returns>
        //public static List<ElementId> GetTaggedLocalElementId(List<IndependentTag> independentTags)
        //{
        //    return independentTags.Select(x => x.TaggedLocalElementId).Distinct().ToList(); // 2020
        //    //return independentTags.Select(x => x.GetTaggedLocalElementIds().FirstOrDefault()).Distinct().ToList(); // 2022
        //}
        ///// <summary>
        ///// 於識別資料中建立參數
        ///// </summary>
        ///// <param name="familyManager"></param>
        ///// <returns></returns>
        //public static FamilyParameter FamilyPara(FamilyManager familyManager)
        //{
        //    return familyManager.AddParameter("指標ID", BuiltInParameterGroup.PG_IDENTITY_DATA, ParameterType.Text, true); // 2020
        //    //return familyManager.AddParameter("指標ID", GroupTypeId.IdentityData, SpecTypeId.String.Text, true); // 2022
        //}
        ///// <summary>
        ///// 轉換單位
        ///// </summary>
        ///// <param name="number"></param>
        ///// <param name="unit"></param>
        ///// <returns></returns>        
        //public static double ConvertFromInternalUnits(double number, string unit)
        //{
        //    if (unit.Equals("meters"))
        //    {
        //        number = UnitUtils.ConvertFromInternalUnits(number, DisplayUnitType.DUT_METERS); // 2020
        //        //number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Meters); // 2022
        //    }
        //    else if (unit.Equals("millimeters"))
        //    {
        //        number = UnitUtils.ConvertFromInternalUnits(number, DisplayUnitType.DUT_MILLIMETERS); // 2020
        //        //number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Millimeters); // 2022
        //    }
        //    return number;
        //}
        ///// <summary>
        ///// 轉換單位
        ///// </summary>
        ///// <param name="number"></param>
        ///// <param name="unit"></param>
        ///// <returns></returns>
        //public static double ConvertToInternalUnits(double number, string unit)
        //{
        //    if (unit.Equals("meters"))
        //    {
        //        number = UnitUtils.ConvertToInternalUnits(number, DisplayUnitType.DUT_METERS); // 2020
        //        //number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Meters); // 2022
        //    }
        //    else if (unit.Equals("millimeters"))
        //    {
        //        number = UnitUtils.ConvertToInternalUnits(number, DisplayUnitType.DUT_MILLIMETERS); // 2020
        //        //number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Millimeters); // 2022
        //    }
        //    return number;
        //}
        /// <summary>
        ///  關閉
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
        /// <summary>
        /// 2022
        /// </summary>
        /// <param name="elemId"></param>
        /// <returns></returns>
        public static List<ElementId> GetTaggedLocalElementId(List<IndependentTag> independentTags)
        {
            return independentTags.Select(x => x.GetTaggedLocalElementIds().FirstOrDefault()).Distinct().ToList(); // 2022
        }
        public static FamilyParameter FamilyPara(FamilyManager familyManager)
        {
            return familyManager.AddParameter("指標ID", GroupTypeId.IdentityData, SpecTypeId.String.Text, true); // 2022
        }
        public static double ConvertFromInternalUnits(double number, string unit)
        {
            if (unit.Equals("meters"))
            {
                number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Meters); // 2022
            }
            else if (unit.Equals("millimeters"))
            {
                number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Millimeters); // 2022
            }
            return number;
        }
        public static double ConvertToInternalUnits(double number, string unit)
        {
            if (unit.Equals("meters"))
            {
                number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Meters); // 2022
            }
            else if (unit.Equals("millimeters"))
            {
                number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Millimeters); // 2022
            }
            return number;
        }
    }
}