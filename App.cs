using System;
using System.Reflection;
using Autodesk.Revit.UI;

namespace FloorRoofTopo
{
    /// <summary>
    /// IExternalApplication that registers a ribbon button for the converter.
    /// </summary>
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon panel in built-in "Add-Ins" tab
                RibbonPanel panel = application.CreateRibbonPanel("Floor/Roof/Topo");

                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                PushButtonData buttonData = new PushButtonData(
                    "ConvertFloorRoofTopo",
                    "Convert\nF↔R↔T",
                    assemblyPath,
                    "FloorRoofTopo.ConvertCommand");

                buttonData.ToolTip = "Chuyển đổi qua lại giữa Floor, Roof và TopoSolid";
                buttonData.LongDescription =
                    "Chọn các phần tử Floor (Sàn), Roof (Mái) hoặc TopoSolid (Địa hình),\n" +
                    "sau đó chuyển đổi chúng sang loại khác.\n\n" +
                    "• Floor ↔ Roof: Hỗ trợ tất cả phiên bản Revit\n" +
                    "• TopoSolid: Chỉ hỗ trợ Revit 2024 trở lên\n" +
                    "• Bảo toàn các điểm nhấc (SlabShapeEditor)";

                PushButton button = panel.AddItem(buttonData) as PushButton;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FloorRoofTopo Startup Error", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
