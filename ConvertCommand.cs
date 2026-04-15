using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FloorRoofTopo
{
    [Transaction(TransactionMode.Manual)]
    public class ConvertCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ── 1. Collect valid elements ──────────────────────────────
                List<Element> validElements = new List<Element>();

                // Check current selection first
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (ConversionEngine.IsValidSourceElement(elem))
                        validElements.Add(elem);
                }

                // If nothing valid selected, prompt user to pick
                if (validElements.Count == 0)
                {
                    try
                    {
                        IList<Reference> refs = uidoc.Selection.PickObjects(
                            ObjectType.Element,
                            new FloorRoofTopoFilter(),
                            "Chọn các Floor, Roof hoặc TopoSolid cần chuyển đổi");

                        foreach (Reference r in refs)
                        {
                            Element elem = doc.GetElement(r);
                            if (ConversionEngine.IsValidSourceElement(elem))
                                validElements.Add(elem);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }
                }

                if (validElements.Count == 0)
                {
                    TaskDialog.Show("Converter",
                        "Không có phần tử Floor, Roof hoặc TopoSolid hợp lệ nào được chọn.");
                    return Result.Failed;
                }

                // ── 2. Determine source type description ───────────────────
                var sourceTypes = validElements
                    .Select(e => ConversionEngine.GetElementTypeName(e))
                    .Distinct()
                    .ToList();
                string sourceDesc = string.Join(", ", sourceTypes);

                // ── 3. Show conversion dialog ──────────────────────────────
                var dialog = new ConvertDialog(
                    sourceDesc,
                    validElements.Count,
                    ConversionEngine.IsTopoSolidAvailable());

                // Set owner to Revit main window
                SetDialogOwner(dialog, commandData);

                if (dialog.ShowDialog() != true)
                    return Result.Cancelled;

                ConvertTarget target = dialog.SelectedTargetType;
                bool deleteSource = dialog.DeleteSource;

                // ── 4. Execute conversions ─────────────────────────────────
                int success = 0;
                int skipped = 0;
                int failed = 0;
                List<string> messages = new List<string>();
                List<ConversionDiagnosticInfo> allDiagnostics = new List<ConversionDiagnosticInfo>();

                using (Transaction tx = new Transaction(doc, "Convert Floor/Roof/Topo"))
                {
                    tx.Start();

                    foreach (Element elem in validElements)
                    {
                        ConversionResult result =
                            ConversionEngine.Convert(doc, elem, target, deleteSource);

                        if (result.Success)
                        {
                            if (ConversionEngine.IsSameType(elem, target))
                                skipped++;
                            else
                                success++;

                            if (!string.IsNullOrEmpty(result.WarningMessage))
                                messages.Add("⚠ " + result.WarningMessage);
                        }
                        else
                        {
                            failed++;
                            messages.Add("❌ " + (result.ErrorMessage ?? "Unknown error"));
                        }

                        if (result.Diagnostics != null)
                            allDiagnostics.Add(result.Diagnostics);
                    }

                    if (failed == validElements.Count)
                    {
                        tx.RollBack();
                    }
                    else
                    {
                        tx.Commit();
                    }
                }

                // ── 5. Show diagnostic results ────────────────────────────
                var diagDialog = new DiagnosticDialog(
                    allDiagnostics, success, skipped, failed);
                SetDialogOwner(diagDialog, commandData);
                diagDialog.ShowDialog();

                return success > 0 ? Result.Succeeded : Result.Failed;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Converter Error", ex.ToString());
                return Result.Failed;
            }
        }

        /// <summary>
        /// Set the WPF dialog's owner to the Revit main window.
        /// Uses reflection for cross-version compatibility.
        /// </summary>
        private static void SetDialogOwner(System.Windows.Window dialog, ExternalCommandData cmdData)
        {
            try
            {
                // Revit 2019+: UIApplication.MainWindowHandle
                var prop = cmdData.Application.GetType()
                    .GetProperty("MainWindowHandle");
                if (prop != null)
                {
                    IntPtr handle = (IntPtr)prop.GetValue(cmdData.Application);
                    var helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                    helper.Owner = handle;
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// Selection filter that only allows Floor, FootPrintRoof, and TopoSolid.
    /// </summary>
    public class FloorRoofTopoFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return ConversionEngine.IsValidSourceElement(elem);
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
