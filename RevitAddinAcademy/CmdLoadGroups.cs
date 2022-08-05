#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class CmdLoadGroups : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            /*
             * Create a variable
             */

            string revitFile = "";

            /*
             * Prompt user to select a file
             */

            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Revit file";
            ofd.InitialDirectory = @"c:\";
            ofd.Filter = "Revit files (*.rvt)|*.rvt";

            if(ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;

            revitFile = ofd.FileName;

            /*
             * If file is selected, open that file, and activate it
             */

            UIDocument newUIDoc = uiapp.OpenAndActivateDocument(revitFile);
            Document newDoc = newUIDoc.Document;

            /*
             * Creates new document variable, then apply FilteredElementCollector
             */

            FilteredElementCollector collector = new FilteredElementCollector(newDoc);
            collector.OfCategory(BuiltInCategory.OST_IOSModelGroups);
            collector.WhereElementIsNotElementType();

            /*
             * Loop through the results (generator list of Element IDs
             */

            List<ElementId> groupIdList = new List<ElementId>();
            foreach (Element curElem in collector)
            {
                groupIdList.Add(curElem.Id);
            }

            Transform transform = null;
            CopyPasteOptions options = new CopyPasteOptions();

            /*
             * Create a transaction
             */

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Loads groups");

                /*
                 * Copy from source to destination document
                 */

                ElementTransformUtils.CopyElements(newDoc, groupIdList, doc, transform, options);
                t.Commit();
            }

            try
            {
                /*
                 * Save and close
                 */

                uiapp.OpenAndActivateDocument(doc.PathName);
                newUIDoc.SaveAndClose();
            }
            catch (Exception)
            { }

            TaskDialog.Show("Complete", "Loaded " + groupIdList.Count.ToString() + " groups into the current model.");
            return Result.Succeeded;
        }
    }
}