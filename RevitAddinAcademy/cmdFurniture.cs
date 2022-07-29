#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdFurniture : IExternalCommand
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

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.Multiselect = false;
            //dialog.Filter = "Excel files | *.xls*";
            dialog.Filter = "Excel files | *.xlsx; *.xls; *.xlsm | All files | *.*";

            if (dialog.ShowDialog() != Form.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file.");
                return Result.Failed;
            }

            string excelFile = dialog.FileName;

            return Result.Succeeded;
        }
    }
}