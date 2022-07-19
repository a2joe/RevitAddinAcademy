#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdProjectSetup : IExternalCommand
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
            dialog.InitialDirectory = @"c:\";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xls*";

            string filePath = "";
            //string[] filePaths;

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file.");
                return Result.Failed;
            }
            
            filePath = dialog.FileName;
            int levelCounter = 0;
            int sheetCounter = 0;   

           
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook excelWB = excelapp.Workbooks.Open(filePath);

            Excel.Worksheet excelWs1 = GetExcelWorksheeetByName(excelWB, "Levels");
            Excel.Worksheet excelWs2 = GetExcelWorksheeetByName(excelWB, "Sheets");


            List<LevelStruct> levelData = GetLevelDataFromExcel(excelWs1);
            List<SheetStruct> sheetData = GetSheetDataFromExcel(excelWs2)
            

            excelWB.Close();
            excelapp.Quit();



            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Levels");

                ViewFamilyType planVFT = GetViewFamily(doc, "plan");
                ViewFamilyType rcpVFT = GetViewFamily(doc, "rcp");

                foreach(LevelStruct level in levelData)
                {
                    Level curLevel = Level.Create(doc, curLevel.LevelElev);
                    curLevel.Name = curLevel.LevelName;

                    ViewPlan curFloorPlan = ViewPlan.Create(doc, planVFT, newLevel.Id);
                    ViewPlan curRCP = ViewPlan.CreateAreaPlan(doc.rcpVFT.Id, newLevel.Id);

                    curRCP.Name = curRCP.Name + " RCP";

                    ViewPlan.Create()
                }

                //FilteredElementCollector.ReferenceEquals..

                foreach(SheetStruct curSheet in sheetData)
                {
                    ViewSheet newSheet = ViewSheet.Create(doc, FilteredElementCollector.FirElementId());

                }
                t.Commit();
            }




            return Result.Succeeded;
        }

        private ViewFamilyType GetViewFamily(Document doc, string type)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            foreach(ViewFamilyType vft in collector)
            {
                if(vft.ViewFamily == ViewFamily.FloorPlan && type == "plan")
                {
                    return vft;
                }
                else if(vft.ViewFamily == ViewFamily.CeilingPlan && type == "rcp")
                {
                    return vft;
                }

            }
            return null;
        }

        private List<SheetStruct> GetSheetDataFromExcel(Excel.Worksheet excelWs)
        {
            //int rowcountlev = excelRnglevel.Rows.Count;
            List<LevelStruct> returnList = new List<LevelStruct>();

            Excel.Range excelRngSh = excelWSsheet.UsedRange;
            int rowcountsh = excelRngSh.Rows.Count;
            for (int i = 2; i <= rowcountsh; i++)
            {
                Excel.Range data1 = excelWs.Cells[i, 1];
                Excel.Range data2 = excelWs.Cells[i, 2];
                Excel.Range data3 = excelWs.Cells[i, 3];
                Excel.Range data4 = excelWs.Cells[i, 4];
                Excel.Range data5 = excelWs.Cells[i, 5];

                SheetStruct curSheet = new SheetStruct();
                curSheet.SheetNumber = data1.Value.ToString();
                curSheet.SheetName = data2.Value.ToString();
                curSheet.SheetView = data3.Value.ToString();
                curSheet.DrawnBy = data4.Value; 
                curSheet.CheckedBy = data5.Value;

                string datashname = sheetname.Value.ToString();
                string datashnum = sheetnum.Value.ToString();

                LevelStruct curLevel = new LevelStruct(levelName, levelElev);
                returnList.Add(curLevel);

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Sheets");

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                    collector.WhereElementIsElementType();

                    ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    curSheet.SheetNumber = datashnum;
                    curSheet.Name = datashname;

                    t.Commit();
                }
            }
        }

        private List<LevelStruct> GetLevelDataFromExcel(Excel.Worksheet excelWs1)
        {
            throw new NotImplementedException();
        }

        private Excel.Worksheet GetExcelWorksheeetByName(Excel.Workbook curWb, string wsName)
        {
            foreach(Excel.Worksheet ws in curWb.Worksheets)
            {
                if(ws.Name == wsName)
                {
                    return ws;
                }
            }

            return null;
        }

        private struct LevelStruct
        {
            public string LevelName;
            public double LevelElev;

            public LevelStruct(string name, double elev)
            {
                LevelName = name;
                LevelElev = elev;
            }
        }

        private struct SheetStruct
        {
            public string SheetNumber;
            public string SheetName;
            public string SheetView;
            public string DrawnBy;
            public string CheckedBy;
            //public string SheetCat;

            public SheetStruct(string number, string name, string view, string db, string cb)
            {
                SheetNumber = number;
                SheetName = name;
                SheetView = view;
                DrawnBy = db;
                CheckedBy = cb;

                //SheetCat = "DOCUMENT";
            }
        }
    }
}