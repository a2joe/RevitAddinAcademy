#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdLesson04 : IExternalCommand
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements:");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            Level curLevel = GetLevelByName(doc, "Level 1");

            MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit stuff");

                foreach (Element element in pickList)
                {
                    if (element is CurveElement)
                    {
                        CurveElement curve = (CurveElement)element;
                        CurveElement curve2 = element as CurveElement;

                        curveList.Add(curve);

                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;
                        GraphicsStyle curGS2 = curve2.LineStyle as GraphicsStyle;
                        ElementId curID = curve.Id as ElementId;
                        ElementId curID2 = curve2.Id as ElementId;

                        Curve curCurve = curve.GeometryCurve;
                        XYZ startPoint = curCurve.GetEndPoint(0);
                        XYZ endPoint = curCurve.GetEndPoint(1);

                        switch (curGS.Name)
                        {
                            case "A-GLAZ":
                                Debug.Print("A-GLAZ" + curID + " " + curID2);
                                break;

                            case "A-WALL":
                                Debug.Print("A-WALL" + curID + " " + curID2);
                                break;

                            case "M-DUCT":
                                Debug.Print("M-DUCT" + curID + " " + curID2);
                                break;

                            case "P-PIPE":
                                Debug.Print("P-PIPE" + curID + " " + curID2); 
                                break;

                            case "<Medium>":

                                Debug.Print("found a medium line");
                                break;

                            case "<Thin lines>":
                                Debug.Print("found a thin line");
                                break;

                            case "<Wide Lines>":
                                Pipe newPipe = Pipe.Create(
                                    doc,
                                    curSystemType.Id,
                                    curPipeType.Id,
                                    curLevel.Id,
                                    startPoint,
                                    endPoint);
                                break;

                            default:
                                Debug.Print("Found something else");
                                break;
                        }

                        if (curGS.Name == "A-GLAZ" ||
                            curGS.Name == "A-WALL" ||
                            curGS.Name == "M-DUCT" ||
                            curGS.Name == "P-PIPE")
                        {
                            Wall newWall = Wall.Create(doc, curCurve, curID, curID2, 15, 0, false, false);
                        }

                        //Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);


                        Debug.Print(curGS.Name);
                    }
                }

                t.Commit();
            }


            TaskDialog.Show("complete", curveList.Count.ToString());
            return Result.Succeeded;
        }

        //private Wall Wall.Create();

        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element curElem in collector)
            {
                WallType wallType = curElem as WallType;

                if (wallType.Name == wallTypeName)
                    return wallType;
            }
            return null;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element curElem in collector)
            {
                Level level = curElem as Level;

                if (level.Name == levelName)
                    return level;
            }
            return null;
        }

        private MEPSystemType GetSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in collector)
            {
                MEPSystemType curType = curElem as MEPSystemType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }

        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element curElem in collector)
            {
                PipeType curType = curElem as PipeType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
        public void GetInfo_WallType(WallType wallType)
        {
            string message;
            // Reports the nature of the wall
            message = "The wall type kind is : " + wallType.Kind.ToString();

            // Get the overall thickness of this type of wall.
            message += "\nThe wall type Width is : " + wallType.Width.ToString();

            TaskDialog.Show("Revit", message);
        }
    }
}