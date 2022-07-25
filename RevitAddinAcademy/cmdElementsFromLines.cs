#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RevitAddinAcademy
{
    /*
     * Stolen from MichalZaperty
     */

    [Transaction(TransactionMode.Manual)]
    public class cmdElementsFromLines : IExternalCommand
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


            IList<Element> picklist = uidoc.Selection.PickElementsByRectangle("Select Elements");

            List<CurveElement> curveList = new List<CurveElement>();

            Level level = GetLevel(doc, "Level 1");

            WallType wallGlaz = GetWallType(doc, "Storefront");
            WallType wallArch = GetWallType(doc, @"Generic - 8""");


            PipeType pipeType = GetPipe(doc, "Default");
            MEPSystemType pipeSystem = GetMEPSystem(doc, "Domestic Hot Water");
            DuctType ductType = GetDuct(doc, "Default");
            MEPSystemType ductSystem = GetMEPSystem(doc, "Return Air");

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit stuff");
                try
                {
                    foreach (Element e in picklist)
                    {
                        if (e is CurveElement)
                        {
                            CurveElement curve = (CurveElement)e;
                            curveList.Add(curve);


                            GraphicsStyle curveGS = curve.LineStyle as GraphicsStyle;

                            switch (curveGS.Name)
                            {
                                case "A-GLAZ":
                                    //Wall newWall = Wall.Create(
                                    //doc,
                                    //curCurve,
                                    //curWallType.Id,
                                    //curLevel.Id,
                                    //15,
                                    //0,
                                    //false,
                                    //false
                                    //);
                                    Wall newGWall = Wall.Create(
                                        doc,
                                        curve.GeometryCurve,
                                        wallGlaz.Id,
                                        level.Id,
                                        5,
                                        0,
                                        false,
                                        false
                                        );
                                    Debug.WriteLine(curveGS.Name.ToString());
                                    break;

                                case "A-WALL":
                                    Wall newAWall = Wall.Create(
                                        doc,
                                        curve.GeometryCurve,
                                        wallArch.Id,
                                        level.Id,
                                        5,
                                        0,
                                        false,
                                        false
                                        );
                                    Debug.WriteLine(curveGS.Name.ToString());
                                    break;

                                case "M-DUCT":
                                    XYZ startPoint = curve.GeometryCurve.GetEndPoint(0);
                                    XYZ endPoint = curve.GeometryCurve.GetEndPoint(1);
                                    Duct newDuct = Duct.Create(
                                        doc,
                                        ductSystem.Id,
                                        ductType.Id,
                                        level.Id,
                                        startPoint,
                                        endPoint
                                        );
                                    Debug.WriteLine(curveGS.Name.ToString());

                                    break;

                                case "P-PIPE":
                                    XYZ startPPoint = curve.GeometryCurve.GetEndPoint(0);
                                    XYZ endPPoint = curve.GeometryCurve.GetEndPoint(1);
                                    Pipe newPipe = Pipe.Create(
                                        doc,
                                        pipeSystem.Id,
                                        pipeType.Id,
                                        level.Id,
                                        startPPoint,
                                        endPPoint
                                        );
                                    Debug.WriteLine(curveGS.Name.ToString());
                                    break;
                            }
                        }

                        ICollection<ElementId> deletedIdSet = doc.Delete(e.Id);


                    }
                }

                catch (Exception e)
                {
                    // Debug.WriteLine(e.ToString());
                    // Debug.WriteLine(e.Message.ToString());
                    // Debug.WriteLine(curveList.ToString());

                }


                t.Commit();
            }





            return Result.Succeeded;
        }
        private WallType GetWallType(Document doc, string input)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element e in collector)
            {
                WallType wt = e as WallType;
                if (wt.Name == input)
                    return wt;
            }
            return null;
        }
        private Level GetLevel(Document doc, string input)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element e in collector)
            {
                Level curElement = e as Level;
                if (curElement.Name == input)
                    return curElement;
            }
            return null;
        }
        private MEPSystemType GetMEPSystem(Document doc, string input)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element e in collector)
            {
                MEPSystemType curElement = e as MEPSystemType;
                if (curElement.Name == input)
                    return curElement;
            }
            return null;
        }
        private PipeType GetPipe(Document doc, string input)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element e in collector)
            {
                PipeType curElement = e as PipeType;
                if (curElement.Name == input)
                    return curElement;
            }
            return null;
        }
        private DuctType GetDuct(Document doc, string input)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element e in collector)
            {
                DuctType curElement = e as DuctType;
                if (curElement.Name == input)
                    return curElement;
            }
            return null;
        }
    }
}