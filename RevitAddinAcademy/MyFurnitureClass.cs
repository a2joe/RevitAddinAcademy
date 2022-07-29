using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;

namespace RevitAddinAcademy
{
    internal class MyFurnitureClass
    {
    }
    public class Furniture
    {
        public List<Furniture> FurnitureList { get; set; }

        public Furniture(List<Furniture> furniture)
        {
            FurnitureList = furniture;
        }

        public int GetFurnitureCount()
        {
            return FurnitureList.Count;
        }
    }
    public static class Utilities
    {
        public static string GetTextFromClass()
        {
            return "All I got was this lousy t-shirt from my static class";
        }

    }
    public static FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
    {

        FilteredElementCollector collector = new FilteredElementCollector(doc);
        collector.OfClass(typeof(Family));

        foreach (Element element in collector)
        {
            Family curFamily = element as Family;

            if (curFamily.Name == familyName)
            {
                ISet<ElementId> famSymbolIdList = curFamily.GetFamilySymbolIds();

                foreach (ElementId famSymbolId in famSymbolIdList)
                {
                    FamilySymbol curFS = doc.GetElement(famSymbolId) as FamilySymbol;

                    if (curFS.Name == typeName)
                        return curFS;
                }
            }
        }
        return null;
    }
}