#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdEmployees : IExternalCommand
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

            Employee emp1 = new Employee("Joe", 24, "blue,red,white");
            Employee emp2 = new Employee("Mary", 36, "greenred,brown");
            Employee emp3 = new Employee("Felix", 45, "gray, beige");

            List<Employee> empList = new List<Employee>();
            empList.Add(emp1);
            empList.Add(emp2);
            empList.Add(emp3);
;

            Employees allEmployee = new Employees(empList);

            Debug.Print("There are " + allEmployee.GetEmployeeCount() + " employees.");

            Debug.Print(Utilities.GetTextFromClass());

            using(Transaction t = new Transaction(doc))
            {
                FamilySymbol curFS = Utilities.GetFamilySymbolByName(doc, "Desk", @"60"" x 30""");
                List<SpatialElement> roomList = Utilities.GetAllRooms(doc);
                foreach (SpatialElement curRoom in roomList)
                {
                    LocationPoint roomLocation = curRoom.Location as LocationPoint;
                    XYZ roomPoint = roomLocation.Point;



                    FamilyInstance curFI = doc.Create.NewFamilyInstance(roomPoint, curFS, StructuralType.NonStructural);
                    double area = Utilities.GetParamValueAsDouble(curRoom, "Area");

                    Utilities.SetParamValue(curRoom, "Comments", "This is a comment");
                }
            }

            
            return Result.Succeeded;
        }
    }
}