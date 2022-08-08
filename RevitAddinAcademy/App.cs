#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;
#endregion

namespace RevitAddinAcademy
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            // step 1: create ribbon tab
            try
            {
                a.CreateRibbonTab("Revit Add-in \nAcademy");

            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

            // step 2: create ribbon panel
            //RibbonPanel curPanel = a.CreateRibbonPanel("Test Tab", "Test Panel");
            RibbonPanel curPanel = CreateRibbonPanel(a, "Revit Add-in \nAcademy", "Revit Tools");
            // step 3: create button data instances
            PushButtonData pData1 = new PushButtonData("button1", "Delete Backups", GetAssemblyName(), "RevitAddinAcademy.cmdDeleteBackups");
            PushButtonData pData2 = new PushButtonData("button2", "Elements from Lines", GetAssemblyName(), "RevitAddinAcademy.cmdElementsFromLines");
            PushButtonData pData3 = new PushButtonData("button3", "Insert Furniture", GetAssemblyName(), "RevitAddinAcademy.CmdInsertFurniture");
            PushButtonData pData4 = new PushButtonData("button4", "Project Setup", GetAssemblyName(), "RevitAddinAcademy.cmdProjectSetup");
            PushButtonData pData5 = new PushButtonData("button5", "Walls from Lines", GetAssemblyName(), "RevitAddinAcademy.cmdWallsFromLines");
            PushButtonData pData6 = new PushButtonData("button6", "Google", GetAssemblyName(), "https://google.com");
            PushButtonData pData7 = new PushButtonData("button7", "Load Groups from one file to another", GetAssemblyName(), "RevitAddinAcademy.CmdLoadGroups");
            PushButtonData pData8 = new PushButtonData("button8", "Button 8", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData9 = new PushButtonData("button9", "Button 9", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData10 = new PushButtonData("button10", "Button 10", GetAssemblyName(), "RevitAddinAcademy.Command");




            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button 1");
            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "More Tools");

            // step 4: add images
            pData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_1_16);
            pData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_1_32);

            pData2.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_16);
            pData2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_32);

            pData3.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_16);
            pData3.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_32);
 
            pData4.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_16);
            pData4.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_32);

            pData5.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_16);
            pData5.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_32);

            pData6.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_16);
            pData6.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_32);

            pData7.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_16);
            pData7.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_32);
 
            pData8.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_16);
            pData8.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_32);
 
            pData9.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_16);
            pData9.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_32);
 
            pData10.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_16);
            pData10.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.QuinnEvans_Gradient_32);

            pbData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Michael_16);
            pbData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Michael_32);


            // step 5: add tooltop info

            pData1.ToolTip = "Delete Backups";
            pData2.ToolTip = "Create Elements from Lines";
            pData3.ToolTip = "Insert Furniture";
            pData4.ToolTip = "Project Setup";
            pData5.ToolTip = "Walls from Lines";
            pData6.ToolTip = "Button 6 tool tip";
            pData7.ToolTip = "Button 7 tool tip";
            pData8.ToolTip = "Button 8 tool tip";
            pData9.ToolTip = "Button 8 tool tip";
            pData10.ToolTip = "Button 10 tool tip";
            pbData1.ToolTip = "Group of tools";

            // step 6: create buttons

            PushButton b1 = curPanel.AddItem(pData1) as PushButton;
            PushButton b2 = curPanel.AddItem(pData2) as PushButton;

            curPanel.AddStackedItems(pData3, pData4, pData5);

            SplitButton splitButton1 = curPanel.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData6);
            splitButton1.AddPushButton(pData7);

            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(pData8);
            pulldownButton1.AddPushButton(pData9);
            pulldownButton1.AddPushButton(pData10);


            return Result.Succeeded;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach(RibbonPanel tmpPanel in a.GetRibbonPanels(tabName))
            {
                if(tmpPanel.Name == tabName)
                    return tmpPanel;
            }
            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);

            return returnPanel;
        }


        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        private BitmapImage BitmapToImageSource(System.Drawing.Bitmap bm)
        {
            using(MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }


        }
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
