using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Windows.Media.Imaging;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Reflection;
using Autodesk.Revit.DB.Plumbing;

namespace ArchSmarterUtils
{
    static class Utils
	{
		#region CSV and Text files
		public static List<string[]> ReadCSVFile(string filepath)
		{
			//create list for sheet information
			List<string[]> dataList = new List<string[]>();

			//create reader and read CSV file
			TextFieldParser myReader = new TextFieldParser(filepath);
			myReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
			myReader.SetDelimiters(",");

			//create variable to hold current row of CSV file
			string[] currentRow;

			//loop through data, read each line and put into sheet list array
			while (myReader.EndOfData)
			{
				currentRow = myReader.ReadFields();
				//add data to list
				dataList.Add(currentRow);
			}
			return dataList;
		}

		public static void WriteToTxtFile(string filePath, string fileContents)
		{
			using (StreamWriter writer = File.AppendText(filePath))
			{
				writer.WriteLine(fileContents);
			}
		}
		#endregion

		#region Curves
		public static Line OffsetCurve(Line l, double dist)
		{
			XYZ lineDirection = l.Direction;
			XYZ normal = XYZ.BasisZ.CrossProduct(lineDirection).Normalize();
			XYZ translation = normal.Multiply(dist);

			XYZ startPt = l.GetEndPoint(0).Add(translation);
			XYZ endPt = l.GetEndPoint(1).Add(translation);

			return Line.CreateBound(startPt, endPt);
		}

		public static List<Curve> SimplifyRoomCurves(List<Curve> roomCurves)
		{
			int segCount = roomCurves.Count;
			List<Curve> cl = new List<Curve>();

			// loop through boundaries, check if parallel, if so then create new segment
			for (int i = 0; i <= segCount - 1; i++)
			{
				Curve curSeg = roomCurves[i];
				Curve nextSeg = null;

				if (i == segCount - 1)
				{
					nextSeg = roomCurves[0];
				}
				else
				{
					nextSeg = roomCurves[i + 1];
				}

				XYZ sp = curSeg.GetEndPoint(0);
				XYZ ep = null;

				// check if lines are parallel - if so then join
				if (AreLinesParallel(curSeg, nextSeg))
				{
					// combine into new line
					if (nextSeg.GetEndPoint(1).Equals(sp))
					{
						ep = nextSeg.GetEndPoint(0);
					}
					else
					{
						ep = nextSeg.GetEndPoint(1);
					}
				}

				if (ep != null)
				{
					// create new line
					try
					{
						Line newLine = Line.CreateBound(sp, ep);
						cl.Add(newLine as Curve);
					}
					catch (Exception)
					{
						Debug.Print("Line too small to generate.");
					}
					i++;
				}
				else
				{
					cl.Add(curSeg);
				}
			}

			if (segCount == cl.Count())
			{
				// curves are at their most simplified state - return list
				return cl;
			}
			else
			{
				// recursive loop to simplifly curve list
				return SimplifyRoomCurves(cl);
			}
		}

		public static bool AreLinesParallel(Curve c1, Curve c2)
		{
			double dx1 = c1.GetEndPoint(1).X - c1.GetEndPoint(0).X;
			double dy1 = c1.GetEndPoint(1).Y - c1.GetEndPoint(0).Y;
			double dx2 = c2.GetEndPoint(1).X - c2.GetEndPoint(0).X;
			double dy2 = c2.GetEndPoint(1).Y - c2.GetEndPoint(0).Y;
			double cosAngle = Math.Abs((dx1 * dx2 + dy1 * dy2) / Math.Sqrt((dx1 * dx1 + dy1 * dy1) * (dx2 * dx2 + dy2 * dy2)));

			if (cosAngle != 1)
			{
				return false;
			}
			else
			{
				return true;
			}
		}

		#endregion

		#region Design options
		public static List<DesignOption> getAllDesignOptions(Document curDoc)
		{
			FilteredElementCollector curCol = new FilteredElementCollector(curDoc);
			curCol.OfCategory(BuiltInCategory.OST_DesignOptions);

			List<DesignOption> doList = new List<DesignOption>();
			foreach (DesignOption curOpt in curCol)
			{
				doList.Add(curOpt);
			}

			return doList;
		}

		public static DesignOption getDesignOptionByName(Document curDoc, string designOpt)
		{
			//get all design options
			List<DesignOption> doList = getAllDesignOptions(curDoc);

			foreach (DesignOption curOpt in doList)
			{
				if (curOpt.Name == designOpt)
				{
					return curOpt;
				}
			}

			return null;
		}
		#endregion

		#region DWG files
		public static int getACADColorFromRGB(int r, int g, int b)
		{
			//there's got to be a better way to do this!
			if (r == 0 && g == 0 && b == 0)
			{
				return 7;
			}
			else if (r == 255 && g == 0 && b == 0)
				return 1;
			else if (r == 255 && g == 255 && b == 0)
				return 2;
			else if (r == 0 && g == 255 && b == 0)
				return 3;
			else if (r == 0 && g == 255 && b == 255)
				return 4;
			else if (r == 0 && g == 0 && b == 255)
				return 5;
			else if (r == 255 && g == 0 && b == 255)
				return 6;
			else if (r == 255 && g == 255 && b == 255)
				return 7;
			else if (r == 65 && g == 65 && b == 65)
				return 8;
			else if (r == 128 && g == 128 && b == 128)
				return 9;
			else if (r == 255 && g == 0 && b == 0)
				return 10;
			else if (r == 255 && g == 170 && b == 170)
				return 11;
			else if (r == 189 && g == 0 && b == 0)
				return 12;
			else if (r == 189 && g == 126 && b == 126)
				return 13;
			else if (r == 129 && g == 0 && b == 0)
				return 14;
			else if (r == 129 && g == 86 && b == 86)
				return 15;
			else if (r == 104 && g == 0 && b == 0)
				return 16;
			else if (r == 104 && g == 69 && b == 69)
				return 17;
			else if (r == 79 && g == 0 && b == 0)
				return 18;
			else if (r == 79 && g == 53 && b == 53)
				return 19;
			else if (r == 255 && g == 63 && b == 0)
				return 20;
			else if (r == 255 && g == 191 && b == 170)
				return 21;
			else if (r == 189 && g == 46 && b == 0)
				return 22;
			else if (r == 189 && g == 141 && b == 126)
				return 23;
			else if (r == 129 && g == 31 && b == 0)
				return 24;
			else if (r == 129 && g == 96 && b == 86)
				return 25;
			else if (r == 104 && g == 25 && b == 0)
				return 26;
			else if (r == 104 && g == 78 && b == 69)
				return 27;
			else if (r == 79 && g == 19 && b == 0)
				return 28;
			else if (r == 79 && g == 59 && b == 53)
				return 29;
			else if (r == 255 && g == 127 && b == 0)
				return 30;
			else if (r == 255 && g == 212 && b == 170)
				return 31;
			else if (r == 189 && g == 94 && b == 0)
				return 32;
			else if (r == 189 && g == 157 && b == 126)
				return 33;
			else if (r == 129 && g == 64 && b == 0)
				return 34;
			else if (r == 129 && g == 107 && b == 86)
				return 35;
			else if (r == 104 && g == 52 && b == 0)
				return 36;
			else if (r == 104 && g == 86 && b == 69)
				return 37;
			else if (r == 79 && g == 39 && b == 0)
				return 38;
			else if (r == 79 && g == 66 && b == 53)
				return 39;
			else if (r == 255 && g == 191 && b == 0)
				return 40;
			else if (r == 255 && g == 234 && b == 170)
				return 41;
			else if (r == 189 && g == 141 && b == 0)
				return 42;
			else if (r == 189 && g == 173 && b == 126)
				return 43;
			else if (r == 129 && g == 96 && b == 0)
				return 44;
			else if (r == 129 && g == 118 && b == 86)
				return 45;
			else if (r == 104 && g == 78 && b == 0)
				return 46;
			else if (r == 104 && g == 95 && b == 69)
				return 47;
			else if (r == 79 && g == 59 && b == 0)
				return 48;
			else if (r == 79 && g == 73 && b == 53)
				return 49;
			else if (r == 255 && g == 255 && b == 0)
				return 50;
			else if (r == 255 && g == 255 && b == 170)
				return 51;
			else if (r == 189 && g == 189 && b == 0)
				return 52;
			else if (r == 189 && g == 189 && b == 126)
				return 53;
			else if (r == 129 && g == 129 && b == 0)
				return 54;
			else if (r == 129 && g == 129 && b == 86)
				return 55;
			else if (r == 104 && g == 104 && b == 0)
				return 56;
			else if (r == 104 && g == 104 && b == 69)
				return 57;
			else if (r == 79 && g == 79 && b == 0)
				return 58;
			else if (r == 79 && g == 79 && b == 53)
				return 59;
			else if (r == 191 && g == 255 && b == 0)
				return 60;
			else if (r == 234 && g == 255 && b == 170)
				return 61;
			else if (r == 141 && g == 189 && b == 0)
				return 62;
			else if (r == 173 && g == 189 && b == 126)
				return 63;
			else if (r == 96 && g == 129 && b == 0)
				return 64;
			else if (r == 118 && g == 129 && b == 86)
				return 65;
			else if (r == 78 && g == 104 && b == 0)
				return 66;
			else if (r == 95 && g == 104 && b == 69)
				return 67;
			else if (r == 59 && g == 79 && b == 0)
				return 68;
			else if (r == 73 && g == 79 && b == 53)
				return 69;
			else if (r == 127 && g == 255 && b == 0)
				return 70;
			else if (r == 212 && g == 255 && b == 170)
				return 71;
			else if (r == 94 && g == 189 && b == 0)
				return 72;
			else if (r == 157 && g == 189 && b == 126)
				return 73;
			else if (r == 64 && g == 129 && b == 0)
				return 74;
			else if (r == 107 && g == 129 && b == 86)
				return 75;
			else if (r == 52 && g == 104 && b == 0)
				return 76;
			else if (r == 86 && g == 104 && b == 69)
				return 77;
			else if (r == 39 && g == 79 && b == 0)
				return 78;
			else if (r == 66 && g == 79 && b == 53)
				return 79;
			else if (r == 63 && g == 255 && b == 0)
				return 80;
			else if (r == 191 && g == 255 && b == 170)
				return 81;
			else if (r == 46 && g == 189 && b == 0)
				return 82;
			else if (r == 141 && g == 189 && b == 126)
				return 83;
			else if (r == 31 && g == 129 && b == 0)
				return 84;
			else if (r == 96 && g == 129 && b == 86)
				return 85;
			else if (r == 25 && g == 104 && b == 0)
				return 86;
			else if (r == 78 && g == 104 && b == 69)
				return 87;
			else if (r == 19 && g == 79 && b == 0)
				return 88;
			else if (r == 59 && g == 79 && b == 53)
				return 89;
			else if (r == 0 && g == 255 && b == 0)
				return 90;
			else if (r == 170 && g == 255 && b == 170)
				return 91;
			else if (r == 0 && g == 189 && b == 0)
				return 92;
			else if (r == 126 && g == 189 && b == 126)
				return 93;
			else if (r == 0 && g == 129 && b == 0)
				return 94;
			else if (r == 86 && g == 129 && b == 86)
				return 95;
			else if (r == 0 && g == 104 && b == 0)
				return 96;
			else if (r == 69 && g == 104 && b == 69)
				return 97;
			else if (r == 0 && g == 79 && b == 0)
				return 98;
			else if (r == 53 && g == 79 && b == 53)
				return 99;
			else if (r == 0 && g == 255 && b == 63)
				return 100;
			else if (r == 170 && g == 255 && b == 191)
				return 101;
			else if (r == 0 && g == 189 && b == 46)
				return 102;
			else if (r == 126 && g == 189 && b == 141)
				return 103;
			else if (r == 0 && g == 129 && b == 31)
				return 104;
			else if (r == 86 && g == 129 && b == 96)
				return 105;
			else if (r == 0 && g == 104 && b == 25)
				return 106;
			else if (r == 69 && g == 104 && b == 78)
				return 107;
			else if (r == 0 && g == 79 && b == 19)
				return 108;
			else if (r == 53 && g == 79 && b == 59)
				return 109;
			else if (r == 0 && g == 255 && b == 127)
				return 110;
			else if (r == 170 && g == 255 && b == 212)
				return 111;
			else if (r == 0 && g == 189 && b == 94)
				return 112;
			else if (r == 126 && g == 189 && b == 157)
				return 113;
			else if (r == 0 && g == 129 && b == 64)
				return 114;
			else if (r == 86 && g == 129 && b == 107)
				return 115;
			else if (r == 0 && g == 104 && b == 52)
				return 116;
			else if (r == 69 && g == 104 && b == 86)
				return 117;
			else if (r == 0 && g == 79 && b == 39)
				return 118;
			else if (r == 53 && g == 79 && b == 66)
				return 119;
			else if (r == 0 && g == 255 && b == 191)
				return 120;
			else if (r == 170 && g == 255 && b == 234)
				return 121;
			else if (r == 0 && g == 189 && b == 141)
				return 122;
			else if (r == 126 && g == 189 && b == 173)
				return 123;
			else if (r == 0 && g == 129 && b == 96)
				return 124;
			else if (r == 86 && g == 129 && b == 118)
				return 125;
			else if (r == 0 && g == 104 && b == 78)
				return 126;
			else if (r == 69 && g == 104 && b == 95)
				return 127;
			else if (r == 0 && g == 79 && b == 59)
				return 128;
			else if (r == 53 && g == 79 && b == 73)
				return 129;
			else if (r == 0 && g == 255 && b == 255)
				return 130;
			else if (r == 170 && g == 255 && b == 255)
				return 131;
			else if (r == 0 && g == 189 && b == 189)
				return 132;
			else if (r == 126 && g == 189 && b == 189)
				return 133;
			else if (r == 0 && g == 129 && b == 129)
				return 134;
			else if (r == 86 && g == 129 && b == 129)
				return 135;
			else if (r == 0 && g == 104 && b == 104)
				return 136;
			else if (r == 69 && g == 104 && b == 104)
				return 137;
			else if (r == 0 && g == 79 && b == 79)
				return 138;
			else if (r == 53 && g == 79 && b == 79)
				return 139;
			else if (r == 0 && g == 191 && b == 255)
				return 140;
			else if (r == 170 && g == 234 && b == 255)
				return 141;
			else if (r == 0 && g == 141 && b == 189)
				return 142;
			else if (r == 126 && g == 173 && b == 189)
				return 143;
			else if (r == 0 && g == 96 && b == 129)
				return 144;
			else if (r == 86 && g == 118 && b == 129)
				return 145;
			else if (r == 0 && g == 78 && b == 104)
				return 146;
			else if (r == 69 && g == 95 && b == 104)
				return 147;
			else if (r == 0 && g == 59 && b == 79)
				return 148;
			else if (r == 53 && g == 73 && b == 79)
				return 149;
			else if (r == 0 && g == 127 && b == 255)
				return 150;
			else if (r == 170 && g == 212 && b == 255)
				return 151;
			else if (r == 0 && g == 94 && b == 189)
				return 152;
			else if (r == 126 && g == 157 && b == 189)
				return 153;
			else if (r == 0 && g == 64 && b == 129)
				return 154;
			else if (r == 86 && g == 107 && b == 129)
				return 155;
			else if (r == 0 && g == 52 && b == 104)
				return 156;
			else if (r == 69 && g == 86 && b == 104)
				return 157;
			else if (r == 0 && g == 39 && b == 79)
				return 158;
			else if (r == 53 && g == 66 && b == 79)
				return 159;
			else if (r == 0 && g == 63 && b == 255)
				return 160;
			else if (r == 170 && g == 191 && b == 255)
				return 161;
			else if (r == 0 && g == 46 && b == 189)
				return 162;
			else if (r == 126 && g == 141 && b == 189)
				return 163;
			else if (r == 0 && g == 31 && b == 129)
				return 164;
			else if (r == 86 && g == 96 && b == 129)
				return 165;
			else if (r == 0 && g == 25 && b == 104)
				return 166;
			else if (r == 69 && g == 78 && b == 104)
				return 167;
			else if (r == 0 && g == 19 && b == 79)
				return 168;
			else if (r == 53 && g == 59 && b == 79)
				return 169;
			else if (r == 0 && g == 0 && b == 255)
				return 170;
			else if (r == 170 && g == 170 && b == 255)
				return 171;
			else if (r == 0 && g == 0 && b == 189)
				return 172;
			else if (r == 126 && g == 126 && b == 189)
				return 173;
			else if (r == 0 && g == 0 && b == 129)
				return 174;
			else if (r == 86 && g == 86 && b == 129)
				return 175;
			else if (r == 0 && g == 0 && b == 104)
				return 176;
			else if (r == 69 && g == 69 && b == 104)
				return 177;
			else if (r == 0 && g == 0 && b == 79)
				return 178;
			else if (r == 53 && g == 53 && b == 79)
				return 179;
			else if (r == 63 && g == 0 && b == 255)
				return 180;
			else if (r == 191 && g == 170 && b == 255)
				return 181;
			else if (r == 46 && g == 0 && b == 189)
				return 182;
			else if (r == 141 && g == 126 && b == 189)
				return 183;
			else if (r == 31 && g == 0 && b == 129)
				return 184;
			else if (r == 96 && g == 86 && b == 129)
				return 185;
			else if (r == 25 && g == 0 && b == 104)
				return 186;
			else if (r == 78 && g == 69 && b == 104)
				return 187;
			else if (r == 19 && g == 0 && b == 79)
				return 188;
			else if (r == 59 && g == 53 && b == 79)
				return 189;
			else if (r == 127 && g == 0 && b == 255)
				return 190;
			else if (r == 212 && g == 170 && b == 255)
				return 191;
			else if (r == 94 && g == 0 && b == 189)
				return 192;
			else if (r == 157 && g == 126 && b == 189)
				return 193;
			else if (r == 64 && g == 0 && b == 129)
				return 194;
			else if (r == 107 && g == 86 && b == 129)
				return 195;
			else if (r == 52 && g == 0 && b == 104)
				return 196;
			else if (r == 86 && g == 69 && b == 104)
				return 197;
			else if (r == 39 && g == 0 && b == 79)
				return 198;
			else if (r == 66 && g == 53 && b == 79)
				return 199;
			else if (r == 191 && g == 0 && b == 255)
				return 200;
			else if (r == 234 && g == 170 && b == 255)
				return 201;
			else if (r == 141 && g == 0 && b == 189)
				return 202;
			else if (r == 173 && g == 126 && b == 189)
				return 203;
			else if (r == 96 && g == 0 && b == 129)
				return 204;
			else if (r == 118 && g == 86 && b == 129)
				return 205;
			else if (r == 78 && g == 0 && b == 104)
				return 206;
			else if (r == 95 && g == 69 && b == 104)
				return 207;
			else if (r == 59 && g == 0 && b == 79)
				return 208;
			else if (r == 73 && g == 53 && b == 79)
				return 209;
			else if (r == 255 && g == 0 && b == 255)
				return 210;
			else if (r == 255 && g == 170 && b == 255)
				return 211;
			else if (r == 189 && g == 0 && b == 189)
				return 212;
			else if (r == 189 && g == 126 && b == 189)
				return 213;
			else if (r == 129 && g == 0 && b == 129)
				return 214;
			else if (r == 129 && g == 86 && b == 129)
				return 215;
			else if (r == 104 && g == 0 && b == 104)
				return 216;
			else if (r == 104 && g == 69 && b == 104)
				return 217;
			else if (r == 79 && g == 0 && b == 79)
				return 218;
			else if (r == 79 && g == 53 && b == 79)
				return 219;
			else if (r == 255 && g == 0 && b == 191)
				return 220;
			else if (r == 255 && g == 170 && b == 234)
				return 221;
			else if (r == 189 && g == 0 && b == 141)
				return 222;
			else if (r == 189 && g == 126 && b == 173)
				return 223;
			else if (r == 129 && g == 0 && b == 96)
				return 224;
			else if (r == 129 && g == 86 && b == 118)
				return 225;
			else if (r == 104 && g == 0 && b == 78)
				return 226;
			else if (r == 104 && g == 69 && b == 95)
				return 227;
			else if (r == 79 && g == 0 && b == 59)
				return 228;
			else if (r == 79 && g == 53 && b == 73)
				return 229;
			else if (r == 255 && g == 0 && b == 127)
				return 230;
			else if (r == 255 && g == 170 && b == 212)
				return 231;
			else if (r == 189 && g == 0 && b == 94)
				return 232;
			else if (r == 189 && g == 126 && b == 157)
				return 233;
			else if (r == 129 && g == 0 && b == 64)
				return 234;
			else if (r == 129 && g == 86 && b == 107)
				return 235;
			else if (r == 104 && g == 0 && b == 52)
				return 236;
			else if (r == 104 && g == 69 && b == 86)
				return 237;
			else if (r == 79 && g == 0 && b == 39)
				return 238;
			else if (r == 79 && g == 53 && b == 66)
				return 239;
			else if (r == 255 && g == 0 && b == 63)
				return 240;
			else if (r == 255 && g == 170 && b == 191)
				return 241;
			else if (r == 189 && g == 0 && b == 46)
				return 242;
			else if (r == 189 && g == 126 && b == 141)
				return 243;
			else if (r == 129 && g == 0 && b == 31)
				return 244;
			else if (r == 129 && g == 86 && b == 96)
				return 245;
			else if (r == 104 && g == 0 && b == 25)
				return 246;
			else if (r == 104 && g == 69 && b == 78)
				return 247;
			else if (r == 79 && g == 0 && b == 19)
				return 248;
			else if (r == 79 && g == 53 && b == 59)
				return 249;
			else if (r == 51 && g == 51 && b == 51)
				return 250;
			else if (r == 80 && g == 80 && b == 80)
				return 251;
			else if (r == 105 && g == 105 && b == 105)
				return 252;
			else if (r == 130 && g == 130 && b == 130)
				return 253;
			else if (r == 190 && g == 190 && b == 190)
				return 254;
			else if (r == 255 && g == 255 && b == 255)
				return 255;
			return 0;
		}

		public static List<string> GetDWGLayers(Document curDoc, ImportInstance curLink)
		{
			List<string> layerList = new List<string>();
			Options curOptions = new Options();
			GeometryElement geoElement;
			GeometryInstance geoInstance;
			GeometryElement geoElem2;

			//get geometry from current link
			geoElement = curLink.get_Geometry(curOptions);

			//loop through geometry and put layers into list
			foreach (GeometryObject geoObject in geoElement)
			{
				//convert geoObject to geometry instance
				geoInstance = (GeometryInstance)geoObject;
				geoElem2 = geoInstance.GetInstanceGeometry();

				try
				{
					foreach (GeometryObject curObj in geoElem2)
					{
						//get object's ACAD color
						GraphicsStyle curGraphicStyle;
						curGraphicStyle = (GraphicsStyle)curDoc.GetElement(curObj.GraphicsStyleId);

						//add object to list
						if (layerList.Contains(curGraphicStyle.GraphicsStyleCategory.Name) == false)
						{
							layerList.Add(curGraphicStyle.GraphicsStyleCategory.Name);
						}
					}
				}
				catch (Exception ex)
				{
					Debug.Print("Error: " + ex.Message);
				}
			}
			//sort layerlist alphabetically
			layerList.Sort();

			return layerList;
		}
		#endregion

		#region Elements
		public static bool WasElemDemolished(Element curElem)
		{
			// Get the Phase Demolished property, and check whether it be null
			Phase phaseDemolished = curElem.Document.GetElement(curElem.DemolishedPhaseId) as Phase;
			if (null == phaseDemolished)
			{
				return false;
			}
			else
			{
				// Show the Phase Demolished name to the user.
				return true;
			}
		}

		public static void RotateElement(Document curDoc, Element curElem, double rotateAngle, XYZ rotationPoint)
		{
			//create axis line from rotation point
			Line axisLine = Line.CreateBound(new XYZ(rotationPoint.X, rotationPoint.Y, 0), new XYZ(rotationPoint.X, rotationPoint.Y, 1));

			//rotate element along axis line
			ElementTransformUtils.RotateElement(curDoc, curElem.Id, axisLine, rotateAngle);
		}
		#endregion

		#region Excel
		public static List<string[]> GetDataFromExcel(string excelFile, int numCols)
		{
			// open Excel and read data
			Excel.Application excelApp = new Excel.Application();
			Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
			Excel.Worksheet excelWs = (Excel.Worksheet)excelApp.Worksheets.Item[1];
			Excel.Range excelRng = (Excel.Range)excelWs.UsedRange;

			// get row count
			int rowCount = excelRng.Rows.Count;

			// create list for sheet data
			List<string[]> dataList = new List<string[]>();

			// loop through Excel data and put into data structure
			for (int i = 1; i <= rowCount; i++)
			{
				// create array for row data
				string[] sheetData = new String[numCols];

				// loop through columns and output data to array
				for (int j = 1; j <= numCols; j++)
				{
					// get Excel data
					Excel.Range cellData = (Excel.Range)excelWs.Cells[i, j];

					// add data to array
					sheetData[j - 1] = cellData.Value2.ToString();
				}

				// add sheet data to list
				dataList.Add(sheetData);
			}

			// close Excel and cleanup
			excelWb.Close();
			excelApp.Quit();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWb);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWs);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRng);

			// return list
			return dataList;
		}
		#endregion

		#region Families
		public static string GetFamilyNameFromInstance(FamilyInstance curElem)
		{
			return curElem.Symbol.FamilyName;
		}

		public static FamilyInstance InsertFamily(Document curDoc, Level curLevel, FamilySymbol curFam, XYZ insPoint)
		{
			//insert new family instance
			FamilyInstance nfInst = curDoc.Create.NewFamilyInstance(insPoint, curFam, curLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
			return nfInst;
		}

		public static FamilyInstance InsertFamilyToView(Document curDoc, View curView, FamilySymbol curFam, XYZ insPoint)
		{
			//insert new family instance
			FamilyInstance newFamInst = curDoc.Create.NewFamilyInstance(insPoint, curFam, curView);

			//return family instance
			return newFamInst;
		}
		#endregion

		#region Files and file paths
		public static string GetFolderPathFromFile(string filePath)
		{
			string outputPath;
			outputPath = Path.GetDirectoryName(filePath);

			return outputPath;
		}

		public static string GetLinkedDWGFilepath(Document curDoc, ImportInstance curDWG)
		{
			//get filepath
			ExternalFileReference exFileRef = ExternalFileUtils.GetExternalFileReference(curDoc, curDWG.GetTypeId());

			if (exFileRef != null)
			{
				string curFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(exFileRef.GetAbsolutePath());
				return curFilePath;
			}

			return null;
		}

		public static long GetFileSize(string curFilepath)
		{
			//get file's size (in bytes)
			FileInfo curFile = new FileInfo(curFilepath);
			return curFile.Length;
		}

		public static bool DoesFileExist(string sTestFile)
		{
			long lSize = -1;

			try
			{
				//Get the length of the file
				lSize = FileSystem.GetFileInfo(sTestFile).Length;
			}
			catch (Exception)
			{
				lSize = -1;
			}

			if (lSize > -1)
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		#endregion

		#region Levels
		public static Level CreateNewLevel(Document curDoc, string lName, double lElev)
		{
			//make sure level doesn't exist already
			Level newLevel = Collectors.GetLevelByName(curDoc, lName);
			if (newLevel == null)
			{
				//start transaction since the level doesn't exist yet
				using (Transaction curTrans = new Transaction(curDoc, "Create new level"))
				{
					if (curTrans.Start() == TransactionStatus.Started)
					{
						//create level
						newLevel = Level.Create(curDoc, lElev);
						newLevel.Name = lName;
					}

					//commit changes
					curTrans.Commit();
					return newLevel;
				}
			}
			else
			{
				//level already exist
				return newLevel;
			}
		}
		#endregion

		#region Lines
		public static CategoryNameMap GetAllLineStyles(Document curDoc)
		{
			Category curCat = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
			CategoryNameMap subCat = curCat.SubCategories;
			return subCat;
		}

		public static List<ModelCurve> GetAllModelCurves(Document curDoc)
		{
			FilteredElementCollector col = new FilteredElementCollector(curDoc);
			CurveElementFilter cFilter = new CurveElementFilter(CurveElementType.ModelCurve);

			List<ModelCurve> lList = new List<ModelCurve>();

			foreach (ModelCurve ccurve in col.WherePasses(cFilter))
			{
				lList.Add(ccurve);
			}

			return lList;
		}
		public static List<string> FilterLineStyles(Document curDoc, string ltFilter)
		{
			//get all revit linestyles
			List<string> curLines = GetAllLineStyleNames(curDoc);
			List<string> filterList = new List<string>();

			switch (ltFilter)
			{
				case "area":
					filterList.Add("<Area Boundary>");
					break;
				case "room":
					filterList.Add("<Room Separation>");
					break;
				case "space":
					filterList.Add("<Space Separation>");
					break;
				default:
					curLines.Remove("<Area Boundary>");
					curLines.Remove("<Room Separation>");
					curLines.Remove("<Space Separation>");
					filterList = curLines;
					break;
			}

			return filterList;
		}

		public static List<string> GetAllLineStyleNames(Document curDoc)
		{
			//get all linestyles
			CategoryNameMap tmpCatMap = Collectors.GetAllLineStyles(curDoc);
			List<string> lsList = new List<string>();

			//loop through category name map and put names into list
			foreach (Category curCat in tmpCatMap)
			{
				//exclude sketch and fabric sheets
				if ((curCat.Name != "<Sketch>") && (curCat.Name != "<Fabric Sheets>") && (curCat.Name != "<Fabric Envelope>") && (curCat.Name != "<Demolished>"))
				{
					lsList.Add(curCat.Name);
				}
			}

			//sort list
			lsList.Sort();
			return lsList;
		}

		public static Category CreateNewLinestyle(Document curDoc, string sName, Color nlColor, int nlWeight)
		{
			Category nlStyle = null;
			Category lineCat = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
			//create transaction
			using (Transaction curTrans = new Transaction(curDoc, "Create Linestyle"))
			{
				if (curTrans.Start() == TransactionStatus.Started)
				{
					nlStyle = curDoc.Settings.Categories.NewSubcategory(lineCat, sName);
					nlStyle.LineColor = nlColor;
					nlStyle.SetLineWeight(nlWeight, GraphicsStyleType.Projection);
				}
				//commit changes
				curTrans.Commit();
			}

			return nlStyle;
		}
		#endregion

		#region MEP
		public static int CreatePipesFromLines(Document curDoc, string sysType, string lineStyle, string pipeType, string selLevel)
		{
			int counter = 0;
			//get lines
			List<ModelCurve> ltc = new List<ModelCurve>();
			ltc = Collectors.GetModelLinesByStyle(curDoc, lineStyle);

			//get pipe type
			PipeType curPT = Collectors.GetPipeType(curDoc, pipeType);

			//get system types
			MEPSystemType curST = Collectors.GetMEPSystemTypeByName(curDoc, sysType);

			//get level 
			Level curLevel = Collectors.GetLevelByName(curDoc, selLevel);

			//create the pipes
			if (ltc.Count() > 0)
			{
				using (Transaction t = new Transaction(curDoc, "add pipes"))
				{
					if (t.Start() == TransactionStatus.Started)
					{
						//loop through lines and create pipes
						foreach (ModelCurve curLine in ltc)
						{
							//get line from model curve
							Curve tmpLine = curLine.GeometryCurve;

							if (tmpLine.IsBound == false)
							{
								continue;
							}

							XYZ startPt = tmpLine.GetEndPoint(0);
							XYZ endPt = tmpLine.GetEndPoint(1);

							//create new pipe
							try
							{
								Pipe newPipe = Pipe.Create(curDoc, curST.Id, curPT.Id, curLevel.Id, startPt, endPt);
								counter++;
							}
							catch (Exception ex)
							{
								//alert user - can't create pipe
								TaskDialog.Show("Error", "Could not create pipe.");
								Debug.Print(ex.Message);
							}
						}
					}
					//commit changes
					t.Commit();
				}
			}
			return counter;
		}
		#endregion

		#region Parameters
		public static Parameter GetParameterByName(Element curElem, string paramName)
		{
			foreach (Parameter curParam in curElem.Parameters)
			{
				if (curParam.Definition.Name.ToString() == paramName)
					return curParam;
			}

			return null;
		}

		public static string GetParameterValueString(Element curElem, string paramName)
		{
			Parameter curParam = GetParameterByName(curElem, paramName);

			if (curParam != null)
				return curParam.AsString();

			return string.Empty;

		}

		public static double GetParameterValueDouble(Element curElem, string paramName)
		{
			Parameter curParam = GetParameterByName(curElem, paramName);

			if (curParam != null)
				return curParam.AsDouble();

			return 0;
		}

		public static ElementId GetParameterValueID(Element curElem, string paramName)
		{
			Parameter curParam = GetParameterByName(curElem, paramName);

			if (curParam != null)
				return curParam.AsElementId();

			return ElementId.InvalidElementId;
		}

		public static bool SetParameterValue(Element curElem, string paramName, string value)
        {
			Parameter curParam = GetParameterByName(curElem, paramName);

			if (curParam != null)
            {
				curParam.Set(value);
				return true;
			}

			return false;
				
		}

		public static bool SetParameterValue(Element curElem, string paramName, double value)
		{
			Parameter curParam = GetParameterByName(curElem, paramName);

			if (curParam != null)
			{
				curParam.Set(value);
				return true;
			}

			return false;

		}

		// BIP = Built-In Parameter
		public static string GetBIPValueString(Element curElem, BuiltInParameter bip)
		{
			Parameter curParam = curElem.get_Parameter(bip);

			try
			{
				return curParam.AsString();
			}
			catch (Exception)
			{
				return null;
			}
		}

		public static void SetBIPValueString(Element curElem, BuiltInParameter bip, string newValue)
		{
			Parameter curParam = curElem.get_Parameter(bip);
			try
			{
				curParam.Set(newValue);
			}
			catch (Exception)
			{
				TaskDialog.Show("Error", "Could not change parameter value.");
			}
		}

		public static List<string> GetLinkedParameters(Document curDoc, string parameterName)
		{
			List<string> dpParams = new List<string>();

			//get all links in the current model
			List<RevitLinkInstance> RVTLinks = Collectors.GetAllRVTLinks(curDoc);

			//create list for parameter values
			List<string> paramList = new List<string>();

			//loop through each RVT link and look for DP categories
			foreach (RevitLinkInstance curLink in RVTLinks)
			{
				//get current link's document
				Document tmpDoc = curLink.GetLinkDocument();
				Debug.Print(tmpDoc.PathName);

				//Iterate over all elements, both symbols and model elements, and put them in the dictionary.
				ElementFilter curFilter = new LogicalOrFilter(new ElementIsElementTypeFilter(false), new ElementIsElementTypeFilter(true));
				FilteredElementCollector collector = new FilteredElementCollector(tmpDoc).WherePasses(curFilter);

				//loop through elements
				foreach (Element curElem in collector)
				{
					string curParamValue = GetParameterValueString(curElem, parameterName);

					if (curParamValue != null)
					{
						if (paramList.Contains(curParamValue) == false)
						{
							//add value to list
							paramList.Add(curParamValue);
						}
					}
				}
			}
			return paramList;
		}

		public static void OutputParams(Document curDoc, Element curElem)
		{
			//output parameter name and value for element
			foreach (Parameter curParam in curElem.Parameters)
			{
				Debug.Print(curParam.Definition.Name + " - " + curParam.AsString());
			}
		}

		#endregion

		#region Selection
		//prompt user to select point
		public static XYZ SelectPoint(UIApplication curUIA, string prompt)
		{
			XYZ curPt = curUIA.ActiveUIDocument.Selection.PickPoint(prompt);
			return curPt;
		}

		//prompt user to select an edge
		public static Element SelectEdge(UIApplication curUIA, string prompt)
		{
			Reference curRef = curUIA.ActiveUIDocument.Selection.PickObject(ObjectType.Edge, prompt);
			return curUIA.ActiveUIDocument.Document.GetElement(curRef.ElementId);
		}

		//prompt user to select linked DWG
		public static ImportInstance SelectDWG(UIApplication curUIA)
		{
			Element curElem = SelectElement(curUIA, "Select linked DWG file");
			if (curElem.GetType() == typeof(ImportInstance))
			{
				ImportInstance curLink = (ImportInstance)curElem;
				return curLink;
			}
			return null;
		}

		//prompt user to select a single element - provide current UIApplication, returns element
		public static Element SelectElement(UIApplication curUIA, string prompt)
		{
			Reference curRef;
			curRef = curUIA.ActiveUIDocument.Selection.PickObject(ObjectType.Element, prompt);
			return curUIA.ActiveUIDocument.Document.GetElement(curRef.ElementId);
		}

		//prompt user to select a wall element - provide current UIApplication, returns element
		public static Wall SelectWall(UIApplication curUIA)
		{
			Reference curRef;
			curRef = curUIA.ActiveUIDocument.Selection.PickObject(ObjectType.Element, "Select wall");

			Element curElem = curUIA.ActiveUIDocument.Document.GetElement(curRef.ElementId);

			if (curElem.GetType() == typeof(Wall))
			{
				return (Wall)curElem;
			}
			else
			{
				TaskDialog.Show("Error", "Please select a wall element");
			}
			return null;
		}

		//promt user to select multiple elements
		public static List<Element> SelectElements(UIApplication curUIA, string prompt)
		{
			IList<Reference> curRefs = curUIA.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, prompt);

			//loop through references and add to element list
			List<Element> curElems = new List<Element>();

			foreach (Reference tmp in curRefs)
			{
				curElems.Add(curUIA.ActiveUIDocument.Document.GetElement(tmp.ElementId));
			}
			return curElems;
		}

		public static Document SelectLinkedDoc(UIDocument uiDoc)
		{
			Document curDoc = uiDoc.Document;

			// prompt user to select linked file
			ElementId curElemID = uiDoc.Selection.PickObject(ObjectType.Element, "Select linked Revit file:").ElementId;

			// get element from element ID
			Element curElem = curDoc.GetElement(curElemID);

			// check that selected element is Revit link
			if (curElem.GetType() == typeof(RevitLinkInstance))
			{
				RevitLinkInstance curLink = (RevitLinkInstance)curElem;

				return curLink.GetLinkDocument();
			}
			return null;
		}
		#endregion

		#region Sheets
		public static bool DoesModelContainSheets(Document curDoc)
		{
			// get all sheets
			List<ViewSheet> sheets = Collectors.GetAllSheets(curDoc);

			try
			{
				// check sheet count
				if (sheets.Count() > 0)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
			catch (Exception)
			{
				return false;
			}
		}

		public static int DeleteAllSheets(Document curDoc)
		{
			int counter = 0;

			//get all sheets
			List<ViewSheet> vList = Collectors.GetAllSheets(curDoc);

			//create transaction
			using (Transaction t = new Transaction(curDoc, "Delete All Sheets"))
			{
				if (t.Start() == TransactionStatus.Started)
				{
					//loop through sheets - if sheet is not current then delete
					try
					{
						foreach (ViewSheet curView in vList)
						{
							if (curDoc.ActiveView.Id.Compare(curView.Id) != 0)
							{
								try
								{
									//delete sheet
									curDoc.Delete(curView.Id);
									//increment counter
									counter++;
								}
								catch (Exception ex)
								{
									Debug.Print(ex.Message);
								}
							}
						}
					}
					catch (Exception ex)
					{
						Debug.Print(ex.Message);
					}
				}
				//commit changes
				t.Commit();
				t.Dispose();
			}

			return counter;
		}


		#endregion

		#region Strings
		public static string CleanString(String strIn)
		{
			//Replace invalid characters with empty strings.
			try
			{
				return Regex.Replace(strIn, @"[^\w\ \.@%-]", "");
				//If we timeout when replacing invalid characters, we should return String.Empty.
			}
			catch (Exception ex)
			{
				Debug.Print(ex.Message);
				return string.Empty;
			}
		}
		#endregion

		#region Text Notes & Fonts
		public static TextNoteType CreateNewTextNoteType(Document curDoc, string name, string font, double textSize)
		{
			TextNoteType curTextNoteType;

			//get first text note type
			TextNoteType tmpTextNoteType = Collectors.GetAllTextNoteTypes(curDoc).First();

			//copy note type
			curTextNoteType = (TextNoteType)tmpTextNoteType.Duplicate(name);

			//change properties
			SetParameterValue(curTextNoteType, "Text Font", font);
			SetParameterValue(curTextNoteType, "Leader Arrowhead", "Arrow Filled 30 Degree");
			SetParameterValue(curTextNoteType, "Text Size", textSize);

			return curTextNoteType;
		}

		public static List<TextNote> GetAllTextNotes(Document curDoc)
		{
			FilteredElementCollector col = new FilteredElementCollector(curDoc);
			col.OfClass(typeof(TextNote));

			List<TextNote> noteList = new List<TextNote>();
			foreach (TextNote note in col)
			{
				noteList.Add(note);
			}

			return noteList;
		}

		public static List<TextNoteType> GetAllTextNoteTypes(Document curDoc)
		{
			List<TextNoteType> tnList = new List<TextNoteType>();

			FilteredElementCollector tsCol = new FilteredElementCollector(curDoc);
			tsCol.OfClass(typeof(TextNoteType));

			foreach (TextNoteType curStyle in tsCol)
			{
				tnList.Add(curStyle);
			}
			return tnList;
		}

		public static List<string> GetAllTextNoteFonts(Document curDoc)
		{
			List<string> fList = new List<string>();
			List<TextNoteType> tnList = GetAllTextNoteTypes(curDoc);

			//loop through text styles and add font to list
			foreach (TextNoteType curStyle in tnList)
			{
				string curFont = GetParameterValueString((Element)curStyle, "Text Font");
				if ((curFont != "") && (!fList.Contains(curFont)))
				{
					fList.Add(curFont);
				}

			}
			return fList;
		}

		public static List<DimensionType> GetAllDimStyles(Document curDoc)
		{
			List<DimensionType> dsList = new List<DimensionType>();
			//get all dimension fonts in current model file
			FilteredElementCollector dsCol = new FilteredElementCollector(curDoc);
			dsCol.OfClass(typeof(DimensionType));

			//loop through dim styles and add font to list
			foreach (DimensionType curStyle in dsCol)
			{
				dsList.Add(curStyle);
			}

			return dsList;
		}

		public static List<string> GetAllDimFonts(Document curDoc)
		{
			List<string> fList = new List<string>();
			//get all dimension fonts in current model file
			List<DimensionType> dsList = GetAllDimStyles(curDoc);

			//loop through dim styles and add font to list
			foreach (DimensionType curStyle in dsList)
			{
				string curFont = GetParameterValueString((Element)curStyle, "Dimension");
				//add to list
				if ((curFont != "") && (!fList.Contains(curFont)))
				{
					fList.Add(curFont);
				}
			}
			return fList;
		}

		public static List<TextElementType> GetAllLabelStyles(Document curDoc)
		{
			//get all label fonts in current model file
			FilteredElementCollector lsCol = new FilteredElementCollector(curDoc);
			lsCol.OfClass(typeof(TextElementType));

			//create list for label styles
			List<TextElementType> lsList = new List<TextElementType>();

			//loop through label styles and add font to list
			foreach (TextElementType curStyle in lsCol)
			{
				lsList.Add(curStyle);
			}

			return lsList;
		}

		public static List<string> GetAllLabelFonts(Document curDoc)
		{
			List<string> fList = new List<string>();

			//get all label fonts in current model file
			List<TextElementType> lsList = GetAllLabelStyles(curDoc);

			//loop through label styles and add font to list
			foreach (TextElementType curStyle in lsList)
			{
				string curFont = GetParameterValueString((Element)curStyle, "Label");
				if ((curFont != "") && (fList.Contains(curFont)))
				{
					fList.Add(curFont);
				}
			}

			return fList;
		}

		public static string GetFont(Element curElem)
		{
			string curFont = GetParameterValueString((Element)curElem, "Text Font");
			return curFont;
		}
		#endregion

		#region Worksets
		public static void CreateNewWorkset(Document curDoc, string wsName)
		{
			//create new workset
			using (Transaction curTrans = new Transaction(curDoc, "Create Workset"))
			{
				if (curTrans.Start() == TransactionStatus.Started)
				{
					Workset.Create(curDoc, wsName);
					curTrans.Commit();
				}
			}
		}
		public static Boolean RenameWorkset(Document curDoc, Workset curWS, string nName)
		{
			try
			{
				WorksetTable.RenameWorkset(curDoc, curWS.Id, nName);
				return true;
			}
			catch (Exception ex)
			{
				Debug.Print(ex.Message);
				return false;
			}
		}
		public static List<Workset> GetAllUserWorksets(Document curDoc)
		{
			FilteredWorksetCollector wsCol = new FilteredWorksetCollector(curDoc);
			wsCol.OfKind(WorksetKind.UserWorkset);

			List<Workset> wsList = new List<Workset>();

			foreach (Workset ws in wsCol)
			{
				wsList.Add(ws);
			}

			return wsList;
		}

		public static bool SetElementWorkset(Element curElem, WorksetId newWorksetID)
		{
			//change element's workset to new workset
			Parameter wsParam = null;

			try
			{
				wsParam = curElem.GetParameters("Workset").First();
			}
			catch (Exception ex)
			{
				Debug.Print(ex.Message);
			}

			if (wsParam != null)
			{
				wsParam.Set(newWorksetID.IntegerValue);
				return true;
			}

			return false;
		}

		public static string GetWorksetName(Document curDoc, ElementId worksetID)
		{
			//get all user created worksets in current file
			FilteredWorksetCollector worksetCollector = new FilteredWorksetCollector(curDoc);
			worksetCollector.OfKind(WorksetKind.UserWorkset);

			//loop through worksets and check if specified workset exists
			foreach (Workset curWorkset in worksetCollector)
			{
				if (curWorkset.Id.Equals(worksetID))
				{
					return curWorkset.Name;
				}
			}

			return null;
		}
		#endregion
	}
}