using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Reflection;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
namespace BTOOLS_PLOT
{
    public class settingGet
	{

		//'глобальное объявление переменых внутри этого файла
		//'***********************************************'
		//public static Document acDoc = Application.DocumentManager.MdiActiveDocument;
		//public static Database acCurDb = acDoc.Database;
		public static settingXML getSettingAll()
		{
			Document acDoc = Application.DocumentManager.MdiActiveDocument;
			Database acCurDb = acDoc.Database;
			settingXML setXML = new settingXML();
			//string pathDLL = new (System.Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath;
			System.Uri uri = new Uri(Assembly.GetExecutingAssembly().CodeBase);
			string pathDLL = uri.AbsolutePath;
			string pathDLLCode = Uri.UnescapeDataString(pathDLL);
			string fileXML = System.IO.Path.ChangeExtension(pathDLLCode, ".xml");

			XmlSerializer serilizer = new XmlSerializer(typeof(settingXML));
			if (File.Exists(fileXML)) {
				acDoc.Editor.WriteMessage("Загрузка настроек по пути:");
				acDoc.Editor.WriteMessage("\n");
				acDoc.Editor.WriteMessage(fileXML);
				acDoc.Editor.WriteMessage("\n");
				using (FileStream writter2 = new FileStream(fileXML, FileMode.Open))
				{
					setXML = (settingXML)serilizer.Deserialize(writter2);
				}
			}
			else
			{
				acDoc.Editor.WriteMessage("Создание настроек по пути:");
				acDoc.Editor.WriteMessage("\n");
				acDoc.Editor.WriteMessage(fileXML);
				acDoc.Editor.WriteMessage("\n");

				setXML.waysavetoext = ".ps";
				setXML.waysavetopdf = "D:\\plot2pdf\\PS\\";

				setXML.colorWB = "monochrome.ctb";
				setXML.colorRGB = "acad.ctb";
				setXML.specTextColor = "RGB";
                setXML.prefixSheet = "Лист ";
                setXML.connectorSheet = ". ";
                setXML.suffixSheet = "";
				setXML.swapNameAndNum = true;


                settingXML.setSheet fSheet = new settingXML.setSheet();
					
				fSheet.printPortrait = "PDF-portrait.pc3";
				fSheet.printLandscape = "PDF-portrait.pc3";
				fSheet.heightPlot = -1;
				fSheet.weightPlot = -1;
				fSheet.nameSheet = "A4";
				fSheet.height = 210;
				fSheet.weight = 297;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A3";
				fSheet.height = 297;
				fSheet.weight = 420;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A2";
				fSheet.height = 420;
				fSheet.weight = 594;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A1";
				fSheet.height = 594;
				fSheet.weight = 841;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A0";
				fSheet.height = 841;
				fSheet.weight = 1189;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x3";
				fSheet.height = 297;
				fSheet.weight = 630;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x4";
				fSheet.height = 297;
				fSheet.weight = 841;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x5";
				fSheet.height = 297;
				fSheet.weight = 1051;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x6";
				fSheet.height = 297;
				fSheet.weight = 1261;
				fSheet.heightPlot = 950;
				fSheet.weightPlot = 1261;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x7";
				fSheet.heightPlot = -1;
				fSheet.weightPlot = -1;
				fSheet.height = 297;
				fSheet.weight = 1471;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x8";
				fSheet.height = 297;
				fSheet.weight = 1682;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A4x9";
				fSheet.height = 297;
				fSheet.weight = 1892;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A3x3";
				fSheet.height = 420;
				fSheet.weight = 891;
				fSheet.heightPlot = -1;
				fSheet.weightPlot = -1;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A3x4";
				fSheet.heightPlot = -1;
				fSheet.weightPlot = -1;
				fSheet.height = 420;
				fSheet.weight = 1189;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A3x5";
				fSheet.height = 420;
				fSheet.weight = 1486;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A3x6";
				fSheet.height = 420;
				fSheet.weight = 1783;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A3x7";
				fSheet.height = 420;
				fSheet.weight = 2080;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A2x3";
				fSheet.height = 594;
				fSheet.weight = 1260;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A2x4";
				fSheet.height = 594;
				fSheet.weight = 1682;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A2x5";
				fSheet.height = 594;
				fSheet.weight = 2102;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A1x3";
				fSheet.height = 841;
				fSheet.weight = 1783;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A1x4";
				fSheet.height = 841;
				fSheet.weight = 2378;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A0x2";
				fSheet.height = 1189;
				fSheet.weight = 1682;
				setXML.listFS.Add(fSheet);
				fSheet.nameSheet = "A0x3";
				fSheet.height = 1189;
				fSheet.weight = 2523;
				setXML.listFS.Add(fSheet);

				using (FileStream writter = File.Open(fileXML, FileMode.OpenOrCreate))
				{
					serilizer.Serialize(writter, setXML);
				}		
			}
			return setXML;
		}
	}
}





