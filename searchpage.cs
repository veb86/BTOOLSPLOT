using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;


namespace BTOOLS_PLOT
{
    //public static class AcDbLinqExtensionMethods
    //{
    //    /// <summary>
    //    /// Get all references to the given BlockTableRecord, including 
    //    /// references to anonymous dynamic BlockTableRecords.
    //    /// </summary>

    //    public static IEnumerable<BlockReference> GetBlockReferences(
    //       this BlockTableRecord btr,
    //       OpenMode mode = OpenMode.ForRead,
    //       bool directOnly = true)
    //    {

    //        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
    //        ed.WriteMessage("запустил");

    //        if (btr == null)
    //            throw new ArgumentNullException("btr");
    //        var tr = btr.Database.TransactionManager.TopTransaction;
    //        if (tr == null)
    //            throw new InvalidOperationException("No transaction");
    //        var ids = btr.GetBlockReferenceIds(directOnly, true);
    //        int cnt = ids.Count;
    //        for (int i = 0; i < cnt; i++)
    //        {
    //            yield return (BlockReference)
    //               tr.GetObject(ids[i], mode, false, false);
    //        }
    //        if (btr.IsDynamicBlock)
    //        {
    //            BlockTableRecord btr2 = null;
    //            var blockIds = btr.GetAnonymousBlockIds();
    //            cnt = blockIds.Count;
    //            for (int i = 0; i < cnt; i++)
    //            {
    //                btr2 = (BlockTableRecord)tr.GetObject(blockIds[i],
    //                   OpenMode.ForRead, false, false);
    //                ids = btr2.GetBlockReferenceIds(directOnly, true);
    //                int cnt2 = ids.Count;
    //                for (int j = 0; j < cnt2; j++)
    //                {
    //                    yield return (BlockReference)
    //                       tr.GetObject(ids[j], mode, false, false);
    //                }
    //            }
    //        }
    //    }
    //}
    public class searchpage
    {
        //Специальное имя определяющее уникальный блок
       // public const string specialName = "progSheetFormat";
        //Специальное имя определяющее уникальный блок имени
        public const string specialBlockName = "_progSheetFormat";
        //Специальное имя определяющее уникальный блок штамп имени
        public const string specialShtampName = "_shprogSheetFormat";

        //Список блоков со специальными уникальным именем
        

        //Структура и список блоков 
        public struct blockSpecialSheet
        {
            public BlockReference sheetsBlock;      // блок листа
            public BlockReference sheetsStamp;      // блок штампа
            public bool isPortrait;                 // это портрет?
            public string nameSheetCode;            // имя шифра листа
            public string nameSheetNum;             // имя номера листа
            public int hor;                         // размер горизонтального листа
            public int vert;                        // размер вертикального листа
            public Point3d LT;                      // верхний левый угол
            public Point3d BR;                      // нижний правый угол
            public Double mashtab;
           // public int height;
           // public int weight;
            public string printPortrait;
            public string printLandscape;
            public string nameSheet;
            public int weightPlot;
            public int heightPlot;
            public bool isRGB;                 // печатаем в цвете?


            //public string printPortrait;
            //public string printLandscape;
            //public Double weightPlot;
            //public Double heightPlot;
        }
        //public List<blockSpecialSheet> listSheet = new List<blockSpecialSheet>();
        public static ObjectId CreateLayer(String layerName)
        {
            ObjectId layerId;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                // Проверяем нет ли еще слоя с таким именем в чертеже
                if (lt.Has(layerName))
                {
                    layerId = lt[layerName];
                }
                else
                {
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName; // Задаем имя слоя
                    ltr.IsPlottable = false;
                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 255);
                    layerId = lt.Add(ltr);
                    trans.AddNewlyCreatedDBObject(ltr, true);
                }
                trans.Commit();
            }
            return layerId;
        }

        public static void AddLine(Point3d LLTT, Point3d BBRR)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord currentSpace = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                

                // Create a line that starts at 5,5 and ends at 12,3
                using (Line acLine = new Line(LLTT, BBRR))
                {
                    acLine.LayerId = CreateLayer("!!!printerLayerTemp");
                    acLine.LineWeight = LineWeight.LineWeight100;
                    acLine.Color=Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 256);
                    // Add the new object to the block table record and the transaction
                    currentSpace.AppendEntity(acLine);
                    acTrans.AddNewlyCreatedDBObject(acLine, true);
                }

                // Save the new object to the database
                acTrans.Commit();
            }
        }

        public static bool isShtamInSheet(Point3d iLT, Point3d iBR, Point3d ibShtamp)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            //acDoc.Editor.WriteMessage("\n" + " -внутри штамп???- " + "\n");
            //acDoc.Editor.WriteMessage(" -корд штампа- Х- " + ibShtamp.X + " -У- " + ibShtamp.Y + "\n");
            //acDoc.Editor.WriteMessage(" -корд LT- Х- " + iLT.X + " -У- " + iLT.Y + "\n");
            //acDoc.Editor.WriteMessage(" -корд BR- Х- " + iBR.X + " -У- " + iBR.Y + "\n");
            if (iLT.X < ibShtamp.X && ibShtamp.X < iBR.X)  {
                if (iLT.Y > ibShtamp.Y && ibShtamp.Y > iBR.Y)
                {
                    //acDoc.Editor.WriteMessage(" -штамп_ВНУТРИ- ");
                    return true;
                } else { return false; }
            } else { return false; }
        }

        public static bool isRGBSheet(Point3d iLT, Point3d iBR, Scale3d masht, settingXML sXML)
        {
            // Get the current document editor
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
           
            Point3d zonaRGBTL = new Point3d(iLT.X, (iBR.Y + 5*masht.Y), 0);
            Point3d zonaRGBBR = new Point3d((iLT.X+20 * masht.X), iBR.Y, 0);
            
            bool isRGB = false;
            
            //acDoc.Editor.WriteMessage(masht.X.ToString() + " -текст- " + masht.Y.ToString() + " - " + masht.Z.ToString());

            // Create a crossing window from (2,2,0) to (10,8,0)
            // Start a transaction

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                PromptSelectionResult acSSPrompt;
                acSSPrompt = acDocEd.SelectWindow(zonaRGBTL, zonaRGBBR);

                // If the prompt status is OK, objects were selected
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;
                    // Step through the objects in the selection set
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            MText acMText = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as MText;
                            if (acMText != null)
                                if (acMText.Text == sXML.specTextColor) 
                                    isRGB = true;

                            DBText ntext = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as DBText;
                            if (ntext != null)
                                if (ntext.TextString == sXML.specTextColor)
                                    isRGB = true;
                        }
                    }
                };

                //acDoc.Editor.WriteMessage(" -норм листа- Х- " + bShtamp.Position.X + " -У- " + bShtamp.Position.Y + "\n");
                //acDoc.Editor.WriteMessage(" -корд LT- Х- " + zonaNumTL.X + " -У- " + zonaNumTL.Y + "\n");
                //acDoc.Editor.WriteMessage(" -корд BR- Х- " + zonaNumBR.X + " -У- " + zonaNumBR.Y + "\n");

            }

            return isRGB;
        }

        public static string getNameSheetCode(BlockReference bShtamp, settingXML sXML)
        {
            // Get the current document editor
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
            string strCode = "";
            string strCodeFirst = "";
            string strCodeSecond = "";
            string[] splitShtamp = bShtamp.Name.Split('_');

            Point3d zonaCodeTL = new Point3d(Convert.ToInt32(splitShtamp[1]), Convert.ToInt32(splitShtamp[2]), 0);
            Point3d zonaCodeBR = new Point3d(Convert.ToInt32(splitShtamp[3]), Convert.ToInt32(splitShtamp[4]), 0);

            Point3d zonaNumTL = new Point3d(Convert.ToInt32(splitShtamp[5]), Convert.ToInt32(splitShtamp[6]), 0);
            Point3d zonaNumBR = new Point3d(Convert.ToInt32(splitShtamp[7]), Convert.ToInt32(splitShtamp[8]), 0);

            Scale3d masht = bShtamp.ScaleFactors;

            //acDoc.Editor.WriteMessage(masht.X.ToString() + " -текст- " + masht.Y.ToString() + " - " + masht.Z.ToString());

            // Create a crossing window from (2,2,0) to (10,8,0)
            // Start a transaction

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                // ИЩЕМ ИМЯ или ШИФР
                BlockReference blkRef = acTrans.GetObject(bShtamp.ObjectId, OpenMode.ForRead) as BlockReference;
                if (blkRef != null)
                {
                    BlockTableRecord btr = acTrans.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId btrObj in btr)
                    {
                        //acDoc.Editor.WriteMessage("\n" + btrObj.ToString() + "\n");

                        //************* Оброботка объекта если это МТЕКСТ
                        MText acRefMText = acTrans.GetObject(btrObj, OpenMode.ForRead) as MText;
                        if (acRefMText != null)
                        {
                            Point3d acMTextCoord = acRefMText.Location;
                            //acDoc.Editor.WriteMessage("\n" + " -внутри шифра кодов- " + "\n");
                            //acDoc.Editor.WriteMessage(" -корд шифра- Х- " + acMTextCoord.X + " -У- " + acMTextCoord.Y + "\n");
                            //acDoc.Editor.WriteMessage(" -корд LT- Х- " + zonaCodeTL.X + " -У- " + zonaCodeTL.Y + "\n");
                            //acDoc.Editor.WriteMessage(" -корд BR- Х- " + zonaCodeBR.X + " -У- " + zonaCodeBR.Y + "\n");

                            if (-zonaCodeTL.X < acMTextCoord.X && acMTextCoord.X < -zonaCodeBR.X)
                            {
                                if (zonaCodeTL.Y > acMTextCoord.Y && acMTextCoord.Y > zonaCodeBR.Y)
                                {
                                    strCodeFirst = acRefMText.Text;
                                    //acDoc.Editor.WriteMessage(" -штамп_шифр- ");
                                    //acDoc.Editor.WriteMessage(acRefMText.Text);
                                }
                            }
                        }

                        //************* Оброботка объекта если это однострочный текст
                        DBText ntext = acTrans.GetObject(btrObj, OpenMode.ForWrite) as DBText;
                        if (ntext != null)
                        {
                            Point3d acnTextCoord = ntext.Position;
                            if (-zonaCodeTL.X < acnTextCoord.X && acnTextCoord.X < -zonaCodeBR.X)
                            {
                                if (zonaCodeTL.Y > acnTextCoord.Y && acnTextCoord.Y > zonaCodeBR.Y)
                                {
                                    strCodeFirst = ntext.TextString;
                                }
                            }
                        }
                    };
                    acDoc.Editor.WriteMessage("\n");
                }

                //string strCode = "";

                //********************* ПОИСК ИМЕНИ ВНУТРИ БЛОКА

                bool noCode = true;
                BlockTableRecord currentSpace = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                ////****ИЩЕМ что то, переписываю место поиска уже забыл что искал
                ///
                //noCode = true;
                foreach (ObjectId entId in currentSpace)
                {
                    if (entId.ObjectClass == RXClass.GetClass(typeof(MText)))
                    { // если это не блок то пропускаем данный элемент

                        MText mtobj = acTrans.GetObject(entId, OpenMode.ForRead) as MText;
                        if (isShtamInSheet(new Point3d(bShtamp.Position.X - zonaCodeTL.X * masht.X, bShtamp.Position.Y + zonaCodeTL.Y * masht.Y, 0), new Point3d(bShtamp.Position.X - zonaCodeBR.X * masht.X, bShtamp.Position.Y + zonaCodeBR.Y * masht.Y, 0), mtobj.Location))
                        {
                            strCodeFirst = strCodeFirst + mtobj.Text;
                            //strCode = "Лиsdfsdfsdfsdfsdfsdfsdfст " + strCode + ". " + mtobj.Text;
                            noCode = false;
                        }
                    }

                    if (entId.ObjectClass == RXClass.GetClass(typeof(DBText)))
                    { // если это не блок то пропускаем данный элемент

                        DBText dbobj = acTrans.GetObject(entId, OpenMode.ForRead) as DBText;
                        if (isShtamInSheet(new Point3d(bShtamp.Position.X - zonaCodeTL.X * masht.X, bShtamp.Position.Y + zonaCodeTL.Y * masht.Y, 0), new Point3d(bShtamp.Position.X - zonaCodeBR.X * masht.X, bShtamp.Position.Y + zonaCodeBR.Y * masht.Y, 0), dbobj.Position))
                        {
                            //strCode = strCode + dbobj.TextString;
                            strCodeFirst = strCodeFirst + dbobj.TextString;
                            noCode = false;
                        }


                    }
                }

                ///***ИЩЕМ ВТОРУЮ ЧАСТЬ ТЕКСТА А ИМЕННО НОМЕР СТАРНИЦЫ, НОМЕР СТРАНИЦЫ НЕ МОЖЕТ БЫТЬ В БЛОКЕ
                ///
                ///
                bool noNum = true;
                foreach (ObjectId entId in currentSpace)
                {
                    if (entId.ObjectClass == RXClass.GetClass(typeof(MText)))
                    { // если это не блок то пропускаем данный элемент

                        MText mtobj = acTrans.GetObject(entId, OpenMode.ForRead) as MText;
                        if (isShtamInSheet(new Point3d(bShtamp.Position.X - zonaNumTL.X * masht.X, bShtamp.Position.Y + zonaNumTL.Y * masht.Y, 0), new Point3d(bShtamp.Position.X - zonaNumBR.X * masht.X, bShtamp.Position.Y + zonaNumBR.Y * masht.Y, 0), mtobj.Location))
                        {
                            strCodeSecond = mtobj.Text;
                            noNum = false;
                        }
                    }

                    if (entId.ObjectClass == RXClass.GetClass(typeof(DBText)))
                    { // если это не блок то пропускаем данный элемент

                        DBText dbobj = acTrans.GetObject(entId, OpenMode.ForRead) as DBText;
                        if (isShtamInSheet(new Point3d(bShtamp.Position.X - zonaNumTL.X * masht.X, bShtamp.Position.Y + zonaNumTL.Y * masht.Y, 0), new Point3d(bShtamp.Position.X - zonaNumBR.X * masht.X, bShtamp.Position.Y + zonaNumBR.Y * masht.Y, 0), dbobj.Position))
                        {
                            strCodeSecond = dbobj.TextString;
                            noNum = false;
                        }
                    }
                }
            }

            //acDoc.Editor.WriteMessage("Дополнительная часть шифра не обнаружена!");
            strCodeFirst = strCodeFirst.Replace("\r\n", string.Empty);
            strCodeSecond = strCodeSecond.Replace("\r\n", string.Empty);
            if (strCodeFirst == "")
            {
                strCodeFirst = "NONAME";
            }
            acDoc.Editor.WriteMessage("   -1я часть текста (имя или шифр):" + strCodeFirst + "\n");
            if (strCodeSecond == "")
            {
                //acDoc.Editor.WriteMessage("Номер листа не обнаружен!");
                strCodeSecond = "000";
            }
            acDoc.Editor.WriteMessage("   -2я часть текста (номер):" + strCodeSecond + "\n");


            if (sXML.swapNameAndNum) 
            {
                strCode = sXML.prefixSheet + strCodeSecond + sXML.connectorSheet + strCodeFirst + sXML.suffixSheet;
                acDoc.Editor.WriteMessage("   -префикс + 2я часть текста + разъеденитель + 1я часть текста + суффикс (swapNameAndNum=true)" + "\n");
            }
            else {
                strCode = sXML.prefixSheet + strCodeFirst + sXML.connectorSheet + strCodeSecond + sXML.suffixSheet;
                acDoc.Editor.WriteMessage("   -префикс + 1я часть текста + разъеденитель + 2я часть текста + суффикс (swapNameAndNum=false)" + "\n");
            }
            acDoc.Editor.WriteMessage("   -префикс, разъеденитель, суффикс, swapNameAndNum меняются в настройках" + "\n");
            acDoc.Editor.WriteMessage("***-полученное имя листа:" + strCode + "\n");
            return strCode;
        }


        public static string getNameSpecNameSheetCode(BlockReference bSheet, int hor, int vert, settingXML sXML)
        {
            // Get the current document editor
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

            Point3d stSheet = bSheet.Position;
            string strCodeFirst = "";
            string strCodeSecond = "";
            Scale3d mashtab = bSheet.ScaleFactors;
            //acDoc.Editor.WriteMessage(mashtab.X.ToString() + " - внутри - " + mashtab.Y.ToString() + " - " + mashtab.Z.ToString());
            Point3d BRSheet = bSheet.Position;
            Point3d LTSheet = new Point3d(BRSheet.X - hor * mashtab.X, BRSheet.Y + vert * mashtab.Y, BRSheet.Z);
            Point3d zonaSpecLT = new Point3d(BRSheet.X - 5 * mashtab.X, LTSheet.Y, BRSheet.Z);

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {


                string strCode = "";
                bool noCode = true;

                //Выбираем все элементы на чертеже
                //using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                //{
                //acDoc.Editor.WriteMessage("BRSheet - " + BRSheet.X.ToString() + " - " + BRSheet.Y.ToString());
                //acDoc.Editor.WriteMessage("zonaSpecLT - " + zonaSpecLT.X.ToString() + " - " + zonaSpecLT.Y.ToString());
                //AddLine(BRSheet, zonaSpecLT);

                BlockTableRecord currentSpace = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId entId in currentSpace)
                {
                    //acDoc.Editor.WriteMessage("ObjectId - " + entId.ObjectClass.Name + " ---------- ");

                    if (entId.ObjectClass == RXClass.GetClass(typeof(MText)))
                    { // если это не блок то пропускаем данный элемент
                        ////acDoc.Editor.WriteMessage("Мтекст - ");
                        MText mtobj = acTrans.GetObject(entId, OpenMode.ForRead) as MText;
                        //acDoc.Editor.WriteMessage("mtobj.Location - " + mtobj.Location.X.ToString() + " - " + mtobj.Location.Y.ToString());
                        if (isShtamInSheet(zonaSpecLT, BRSheet, mtobj.Location))
                        {
                            //acDoc.Editor.WriteMessage("mtobj - " + mtobj.Location.X.ToString() + " - " + mtobj.Location.Y.ToString());
                            strCode = mtobj.Text;
                            //acDoc.Editor.WriteMessage("mtobj - " + strCode);
                            if (strCode.IndexOf(specialShtampName, 0) > -1)
                            {
                                noCode = false;
                            };
                        }

                    }
                    if (entId.ObjectClass == RXClass.GetClass(typeof(DBText)))
                    { // если это не блок то пропускаем данный элемент

                        DBText dbobj = acTrans.GetObject(entId, OpenMode.ForRead) as DBText;
                        //acDoc.Editor.WriteMessage("одинтекст - ");
                        if (isShtamInSheet(zonaSpecLT, BRSheet, dbobj.Position))
                        {
                            //acDoc.Editor.WriteMessage("dbobj - " + dbobj.Position.X.ToString() + " - " + dbobj.Position.Y.ToString());
                            strCode = dbobj.TextString;
                            //acDoc.Editor.WriteMessage("dbobj - " + strCode);
                            if (strCode.IndexOf(specialShtampName, 0) > -1)
                            {
                                noCode = false;
                            };

                        }

                    }
                }

                if (noCode==false)
                //{
                //    strCode = "NONAME";
                //}
                //else
                {

                    string[] splitShtamp = strCode.Split('_');

                    Point3d zonaCodeTL = new Point3d(Convert.ToInt32(splitShtamp[1]), Convert.ToInt32(splitShtamp[2]), 0);
                    Point3d zonaCodeBR = new Point3d(Convert.ToInt32(splitShtamp[3]), Convert.ToInt32(splitShtamp[4]), 0);

                    Point3d zonaNumTL = new Point3d(Convert.ToInt32(splitShtamp[5]), Convert.ToInt32(splitShtamp[6]), 0);
                    Point3d zonaNumBR = new Point3d(Convert.ToInt32(splitShtamp[7]), Convert.ToInt32(splitShtamp[8]), 0);

                    strCode = "";
                    // смотрим код шифр штампа 
                    //PromptSelectionResult acSSPrompt;


                    //  BlockTableRecord currentSpace = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    ////****ИЩЕМ что то, переписываю место поиска уже забыл что искал
                    ///
                    noCode = true;
                    foreach (ObjectId entId in currentSpace)
                    {
                        if (entId.ObjectClass == RXClass.GetClass(typeof(MText)))
                        { // если это не блок то пропускаем данный элемент

                            MText mtobj = acTrans.GetObject(entId, OpenMode.ForRead) as MText;
                            if (isShtamInSheet(new Point3d(BRSheet.X - zonaCodeTL.X * mashtab.X, BRSheet.Y + zonaCodeTL.Y * mashtab.Y, 0), new Point3d(BRSheet.X - zonaCodeBR.X * mashtab.X, BRSheet.Y + zonaCodeBR.Y * mashtab.Y, 0), mtobj.Location))
                            {
                                strCodeFirst = mtobj.Text;
                                noCode = false;
                            }
                        }

                        if (entId.ObjectClass == RXClass.GetClass(typeof(DBText)))
                        { // если это не блок то пропускаем данный элемент

                            DBText dbobj = acTrans.GetObject(entId, OpenMode.ForRead) as DBText;
                            if (isShtamInSheet(new Point3d(BRSheet.X - zonaCodeTL.X * mashtab.X, BRSheet.Y + zonaCodeTL.Y * mashtab.Y, 0), new Point3d(BRSheet.X - zonaCodeBR.X * mashtab.X, BRSheet.Y + zonaCodeBR.Y * mashtab.Y, 0), dbobj.Position))
                            {
                                strCodeFirst = dbobj.TextString;
                                noCode = false;
                            }


                        }
                    }

                    ////****ИЩЕМ ЕЩЕ что то, переписываю место поиска уже забыл что искал
                    ///
                    bool noNum = true;
                    foreach (ObjectId entId in currentSpace)
                    {
                        if (entId.ObjectClass == RXClass.GetClass(typeof(MText)))
                        { // если это не блок то пропускаем данный элемент

                            MText mtobj = acTrans.GetObject(entId, OpenMode.ForRead) as MText;
                            if (isShtamInSheet(new Point3d(BRSheet.X - zonaNumTL.X * mashtab.X, BRSheet.Y + zonaNumTL.Y * mashtab.Y, 0), new Point3d(BRSheet.X - zonaNumBR.X * mashtab.X, BRSheet.Y + zonaNumBR.Y * mashtab.Y, 0), mtobj.Location))
                            {
                                strCodeSecond = mtobj.Text;
                                noNum = false;
                            }
                        }

                        if (entId.ObjectClass == RXClass.GetClass(typeof(DBText)))
                        { // если это не блок то пропускаем данный элемент

                            DBText dbobj = acTrans.GetObject(entId, OpenMode.ForRead) as DBText;
                            if (isShtamInSheet(new Point3d(BRSheet.X - zonaNumTL.X * mashtab.X, BRSheet.Y + zonaNumTL.Y * mashtab.Y, 0), new Point3d(BRSheet.X - zonaNumBR.X * mashtab.X, BRSheet.Y + zonaNumBR.Y * mashtab.Y, 0), dbobj.Position))
                            {
                                strCodeSecond = dbobj.TextString;
                                noNum = false;
                            }
                        }
                    }
                }

                strCodeFirst = strCodeFirst.Replace("\r\n", string.Empty);
                strCodeSecond = strCodeSecond.Replace("\r\n", string.Empty);

                //acDoc.Editor.WriteMessage("Дополнительная часть шифра не обнаружена!");
                if (strCodeFirst == "")
                {
                    strCodeFirst = "NONAME";
                }
                acDoc.Editor.WriteMessage("   -1я часть текста (имя или шифр):" + strCodeFirst + "\n");
                if (strCodeSecond == "")
                {
                    //acDoc.Editor.WriteMessage("Номер листа не обнаружен!");
                    strCodeSecond = "000";
                }
                acDoc.Editor.WriteMessage("   -2я часть текста (номер):" + strCodeSecond + "\n");


                if (sXML.swapNameAndNum)
                {
                    strCode = sXML.prefixSheet + strCodeSecond + sXML.connectorSheet + strCodeFirst + sXML.suffixSheet;
                    acDoc.Editor.WriteMessage("   -префикс + 2я часть текста + разъеденитель + 1я часть текста + суффикс (swapNameAndNum=true)" + "\n");
                }
                else
                {
                    strCode = sXML.prefixSheet + strCodeFirst + sXML.connectorSheet + strCodeSecond + sXML.suffixSheet;
                    acDoc.Editor.WriteMessage("   -префикс + 1я часть текста + разъеденитель + 2я часть текста + суффикс (swapNameAndNum=false)" + "\n");
                }
                acDoc.Editor.WriteMessage("   -префикс, разъеденитель, суффикс, swapNameAndNum меняются в настройках" + "\n");
                acDoc.Editor.WriteMessage("***-полученное имя листа:" + strCode + "\n");
                return strCode;
            }
        }


        //public static string getNameSpecNameSheetCode(BlockReference bSheet, int hor, int vert)
        //{
        //    // Get the current document editor
        //    Document acDoc = Application.DocumentManager.MdiActiveDocument;
        //    Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
        //    Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

        //    Point3d stSheet = bSheet.Position;

        //    Scale3d mashtab = bSheet.ScaleFactors;
        //    //acDoc.Editor.WriteMessage(mashtab.X.ToString() + " - внутри - " + mashtab.Y.ToString() + " - " + mashtab.Z.ToString());
        //    Point3d BRSheet = bSheet.Position;
        //    Point3d LTSheet = new Point3d(BRSheet.X - hor * mashtab.X, BRSheet.Y + vert * mashtab.Y, BRSheet.Z);
        //    Point3d zonaSpecLT = new Point3d(BRSheet.X - 5 * mashtab.X, LTSheet.Y, BRSheet.Z);

        //    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        //    {


        //        string strCode = "";
        //        PromptSelectionResult acSSPrompt;
        //        acSSPrompt = acDocEd.SelectWindow(BRSheet, zonaSpecLT);
        //        bool noCode = true;
        //        // If the prompt status is OK, objects were selected
        //        if (acSSPrompt.Status == PromptStatus.OK)
        //        {
        //            SelectionSet acSSet = acSSPrompt.Value;
        //            // Step through the objects in the selection set

        //            foreach (SelectedObject acSSObj in acSSet)
        //            {
        //                if (acSSObj != null)
        //                {
        //                    MText acMText = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as MText;
        //                    if (acMText != null)
        //                    {
        //                        strCode = acMText.Text;
        //                        if (strCode.IndexOf(specialShtampName, 0) > -1)
        //                        {
        //                            noCode = false;
        //                        };
        //                    }
        //                    DBText ntext = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as DBText;
        //                    if (ntext != null)
        //                    {
        //                        strCode = ntext.TextString;
        //                        if (strCode.IndexOf(specialShtampName, 0) > -1)
        //                        {
        //                            noCode = false;
        //                        };
        //                        noCode = false;
        //                    }
        //                }
        //            }

        //        };

        //        if (noCode)
        //        {
        //            strCode = "NONAME";
        //        }
        //        else
        //        {

        //            string[] splitShtamp = strCode.Split('_');

        //            Point3d zonaCodeTL = new Point3d(Convert.ToInt32(splitShtamp[1]), Convert.ToInt32(splitShtamp[2]), 0);
        //            Point3d zonaCodeBR = new Point3d(Convert.ToInt32(splitShtamp[3]), Convert.ToInt32(splitShtamp[4]), 0);

        //            Point3d zonaNumTL = new Point3d(Convert.ToInt32(splitShtamp[5]), Convert.ToInt32(splitShtamp[6]), 0);
        //            Point3d zonaNumBR = new Point3d(Convert.ToInt32(splitShtamp[7]), Convert.ToInt32(splitShtamp[8]), 0);

        //            strCode = "";
        //            // смотрим код шифр штампа 
        //            //PromptSelectionResult acSSPrompt;
        //            acSSPrompt = acDocEd.SelectWindow(new Point3d(BRSheet.X - zonaCodeTL.X * mashtab.X, BRSheet.Y + zonaCodeTL.Y * mashtab.Y, 0), new Point3d(BRSheet.X - zonaCodeBR.X * mashtab.X, BRSheet.Y + zonaCodeBR.Y * mashtab.Y, 0));

        //            // If the prompt status is OK, objects were selected
        //            if (acSSPrompt.Status == PromptStatus.OK)
        //            {
        //                SelectionSet acSSet = acSSPrompt.Value;
        //                // Step through the objects in the selection set
        //                noCode = true;
        //                foreach (SelectedObject acSSObj in acSSet)
        //                {
        //                    if (acSSObj != null)
        //                    {
        //                        MText acMText = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as MText;
        //                        if (acMText != null)
        //                        {
        //                            strCode = strCode + acMText.Text;
        //                            noCode = false;
        //                            //acDoc.Editor.WriteMessage(acMText.Text);
        //                            //acDoc.Editor.WriteMessage("\n");
        //                        }
        //                        DBText ntext = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as DBText;
        //                        if (ntext != null)
        //                        {
        //                            strCode = strCode + ntext.TextString;
        //                            noCode = false;
        //                        }
        //                    }
        //                }
        //                if (noCode)
        //                {
        //                    //acDoc.Editor.WriteMessage("Дополнительная часть шифра не обнаружена!");
        //                    if (strCode.Length < 1)
        //                    {
        //                        strCode = strCode + "NONAME";
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                //acDoc.Editor.WriteMessage("Дополнительная часть шифра не обнаружена!");
        //                if (strCode.Length < 1)
        //                {
        //                    strCode = strCode + "NONAME";
        //                }
        //            }

        //            //acDoc.Editor.WriteMessage(" -норм листа- Х- " + bShtamp.Position.X + " -У- " + bShtamp.Position.Y + "\n");
        //            //acDoc.Editor.WriteMessage(" -корд LT- Х- " + zonaNumTL.X + " -У- " + zonaNumTL.Y + "\n");
        //            //acDoc.Editor.WriteMessage(" -корд BR- Х- " + zonaNumBR.X + " -У- " + zonaNumBR.Y + "\n");

        //            PromptSelectionResult acSSPrompt2 = acDocEd.SelectWindow(new Point3d(BRSheet.X - zonaNumTL.X * mashtab.X, BRSheet.Y + zonaNumTL.Y * mashtab.Y, 0), new Point3d(BRSheet.X - zonaNumBR.X * mashtab.X, BRSheet.Y + zonaNumBR.Y * mashtab.Y, 0));

        //            // If the prompt status is OK, objects were selected
        //            if (acSSPrompt2.Status == PromptStatus.OK)
        //            {
        //                SelectionSet acSSet2 = acSSPrompt2.Value;
        //                // Step through the objects in the selection set
        //                bool noNum = true;
        //                foreach (SelectedObject acSSObj2 in acSSet2)
        //                {
        //                    if (acSSObj2 != null)
        //                    {
        //                        MText acMText2 = acTrans.GetObject(acSSObj2.ObjectId, OpenMode.ForWrite) as MText;

        //                        // ПОЛУЧАЕМ номер листа
        //                        //acDoc.Editor.WriteMessage("\n" + " - номер листа - " + "\n");
        //                        //acDoc.Editor.WriteMessage(" - норм листа- Х- " + acMText2.Location.X + " -У- " + acMText2.Location.Y + "\n");
        //                        //acDoc.Editor.WriteMessage(" -корд LT- Х- " + zonaNumTL.X + " -У- " + zonaNumTL.Y + "\n");
        //                        //acDoc.Editor.WriteMessage(" -корд BR- Х- " + zonaNumBR.X + " -У- " + zonaNumBR.Y + "\n");
        //                        if (acMText2 != null)
        //                        {
        //                            strCode = strCode + "_" + acMText2.Text.Replace("л.", "");
        //                            noNum = false;
        //                        }

        //                        DBText ntext2 = acTrans.GetObject(acSSObj2.ObjectId, OpenMode.ForWrite) as DBText;
        //                        if (ntext2 != null)
        //                        {
        //                            strCode = strCode + "_" + ntext2.TextString.Replace("л.", "");
        //                            noNum = false;
        //                        }

        //                    }
        //                }
        //                if (noNum)
        //                {
        //                    //acDoc.Editor.WriteMessage("Номер листа не обнаружен!");
        //                    strCode = strCode + "_0";
        //                }
        //            }
        //            else
        //            {
        //                //acDoc.Editor.WriteMessage("Номер листа не обнаружен!");
        //                strCode = strCode + "_0";
        //            }

        //        }
        //        return strCode;
        //    }
        //}


        public static bool getBThisShtamp(BlockReference bSheet, BlockReference bShtamp, out blockSpecialSheet iInfSheet, settingXML sXML)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            iInfSheet = new blockSpecialSheet();

            string[] split = bSheet.Name.Split('_');

            if (Convert.ToInt32(split[1]) == 1)
            {
                iInfSheet.isPortrait = true;
                iInfSheet.hor = Convert.ToInt32(split[2]);
                iInfSheet.vert = Convert.ToInt32(split[3]);
            }  
            else
            {
                iInfSheet.isPortrait = false;
                iInfSheet.hor = Convert.ToInt32(split[3]);
                iInfSheet.vert = Convert.ToInt32(split[2]);
            };

            Scale3d mashtab = bSheet.ScaleFactors;
            //acDoc.Editor.WriteMessage(mashtab.X.ToString() + " - внутри - " + mashtab.Y.ToString() + " - " + mashtab.Z.ToString());
            iInfSheet.mashtab = mashtab.X;
            iInfSheet.BR = bSheet.Position;
            iInfSheet.LT = new Point3d(iInfSheet.BR.X - iInfSheet.hor * mashtab.X, iInfSheet.BR.Y + iInfSheet.vert * mashtab.Y, iInfSheet.BR.Z);
            bool res = false;
            if (bShtamp != null)
            {
                if (isShtamInSheet(iInfSheet.LT, iInfSheet.BR, bShtamp.Position))
                {
                    acDoc.Editor.WriteMessage("Анализ следующего листа: СПЕЦ.ШТАМПА ЕСТЬ!" + "\n");
                    res = true;
                    iInfSheet.sheetsStamp = bShtamp;
                    iInfSheet.nameSheetCode = getNameSheetCode(bShtamp, sXML);
                    
                };
            } else {
                acDoc.Editor.WriteMessage("Анализ следующего листа: СПЕЦ.ШТАМПА НЕТ!" + "\n");
                res = true;
                iInfSheet.sheetsStamp = null;
                iInfSheet.nameSheetCode = getNameSpecNameSheetCode(bSheet, iInfSheet.hor, iInfSheet.vert, sXML);
            }

            iInfSheet.isRGB = isRGBSheet(iInfSheet.LT, iInfSheet.BR, mashtab, sXML);
            iInfSheet.printLandscape = "";
            iInfSheet.printPortrait = "";
            foreach (settingXML.setSheet br in sXML.listFS)
            {
                //acDoc.Editor.WriteMessage(" - ПОЛНЫЙ - " + "\n");
                //acDoc.Editor.WriteMessage(iInfSheet.vert + " - " + br.height + " - " +iInfSheet.hor + " - " + br.weight);

                if ((iInfSheet.vert == br.height) && (iInfSheet.hor == br.weight))
                {
                    iInfSheet.printLandscape = br.printLandscape;
                    iInfSheet.printPortrait = br.printPortrait;
                    iInfSheet.weightPlot = br.weightPlot;
                    iInfSheet.heightPlot = br.heightPlot;
                    //acDoc.Editor.WriteMessage(" - Первый - " + "\n");
                }
                if ((iInfSheet.hor == br.height) && (iInfSheet.vert == br.weight))
                {
                    iInfSheet.printLandscape = br.printLandscape;
                    iInfSheet.printPortrait = br.printPortrait;
                    iInfSheet.weightPlot = br.weightPlot;
                    iInfSheet.heightPlot = br.heightPlot;
                    //acDoc.Editor.WriteMessage(" - Второй - " + "\n");
                }
                //acDoc.Editor.WriteMessage(br.nameSheet + " - " + br.weight + " - " + br.height);
                //acDoc.Editor.WriteMessage("\n");
            }
            //bool res = true;
            //acDoc.Editor.WriteMessage(mashtab.X.ToString() + " - наружу - " + mashtab.Y.ToString() + " - " + mashtab.Z.ToString());

            return res;
        }

        [CommandMethod("BTOOLSPLOTALL")]
        public void BTOOLSPLOTALL()
        {
            BTOOLSPLOT(true,false);
        }

        [CommandMethod("BTOOLSPLOTAREA")]
        public void BTOOLSPLOTAREA()
        {
            BTOOLSPLOT(false,false);
        }
        [CommandMethod("BTOOLSPLOTALLPDF")]
        public void BTOOLSPLOTALLPDF()
        {
            BTOOLSPLOT(true, true);
        }
        [CommandMethod("BTOOLSPLOTAREAPDF")]
        public void BTOOLSPLOTAREAPDF()
        {
            BTOOLSPLOT(false,true);
        }
        public void BTOOLSPLOT(bool allPlot, bool isConverttoPDF)
        {
            // 1. Определить необходимые блоки на модели/листах
            // 2. Создать список листов и заполнить характеристик
            // 3. Распечатать на принтере

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

            //Загрузка настроек из XML
            settingXML setXML = new settingXML();
            setXML = (settingXML)settingGet.getSettingAll();

            //foreach (settingXML.setSheet br in setXML.listFS)
            //{
            //    acDoc.Editor.WriteMessage(br.nameSheet + " - " + br.weightPlot + " - " + br.heightPlot);
            //    acDoc.Editor.WriteMessage("\n");
            //}
            //******/


            //// Получаем необходимые нам блоки
            ////List<BlockReference> listAllBlocksSheet = new List<BlockReference>();
            List<BlockReference> listBlocksSheets = new List<BlockReference>();
            List<BlockReference> listBlocksShtamp = new List<BlockReference>();

            acDoc.Editor.WriteMessage("Пользователем выбрана следующая обработка:");
            // Организуем выборку всех элементов на всем чертеже или выбделеной области
            if (allPlot)
            {
                //Выбираем все элементы на чертеже
                using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTableRecord currentSpace = trans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entId in currentSpace)
                    {
                        if (entId.ObjectClass != RXClass.GetClass(typeof(BlockReference))) continue; // если это не блок то пропускаем данный элемент
                        
                        BlockReference bobj = trans.GetObject(entId, OpenMode.ForRead) as BlockReference;
                        //Смотрим имя данного блока совпадает с специальным именем определяющих лист
                        if (bobj.Name.IndexOf(specialBlockName, 0) > -1)
                        {
                            listBlocksSheets.Add(bobj);
                        };
                        //Смотрим имя данного блока совпадает с специальным именем определяющих штамп
                        if (bobj.Name.IndexOf(specialShtampName, 0) > -1)
                        {
                            listBlocksShtamp.Add(bobj);
                        };

                    }
                }
                acDoc.Editor.WriteMessage(" -----всего чертежа!");
            }
            else
            {
                //организуем выборку элементов по выделению
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                    {
                        foreach (SelectedObject acSSObj in acSSPrompt.Value)
                        {
                            if (acSSObj != null)
                            {

                                BlockReference blkRef = trans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                                if (blkRef != null)
                                {
                                    //Смотрим имя данного блока совпадает с специальным именем определяющих лист
                                    if (blkRef.Name.IndexOf(specialBlockName, 0) > -1)
                                    {
                                        listBlocksSheets.Add(blkRef);
                                    };
                                    //Смотрим имя данного блока совпадает с специальным именем определяющих штамп
                                    if (blkRef.Name.IndexOf(specialShtampName, 0) > -1)
                                    {
                                        listBlocksShtamp.Add(blkRef);
                                    };
                                }
                            }
                        }
                    }
                }
                acDoc.Editor.WriteMessage("-----выделеных элементов");
            }
            acDoc.Editor.WriteMessage("\n");
            acDoc.Editor.WriteMessage(" ");
            acDoc.Editor.WriteMessage("\n");

            acDoc.Editor.WriteMessage("Найдено: \n");
            acDoc.Editor.WriteMessage("   -количество листов: " + listBlocksSheets.Count + "шт. \n");
            acDoc.Editor.WriteMessage("   -количество штампов: " + listBlocksShtamp.Count + "шт. \n");
            //foreach (BlockReference brObj in listAllBlocksSheet)
            //{
            //    acDoc.Editor.WriteMessage(" - " + brObj.Name + " - ");
            //    acDoc.Editor.WriteMessage("\n");
            //}

            ////******////

            ////****** Получаем список листов со штампами
            
            acDoc.Editor.WriteMessage("\n");
            acDoc.Editor.WriteMessage(" ");
            acDoc.Editor.WriteMessage("\n");

            //acDoc.Editor.WriteMessage("Анализ совмещения листа и штампа: \n");
            List<blockSpecialSheet> listBSheet = new List<blockSpecialSheet>();
            //** Определение блоков которые относятся к штампу и объеденение их к листу, с добавлением настроек
            foreach (BlockReference brObj in listBlocksSheets)
            {
                    bool noShtamp = true;
                    blockSpecialSheet iBSheet = new blockSpecialSheet();
                    if (listBlocksShtamp.Count <1) {
                    //acDoc.Editor.WriteMessage("*Анализирую следующий лист:");
                    if (getBThisShtamp(brObj, null, out iBSheet, setXML))
                        {
                            //acDoc.Editor.WriteMessage(" ДО Цикла Лист  getBThisShtamp(brObj, null, out iBSheet, setXML) : " + iBSheet.nameSheetCode + " ");
                            listBSheet.Add(iBSheet);
                            noShtamp = false;
                            
                            //acDoc.Editor.WriteMessage("штамп НЕ НАЙДЕН");
                        }
                    }

                    foreach (BlockReference brObjSh in listBlocksShtamp)
                    {
                        //acDoc.Editor.WriteMessage(" - " + brObjSh.BlockName + " - ");
                        if (getBThisShtamp(brObj, brObjSh, out iBSheet, setXML)) 
                            {
                                //acDoc.Editor.WriteMessage("Лист  getBThisShtamp(brObj, brObjSh, out iBSheet, setXML)   : " + iBSheet.nameSheetCode + " -");
                                listBSheet.Add(iBSheet);
                                noShtamp = false;
                               // if (iBSheet.sheetsStamp != null) acDoc.Editor.WriteMessage(" штамп обноружен!");
                               // else acDoc.Editor.WriteMessage("штамп НЕ НАЙДЕН ");
                            };
                    }  
                    if (noShtamp) {
                        if (getBThisShtamp(brObj, null, out iBSheet, setXML))
                        {
                            //acDoc.Editor.WriteMessage("Лист getBThisShtamp(brObj, null, out iBSheet, setXML)  : " + iBSheet.nameSheetCode + " ");
                            listBSheet.Add(iBSheet);
                            // noShtamp = false;
                            //acDoc.Editor.WriteMessage("штамп НЕ НАЙДЕН");
                        }
                        //else { iBSheet.nameSheetCode = "NONAME_0";  }
                        
                        //listBSheet.Add(iBSheet);
                        //acDoc.Editor.WriteMessage(", но имя листа не Обноружено");
                    }
                acDoc.Editor.WriteMessage("\n");
            };

            acDoc.Editor.WriteMessage("\n");

            acDoc.Editor.WriteMessage("Печать начата. Статус: ");
            acDoc.Editor.WriteMessage("\n");

            //Начинаем печатать
            foreach (blockSpecialSheet iBSObj in listBSheet)
            {
                //подготавливаем имена файлов, убираем запрещенные символы и корректируем длину
                acDoc.Editor.WriteMessage("Лист: " + iBSObj.nameSheetCode);
                string file_name = iBSObj.nameSheetCode;
                file_name = file_name.Replace("\\", "_");
                file_name = file_name.Replace("/", "_");
                file_name = file_name.Replace("<", "_");
                file_name = file_name.Replace(">", "_");
                file_name = file_name.Replace("?", "_");
                file_name = file_name.Replace("_P", "_");
                file_name = file_name.Replace(((char)34).ToString(), "_");
                file_name = file_name.Replace("|", "_");
                file_name = file_name.Replace("*", "_");
                file_name = file_name.Replace(":", "_");
                
                //if (file_name.Length > 20)
                //{
                //    file_name = file_name.Substring(0, 20);
                //};

                //acDoc.Editor.WriteMessage(" - " + iBSObj.printLandscape);
                //acDoc.Editor.WriteMessage(" - " + iBSObj.isRGB);
                //acDoc.Editor.WriteMessage(" - " + iBSObj.heightPlot + " - ");

                int b = 0;
                int ссс = 0;
                bool isExists = false;
                while (System.IO.File.Exists(setXML.waysavetopdf + file_name + setXML.waysavetoext))
                {
                    ссс = b.ToString().Length;
                    //ed.WriteMessage("\n" + b);
                    //ed.WriteMessage("\n длина = " + ссс);
                    //ed.WriteMessage("\n");

                    if (isExists) { 
                        int x1 = file_name.Length - b.ToString().Length - 2;
                        file_name = file_name.Remove(x1);
                        //ed.WriteMessage("\nвввввв - " + file_name);
                    }
                    b++;
                    //ed.WriteMessage("\n" + a);
                    file_name = file_name + "(" + b + ")";
                    isExists = true;
                }
                //acDoc.Editor.WriteMessage(" имя файла: " + file_name);

                string nameprinter = "";

                if (iBSObj.isPortrait)
                {
                    nameprinter = iBSObj.printPortrait;
                    acDoc.Editor.WriteMessage(" - портрет(" + nameprinter + ")");
                }
                else
                {
                    nameprinter = iBSObj.printLandscape;
                    acDoc.Editor.WriteMessage(" - альбом(" + nameprinter + ")");
                }
                string colorstyle = setXML.colorWB;
                if (iBSObj.isRGB)
                {
                    colorstyle = setXML.colorRGB;
                    acDoc.Editor.WriteMessage(" - цветной");
                } else { acDoc.Editor.WriteMessage(" - ЧБ"); }

                //проверяем конвертировать в ПДВ, Да или нет
                //string nameEXT = "";
                //if isConverttoPDF {
                //   nameEXT = setXML.waysavetoext;

                //}; 

                if (nameprinter.Trim() == "")
                {
                    acDoc.Editor.WriteMessage("\n");
                    acDoc.Editor.WriteMessage("ОТМЕНА ПЕЧАТИ. ИМЯ ПРИНТЕРА НЕ НАЙДЕНО: причина не совпадение размера в блоке с заполненым в настройках");
                    acDoc.Editor.WriteMessage("\n");
                } else {
                     if (BTOOLS_PLOT.PlotWindowAreaToPDF.PlotWindowArea(iBSObj, nameprinter, colorstyle, setXML.waysavetopdf + file_name + setXML.waysavetoext, isConverttoPDF))
                    {
                        AddLine(iBSObj.LT, iBSObj.BR); //рисуем линию
                        acDoc.Editor.WriteMessage(" - ОК!");
                        acDoc.Editor.WriteMessage("\n");

                    } 
                        
                }


            }
        }
    }
}

