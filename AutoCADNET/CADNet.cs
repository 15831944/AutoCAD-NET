/** HASKELL - Atlanta
 *  Created: April 16, 2016
 * 
 *  Revision: 2.0.0.0
 *  (c) Copyright by The Haskell Company
 * 
 *  Original Author: Adam Dunigan
 *  Original Contact: adam.dunigan@haskell.com
 **/

using ACSMCOMPONENTS22Lib;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoCADCommands
{
    public class Commands
    {
        #region Global Variables
        //Constants [PLANVIEWTABLE]
        //
        //Table formatting constants
        const double colWidth = 7.5;
        const double rowHeight = 1.0;
        const double textHeight = 0.5;
        const CellAlignment cellAlign = CellAlignment.MiddleCenter;

        //Constants [Change_Attribute_Width]
        //
        //Persistent data for QoL
        double mTextWidth;
        #endregion

        [CommandMethod("AdamHelp")]
        public void AdamTutorials()
        {
            ////////////DISPLAYS THE DIALOG BOX//////////////////////////
            Tutorial_Dialogs.Tutorials_by_Adam_Form TbA = new Tutorial_Dialogs.Tutorials_by_Adam_Form();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(null, TbA, false);
            ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^////////////////////////////////// 
        }

        [CommandMethod("AutoSysVars")]
        public void AutoSysVars()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("ANNOALLVISIBLE", 1);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("ATTDIA", 1);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 3);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BVMODE", 0);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FIELDEVAL", 31);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FILEDIA", 1);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LAYOUTTAB", 1);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("TABMODE", 0);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("VISRETAIN", 1);
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSNAPCOORD", 0);

            doc.SendStringToExecute("ATTDISP N ", true, false, true);
            doc.SendStringToExecute("FILETABS ", true, false, true);
        }

        [CommandMethod("PartLookup")]
        public void PartLookup()
        {
            ////////////DISPLAYS THE DIALOG BOX//////////////////////////
            Part_Lookup_Dialogs.Part_Lookup_Form PLF = new Part_Lookup_Dialogs.Part_Lookup_Form();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(null, PLF, false);
            ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^////////////////////////////////// 
        }

        [CommandMethod("Change_Attribute_Width")]
        public void ChangeAttributeWidth()
        {
            Document doc;
            Database db;
            Editor ed;

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            PromptDoubleOptions pdOpts = new PromptDoubleOptions("\nEnter new text width:");
            pdOpts.DefaultValue = mTextWidth;
            pdOpts.AllowZero = false;
            pdOpts.AllowNone = false;
            pdOpts.AllowNegative = false;
            PromptDoubleResult pdRes = ed.GetDouble(pdOpts);

            if (pdRes.Status == PromptStatus.OK)
            {
                mTextWidth = pdRes.Value;
                DocumentLock loc = doc.LockDocument();
                using (loc)
                {
                    using (Transaction transaction = db.TransactionManager.StartTransaction())
                    {
                        ed.WriteMessage("\nSelect multiline attributes to alter...\n");
                        PromptSelectionResult SPrompt = ed.GetSelection();
                        if (SPrompt.Status == PromptStatus.OK)
                        {
                            SelectionSet SSet = SPrompt.Value;
                            foreach (SelectedObject acSO in SSet)
                            {
                                if (acSO != null)
                                {
                                    Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
                                    if (ent != null)
                                    {                                   
                                        BlockReference blkRef = ent as BlockReference;

                                        if (blkRef != null)
                                        {
                                            foreach (ObjectId attId in blkRef.AttributeCollection)
                                            {
                                                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                                if (attRef.Tag == "NAME")
                                                {
                                                    attRef.WidthFactor = mTextWidth;
                                                    if(attRef.IsMTextAttribute)
                                                    {
                                                        attRef.UpdateMTextAttribute();
                                                    }                                                    
                                                }
                                                if (attRef.Tag == "CONNECTION")
                                                {
                                                    attRef.WidthFactor = mTextWidth;
                                                    if (attRef.IsMTextAttribute)
                                                    {
                                                        attRef.UpdateMTextAttribute();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        transaction.Commit();
                    }
                }
            }
        }

        [CommandMethod("SuperPurge")]
        public void SuperPurge()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int count = PurgeDatabase(db);

            ed.WriteMessage("\nPurged {0} object{1} from the current database.", count, count == 1 ? "" : "s");

            var countZL = db.ConditionalErase(
                e =>
            {
                var cur = e as Curve;
                return
                    cur != null && cur.GetDistanceAtParameter(cur.EndParam) < Tolerance.Global.EqualPoint;
            });

            ed.WriteMessage("\nErased {0} zero-length entities.", countZL);

            int countTXT = PurgeEmptyText(db);

            ed.WriteMessage("\nErased {0} empty text object{1}.", countTXT, countTXT == 1 ? "" : "s");

            doc.SendStringToExecute("-purge O ", true, false, true);
        }            
        
        [CommandMethod("FootprintLookup")]
        public void FootprintLookup()
        {       
                ////////////DISPLAYS THE DIALOG BOX//////////////////////////
            Footprint_Lookup_Dialogs.Footprint_Lookup_Form FLF = new Footprint_Lookup_Dialogs.Footprint_Lookup_Form();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(null, FLF, false);            
                ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^//////////////////////////////////                                  
        }

        [CommandMethod("UpdateTitleBlock")]
        public void UpdateTitleBlock()
        {
            ////////////DISPLAYS THE DIALOG BOX//////////////////////////
            Update_Title_Block_Dialogs.Update_Title_Block_Form UTB = new Update_Title_Block_Dialogs.Update_Title_Block_Form();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(null, UTB, false);
            ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^////////////////////////////////// 
        }

        [CommandMethod("TitleBlockEdit")]
        public void TitleBlockEdit()
        {
            Document doc;
            Database db;
            Editor ed;

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            DocumentLock loc = doc.LockDocument();
            using (loc)
            {
                using (Transaction transaction = db.TransactionManager.StartTransaction())
                {           
                    ed.WriteMessage("Select Title Block(s):\n");
                    PromptSelectionResult SPrompt = ed.GetSelection();

                    if(SPrompt.Status == PromptStatus.OK)
                    {                        
                        SelectionSet SSet = SPrompt.Value;

                        string projnum = "";
                        string sheetnum = "";
                        string line = "";

                        foreach (SelectedObject acSO in SSet)
                        {
                            if (acSO != null)
                            {
                                Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
                                if (ent != null)
                                {
                                    BlockReference blkRef = ent as BlockReference;

                                    if (blkRef != null)
                                    {                                        
                                        foreach (ObjectId attId in blkRef.AttributeCollection)
                                        {
                                            AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                            if (attRef.Tag == "PROJECTNUM")
                                            {
                                                if (projnum == "")
                                                {
                                                    PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect project number: ");
                                                    PromptEntityResult prompt1 = ed.GetEntity(peOpts1);

                                                    if (prompt1.Status == PromptStatus.OK)
                                                    {
                                                        ObjectId objId1 = prompt1.ObjectId;

                                                        if (objId1 != null)
                                                        {
                                                            string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
                                                            projnum = objIDString1;
                                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%";
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + projnum + @">%).TextString>%";
                                                }
                                            }
                                            else if (attRef.Tag == "SHEET_NUMBER")
                                            {
                                                if (line == "")
                                                {
                                                    PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect line number: ");
                                                    PromptEntityResult prompt1 = ed.GetEntity(peOpts1);
                                                    if (prompt1.Status == PromptStatus.OK)
                                                    {
                                                        ObjectId objId1 = prompt1.ObjectId;

                                                        if (objId1 != null)
                                                        {
                                                            string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
                                                            line = objIDString1;
                                                            PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
                                                            PromptEntityResult prompt = ed.GetEntity(peOpts);

                                                            if (prompt.Status == PromptStatus.OK)
                                                            {
                                                                ObjectId objId = prompt.ObjectId;

                                                                if (objId != null)
                                                                {
                                                                    string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
                                                                    sheetnum = objIDString;
                                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
                                                    PromptEntityResult prompt = ed.GetEntity(peOpts);

                                                    if (prompt.Status == PromptStatus.OK)
                                                    {
                                                        ObjectId objId = prompt.ObjectId;

                                                        if (objId != null)
                                                        {
                                                            string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
                                                            sheetnum = objIDString;
                                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + line + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%";
                                                        }
                                                    }
                                                }
                                            }
                                            else if (attRef.Tag == "SHEET_TITLE")
                                            {
                                                PromptStringOptions PSOpts = new PromptStringOptions("\nEnter sheet title: ");
                                                PSOpts.AllowSpaces = true;
                                                PromptResult PResult = ed.GetString(PSOpts);

                                                attRef.TextString = PResult.StringResult;
                                            }
                                        }                                        
                                    }
                                }
                            }                            
                        }

                    }                    
                    transaction.Commit();
                }                
            }
            doc.SendStringToExecute("Regenall ", true, false, false);
        }

        [CommandMethod("Renumber")]
        public void Renumber()
        {
            Document doc;
            Database db;
            Editor ed;

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            PromptKeywordOptions pkOpts = new PromptKeywordOptions("\nSelect type to renumber?");
            pkOpts.Keywords.Add("Wire");
            pkOpts.Keywords.Add("Component");
            pkOpts.AllowArbitraryInput = false;
            pkOpts.AllowNone = true;
            PromptResult firstPrompt = ed.GetKeywords(pkOpts);

            if(firstPrompt.Status == PromptStatus.OK)
            {
                if(firstPrompt.StringResult.Equals("Wire"))
                {
                    DocumentLock loc = doc.LockDocument();
                    using (loc)
                    {
                        using (Transaction transaction = db.TransactionManager.StartTransaction())
                        {
                            ed.WriteMessage("Select Wire Numbers:\n");
                            PromptSelectionResult SPrompt = ed.GetSelection();

                            if (SPrompt.Status == PromptStatus.OK)
                            {
                                SelectionSet SSet = SPrompt.Value;

                                PromptKeywordOptions pkCompConfOpts = new PromptKeywordOptions("\nUse wire configuration?");
                                pkCompConfOpts.Keywords.Add("Yes");
                                pkCompConfOpts.Keywords.Add("No");
                                pkCompConfOpts.AllowArbitraryInput = false;
                                pkCompConfOpts.AllowNone = true;
                                PromptResult pkCompConfPrmpt = ed.GetKeywords(pkCompConfOpts);

                                if(pkCompConfPrmpt.Status == PromptStatus.OK)
                                {
                                    if (pkCompConfPrmpt.StringResult.Equals("Yes"))
                                    {
                                        PromptEntityOptions pOpts = new PromptEntityOptions("\nSelect wire configuration: ");
                                        PromptEntityResult pmpt = ed.GetEntity(pOpts);

                                        if (pmpt.Status == PromptStatus.OK)
                                        {
                                            string objString = Regex.Replace(pmpt.ObjectId.ToString(), @"[^0-9a-zA-Z]+", "");

                                            PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
                                            PromptEntityResult prompt = ed.GetEntity(peOpts);

                                            if (prompt.Status == PromptStatus.OK)
                                            {
                                                PromptIntegerOptions PSIOpts = new PromptIntegerOptions("\nEnter suffix character count: ");
                                                PSIOpts.AllowNegative = false;
                                                PSIOpts.AllowNone = false;
                                                PSIOpts.AllowZero = true;
                                                PromptIntegerResult PIResult = ed.GetInteger(PSIOpts);

                                                if (PIResult.Status == PromptStatus.OK)
                                                {
                                                    PromptStringOptions PStringOpts = new PromptStringOptions("\nEnter any characters to include before suffix: ");
                                                    PStringOpts.AllowSpaces = true;
                                                    PromptResult PStringResult = ed.GetString(PStringOpts);

                                                    if(PStringResult.Status == PromptStatus.OK)
                                                    {
                                                        ObjectId objId = prompt.ObjectId;

                                                        foreach (SelectedObject acSO in SSet)
                                                        {
                                                            if (acSO != null)
                                                            {
                                                                Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
                                                                if (ent != null)
                                                                {
                                                                    BlockReference blkRef = ent as BlockReference;

                                                                    if (blkRef != null)
                                                                    {
                                                                        foreach (ObjectId attId in blkRef.AttributeCollection)
                                                                        {
                                                                            AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                                                            if (attRef.Tag == "WIRENO")
                                                                            {
                                                                                if (objId != null)
                                                                                {
                                                                                    string suffix = attRef.TextString.Substring(attRef.TextString.Length - PIResult.Value, PIResult.Value);
                                                                                    string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
                                                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objString + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%" + PStringResult.StringResult + suffix;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }                                                    
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
                                        PromptEntityResult prompt = ed.GetEntity(peOpts);

                                        if (prompt.Status == PromptStatus.OK)
                                        {
                                            PromptIntegerOptions PSIOpts = new PromptIntegerOptions("\nEnter suffix character count: ");
                                            PSIOpts.AllowNegative = false;
                                            PSIOpts.AllowNone = false;
                                            PSIOpts.AllowZero = true;
                                            PromptIntegerResult PIResult = ed.GetInteger(PSIOpts);

                                            if (PIResult.Status == PromptStatus.OK)
                                            {
                                                PromptStringOptions PStringOpts = new PromptStringOptions("\nEnter any characters to include before suffix: ");
                                                PStringOpts.AllowSpaces = true;
                                                PromptResult PStringResult = ed.GetString(PStringOpts);

                                                if (PStringResult.Status == PromptStatus.OK)
                                                {
                                                    ObjectId objId = prompt.ObjectId;

                                                    foreach (SelectedObject acSO in SSet)
                                                    {
                                                        if (acSO != null)
                                                        {
                                                            Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
                                                            if (ent != null)
                                                            {
                                                                BlockReference blkRef = ent as BlockReference;

                                                                if (blkRef != null)
                                                                {
                                                                    foreach (ObjectId attId in blkRef.AttributeCollection)
                                                                    {
                                                                        AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                                                        if (attRef.Tag == "WIRENO")
                                                                        {
                                                                            if (objId != null)
                                                                            {
                                                                                string suffix = attRef.TextString.Substring(attRef.TextString.Length - PIResult.Value, PIResult.Value);
                                                                                string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
                                                                                attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%" + PStringResult.StringResult + suffix;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            transaction.Commit();
                        }
                    }
                }
                else if(firstPrompt.StringResult.Equals("Component"))
                {
                    DocumentLock loc = doc.LockDocument();
                    using (loc)
                    {
                        using (Transaction transaction = db.TransactionManager.StartTransaction())
                        {
                            ed.WriteMessage("Select Components:\n");
                            PromptSelectionResult SPrompt = ed.GetSelection();

                            if (SPrompt.Status == PromptStatus.OK)
                            {
                                SelectionSet SSet = SPrompt.Value;

                                PromptKeywordOptions pkCompConfOpts = new PromptKeywordOptions("\nUse component configuration?");
                                pkCompConfOpts.Keywords.Add("Yes");
                                pkCompConfOpts.Keywords.Add("No");
                                pkCompConfOpts.AllowArbitraryInput = false;
                                pkCompConfOpts.AllowNone = true;
                                PromptResult pkCompConfPrmpt = ed.GetKeywords(pkCompConfOpts);

                                if(pkCompConfPrmpt.Status == PromptStatus.OK)
                                {
                                    if (pkCompConfPrmpt.StringResult.Equals("Yes"))
                                    {
                                        PromptEntityOptions pOpts = new PromptEntityOptions("\nSelect component configuration: ");
                                        PromptEntityResult pmpt = ed.GetEntity(pOpts);

                                        if (pmpt.Status == PromptStatus.OK)
                                        {
                                            string objString = Regex.Replace(pmpt.ObjectId.ToString(), @"[^0-9a-zA-Z]+", "");

                                            PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
                                            PromptEntityResult prompt = ed.GetEntity(peOpts);

                                            if (prompt.Status == PromptStatus.OK)
                                            {
                                                PromptIntegerOptions PSIPOpts = new PromptIntegerOptions("\nEnter prefix character count: ");
                                                PSIPOpts.AllowNegative = false;
                                                PSIPOpts.AllowNone = false;
                                                PSIPOpts.AllowZero = true;
                                                PromptIntegerResult PIPResult = ed.GetInteger(PSIPOpts);

                                                if (PIPResult.Status == PromptStatus.OK)
                                                {
                                                    PromptIntegerOptions PSISOpts = new PromptIntegerOptions("\nEnter suffix character count: ");
                                                    PSISOpts.AllowNegative = false;
                                                    PSISOpts.AllowNone = false;
                                                    PSISOpts.AllowZero = true;
                                                    PromptIntegerResult PISResult = ed.GetInteger(PSISOpts);

                                                    if (PISResult.Status == PromptStatus.OK)
                                                    {
                                                        PromptStringOptions PStringOpts = new PromptStringOptions("\nEnter any characters to include before suffix: ");
                                                        PStringOpts.AllowSpaces = true;
                                                        PromptResult PStringResult = ed.GetString(PStringOpts);

                                                        if (PStringResult.Status == PromptStatus.OK)
                                                        {
                                                            ObjectId objId = prompt.ObjectId;

                                                            foreach (SelectedObject acSO in SSet)
                                                            {
                                                                if (acSO != null)
                                                                {
                                                                    Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
                                                                    if (ent != null)
                                                                    {
                                                                        BlockReference blkRef = ent as BlockReference;

                                                                        if (blkRef != null)
                                                                        {
                                                                            foreach (ObjectId attId in blkRef.AttributeCollection)
                                                                            {
                                                                                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                                                                if (attRef.Tag == "TAG")
                                                                                {
                                                                                    if (objId != null)
                                                                                    {
                                                                                        string prefix = attRef.TextString.Substring(0, PIPResult.Value);
                                                                                        string suffix = attRef.TextString.Substring(attRef.TextString.Length - PISResult.Value, PISResult.Value);
                                                                                        string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
                                                                                        attRef.TextString = prefix + @"%<\AcObjProp Object(%<\_ObjId " + objString + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%" + PStringResult.StringResult + suffix;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
                                        PromptEntityResult prompt = ed.GetEntity(peOpts);

                                        if (prompt.Status == PromptStatus.OK)
                                        {
                                            PromptIntegerOptions PSIPOpts = new PromptIntegerOptions("\nEnter prefix character count: ");
                                            PSIPOpts.AllowNegative = false;
                                            PSIPOpts.AllowNone = false;
                                            PSIPOpts.AllowZero = true;
                                            PromptIntegerResult PIPResult = ed.GetInteger(PSIPOpts);

                                            if (PIPResult.Status == PromptStatus.OK)
                                            {
                                                PromptIntegerOptions PSISOpts = new PromptIntegerOptions("\nEnter suffix character count: ");
                                                PSISOpts.AllowNegative = false;
                                                PSISOpts.AllowNone = false;
                                                PSISOpts.AllowZero = true;
                                                PromptIntegerResult PISResult = ed.GetInteger(PSISOpts);

                                                if (PISResult.Status == PromptStatus.OK)
                                                {
                                                    PromptStringOptions PStringOpts = new PromptStringOptions("\nEnter any characters to include before suffix: ");
                                                    PStringOpts.AllowSpaces = true;
                                                    PromptResult PStringResult = ed.GetString(PStringOpts);

                                                    if (PStringResult.Status == PromptStatus.OK)
                                                    {
                                                        ObjectId objId = prompt.ObjectId;

                                                        foreach (SelectedObject acSO in SSet)
                                                        {
                                                            if (acSO != null)
                                                            {
                                                                Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
                                                                if (ent != null)
                                                                {
                                                                    BlockReference blkRef = ent as BlockReference;

                                                                    if (blkRef != null)
                                                                    {
                                                                        foreach (ObjectId attId in blkRef.AttributeCollection)
                                                                        {
                                                                            AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                                                            if (attRef.Tag == "TAG")
                                                                            {
                                                                                if (objId != null)
                                                                                {
                                                                                    string prefix = attRef.TextString.Substring(0, PIPResult.Value);
                                                                                    string suffix = attRef.TextString.Substring(attRef.TextString.Length - PISResult.Value, PISResult.Value);
                                                                                    string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
                                                                                    attRef.TextString = prefix + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%" + PStringResult.StringResult + suffix;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            transaction.Commit();
                        }
                    }
                }
                doc.SendStringToExecute("Regenall ", true, false, false);
            }          
        }

        [CommandMethod("IODescriptions")]
        public void IODescriptions()
        {
            Document doc;
            Database db;
            Editor ed;

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            //Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select Excel file with descriptions:", null, ".xls; .xlsx", "ExcelFiles", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple | Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
            //System.Windows.Forms.DialogResult dr = ofd.ShowDialog();

            //if (dr != System.Windows.Forms.DialogResult.OK)
            //{
            //    ed.WriteMessage("Invalid selection.");               
            //    return;
            //}

            PromptEntityOptions pEntOpts = new PromptEntityOptions("\nSelect IO Card: ");
            PromptEntityResult EntPmpt = ed.GetEntity(pEntOpts);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                if (EntPmpt.Status == PromptStatus.OK)
                {
                    ObjectId objId = EntPmpt.ObjectId;
                    if (objId != null)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (ent != null)
                        {
                            BlockReference blkRef = ent as BlockReference;
                            if (blkRef != null)
                            {
                                PromptStringOptions PAckOpts = new PromptStringOptions("\nCopy Excel data to Clipboard and press Enter: ");
                                PAckOpts.AllowSpaces = true;
                                PromptResult PAckResult = ed.GetString(PAckOpts);

                                if (PAckResult.Status == PromptStatus.OK)
                                {
                                    string[] lines;
                                    var obj = (System.Windows.Forms.DataObject)System.Windows.Forms.Clipboard.GetDataObject();                                    
                                    if(obj != null)
                                    {
                                        var formats = obj.GetFormats();
                                        bool excelDataFound = false;
                                        foreach(var name in formats)
                                        {
                                            if(name.Contains("XML Spreadsheet"))
                                            {
                                                excelDataFound = true;
                                                break;
                                            }
                                        }

                                        if(excelDataFound)
                                        {
                                            string clipboardText = obj.GetText(); //.ToString();                                            
                                            lines = clipboardText.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                                            foreach (ObjectId attId in blkRef.AttributeCollection)
                                            {
                                                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                                                for (int i = 0; i < lines.Length; i++)
                                                {
                                                    if(attRef.Tag.Contains("PORT" + (i + 1).ToString() +"DESC") && lines.Length >= i)                                                    
                                                    {
                                                        attRef.TextString = lines[i];
                                                    }
                                                }                                                    
                                            }
                                        }
                                        else
                                        {
                                            ed.WriteMessage("No data from Excel found on clipboard.\n");
                                        }
                                    }
                                    
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }
            doc.SendStringToExecute("Regenall ", true, false, false);
        }

        [CommandMethod("IOExtract")]
        public void IOExtract()
        {
            ////////////DISPLAYS THE DIALOG BOX//////////////////////////
            IO_Extract_Dialogs.IO_Extract_Form IEF = new IO_Extract_Dialogs.IO_Extract_Form();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(null, IEF, false);            
            ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^////////////////////////////////// 
        }
       
        //Helper Functions
        //
        //
        //----------------

        //Helpers [PLANVIEWTABLE]
        //Set text height and alignment of specific cells as well as inserting the text
        public static void SetCellText(Table tb, int row, int col, string value)
        {
            tb.Cells[row, col].Alignment = cellAlign;
            tb.Cells[row, col].TextHeight = textHeight;
            tb.Cells[row, col].TextString = value;
        }
        //------------------------------------------------------------------------

        //Helpers [SUPERPURGE]
        //Catalogs and deletes all of the regapps in the database
        private static int PurgeDatabase(Database db)
        {
            int idCount = 0;

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                ObjectIdCollection idsToPurge = new ObjectIdCollection();

                RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);

                foreach (ObjectId raId in rat)
                {
                    if (raId.IsValid)
                    {
                        idsToPurge.Add(raId);
                    }
                }

                db.Purge(idsToPurge);  

                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                ed.WriteMessage("\nRegistered application being purged: ");

                foreach (ObjectId id in idsToPurge)
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForWrite);

                    RegAppTableRecord ratr = obj as RegAppTableRecord;
                    if (ratr != null)
                    {
                        ed.WriteMessage("\"{0}\" ", ratr.Name);
                    }

                    obj.Erase();
                }

                idCount = idsToPurge.Count;
                tr.Commit();
            }
            return idCount;
        }

        //Purges empty text from the document
        private static int PurgeEmptyText(Database db)
        {
            int count = 0;

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId objId in btr)
                {
                    var DBtext = tr.GetObject(objId, OpenMode.ForRead) as DBText;
                    var Mtext = tr.GetObject(objId, OpenMode.ForRead) as MText;

                    if (DBtext != null && DBtext.TextString == "")
                    {
                        DBtext.Erase();
                        count++;
                    }
                    else if (Mtext != null && Mtext.Text == "")
                    {
                        Mtext.Erase();
                        count++;                        
                    }
                }
            }
            return count;
        }        
        //------------------------------------------------------------------------

        //Helpers [WRITESOW]
        //Catalog for the contractual line items
        public static string SOWCatalog(string blName)
        {
            string sentence;
            Dictionary<string, string> catalog = new Dictionary<string, string>();
            catalog.Add("ETHERNET", @"Contractor shall provide and install Network {2} conduit and {1} cable {3} to the {0} as indicated on drawing #### and #### and terminate.  Contractor shall test each cable and provide results to the Owner.  Refer to vendor drawings #### for termination point.  Reuse existing conduit and wireway where possible.");
            catalog.Add("ELECTRICALDROP", @"Contractor shall provide and install {1} {3} conduit and {2} wiring {4} to the {0} as indicated on drawing #### and terminate.  Reuse existing conduit and wireway where possible.");
            catalog.Add("ARMORSTART", @"Contractor shall provide and install {1} {3} conduit and wiring {4} to {0} as indicated on drawing #### and terminate.  Reuse existing conduit and wireway where possible.");
            catalog.Add("DISCONNECT", @"Contractor shall provide and install {1} {3} conduit and wiring {2} to {0} as indicated on drawing #### and terminate.  Reuse existing conduit and wireway where possible.
Contractor shall provide and install 24VDC rigid conduit and (2) wires {2} to the local disconnects as indicated on drawing #### and terminate.  Refer to #### for termination points.  Reuse existing conduit and wireway where possible.");
            catalog.Add("OPERATORSTATION", @"Contractor shall install push button Operator Station provided by the Owner's System Integrator as indicated on drawing ####.
Contractor shall provide and install {1} {2} conduit and wiring {3} to the Operator Station as indicated on drawing #### and terminate.  Refer to #### for termination points.  Reuse existing conduit and wireway where possible.");
            catalog.Add("JUNCTIONBOX", @"Contractor shall provide and install {1} {2} and wiring {3} to the Junction Box as indicated on drawing #### and terminate.  Refer to #### for termination points.  Reuse existing conduit and wireway where possible.");
            catalog.Add("RECEPTACLE", @"Contractor shall provide and install {1} convenience outlet as indicated on drawing ####.
Contractor shall provide and install {1} rigid conduit and {2} wiring {3} to the convenience outlet as indicated on drawing #### and terminate.  Reuse existing conduit and wireway where possible.");
            catalog.Add("DEVICEINPUTS", @"Contractor shall provide and install {1} rigid conduit and {2} wiring {3} to {0} as indicated on drawing #### and terminate.  Reuse existing conduit and wireway where possible.");
            catalog.Add("DEVICEOUTPUTS", @"Contractor shall provide and install {1} rigid conduit and {2} wiring {3} to {0} as indicated on drawing #### and terminate.  Reuse existing conduit and wireway where possible.");

            catalog.TryGetValue(blName, out sentence);
            return sentence;
        }
        //------------------------------------------------------------------------
    }        
    
    //EXTENSION CLASS
    public static class Extensions
    {
        //Extension [SUPERPURGE]
        //Deletes objects if they do not meet the global size tolerance.
        public static int ConditionalErase(this Database db, Func<Entity, bool> f, bool countOnly = false)
        {
            int count = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent != null && f(ent))
                        {
                            if (!countOnly)
                            {
                                ent.UpgradeOpen();
                                ent.Erase();
                                ent.DowngradeOpen();
                            }
                        count++;
                        }
                    }
                }
                tr.Commit();
            }
            return count;
        }
        //----------------------------------------------------------------------------
    }
}

namespace AutoCADBackground
{
    public class BackgroundProcesses : IExtensionApplication
    {
        //Global Variables
        //
        //
        //---------
        Document doc;
        Database db;
        Editor ed;
        Transaction tr;
        ObjectIdCollection ids = new ObjectIdCollection();

        List<string> propertyNames = new List<string>();
        string openFileName;        

        AcSmCustomPropertyBag cpb = new AcSmCustomPropertyBag();

        public void Initialize()
        {
            //Registers the event for when a new document is created.  This event registers all other events.
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.DocumentCreated += new DocumentCollectionEventHandler(regEvents);
        }

        public void Terminate()
        {
            //Do something 
        }

        public void regEvents(object senderObj, DocumentCollectionEventArgs docCrtEvtArgs)
        {
            // Get the current document and database for registering the events
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;

            //db.InsertEnded += new EventHandler(OnInsertEnded);
            //db.BeginSave += new DatabaseIOEventHandler(OnBeginSave);
            //db.ObjectModified += new ObjectEventHandler(OnObjectModify);
            //doc.CommandEnded += new CommandEventHandler(OnCommandEnd);
            //doc.CommandWillStart += new CommandEventHandler(OnBeginSave);
            
            //Doc.BeginDocumentClose += new DocumentBeginCloseEventHandler(docBeginDocClose);
        }

        private void OnInsertEnded(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.Internal.Utils.EntLast();
        }

        public void OnBeginSave(object senderObj, DatabaseIOEventArgs dbIOEvtArgs)
        {
            //Ensures the current doc, db, and ed are for the active document
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;
            AcSmSheet sheet = new AcSmSheet();

            string SSPath = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("SSFOUND").ToString();
            if(!String.IsNullOrEmpty(SSPath))
            {
                AcSmSheetSetMgr mgr = new AcSmSheetSetMgr();
                AcSmDatabase SSdb = new AcSmDatabase();

                try
                {

                    SSdb = mgr.OpenDatabase(SSPath, true);
                }
                catch
                {
                    ed.WriteMessage("Database does not exist or is already open.\n");
                    try
                    {
                        SSdb = mgr.FindOpenDatabase(SSPath);
                    }
                    catch
                    {
                        ed.WriteMessage("Could not open database.");
                    }                     
                }

                if (SSdb.GetLockStatus() == AcSmLockStatus.AcSmLockStatus_UnLocked)
                {
                    SSdb.LockDb(SSdb);
                }                

                try
                {
                    //Get sheet set from open database
                    AcSmSheetSet SS = new AcSmSheetSet();
                    SS = SSdb.GetSheetSet();

                    //Get custom property bag for the sheet set and enumerate and verify                    
                    cpb = SS.GetCustomPropertyBag();
                    bool updateRefs = false;
                    updateRefs = PropertyEnumerator(cpb.GetPropertyEnumerator());

                    if(updateRefs)
                    {
                        for(int i = 1; i <= 20; i++)
                        {
                            propertyNames.Add("Wire Reference " + i.ToString());
                        }                        
                        AcSmCustomPropertyValue customPropertyValue = new AcSmCustomPropertyValue();
                        customPropertyValue.InitNew((IAcSmPersist)SS);
                        customPropertyValue.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP);
                        for(int i = 0; i < 20; i++)
                        {
                            cpb.SetProperty(propertyNames[i], customPropertyValue);
                        }                                           
                    }

                    AcSmCustomPropertyValue cpv = new AcSmCustomPropertyValue();
                    cpv.InitNew((IAcSmPersist)SS);
                    cpv.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP);
                    cpb.SetProperty("testprop", cpv);

                    cpb = SS.GetCustomPropertyBag();

                    //Get and enumerate sheet enumerator while searching for saved file
                    IAcSmEnumComponent sheetenum = SS.GetSheetEnumerator();
                    openFileName = Path.GetFileName(doc.Name);
                    //ed.WriteMessage(openFileName);
                    Enumerator(sheetenum, openFileName);                    
                }
                catch
                {
                    ed.WriteMessage("Error");
                }
                
                if (SSdb.GetLockStatus() == AcSmLockStatus.AcSmLockStatus_Locked_Local)
                {
                    SSdb.UnlockDb(SSdb);
                }
                //mgr.Close(SSdb);
            }
        }

        public bool PropertyEnumerator(IAcSmEnumProperty iter)
        {
            string propName;
            bool refsFound = false;
            AcSmCustomPropertyValue propValue = new AcSmCustomPropertyValue();
            iter.Next(out propName, out propValue);
            while (!String.IsNullOrEmpty(propName))
            {
                if(propName.Equals("Wire Reference 20") && propValue.GetFlags() == PropertyFlags.CUSTOM_SHEET_PROP)
                {
                    refsFound = true;
                }                
                iter.Next(out propName, out propValue);
            }
            return refsFound;
        }

        public void Enumerator(IAcSmEnumComponent iter, string findString)
        {
            IAcSmEnumProperty enumProp = cpb.GetPropertyEnumerator();
            AcSmCustomPropertyValue cpv = new AcSmCustomPropertyValue();

            IAcSmComponent item = iter.Next();
            while (item != null)
            {
                string type = item.GetTypeName();
                if (type.Equals("AcSmSheet"))
                {
                    AcSmSheet sh = (AcSmSheet)item;

                    //Update properties of each sheet if necessary
                    AcSmCustomPropertyBag shcpb = sh.GetCustomPropertyBag();
                    AcSmCustomPropertyValue shcpv = new AcSmCustomPropertyValue();
                    for (int i = 0; i < 20; i++)
                    {
                        shcpv.InitNew((IAcSmPersist)sh);
                        shcpv.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP);
                        shcpb.SetProperty(propertyNames[i], shcpv);
                    }

                    //Check if the name of the current sheet file matches the one opened in AutoCAD
                    string shName = sh.GetName();
                    AcSmAcDbLayoutReference layout = sh.GetLayout();
                    string path = layout.GetFileName();

                    if (findString.ToUpper().Equals(Path.GetFileName(path).ToUpper()))
                    {
                        ed.WriteMessage("Found");
                        tr = doc.TransactionManager.StartTransaction();
                        using (tr)
                        {
                            BlockTable blkTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            foreach (ObjectId id in blkTable)
                            {
                                if(id != null)
                                {
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                                    if(btr != null)
                                    {
                                        ObjectIdCollection ObjIdC = btr.GetBlockReferenceIds(false, false);
                                        foreach (ObjectId ObjId in ObjIdC)
                                        {
                                            string wirenumber = "";
                                            string refID = "";
                                            string referenceID = "";

                                            BlockReference br = (BlockReference)tr.GetObject(ObjId, OpenMode.ForRead);
                                            if (br != null)
                                            {
                                                if (br.Name.Equals("WD_WNH") || btr.Name.Equals("WD_WNV"))
                                                {
                                                    foreach (ObjectId attId in br.AttributeCollection)
                                                    {
                                                        AttributeReference ar = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                                                        if(ar != null)
                                                        {                                                           
                                                            if (ar.TextString.Equals("1"))
                                                            {                                                                
                                                                refID = ar.Tag;
                                                            }
                                                            if (ar.Tag.Equals("WIRENO"))
                                                            {                                                                
                                                                wirenumber = ar.TextString;
                                                            }                                                            
                                                            if (!String.IsNullOrEmpty(refID))
                                                            {                                                                
                                                                switch(refID)
                                                                {
                                                                    case "WR1":
                                                                        referenceID = "Wire Reference 1";
                                                                        break;
                                                                    case "WR2":
                                                                        referenceID = "Wire Reference 2";
                                                                        break;
                                                                    case "WR3":
                                                                        referenceID = "Wire Reference 3";
                                                                        break;
                                                                    case "WR4":
                                                                        referenceID = "Wire Reference 4";
                                                                        break;
                                                                    case "WR5":
                                                                        referenceID = "Wire Reference 5";
                                                                        break;
                                                                    case "WR6":
                                                                        referenceID = "Wire Reference 6";
                                                                        break;
                                                                    case "WR7":
                                                                        referenceID = "Wire Reference 7";
                                                                        break;
                                                                    case "WR8":
                                                                        referenceID = "Wire Reference 8";
                                                                        break;
                                                                    case "WR9":
                                                                        referenceID = "Wire Reference 10";
                                                                        break;
                                                                    case "WR10":
                                                                        referenceID = "Wire Reference 10";
                                                                        break;
                                                                    case "WR11":
                                                                        referenceID = "Wire Reference 11";
                                                                        break;
                                                                    case "WR12":
                                                                        referenceID = "Wire Reference 12";
                                                                        break;
                                                                    case "WR13":
                                                                        referenceID = "Wire Reference 13";
                                                                        break;
                                                                    case "WR14":
                                                                        referenceID = "Wire Reference 14";
                                                                        break;
                                                                    case "WR15":
                                                                        referenceID = "Wire Reference 15";
                                                                        break;
                                                                    case "WR16":
                                                                        referenceID = "Wire Reference 16";
                                                                        break;
                                                                    case "WR17":
                                                                        referenceID = "Wire Reference 17";
                                                                        break;
                                                                    case "WR18":
                                                                        referenceID = "Wire Reference 18";
                                                                        break;
                                                                    case "WR19":
                                                                        referenceID = "Wire Reference 19";
                                                                        break;
                                                                    case "WR20":
                                                                        referenceID = "Wire Reference 20";
                                                                        break;
                                                                }
                                                            }                                                        
                                                            if(!String.IsNullOrEmpty(referenceID) && !String.IsNullOrEmpty(wirenumber))
                                                            {
                                                                //Update sheet properties
                                                                AcSmCustomPropertyBag shcpbWR = sh.GetCustomPropertyBag();
                                                                AcSmCustomPropertyValue shcpvWR = new AcSmCustomPropertyValue();
                                                                shcpvWR.InitNew((IAcSmPersist)sh);
                                                                shcpvWR.SetFlags(PropertyFlags.CUSTOM_SHEET_PROP);
                                                                shcpvWR.SetValue((Object)wirenumber);
                                                                shcpbWR.SetProperty(referenceID, shcpvWR);  
                                                            }                                                                                                                              
                                                            
                                                        }                                                        
                                                    }
                                                }
                                            }                                            
                                        }
                                    }                                    
                                }                                
                            }
                        }
                    }
                }
                else if (type.Equals("AcSmSubset"))
                {
                    AcSmSubset subset = (AcSmSubset)item;
                    string subsetName = subset.GetName();

                    if (!String.IsNullOrEmpty(subsetName))
                    {
                        IAcSmEnumComponent subsetEnum = (IAcSmEnumComponent)subset.GetSheetEnumerator();
                        Enumerator(subsetEnum, openFileName);
                    }
                }
                item = iter.Next();
            }
        }
        //}
        //    //ed.WriteMessage(SSPath);

        //    //tr = doc.TransactionManager.StartTransaction();
        //    //using (tr)
        //    //{
        //    //    try
        //    //    {
        //    //        obj = (AXDBLib.AcadObject)LM.GetLayoutId(LM.CurrentLayout).GetObject(OpenMode.ForRead).AcadObject;
        //    //        sheetSetMgr.GetSheetFromLayout(obj, out sheet);
        //    //    }
        //    //    catch
        //    //    {

        //    //    }
        //    //}

        //    //doc.SendStringToExecute("LINE 0 5 ", true, false, true);
        //    //db.BeginSave -= new DatabaseIOEventHandler(OnBeginSave);
        //    //db.BeginSave += new DatabaseIOEventHandler(OnBeginSave);
            
        //}

        //public void OnObjectModify(object senderObj, ObjectEventArgs dbObjModEvtArgs)
        //{
        //    doc = Application.DocumentManager.MdiActiveDocument;
        //    db = doc.Database;
        //    ed = doc.Editor;
        //    //if(dbObjModEvtArgs.DBObject.IsModified)
        //    //{
        //    //    if(!ids.Contains(dbObjModEvtArgs.DBObject.ObjectId))
        //    //    {
        //    //        ids.Add(dbObjModEvtArgs.DBObject.ObjectId);
        //    //    }
        //    //}   
        //    tr = db.TransactionManager.StartTransaction();
        //    try
        //    {
        //        using (tr)
        //        {
        //            DBObject dbObj = tr.GetObject(dbObjModEvtArgs.DBObject.ObjectId, OpenMode.ForRead);
        //            BlockReference br = dbObj as BlockReference;
        //            if (br != null)
        //            {
        //                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.IsDynamicBlock ? br.DynamicBlockTableRecord : br.BlockTableRecord, OpenMode.ForRead);

        //                string blockName = btr.Name.ToUpper();
        //                ed.WriteMessage(blockName);
        //                if (blockName.Equals("WD_WNH"))
        //                {
        //                    ed.WriteMessage("test OK");
        //                }
        //            }
        //            tr.Commit();
        //        }
        //    }
        //    catch
        //    {

        //    }
        //}

        //public void OnCommandEnd(object sender, CommandEventArgs CmdEvtArgs)
        //{
        //    doc = Application.DocumentManager.MdiActiveDocument;
        //    db = doc.Database;
        //    ed = doc.Editor;

        //    if (ids != null)
        //    {
        //        tr = db.TransactionManager.StartTransaction();
        //        using (tr)
        //        {
        //            ed.WriteMessage("1");
        //            foreach (ObjectId id in ids)
        //            {
        //                WireRefUpdate(doc, tr, id);
        //            }
        //            tr.Commit();
        //        }
        //        ids.Clear();
        //    }
        //}

        //public void WireRefUpdate(Document WRUdoc, Transaction WRUtr, ObjectId WRUid)
        //{
        //    ed = WRUdoc.Editor;
        //    ed.WriteMessage("test");

        //    DBObject obj = tr.GetObject(WRUid, OpenMode.ForRead);
        //    BlockReference br = obj as BlockReference;
        //    if(br != null)
        //    {
        //        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.IsDynamicBlock ? br.DynamicBlockTableRecord : br.BlockTableRecord, OpenMode.ForRead);

        //        string blockName = btr.Name.ToUpper();
        //        ed.WriteMessage(blockName);
        //        if(blockName.Equals("WD_WNH"))
        //        {
        //            ed.WriteMessage("test OK");
        //        }
        //    }
        //}

        //public void docBeginDocClose(object senderObj, DocumentBeginCloseEventArgs docBegClsEvtArgs)
        //{
        //    // Display a message box prompting to continue closing the document
        //    if (System.Windows.Forms.MessageBox.Show(
        //                         "The document is about to be closed." +
        //                         "\nDo you want to continue?",
        //                         "Close Document",
        //                         System.Windows.Forms.MessageBoxButtons.YesNo) ==
        //                         System.Windows.Forms.DialogResult.No)
        //    {
        //        docBegClsEvtArgs.Veto();
        //    }
        //}
    }
}

//[CommandMethod("CreateSheetSet")] 
//public void CreateSheetSet()
//{
//    Document doc1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
//    Database db1 = doc1.Database;
//    Editor ed1 = doc1.Editor;
//    HostApplicationServices hs = HostApplicationServices.Current;

//    string ProjectPath = Path.GetFullPath(Path.Combine(db1.Filename, @"..\..\..\..\..\"));
//    string ProjectDirectory = Path.GetFileName(ProjectPath.TrimEnd(Path.DirectorySeparatorChar));
//    string ProjectNumber = ProjectDirectory.Substring(0, 5);            

//    AcSmSheetSetMgr SheetSetMgr = new AcSmSheetSetMgr();
//    AcSmDatabase SheetSetDB = new AcSmDatabase();
//    SheetSetDB = SheetSetMgr.CreateDatabase(ProjectPath + @"201 Design\X\Automation\" + ProjectNumber + @" Sheet Set.dst", "", true);

//    AcSmSheetSet SheetSet = SheetSetDB.GetSheetSet();
//    try
//    {
//        SheetSetDB.LockDb(SheetSetDB);
//    }
//    catch { }
//    SheetSet.SetName(ProjectNumber + " Sheet Set");

//    AcSmSubset CPSubset = (AcSmSubset) SheetSet.CreateSubset(ProjectNumber + " Control Panel Layouts", "");
//    AcSmSubset PVSubset = (AcSmSubset) SheetSet.CreateSubset(ProjectNumber + " Planview Layouts", "");

//    DirectoryInfo CPDWGInfo = new DirectoryInfo(ProjectPath + @"201 Design\X\Automation\Control Panel Drawings");
//    DirectoryInfo PVDWGInfo = new DirectoryInfo(ProjectPath + @"201 Design\X\Automation\Planview Drawings");
//    FileInfo[] CPDWGfiles = CPDWGInfo.GetFiles();
//    FileInfo[] PVDWGfiles = PVDWGInfo.GetFiles();

//    foreach (FileInfo file in CPDWGfiles)
//    {
//        Document doc;
//        Database db;
//        Editor ed;
//        try
//        {
//            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Open(ProjectPath + @"201 Design\X\Automation\Control Panel Drawings\" + file.Name);
//            db = doc.Database;
//            ed = doc.Editor;
//        }
//        catch
//        {
//            continue;
//        }

//        using (Transaction tr = db.TransactionManager.StartTransaction())
//        {
//            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false) as DBDictionary;
//            foreach (DBDictionaryEntry dictEnt in layoutDict)
//            {
//                if (dictEnt.Key == "Model")
//                    continue;
//                AcSmSheet sheet = new AcSmSheet();
//                AcSmAcDbLayoutReference layoutReference = new AcSmAcDbLayoutReference();
//                layoutReference.InitNew(SheetSet);
//                layoutReference.SetFileName(ProjectPath + @"201 Design\X\Automation\Control Panel Drawings\" + file.Name); 
//                layoutReference.SetName(dictEnt.Key);          
//                sheet = SheetSet.ImportSheet(layoutReference);
//                sheet.SetNumber(file.Name.Substring(6, 10));
//                sheet.SetTitle(ProjectNumber + " " + file.Name.Substring(17, file.Name.Length - 21));
//                CPSubset.InsertComponentAfter(sheet, null);                     
//            }
//        }
//        doc.CloseAndDiscard();
//    }
//    foreach (FileInfo file in PVDWGfiles)
//    {
//        Document doc;
//        Database db;
//        Editor ed;
//        try
//        {
//            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Open(ProjectPath + @"201 Design\X\Automation\Planview Drawings\" + file.Name);
//            db = doc.Database;
//            ed = doc.Editor;
//        }
//        catch
//        {
//            continue;
//        }

//        using (Transaction tr = db.TransactionManager.StartTransaction())
//        {
//            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false) as DBDictionary;
//            foreach (DBDictionaryEntry dictEnt in layoutDict)
//            {
//                if (dictEnt.Key == "Model")
//                    continue;
//                AcSmSheet sheet = new AcSmSheet();
//                AcSmAcDbLayoutReference layoutReference = new AcSmAcDbLayoutReference();
//                layoutReference.InitNew(SheetSet);
//                layoutReference.SetFileName(ProjectPath + @"201 Design\X\Automation\Planview Drawings\" + file.Name); 
//                layoutReference.SetName(dictEnt.Key);          
//                sheet = SheetSet.ImportSheet(layoutReference);
//                sheet.SetNumber(file.Name.Substring(6, 10));
//                sheet.SetTitle(ProjectNumber + " " + dictEnt.Key);
//                PVSubset.InsertComponentAfter(sheet, null);
//            }
//        }
//        doc.CloseAndDiscard();
//    }

//    try
//    {
//        SheetSetDB.UnlockDb(SheetSetDB);
//    }
//    catch { }                                    
//}

//[CommandMethod("Monsanto")]
//public void Monsanto()
//{
//    Document doc;
//    Database db;
//    Editor ed;

//    doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
//    db = doc.Database;
//    ed = doc.Editor;

//    DocumentLock loc = doc.LockDocument();
//    using (loc)
//    {
//        using (Transaction transaction = db.TransactionManager.StartTransaction())
//        {
//            ed.WriteMessage("Select Title Block:\n");
//            PromptSelectionResult SPrompt = ed.GetSelection();

//            if (SPrompt.Status == PromptStatus.OK)
//            {
//                SelectionSet SSet = SPrompt.Value;
//                string projnum = "";
//                string sheetnum = "";
//                string plant = "";
//                string zone = "";
//                string type = "";
//                string line = "";
//                string location = "";

//                foreach (SelectedObject acSO in SSet)
//                {
//                    if (acSO != null)
//                    {
//                        Entity ent = transaction.GetObject(acSO.ObjectId, OpenMode.ForRead) as Entity;
//                        if (ent != null)
//                        {
//                            BlockReference blkRef = ent as BlockReference;

//                            if (blkRef != null)
//                            {
//                                foreach (ObjectId attId in blkRef.AttributeCollection)
//                                {
//                                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

//                                    if (attRef.Tag == "PROJECTNUM")
//                                    {
//                                        if (projnum == "")
//                                        {
//                                            PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect project number: ");
//                                            PromptEntityResult prompt1 = ed.GetEntity(peOpts1);

//                                            if (prompt1.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId1 = prompt1.ObjectId;

//                                                if (objId1 != null)
//                                                {
//                                                    string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    projnum = objIDString1;
//                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%";
//                                                }
//                                            }
//                                        }
//                                        else
//                                        {
//                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + projnum + @">%).TextString>%";
//                                        }
//                                    }
//                                    else if (attRef.Tag == "SHEET_NUMBER")
//                                    {
//                                        if (line == "")
//                                        {
//                                            PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect line number: ");
//                                            PromptEntityResult prompt1 = ed.GetEntity(peOpts1);
//                                            if (prompt1.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId1 = prompt1.ObjectId;

//                                                if (objId1 != null)
//                                                {
//                                                    string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    line = objIDString1;
//                                                    PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
//                                                    PromptEntityResult prompt = ed.GetEntity(peOpts);

//                                                    if (prompt.Status == PromptStatus.OK)
//                                                    {
//                                                        ObjectId objId = prompt.ObjectId;

//                                                        if (objId != null)
//                                                        {
//                                                            string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                            sheetnum = objIDString;
//                                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%";
//                                                        }
//                                                    }
//                                                }
//                                            }
//                                        }
//                                        else
//                                        {
//                                            PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect page number: ");
//                                            PromptEntityResult prompt = ed.GetEntity(peOpts);

//                                            if (prompt.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId = prompt.ObjectId;

//                                                if (objId != null)
//                                                {
//                                                    string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    sheetnum = objIDString;
//                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + line + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%";
//                                                }
//                                            }
//                                        }
//                                    }
//                                    else if (attRef.Tag == "SHEET_TITLE")
//                                    {
//                                        PromptStringOptions PSOpts = new PromptStringOptions("\nEnter sheet title: ");
//                                        PSOpts.AllowSpaces = true;
//                                        PromptResult PResult = ed.GetString(PSOpts);

//                                        attRef.TextString = PResult.StringResult;
//                                    }
//                                    else if (attRef.Tag == "PLANT#")
//                                    {
//                                        if (plant == "")
//                                        {
//                                            PromptEntityOptions peOpts = new PromptEntityOptions("\nSelect plant number: ");
//                                            PromptEntityResult prompt = ed.GetEntity(peOpts);

//                                            if (prompt.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId = prompt.ObjectId;

//                                                if (objId != null)
//                                                {
//                                                    string objIDString = Regex.Replace(objId.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    plant = objIDString;
//                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString + @">%).TextString>%";
//                                                }
//                                            }
//                                        }
//                                        else
//                                        {
//                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + plant + @">%).TextString>%";
//                                        }
//                                    }
//                                    else if (attRef.Tag == "ZONE")
//                                    {
//                                        if (zone == "")
//                                        {
//                                            PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect zone: ");
//                                            PromptEntityResult prompt1 = ed.GetEntity(peOpts1);

//                                            if (prompt1.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId1 = prompt1.ObjectId;

//                                                if (objId1 != null)
//                                                {
//                                                    string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    zone = objIDString1;
//                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%";
//                                                }
//                                            }
//                                        }
//                                        else
//                                        {
//                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + zone + @">%).TextString>%";
//                                        }

//                                    }
//                                    else if (attRef.Tag == "TYPE")
//                                    {
//                                        if (type == "")
//                                        {
//                                            PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect type: ");
//                                            PromptEntityResult prompt1 = ed.GetEntity(peOpts1);

//                                            if (prompt1.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId1 = prompt1.ObjectId;

//                                                if (objId1 != null)
//                                                {
//                                                    string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    type = objIDString1;
//                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%";
//                                                }
//                                            }
//                                        }
//                                        else
//                                        {
//                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + type + @">%).TextString>%";
//                                        }
//                                    }
//                                    else if (attRef.Tag == "PLANT_LOCATION")
//                                    {
//                                        if (location == "")
//                                        {
//                                            PromptEntityOptions peOpts1 = new PromptEntityOptions("\nSelect plant location: ");
//                                            PromptEntityResult prompt1 = ed.GetEntity(peOpts1);

//                                            if (prompt1.Status == PromptStatus.OK)
//                                            {
//                                                ObjectId objId1 = prompt1.ObjectId;

//                                                if (objId1 != null)
//                                                {
//                                                    string objIDString1 = Regex.Replace(objId1.ToString(), @"[^0-9a-zA-Z]+", "");
//                                                    location = objIDString1;
//                                                    attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + objIDString1 + @">%).TextString>%";
//                                                }
//                                            }
//                                        }
//                                        else
//                                        {
//                                            attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + location + @">%).TextString>%";
//                                        }
//                                    }
//                                    else if (attRef.Tag.Equals("DESIGNDATE"))
//                                    {
//                                        attRef.TextString = @"3/21/17";
//                                    }
//                                    else if (attRef.Tag.Equals("DRAWNDATE"))
//                                    {
//                                        attRef.TextString = @"3/21/17";
//                                    }
//                                    else if (attRef.Tag.Equals("CHECKEDBY"))
//                                    {
//                                        attRef.TextString = "";
//                                    }
//                                    else if (attRef.Tag.Equals("DESIGNEDBY"))
//                                    {
//                                        attRef.TextString = @"ACD";
//                                    }
//                                }

//                                foreach (ObjectId attId in blkRef.AttributeCollection)
//                                {
//                                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

//                                    if (attRef.Tag == "DRAWING_NUM")
//                                    {
//                                        attRef.TextString = @"%<\AcObjProp Object(%<\_ObjId " + projnum + @">%).TextString>%" + "-" + @"%<\AcObjProp Object(%<\_ObjId " + type + @">%).TextString>%" + "-" + @"%<\AcObjProp Object(%<\_ObjId " + line + @">%).TextString>%" + @"%<\AcObjProp Object(%<\_ObjId " + sheetnum + @">%).TextString>%";
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }

//            }
//            transaction.Commit();
//        }
//    }
//    doc.SendStringToExecute("Regenall ", true, false, false);
//}

//[CommandMethod("PlanviewTable")]
//public void PlanviewTable()
//{
//    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
//    Database db = doc.Database;
//    Editor ed = doc.Editor;

//    PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
//    pKeyOpts.Message = "\nSelect the block type ";
//    pKeyOpts.Keywords.Add("ElectricalDrop");
//    pKeyOpts.Keywords.Add("EThernet");
//    pKeyOpts.Keywords.Add("ArmorStart");
//    pKeyOpts.Keywords.Add("DIsconnect");
//    pKeyOpts.Keywords.Add("OperatorStation");
//    pKeyOpts.Keywords.Add("JunctionBox");
//    pKeyOpts.Keywords.Add("Receptacle");
//    pKeyOpts.Keywords.Add("DEviceInputs");
//    pKeyOpts.Keywords.Add("DeviceOutputs");
//    pKeyOpts.AllowNone = true;

//    PromptResult pr = ed.GetKeywords(pKeyOpts);

//    if (pr.Status != PromptStatus.OK)
//    {
//        ed.WriteMessage("No block was provided.\n");
//        return;
//    }

//    ed.WriteMessage("\nCommand got the following block: {0}\n", pr.StringResult);

//    string blockToFind = pr.StringResult;

//    Transaction tr = doc.TransactionManager.StartTransaction();
//    using (tr)
//    {
//        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);

//        if (!bt.Has(blockToFind))
//        {
//            ed.WriteMessage("\nBlock " + blockToFind + " does not exist.");
//        }
//        else
//        {
//            StringCollection colNames = new StringCollection();

//            BlockTableRecord bd = (BlockTableRecord)tr.GetObject(bt[blockToFind], OpenMode.ForRead);
//            foreach (ObjectId adId in bd)
//            {
//                DBObject adObj = tr.GetObject(adId, OpenMode.ForRead);

//                AttributeDefinition ad = adObj as AttributeDefinition;
//                if (ad != null)
//                {
//                    colNames.Add(ad.Tag);
//                }
//            }
//            if (colNames.Count == 0)
//            {
//                ed.WriteMessage("\nThe block " + blockToFind + " contains no attributes definitions.");
//            }
//            else
//            {
//                PromptPointResult ppr = ed.GetPoint("\nEnter table insertion point: ");

//                if (ppr.Status == PromptStatus.OK)
//                {
//                    Table tb = new Table();
//                    tb.TableStyle = db.Tablestyle;
//                    tb.SetSize(1, colNames.Count);
//                    tb.SetRowHeight(rowHeight);
//                    tb.SetColumnWidth(colWidth);
//                    tb.Position = ppr.Value;

//                    //i starts at 1 to avoid the BLOCKNAME attribute used to identify blocks in this program
//                    for (int i = 1; i < colNames.Count; i++)
//                    {
//                        SetCellText(tb, 0, i, colNames[i]);
//                    }

//                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

//                    int rowNum = 1;
//                    foreach (ObjectId objId in ms)
//                    {
//                        DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
//                        BlockReference br = obj as BlockReference;
//                        if (br != null)
//                        {
//                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
//                            using (btr)
//                            {
//                                string blockName = "";
//                                foreach (ObjectId arCheck in br.AttributeCollection)
//                                {
//                                    DBObject arOne = tr.GetObject(arCheck, OpenMode.ForRead);

//                                    AttributeReference arRef = arOne as AttributeReference;
//                                    if (arRef != null && arRef.Tag == "BLOCKNAME")
//                                    {
//                                        blockName = arRef.TextString;
//                                        break;
//                                    }
//                                }

//                                if (blockName.ToUpper() == blockToFind.ToUpper())
//                                {
//                                    tb.InsertRows(rowNum, rowHeight, 1);

//                                    int attNum = 0;
//                                    foreach (ObjectId arId in br.AttributeCollection)
//                                    {
//                                        DBObject arObj = tr.GetObject(arId, OpenMode.ForRead);

//                                        AttributeReference ar = arObj as AttributeReference;
//                                        if (ar != null)
//                                        {
//                                            string strCell;
//                                            string strArId = arId.ToString();
//                                            strArId = strArId.Trim(new char[] { '(', ')' });
//                                            strCell = "%<\\AcObjProp Object(" + "%<\\_ObjId " + strArId + ">%).TextString>%";
//                                            SetCellText(tb, rowNum, attNum, strCell);
//                                        }
//                                        attNum++;
//                                    }
//                                    rowNum++;
//                                }
//                            }
//                        }
//                    }
//                    tb.GenerateLayout();
//                    tb.UnmergeCells(tb.Rows[0]);

//                    ms.UpgradeOpen();
//                    ms.AppendEntity(tb);
//                    tr.AddNewlyCreatedDBObject(tb, true);
//                    tr.Commit();
//                }
//            }
//        }
//    }
//}

//[CommandMethod("PlanviewExtract")]
//public void PlanviewExtract()
//{
//    List<string> symbols = new List<string>();
//    List<Tuple<string, string, string, string, string, string, string>> items = new List<Tuple<string, string, string, string, string, string, string>>();

//    string voltage = "";
//    string nameHolder = "";
//    string connection = "";
//    string rating = "";
//    string layer = "";
//    string unique_id = "";

//    symbols.Add("ElectricalDrop");
//    symbols.Add("OperatorStation");
//    symbols.Add("JunctionBox");
//    symbols.Add("Receptacle");
//    symbols.Add("Ethernet");
//    symbols.Add("ArmorStart");
//    symbols.Add("DeviceInputs");
//    symbols.Add("DeviceOutputs");
//    symbols.Add("Disconnect");
//    symbols.Add("Photoeye");
//    symbols.Add("Motor");

//    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
//    Database db = doc.Database;
//    Editor ed = doc.Editor;

//    Transaction tr = db.TransactionManager.StartTransaction();
//    using (tr)
//    {
//        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

//        int i = 0;
//        foreach (string blockname in symbols)
//        {
//            try
//            {
//                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockname], OpenMode.ForRead);
//                ObjectIdCollection objIdCol = btr.GetBlockReferenceIds(true, false);
//                ObjectIdCollection annobjIdCol = btr.GetAnonymousBlockIds();

//                int j = 0;
//                foreach (ObjectId objId in objIdCol)
//                {
//                    BlockReference br = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);
//                    layer = br.Layer;

//                    if (br != null)
//                    {
//                        foreach (ObjectId arId in br.AttributeCollection)
//                        {
//                            DBObject obj = tr.GetObject(arId, OpenMode.ForWrite);
//                            AttributeReference ar = (AttributeReference)obj;

//                            if (ar != null)
//                            {
//                                if (ar.Tag.ToUpper().Equals("VOLTAGE"))
//                                {
//                                    voltage = ar.TextString;
//                                }
//                                else if (ar.Tag.ToUpper().Equals("NAME"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        nameHolder = "None";
//                                    }
//                                    else
//                                    {
//                                        nameHolder = ar.TextString;
//                                    }
//                                }
//                                else if (ar.Tag.ToUpper().Equals("CONNECTION"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        connection = "None";
//                                    }
//                                    else
//                                    {
//                                        connection = ar.TextString;
//                                    }
//                                }
//                                else if (ar.Tag.ToUpper().Equals("RATING"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        rating = "None";
//                                    }
//                                    else
//                                    {
//                                        rating = ar.TextString;
//                                    }
//                                }
//                                else if (ar.Tag.ToUpper().Equals("ID"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        long timestamp = (long)(DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds;
//                                        ar.TextString = timestamp.ToString() + i.ToString() + j.ToString();
//                                    }

//                                    unique_id = ar.TextString;
//                                }
//                            }
//                        }
//                    }

//                    items.Add(new Tuple<string, string, string, string, string, string, string>(blockname, voltage, nameHolder, connection, rating, layer, unique_id));
//                    j++;
//                }

//                foreach (ObjectId objId in annobjIdCol)
//                {
//                    BlockTableRecord annbtr = (BlockTableRecord)tr.GetObject(objId, OpenMode.ForRead);
//                    ObjectIdCollection blkRefIds = annbtr.GetBlockReferenceIds(true, false);

//                    foreach (ObjectId id in blkRefIds)
//                    {
//                        BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
//                        layer = br.Layer;

//                        if (br != null)
//                        {
//                            foreach (ObjectId arId in br.AttributeCollection)
//                            {
//                                DBObject obj = tr.GetObject(arId, OpenMode.ForWrite);
//                                AttributeReference ar = (AttributeReference)obj;

//                                if (ar.Tag.ToUpper().Equals("VOLTAGE"))
//                                {
//                                    voltage = ar.TextString;
//                                }
//                                else if (ar.Tag.ToUpper().Equals("NAME"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        nameHolder = "None";
//                                    }
//                                    else
//                                    {
//                                        nameHolder = ar.TextString;
//                                    }
//                                }
//                                else if (ar.Tag.ToUpper().Equals("CONNECTION"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        connection = "None";
//                                    }
//                                    else
//                                    {
//                                        connection = ar.TextString;
//                                    }
//                                }
//                                else if (ar.Tag.ToUpper().Equals("RATING"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        rating = "None";
//                                    }
//                                    else
//                                    {
//                                        rating = ar.TextString;
//                                    }
//                                }
//                                else if (ar.Tag.ToUpper().Equals("ID"))
//                                {
//                                    if (String.IsNullOrEmpty(ar.TextString))
//                                    {
//                                        long timestamp = (long)(DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds;
//                                        ar.TextString = timestamp.ToString() + i.ToString() + j.ToString();
//                                    }

//                                    unique_id = ar.TextString;
//                                }
//                            }
//                        }

//                        items.Add(new Tuple<string, string, string, string, string, string, string>(blockname, voltage, nameHolder, connection, rating, layer, unique_id));
//                        j++;
//                    }
//                }
//            }
//            catch
//            {
//                continue;
//            }

//            i++;
//        }

//        tr.Commit();
//    }

//    if (items.Count > 0)
//    {
//        Excel.Application ExcelApp = new Excel.Application();
//        ExcelApp.Visible = true;
//        Excel.Workbook ExcelWB = ExcelApp.Workbooks.Add();
//        Excel.Sheets ExcelSheets = ExcelWB.Sheets;

//        for (int i = 0; i < items.Count; i++)
//        {
//            ExcelSheets[1].Cells[i + 1, 1].Value = items[i].Item1;  //blockname
//            ExcelSheets[1].Cells[i + 1, 2].Value = items[i].Item2;  //voltage
//            ExcelSheets[1].Cells[i + 1, 3].Value = items[i].Item3;  //name
//            ExcelSheets[1].Cells[i + 1, 4].Value = items[i].Item4;  //connection
//            ExcelSheets[1].Cells[i + 1, 5].Value = items[i].Item5;  //rating
//            ExcelSheets[1].Cells[i + 1, 6].Value = items[i].Item6;  //layer
//            ExcelSheets[1].Cells[i + 1, 7].Value = items[i].Item7;  //id
//        }
//    }
//}

//[CommandMethod("WriteSOW")]
//public void WriteSOW()
//{
//    List<string> items = new List<string>();
//    List<string> lines = new List<string>();
//    string writePath;
//    string fileName = "Scope of Work Items";

//    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
//    Database db = doc.Database;
//    Editor ed = doc.Editor;

//    PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
//    pKeyOpts.Message = "\nSelect the block type ";
//    pKeyOpts.Keywords.Add("ElectricalDrop");
//    pKeyOpts.Keywords.Add("EThernet");
//    pKeyOpts.Keywords.Add("ArmorStart");
//    pKeyOpts.Keywords.Add("DIsconnect");
//    pKeyOpts.Keywords.Add("OperatorStation");
//    pKeyOpts.Keywords.Add("JunctionBox");
//    pKeyOpts.Keywords.Add("Receptacle");
//    pKeyOpts.Keywords.Add("DEviceInputs");
//    pKeyOpts.Keywords.Add("DeviceOutputs");
//    pKeyOpts.AllowNone = true;

//    PromptResult pr = ed.GetKeywords(pKeyOpts);
//    if (pr.Status != PromptStatus.OK || pr.StringResult == null)
//    {
//        ed.WriteMessage("No block was provided.\n");
//        return;
//    }

//    ed.WriteMessage("\nCommand got the following block: {0}\n", pr.StringResult.ToUpper());

//    string blockToFind = pr.StringResult;
//    fileName = blockToFind.ToUpper() + " " + fileName;

//    Transaction tr = doc.TransactionManager.StartTransaction();
//    using (tr)
//    {
//        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);

//        if (!bt.Has(blockToFind))
//        {
//            ed.WriteMessage("\nBlock " + blockToFind + " does not exist.");
//        }
//        else
//        {
//            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

//            foreach (ObjectId objId in ms)
//            {
//                DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
//                BlockReference br = obj as BlockReference;
//                if (br != null)
//                {
//                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
//                    using (btr)
//                    {
//                        string blockName = "";
//                        foreach (ObjectId arCheck in br.AttributeCollection)
//                        {
//                            DBObject arOne = tr.GetObject(arCheck, OpenMode.ForRead);

//                            AttributeReference arRef = arOne as AttributeReference;
//                            if (arRef != null && arRef.Tag == "BLOCKNAME")
//                            {
//                                blockName = arRef.TextString;
//                                break;
//                            }
//                        }

//                        if (blockName.ToUpper() == blockToFind.ToUpper())
//                        {
//                            foreach (ObjectId arId in br.AttributeCollection)
//                            {
//                                string strArId;
//                                DBObject arObj = tr.GetObject(arId, OpenMode.ForRead);

//                                AttributeReference ar = arObj as AttributeReference;
//                                if (ar != null)
//                                {
//                                    strArId = ar.TextString;
//                                    items.Add(strArId);
//                                }
//                                //Handles conditions of unequal number of attributes within the AutoCAD blocks to avoid indexing error
//                                //in string.Format                                        
//                            }
//                            for (int i = 7 - items.Count; i > 0; i--)
//                            {
//                                items.Add(@"N/A");
//                            }
//                            string SOWitem = SOWCatalog(blockToFind.ToUpper());
//                            SOWitem = string.Format(SOWitem, items[2], items[3], items[4], items[5], items[6]);
//                            lines.Add(SOWitem);
//                            items.Clear();
//                        }
//                    }
//                }
//            }
//            Autodesk.AutoCAD.Windows.SaveFileDialog sfd = new Autodesk.AutoCAD.Windows.SaveFileDialog("File Save As", fileName, "txt", "Saving File", Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.DefaultIsFolder);

//            if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
//            {
//                ed.WriteMessage("\nThere was an error in saving the text file, please try again.");
//                return;
//            }
//            else
//            {
//                writePath = sfd.Filename;
//            }

//            try
//            {
//                using (System.IO.StreamWriter file = new System.IO.StreamWriter(writePath))
//                {
//                    for (int i = 0; i < lines.Count; i++)
//                    {
//                        file.WriteLine((i + 1).ToString() + ") " + lines[i]);
//                        file.WriteLine(" ");
//                    }
//                    lines.Clear();
//                    file.Close();
//                }
//            }
//            catch (IOException e)
//            {
//                ed.WriteMessage(string.Format("\n{0}: The file could not be saved.", e.GetType().Name));
//            }
//        }
//        tr.Commit();
//    }
//}

//[CommandMethod("WireNumbers")]
//public void WireNumbers()
//{
//    ////////////DISPLAYS THE DIALOG BOX//////////////////////////
//    Wire_Number_Dialogs.Wire_Number_Form WNF = new Wire_Number_Dialogs.Wire_Number_Form();
//    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(null, WNF, false);
//    ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^//////////////////////////////////

//}

//[CommandMethod("ComponentNumbers")]
//public void ComponentNumbers()
//{
//    ////////////DISPLAYS THE DIALOG BOX//////////////////////////
//    Component_Number_Dialogs.Component_Number_Form CNF = new Component_Number_Dialogs.Component_Number_Form();
//    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(null, CNF, false);
//    ///////////////^^^^^^^^^^^^^^^^^^^^^^^^^^//////////////////////////////////          
//}