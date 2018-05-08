using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ACSMCOMPONENTS22Lib;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Footprint_Lookup_Dialogs
{   
    public partial class Footprint_Lookup_Form : Form
    {
        Document doc;
        Database db;
        Editor ed;

        List<string> partNumbers = new List<string>();
        List<string> description = new List<string>();
        List<string> drawing = new List<string>();
        List<int> index = new List<int>();
        
        string partNumberToAdd;               
        string descriptionToAdd;
        string drawingToAdd;

        BlockTable bt;
        BlockTableRecord btr;

        public Footprint_Lookup_Form()
        {
            InitializeComponent();
        }

        private void Footprint_Lookup_Form_Load(object sender, EventArgs e)
        {
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            btn_Add.Enabled = false;

            try
            {
                using (FileStream stream = File.Open(@"\\thehaskellco.net\Remote\AtlantaData\_Disciplines\Controls\4 Controls Programming\DVBs and DLLs\Footprints\footprint_lookup.csv", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(',');

                            partNumbers.Add(values[0]);
                            description.Add(values[3]);
                            drawing.Add(values[2]);
                        }
                    }
                }

                //////////////DEBUG///////////////////
                ed.WriteMessage(partNumbers[5] + "/n");

                
            }
            catch
            {

            }
        }

        private void previewKeyDown_txtbx_partNumSearch(object sender, PreviewKeyDownEventArgs e)
        {           
            if(e.KeyCode == Keys.Enter)
            {
                cmbbx_partNumbers.Items.Clear();
                index.Clear();
                //Loop through each part number, if it contains the search content, add it to the combo box
                for (int i = 0; i < partNumbers.Count; i++)
                {
                    if (partNumbers[i].Contains(txtbx_partNumSearch.Text.ToUpper()))
                    {
                        cmbbx_partNumbers.Items.Add(partNumbers[i]);
                        index.Add(i);       //Index is saved to be used to retrieve the descriptions on selection
                    }
                }
            }            
        }

        private void click_btn_Search(object sender, EventArgs e)
        {
            cmbbx_partNumbers.Items.Clear();
            index.Clear();
            //Loop through each part number, if it contains the search content, add it to the combo box
            for(int i = 0; i < partNumbers.Count; i++)
            {
                if(partNumbers[i].Contains(txtbx_partNumSearch.Text.ToUpper()))
                {
                    cmbbx_partNumbers.Items.Add(partNumbers[i]);
                    index.Add(i);       //Index is saved to be used to retrieve the descriptions on selection
                }
            }
                
        }
        
        private void txtChanged_cmbBx_PartNum(object sender, EventArgs e)
        {
            //After text changed, save that as part number, add the description, and add the drawing from the database
            partNumberToAdd = cmbbx_partNumbers.Text;
            descriptionToAdd = description[index[cmbbx_partNumbers.Items.IndexOf(cmbbx_partNumbers.Text)]];
            drawingToAdd = drawing[index[cmbbx_partNumbers.Items.IndexOf(cmbbx_partNumbers.Text)]];
            
            txtBx_Description.Text = descriptionToAdd;

            if((partNumberToAdd != null && partNumberToAdd != "") && ((descriptionToAdd != null && descriptionToAdd != "") || (drawingToAdd != null && drawingToAdd != "")))
            {
                btn_Add.Enabled = true;
            }
        }

        private void click_btn_AddPart(object sender, EventArgs e)
        {           
            DocumentLock loc = doc.LockDocument();

            using (loc)
            {
                drawingToAdd.Replace(@"/", @"\");
                string drawingFileName = drawingToAdd.Substring(drawingToAdd.IndexOf(@"/", (drawingToAdd.IndexOf(@"/") + 1)) + 1, drawingToAdd.Length - drawingToAdd.IndexOf(@"/", (drawingToAdd.IndexOf(@"/") + 1)) - 1);

                Database tmpDb = new Database(false, true);
                tmpDb.ReadDwgFile(@"\\thehaskellco.net\Remote\AtlantaData\_Disciplines\Controls\4 Controls Programming\DVBs and DLLs\Footprints\Panel\" + drawingToAdd + ".dwg", System.IO.FileShare.Read, true, "");
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    if (!bt.Has(drawingFileName))
                    {
                        db.Insert(drawingFileName, tmpDb, true);
                    }
                    tr.Commit();
                }                

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    btr = (BlockTableRecord)bt[drawingFileName].GetObject(OpenMode.ForRead);
                    bool tagFound = false;

                    if(btr != null)
                    {
                        foreach(ObjectId objId in btr)
                        {
                            AttributeDefinition attDef = tr.GetObject(objId, OpenMode.ForRead) as AttributeDefinition;
                            if (attDef != null)
                            {
                                if(attDef.Tag.ToUpper() == "TABLE_PART_NUM")
                                {
                                    tagFound = true;
                                }
                                else if (attDef.Tag.ToUpper() == "TABLE_DESCRIPTION")
                                {
                                    tagFound = true;
                                }
                            }
                        }
                        try
                        {
                            if (!tagFound)
                            {
                                AttributeDefinition acAttDef = new AttributeDefinition();

                                acAttDef.Position = new Point3d(0, 0, 0);
                                acAttDef.Prompt = "Part Number?";
                                acAttDef.Tag = "TABLE_PART_NUM";
                                acAttDef.Invisible = true;
                                acAttDef.Height = 1;
                                acAttDef.Justify = AttachmentPoint.MiddleCenter;

                                AttributeDefinition acAttDef2 = new AttributeDefinition();

                                acAttDef2.Position = new Point3d(0, 0, 0);
                                acAttDef2.Prompt = "Description?";
                                acAttDef2.Tag = "TABLE_DESCRIPTION";
                                acAttDef2.Invisible = true;
                                acAttDef2.Height = 1;
                                acAttDef2.Justify = AttachmentPoint.MiddleCenter;

                                AttributeDefinition acAttDef3 = new AttributeDefinition();

                                acAttDef3.Position = new Point3d(0, -10, 0);
                                acAttDef3.Prompt = "DO NOT CHANGE";
                                acAttDef3.Tag = "AUTO_UPDATE";
                                acAttDef3.TextString = "1";
                                acAttDef3.Invisible = true;
                                acAttDef3.Height = 1;
                                acAttDef3.Justify = AttachmentPoint.MiddleCenter;

                                btr.UpgradeOpen();
                                btr.AppendEntity(acAttDef);
                                btr.AppendEntity(acAttDef2);
                                btr.AppendEntity(acAttDef3);
                                tr.AddNewlyCreatedDBObject(acAttDef, true);
                                tr.AddNewlyCreatedDBObject(acAttDef2, true);
                                tr.AddNewlyCreatedDBObject(acAttDef3, true);
                                btr.DowngradeOpen();
                                tr.Commit();
                            }
                        }
                        catch
                        {
                            ed.WriteMessage("Could not add attributes to block.");
                        }
                    }                    
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    if (btr != null)
                    { 

                        MyBlockJig blockJig = new MyBlockJig();
                        Point3d point;
                        PromptResult res = blockJig.DragMe(btr.ObjectId, out point);

                        if (res.Status == PromptStatus.OK)
                        {
                            BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                            BlockReference insert = new BlockReference(point, btr.ObjectId);

                            curSpace.AppendEntity(insert);
                            tr.AddNewlyCreatedDBObject(insert, true);

                            foreach (ObjectId id in btr)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForWrite);
                                AttributeDefinition attDef = obj as AttributeDefinition;
                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attDef, insert.BlockTransform);                                                                                                                     
                                        insert.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);
                                    }                                    
                                }
                            }
                        }
                        
                    }
                    tr.Commit();
                    doc.SendStringToExecute("regenall ", true, false, false);
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    ObjectId msId = bt[BlockTableRecord.ModelSpace];

                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = (Entity)tr.GetObject(entId, OpenMode.ForWrite);
                        if (ent != null)
                        {
                            BlockReference br = ent as BlockReference;
                            if (br != null)
                            {
                                BlockTableRecord bd = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                                if (bd.Name.ToUpper() == drawingFileName.ToUpper())
                                {
                                    foreach (ObjectId arId in br.AttributeCollection)
                                    {
                                        DBObject obj = tr.GetObject(arId, OpenMode.ForWrite);
                                        AttributeReference ar = obj as AttributeReference;
                                        if (ar != null)
                                        {
                                            if (ar.Tag == "AUTO_UPDATE")
                                            {
                                                if(ar.TextString == "1")
                                                {
                                                    foreach (ObjectId attRefId in br.AttributeCollection)
                                                    {
                                                        DBObject objRef = tr.GetObject(attRefId, OpenMode.ForWrite);
                                                        AttributeReference aRef = objRef as AttributeReference;
                                                        if (aRef.Tag == "TABLE_PART_NUM")
                                                        {
                                                            aRef.UpgradeOpen();
                                                            aRef.TextString = partNumberToAdd;
                                                            aRef.DowngradeOpen();
                                                        }
                                                        else if (aRef.Tag == "TABLE_DESCRIPTION")
                                                        {
                                                            aRef.UpgradeOpen();
                                                            aRef.TextString = descriptionToAdd;
                                                            aRef.DowngradeOpen();
                                                        }
                                                    }
                                                }                                              

                                                ar.TextString = "0";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }
    }

    // This Jig will show the given block during the dragging.
    //=================================================================
    public class MyBlockJig : DrawJig
    {
        public Point3d _point;
        private ObjectId _blockId = ObjectId.Null;
 
        // Shows the block until the user clicks a point.
        // The 1st parameter is the Id of the block definition.
        // The 2nd is the clicked point.
        //---------------------------------------------------------------
        public PromptResult DragMe(ObjectId i_blockId, out Point3d o_pnt)
        {
            _blockId = i_blockId;
            Editor ed =
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
 
            PromptResult jigRes = ed.Drag(this);
            o_pnt = _point;
            return jigRes;
        } 
 
        // Need to override this method.
        // Updating the current position of the block.
        //--------------------------------------------------------------
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jigOpts = new JigPromptPointOptions();
            jigOpts.UserInputControls =
                                (UserInputControls.Accept3dCoordinates |
                                UserInputControls.NullResponseAccepted);
            jigOpts.Message = "Select a point:";
            PromptPointResult jigRes = prompts.AcquirePoint(jigOpts);
 
            Point3d pt = jigRes.Value;
            if (pt == _point)
            return SamplerStatus.NoChange;
 
            _point = pt;
            if (jigRes.Status == PromptStatus.OK)
            return SamplerStatus.OK;
 
            return SamplerStatus.Cancel;
        } 
 
        // Need to override this method.
        // We are showing our block in its current position here.
        //--------------------------------------------------------------
        protected override bool WorldDraw(
                        Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            BlockReference inMemoryBlockInsert =
                                    new BlockReference(_point, _blockId);
            draw.Geometry.Draw(inMemoryBlockInsert);
 
            inMemoryBlockInsert.Dispose();
 
            return true;
        } // WorldDraw() 
    } // class BlockJig
}
