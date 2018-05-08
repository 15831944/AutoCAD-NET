using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ACSMCOMPONENTS22Lib;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System.IO;

namespace IO_Extract_Dialogs
{
    public partial class IO_Extract_Form_Progress : Form
    {
        public Document doc { get; set; }
        public Database db { get; set; }
        public Editor ed { get; set; }
        public List<string> fileList { get; set; }
        public bool smartSearch { get; set; }

        int progress;
        int numFiles;

        BackgroundWorker bw = new BackgroundWorker();

        List<IO_Block> IO_Blocks = new List<IO_Block>();

        List<string> IO_CardNames = new List<string>();

        public IO_Extract_Form_Progress()
        {
            InitializeComponent();
            bw.WorkerSupportsCancellation = true;
            bw.WorkerReportsProgress = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);              
        }

        private void load_FormProgress(object sender, EventArgs e)
        {
            numFiles = fileList.Count;
            progress = 0;
            lbl_Progress.Text = String.Format("Processing Drawing {0} of {1}", progress + 1, numFiles);

            IO_CardNames.Add("1756-IB16");
            IO_CardNames.Add("1756-IB32");
            IO_CardNames.Add("1756-IF16");
            IO_CardNames.Add("1756-OF8");
            IO_CardNames.Add("1756-OB16");
            IO_CardNames.Add("1756-OW16I");

            bw.RunWorkerAsync();
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            foreach(string filePath in fileList)
            {
                progress++;                               

                Database sideDB = new Database(false, true);

                using (sideDB)
                {
                    try
                    {                        
                        sideDB.ReadDwgFile(filePath, System.IO.FileShare.ReadWrite, false, "");
                    }
                    catch (System.Exception)
                    {
                        ed.WriteMessage("Could not open drawing file.");
                        return;
                    }

                    Transaction tr = sideDB.TransactionManager.StartTransaction();
                    using(tr)
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(sideDB.BlockTableId, OpenMode.ForRead);

                        if (smartSearch)
                        {
                            int numObjects = IO_CardNames.Count;
                            int currentObj = 0;

                            foreach (string name in IO_CardNames)
                            {
                                currentObj++;
                                bw.ReportProgress((int)(((float)currentObj / numObjects) * ((float)1 / numFiles) * 100 + ((float)(progress - 1) / numFiles) * 100));

                                try
                                {
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[name], OpenMode.ForRead);
                                    ObjectIdCollection objIdCol = btr.GetBlockReferenceIds(true, false);
                                    ObjectIdCollection annobjIdCol = btr.GetAnonymousBlockIds();

                                    foreach (ObjectId objId in objIdCol)
                                    {
                                        bool IO_block = false;
                                        string[] portDescriptions = new string[32];
                                        string rack = "";
                                        string slot = "";
                                        string partNumber = "";

                                        try
                                        {
                                            BlockReference br = (BlockReference)tr.GetObject(objId, OpenMode.ForRead);
                                            if (br != null)
                                            {
                                                foreach (ObjectId arId in br.AttributeCollection)
                                                {
                                                    DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                                    AttributeReference ar = (AttributeReference)obj;

                                                    if (ar != null)
                                                    {
                                                        if (ar.Tag.ToUpper().Equals("RACK"))
                                                        {
                                                            rack = ar.TextString;
                                                        }
                                                        else if (ar.Tag.ToUpper().Equals("SLOT"))
                                                        {
                                                            slot = ar.TextString;
                                                        }
                                                        else if (ar.Tag.ToUpper().Equals("PARTNUMBER"))
                                                        {
                                                            partNumber = ar.TextString;
                                                        }
                                                        else
                                                        {
                                                            for (int i = 0; i < 32; i++)
                                                            {
                                                                if (ar.Tag.ToUpper().Contains("PORT" + i + "DESC"))
                                                                {
                                                                    IO_block = true;
                                                                    portDescriptions[i] = ar.TextString;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                if (IO_block)
                                                {
                                                    IO_Block to_Store = new IO_Block();
                                                    to_Store.partNumber = partNumber;
                                                    to_Store.rack = rack;
                                                    to_Store.slot = slot;
                                                    to_Store.portDescriptions = portDescriptions;
                                                    IO_Blocks.Add(to_Store);
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                        if (bw.CancellationPending)
                                        {
                                            e.Cancel = true;
                                            return;
                                        }
                                    }

                                    foreach (ObjectId id in annobjIdCol)
                                    {
                                        BlockTableRecord annbtr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                                        ObjectIdCollection blkRefIds = annbtr.GetBlockReferenceIds(true, false);
                                        foreach (ObjectId objId in blkRefIds)
                                        {
                                            bool IO_block = false;
                                            string[] portDescriptions = new string[40];
                                            string rack = "";
                                            string slot = "";
                                            string partNumber = "";

                                            try
                                            {
                                                BlockReference br = (BlockReference)tr.GetObject(objId, OpenMode.ForRead);
                                                if (br != null)
                                                {
                                                    foreach (ObjectId arId in br.AttributeCollection)
                                                    {
                                                        DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                                        AttributeReference ar = (AttributeReference)obj;

                                                        if (ar != null)
                                                        {
                                                            if (ar.Tag.ToUpper().Equals("RACK"))
                                                            {
                                                                rack = ar.TextString;
                                                            }
                                                            else if (ar.Tag.ToUpper().Equals("SLOT"))
                                                            {
                                                                slot = ar.TextString;
                                                            }
                                                            else if (ar.Tag.ToUpper().Equals("PARTNUMBER"))
                                                            {
                                                                partNumber = ar.TextString;
                                                            }
                                                            else
                                                            {
                                                                for (int i = 0; i < 40; i++)
                                                                {
                                                                    if (ar.Tag.ToUpper().Contains("PORT" + i + "DESC"))
                                                                    {
                                                                        IO_block = true;
                                                                        portDescriptions[i] = ar.TextString;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (IO_block)
                                                    {
                                                        IO_Block to_Store = new IO_Block();
                                                        to_Store.partNumber = partNumber;
                                                        to_Store.rack = rack;
                                                        to_Store.slot = slot;
                                                        to_Store.portDescriptions = portDescriptions;
                                                        IO_Blocks.Add(to_Store);
                                                    }
                                                }
                                            }
                                            catch
                                            {
                                                continue;
                                            }

                                            if (bw.CancellationPending)
                                            {
                                                e.Cancel = true;
                                                return;
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            int numObjects = 0;
                            int currentObj = 0;
                            
                            try
                            {
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                                foreach(ObjectId id in btr)
                                {
                                    numObjects++;
                                }

                                foreach (ObjectId objId in btr)
                                {
                                    currentObj++;
                                    bw.ReportProgress((int)(((float)currentObj / numObjects) * ((float)1 / numFiles) * 100 + ((float)(progress - 1) / numFiles) * 100));

                                    bool IO_block = false;
                                    string[] portDescriptions = new string[32];
                                    string rack = "";
                                    string slot = "";
                                    string partNumber = "";

                                    try
                                    {
                                        BlockReference br = (BlockReference)tr.GetObject(objId, OpenMode.ForRead);
                                        if (br != null)
                                        {
                                            foreach (ObjectId arId in br.AttributeCollection)
                                            {
                                                DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                                AttributeReference ar = (AttributeReference)obj;

                                                if (ar != null)
                                                {
                                                    if (ar.Tag.ToUpper().Equals("RACK"))
                                                    {
                                                        rack = ar.TextString;
                                                    }
                                                    else if (ar.Tag.ToUpper().Equals("SLOT"))
                                                    {
                                                        slot = ar.TextString;
                                                    }
                                                    else if (ar.Tag.ToUpper().Equals("PARTNUMBER"))
                                                    {
                                                        partNumber = ar.TextString;
                                                    }
                                                    else
                                                    {
                                                        for (int i = 0; i < 32; i++)
                                                        {
                                                            if (ar.Tag.ToUpper().Contains("PORT" + i + "DESC"))
                                                            {
                                                                IO_block = true;
                                                                portDescriptions[i] = ar.TextString;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if (IO_block)
                                            {
                                                IO_Block to_Store = new IO_Block();
                                                to_Store.partNumber = partNumber;
                                                to_Store.rack = rack;
                                                to_Store.slot = slot;
                                                to_Store.portDescriptions = portDescriptions;
                                                IO_Blocks.Add(to_Store);
                                            }
                                        }

                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                    if (bw.CancellationPending)
                                    {
                                        e.Cancel = true;
                                        return;
                                    }
                                }
                            }
                            catch
                            {
                                continue;
                            }                            
                        }
                    }                                        
                }                               
            }
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lbl_Progress.Text = String.Format("Processing Drawing {0} of {1}", progress, numFiles);
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Visible = false;
            progressBar.Value = 0;            
            IO_Extract_Dialogs.IO_Extract_Form_Selections IEFS = new IO_Extract_Dialogs.IO_Extract_Form_Selections();
            IEFS.IO_Blocks = IO_Blocks;
            this.Visible = false;
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(null, IEFS, false);            
        }

        private void click_Cancel(object sender, EventArgs e)
        {
            bw.CancelAsync();
        }        
    }

    public class IO_Block
    {
        public string[] portDescriptions { get; set; }
        public string partNumber { get; set; }
        public string rack { get; set; }
        public string slot { get; set; }
    }
}
