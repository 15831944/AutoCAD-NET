using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Update_Title_Block_Dialogs
{
    public partial class Update_Title_Block_Form : Form
    {
        const string TITLE_BLOCK_NAME = "Haskell ARCH D";
        readonly string[] STAMPS = { "Preliminary", "Review in Progress", "For Review",
            "DPP", "CPP", "Design Review", "RFQ", "Detailed Design", "IFBB", "Frozen",
            "Frozen for IFB", "IFB", "Frozen for IFC", "IFC", "Start-up", "As-Builts",
            "Record Set", "AFDD", "Permit", "Reference", "Sample", "Demo", "Not Released",
            "For Information Only", "Work in Progress" };

        Document doc;
        Database db;
        Editor ed;

        BackgroundWorker bw_WriteTitleBlocks = new BackgroundWorker();
        BackgroundWorker bw_ReadTitleBlocks = new BackgroundWorker();

        List<Title_Block> Title_Blocks = new List<Title_Block>();
        Title_Block loadedToRevInfo;

        System.Data.DataTable datatable_titleblocks = new System.Data.DataTable();
        BindingSource bindSource_titleblocks = new BindingSource();
        System.Data.DataTable datatable_revisioninfo = new System.Data.DataTable();
        BindingSource bindSource_revisioninfo = new BindingSource();

        DataGridViewColumn blkRefIDcol = new DataGridViewTextBoxColumn();
        DataGridViewColumn pageNumbercol = new DataGridViewTextBoxColumn();
        DataGridViewColumn pageTitlecol = new DataGridViewTextBoxColumn();
        DataGridViewColumn projNumcol = new DataGridViewTextBoxColumn();
        DataGridViewColumn drawnBycol = new DataGridViewTextBoxColumn();
        DataGridViewColumn checkedBycol = new DataGridViewTextBoxColumn();
        DataGridViewComboBoxColumn stampcol = new DataGridViewComboBoxColumn();
        DataGridViewColumn revNumcol = new DataGridViewTextBoxColumn();

        DataGridViewColumn revInfoNumcol = new DataGridViewTextBoxColumn();
        DataGridViewColumn revInfoTitlecol = new DataGridViewTextBoxColumn();        
        DataGridViewColumn revInfoDatecol = new DataGridViewTextBoxColumn();

        public Update_Title_Block_Form()
        {
            InitializeComponent();
            SetupDataGrids();
            SetupComboBoxes();
            SetupButtons();

            btn_Clear_SelectedRevInfo.Enabled = false;

            //Initializes the background worker to read the existing title block data
            bw_ReadTitleBlocks.WorkerSupportsCancellation = false;
            bw_ReadTitleBlocks.WorkerReportsProgress = true;
            bw_ReadTitleBlocks.DoWork += new DoWorkEventHandler(bw_ReadTitleBlocks_DoWork);
            bw_ReadTitleBlocks.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_ReadTitleBlocks_RunWorkerCompleted);
            bw_ReadTitleBlocks.ProgressChanged += new ProgressChangedEventHandler(bw_ReadTitleBlocks_ProgressChanged);
        }

        #region FormEvents
        private void Update_Title_Block_Form_Load(object sender, EventArgs e)
        {
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            bw_ReadTitleBlocks.RunWorkerAsync();
        }
        #endregion

        #region ReadTitleBlocks
        private void bw_ReadTitleBlocks_DoWork(object sender, DoWorkEventArgs e)
        {
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = new BlockTableRecord();
                try
                {
                    btr = (BlockTableRecord)tr.GetObject(bt[TITLE_BLOCK_NAME], OpenMode.ForRead);
                }
                catch(Autodesk.AutoCAD.Runtime.Exception acadexc)
                {
                    if(acadexc.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.KeyNotFound)
                    {
                        MessageBox.Show("Command was unable to find the Haskell ARCH D title block.  Please verify the name of the block in AutoCAD and try again.", "Haskell ARCH D Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tr.Commit();
                    }
                    return;
                }  
                ObjectIdCollection objIdCol = btr.GetBlockReferenceIds(true, false);
                ObjectIdCollection annobjIdCol = btr.GetAnonymousBlockIds();

                //Finds all non-dynamic blocks
                #region FIND NONDYNAMIC
                foreach (ObjectId objId in objIdCol)
                {
                    Title_Blocks.Add(new Title_Block(objId));
                    try
                    {
                        BlockReference br = (BlockReference)tr.GetObject(objId, OpenMode.ForRead);
                        if (br != null)
                        {
                            //Loop through dynamic attributes/properties
                            DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in pc)
                            {
                                if (prop.PropertyName.ToUpper().Equals("STAMP"))
                                {
                                    Title_Blocks[Title_Blocks.Count - 1].Stamp = prop.Value.ToString();
                                }
                            }

                            //Loop through all attributes
                            foreach (ObjectId arId in br.AttributeCollection)
                            {
                                DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                AttributeReference ar = (AttributeReference)obj;

                                if (ar != null)
                                {
                                    string attributeTag = ar.Tag.ToUpper();

                                    if (attributeTag.Equals("SHEET_NUMBER"))
                                    {
                                        Title_Blocks[Title_Blocks.Count - 1].PageNumber = ar.TextString;
                                    }
                                    else if (ar.Tag.ToUpper().Equals("SHEET_TITLE"))
                                    {
                                        Title_Blocks[Title_Blocks.Count - 1].PageTitle = ar.TextString;
                                    }
                                    else if (attributeTag.Equals("DRAWNBY"))
                                    {
                                        Title_Blocks[Title_Blocks.Count - 1].DrawnBy = ar.TextString;
                                    }
                                    else if (attributeTag.Equals("CHECKEDBY"))
                                    {
                                        Title_Blocks[Title_Blocks.Count - 1].CheckedBy = ar.TextString;
                                    }
                                    else if (attributeTag.Equals("REV#"))
                                    {
                                        Title_Blocks[Title_Blocks.Count - 1].RevisionNumber = ar.TextString;
                                    }
                                    else
                                    {
                                        for (int i = 1; i <= 15; i++)
                                        {
                                            if (attributeTag.Equals(i.ToString()))
                                            {
                                                Title_Blocks[Title_Blocks.Count - 1].RevisionInfo[i, 1] = ar.TextString;
                                            }
                                            else if (attributeTag.Equals("ISSUE" + i.ToString() + "-TITLE"))
                                            {
                                                Title_Blocks[Title_Blocks.Count - 1].RevisionInfo[i, 2] = ar.TextString;
                                            }
                                            else if (attributeTag.Equals("ISSUE" + i.ToString() + "-DATE"))
                                            {
                                                Title_Blocks[Title_Blocks.Count - 1].RevisionInfo[i, 3] = ar.TextString;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {

                    }
                }
                #endregion

                //Finds all dynamic blocks
                #region FIND DYNAMIC
                foreach (ObjectId id in annobjIdCol)
                {
                    BlockTableRecord annbtr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    ObjectIdCollection blkRefIds = annbtr.GetBlockReferenceIds(true, false);
                    foreach (ObjectId objId in blkRefIds)
                    {
                        Title_Blocks.Add(new Title_Block(objId));
                        try
                        {
                            BlockReference br = (BlockReference)tr.GetObject(objId, OpenMode.ForRead);
                            if (br != null)
                            {
                                //Loop through dynamic attributes/properties
                                DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
                                foreach (DynamicBlockReferenceProperty prop in pc)
                                {
                                    if (prop.PropertyName.ToUpper().Equals("STAMP"))
                                    {
                                        Title_Blocks[Title_Blocks.Count - 1].Stamp = prop.Value.ToString();
                                    }
                                }

                                //Loop through all attributes
                                foreach (ObjectId arId in br.AttributeCollection)
                                {
                                    DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                    AttributeReference ar = (AttributeReference)obj;

                                    if (ar != null)
                                    {
                                        string attributeTag = ar.Tag.ToUpper();

                                        if (attributeTag.Equals("SHEET_NUMBER"))
                                        {
                                            Title_Blocks[Title_Blocks.Count - 1].PageNumber = ar.TextString;
                                        }
                                        else if (ar.Tag.ToUpper().Equals("SHEET_TITLE"))
                                        {
                                            Title_Blocks[Title_Blocks.Count - 1].PageTitle = ar.TextString;
                                        }
                                        else if (ar.Tag.ToUpper().Equals("PROJECTNUM"))
                                        {
                                            Title_Blocks[Title_Blocks.Count - 1].ProjectNumber = ar.TextString;
                                        }
                                        else if (attributeTag.Equals("DRAWNBY"))
                                        {
                                            Title_Blocks[Title_Blocks.Count - 1].DrawnBy = ar.TextString;
                                        }
                                        else if (attributeTag.Equals("CHECKEDBY"))
                                        {
                                            Title_Blocks[Title_Blocks.Count - 1].CheckedBy = ar.TextString;
                                        }
                                        else if (attributeTag.Equals("REV#"))
                                        {
                                            Title_Blocks[Title_Blocks.Count - 1].RevisionNumber = ar.TextString;
                                        }
                                        else
                                        {
                                            for (int i = 0; i < 15; i++)
                                            {
                                                if (attributeTag.Equals((i + 1).ToString()))
                                                {
                                                    Title_Blocks[Title_Blocks.Count - 1].RevisionInfo[i, 0] = ar.TextString;
                                                    break;
                                                }
                                                else if (attributeTag.Equals("ISSUE" + (i + 1).ToString() + "-TITLE"))
                                                {
                                                    Title_Blocks[Title_Blocks.Count - 1].RevisionInfo[i, 1] = ar.TextString;
                                                    break;
                                                }
                                                else if (attributeTag.Equals("ISSUE" + (i + 1).ToString() + "-DATE"))
                                                {
                                                    Title_Blocks[Title_Blocks.Count - 1].RevisionInfo[i, 2] = ar.TextString;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {

                        }
                    }
                }
                #endregion 
                tr.Commit();
            }
        }

        private void bw_ReadTitleBlocks_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void bw_ReadTitleBlocks_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //Sort list of title block info by page number
            Title_Blocks = Title_Blocks.OrderBy(x => x.PageNumber).ToList();
            //Populate DataGridView with all titleblock info
            dgv_TitleBlocks.DataSource = Title_Blocks;
        }
        #endregion

        #region WriteTitleBlocks
        private void Write_To_TitleBlocks()
        {
            using (doc.LockDocument())
            {
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    foreach (Title_Block tb in Title_Blocks)
                    {
                        try
                        {
                            BlockReference br = (BlockReference)tr.GetObject(tb.BlockReferenceID, OpenMode.ForWrite);
                            if (br != null)
                            {
                                //Loop through dynamic attributes/properties
                                DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
                                foreach (DynamicBlockReferenceProperty prop in pc)
                                {
                                    if (prop.PropertyName.ToUpper().Equals("STAMP"))
                                    {
                                        if (!prop.ReadOnly)
                                        {
                                            prop.Value = tb.Stamp;
                                        }
                                    }
                                }

                                //Loop through all attributes
                                foreach (ObjectId arId in br.AttributeCollection)
                                {
                                    DBObject obj = tr.GetObject(arId, OpenMode.ForWrite);
                                    AttributeReference ar = (AttributeReference)obj;

                                    if (ar != null)
                                    {
                                        string attributeTag = ar.Tag.ToUpper();

                                        if (attributeTag.Equals("SHEET_NUMBER"))
                                        {
                                            ar.TextString = tb.PageNumber;
                                        }
                                        else if (attributeTag.Equals("SHEET_TITLE"))
                                        {
                                            ar.TextString = tb.PageTitle;
                                        }
                                        else if (attributeTag.Equals("PROJECTNUM"))
                                        {
                                            ar.TextString = tb.ProjectNumber;
                                        }
                                        else if (attributeTag.Equals("DRAWNBY"))
                                        {
                                            ar.TextString = tb.DrawnBy;
                                        }
                                        else if (attributeTag.Equals("CHECKEDBY"))
                                        {
                                            ar.TextString = tb.CheckedBy;
                                        }
                                        else if (attributeTag.Equals("REV#"))
                                        {
                                            ar.TextString = tb.RevisionNumber;
                                        }
                                        else
                                        {
                                            for (int i = 0; i < 15; i++)
                                            {
                                                if (attributeTag.Equals((i + 1).ToString()))
                                                {
                                                    if(tb.RevisionInfo[i, 0] == null)
                                                    {
                                                        ar.TextString = "";
                                                    }
                                                    else
                                                    {
                                                        ar.TextString = tb.RevisionInfo[i, 0];
                                                    }                                                    
                                                    break;
                                                }
                                                else if (attributeTag.Equals("ISSUE" + (i + 1).ToString() + "-TITLE"))
                                                {
                                                    if (tb.RevisionInfo[i, 1] == null)
                                                    {
                                                        ar.TextString = "";
                                                    }
                                                    else
                                                    {
                                                        ar.TextString = tb.RevisionInfo[i, 1];
                                                    }
                                                    break;
                                                }
                                                else if (attributeTag.Equals("ISSUE" + (i + 1).ToString() + "-DATE"))
                                                {
                                                    if(tb.RevisionInfo[i, 2] == null)
                                                    {
                                                        ar.TextString = "";
                                                    }
                                                    else
                                                    {
                                                        ar.TextString = tb.RevisionInfo[i, 2];
                                                    }                                                    
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {

                        }
                    }
                    tr.Commit();
                }
            }
        }
        #endregion

        #region ButtonEvents
        private void btn_Execute_Click(object sender, EventArgs e)
        {
            btn_Execute.Text = "Processing...";
            btn_Execute.Enabled = false;            

            Write_To_TitleBlocks();

            btn_Execute.Enabled = true;
            btn_Execute.Text = "Execute";
        }

        private void btn_ApplytoAll_Click(object sender, EventArgs e)
        {
            if (chckB_DrawnBy.Checked == true)
            {
                foreach (Title_Block tb in Title_Blocks)
                {
                    tb.DrawnBy = tb_DrawnBy.Text;
                }
            }

            if (chckB_CheckedBy.Checked == true)
            {
                foreach (Title_Block tb in Title_Blocks)
                {
                    tb.CheckedBy = tb_CheckedBy.Text;
                }
            }

            if (chckB_Stamp.Checked == true)
            {
                foreach (Title_Block tb in Title_Blocks)
                {
                    tb.Stamp = cb_Stamp.Text;
                }
            }

            if (chckB_revNum.Checked == true)
            {
                foreach (Title_Block tb in Title_Blocks)
                {
                    tb.RevisionNumber = tb_revNum.Text;
                }
            }

            dgv_TitleBlocks.DataSource = null;
            //SetupDataGrids();
            dgv_TitleBlocks.DataSource = Title_Blocks;
        }

        private void btn_ApplyToAll_revInfo_Click(object sender, EventArgs e)
        {
            if (rb_first.Checked)
            {
                foreach(Title_Block tb in Title_Blocks)
                {
                    tb.RevisionInfo[0, 0] = "1";
                    tb.RevisionInfo[0, 1] = tb_revTitle.Text;
                    tb.RevisionInfo[0, 2] = tb_revDate.Text;
                }
            }
            else if (rb_last.Checked)
            {
                foreach (Title_Block tb in Title_Blocks)
                {
                    tb.RevisionInfo[14, 0] = "15";
                    tb.RevisionInfo[14, 1] = tb_revTitle.Text;
                    tb.RevisionInfo[14, 2] = tb_revDate.Text;
                }
            }
            else if(rb_firstEmpty.Checked)
            {
                foreach (Title_Block tb in Title_Blocks)
                {
                    for(int i = 0; i < tb.RevisionInfo.GetLength(0); i++)
                    {
                        if(String.IsNullOrEmpty(tb.RevisionInfo[i, 0]) && String.IsNullOrEmpty(tb.RevisionInfo[i, 1]) && String.IsNullOrEmpty(tb.RevisionInfo[i, 2]))
                        {
                            tb.RevisionInfo[i, 0] = (i + 1).ToString();
                            tb.RevisionInfo[i, 1] = tb_revTitle.Text;
                            tb.RevisionInfo[i, 2] = tb_revDate.Text;
                            break;
                        }
                    }                    
                }
            }

            if(loadedToRevInfo != null)
            {
                string[,] revisionInfo = loadedToRevInfo.RevisionInfo;
                int numrows = revisionInfo.GetLength(0);
                int numcols = revisionInfo.GetLength(1);

                dgv_revInfo.DataSource = null;
                //Populate DataGridView with all revision info
                datatable_revisioninfo = ConvertArrayToDatatable(revisionInfo, numrows, numcols);
                bindSource_revisioninfo.DataSource = datatable_revisioninfo;
                dgv_revInfo.DataSource = bindSource_revisioninfo;
            }            
        }

        private void Click_btn_ClearSelectedRevInfo(object sender, EventArgs e)
        {
            loadedToRevInfo.RevisionInfo = new string[15, 3];
        }

        private void btn_Clear_ALLrevInfo_Click(object sender, EventArgs e)
        {
            foreach(Title_Block tb in Title_Blocks)
            {
                tb.RevisionInfo = new string[15, 3];
            }
        }
        #endregion        

        #region CheckboxEvents
        private void chckB_DrawnBy_CheckedChanged(object sender, EventArgs e)
        {
            if (chckB_DrawnBy.Checked == true || chckB_CheckedBy.Checked == true || chckB_Stamp.Checked == true || chckB_revNum.Checked == true)
            {
                btn_ApplytoAll.Enabled = true;
            }
            else
            {
                btn_ApplytoAll.Enabled = false;
            }
        }

        private void chckB_CheckedBy_CheckedChanged(object sender, EventArgs e)
        {
            if (chckB_DrawnBy.Checked == true || chckB_CheckedBy.Checked == true || chckB_Stamp.Checked == true || chckB_revNum.Checked == true)
            {
                btn_ApplytoAll.Enabled = true;
            }
            else
            {
                btn_ApplytoAll.Enabled = false;
            }
        }

        private void chckB_Stamp_CheckedChanged(object sender, EventArgs e)
        {
            if (chckB_DrawnBy.Checked == true || chckB_CheckedBy.Checked == true || chckB_Stamp.Checked == true || chckB_revNum.Checked == true)
            {
                btn_ApplytoAll.Enabled = true;
            }
            else
            {
                btn_ApplytoAll.Enabled = false;
            }
        }

        private void chckB_revNum_CheckedChanged(object sender, EventArgs e)
        {
            if (chckB_DrawnBy.Checked == true || chckB_CheckedBy.Checked == true || chckB_Stamp.Checked == true || chckB_revNum.Checked == true)
            {
                btn_ApplytoAll.Enabled = true;
            }
            else
            {
                btn_ApplytoAll.Enabled = false;
            }
        }
        #endregion

        #region RadioButtonEvents
        private void rb_first_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_first.Checked || rb_last.Checked || rb_firstEmpty.Checked)
            {
                btn_ApplyToAll_revInfo.Enabled = true;
            }
            else
            {
                btn_ApplyToAll_revInfo.Enabled = false;
            }
        }

        private void rb_last_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_first.Checked || rb_last.Checked || rb_firstEmpty.Checked)
            {
                btn_ApplyToAll_revInfo.Enabled = true;
            }
            else
            {
                btn_ApplyToAll_revInfo.Enabled = false;
            }
        }

        private void rb_firstEmpty_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_first.Checked || rb_last.Checked || rb_firstEmpty.Checked)
            {
                btn_ApplyToAll_revInfo.Enabled = true;
            }
            else
            {
                btn_ApplyToAll_revInfo.Enabled = false;
            }
        }
        #endregion

        #region DataGridEvents
        private void selectionChanged_dgv_TitleBlocks(object sender, EventArgs e)
        {
            bindSource_revisioninfo.ResetBindings(true);
            if (dgv_TitleBlocks.CurrentRow != null && dgv_TitleBlocks.SelectedRows != null)
            {
                if (dgv_TitleBlocks.CurrentRow.Index >= 0 && dgv_TitleBlocks.CurrentRow.Index < Title_Blocks.Count && dgv_TitleBlocks.SelectedRows.Count >= 1)
                {
                    loadedToRevInfo = Title_Blocks[dgv_TitleBlocks.CurrentRow.Index];
                    if(loadedToRevInfo != null)
                    {
                        btn_Clear_SelectedRevInfo.Enabled = true;
                    }
                    else
                    {
                        btn_Clear_SelectedRevInfo.Enabled = false;
                    }
                    lbl_dgv_revInfo_Title.Text = "Revision Info for: " + loadedToRevInfo.PageNumber;
                    string[,] revisionInfo = loadedToRevInfo.RevisionInfo;
                    int numrows = revisionInfo.GetLength(0);
                    int numcols = revisionInfo.GetLength(1);

                    //Populate DataGridView with all revision info
                    datatable_revisioninfo = ConvertArrayToDatatable(revisionInfo, numrows, numcols);
                    bindSource_revisioninfo.DataSource = datatable_revisioninfo;
                    dgv_revInfo.DataSource = bindSource_revisioninfo;
                }
            }
        }

        private void cellValueChanged_dgv_revInfo(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                loadedToRevInfo.RevisionInfo[e.RowIndex, e.ColumnIndex] = dgv_revInfo.CurrentCell.EditedFormattedValue.ToString();
            }
            else
            {
                dgv_revInfo.DataSource = null;
                dgv_revInfo.DataSource = bindSource_revisioninfo;
            }
        }

        private void dgv_revInfo_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                if (Clipboard.ContainsText())
                {
                    string pasteData = Clipboard.GetText();
                    string[] splitData = pasteData.Split('\t');
                    int column = dgv_revInfo.CurrentCell.ColumnIndex;
                    int row = dgv_revInfo.CurrentCell.RowIndex;

                    if (column >= 0 && row >= 0)
                    {
                        int i = 0;
                        foreach (string s in splitData)
                        {
                            try
                            {
                                loadedToRevInfo.RevisionInfo[row, column + i] = s;
                                i++;
                            }
                            catch (IndexOutOfRangeException)   //Attempted paste too far to the right
                            {
                                break;
                            }
                            finally
                            {
                                string[,] revisionInfo = loadedToRevInfo.RevisionInfo;
                                int numrows = revisionInfo.GetLength(0);
                                int numcols = revisionInfo.GetLength(1);

                                dgv_revInfo.DataSource = null;
                                //Populate DataGridView with all revision info
                                datatable_revisioninfo = ConvertArrayToDatatable(revisionInfo, numrows, numcols);
                                bindSource_revisioninfo.DataSource = datatable_revisioninfo;
                                dgv_revInfo.DataSource = bindSource_revisioninfo;
                            }
                        }
                    }
                }
            }
            else if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewCell cell in dgv_revInfo.SelectedCells)
                {
                    loadedToRevInfo.RevisionInfo[cell.RowIndex, cell.ColumnIndex] = "";
                }

                string[,] revisionInfo = loadedToRevInfo.RevisionInfo;
                int numrows = revisionInfo.GetLength(0);
                int numcols = revisionInfo.GetLength(1);

                dgv_revInfo.DataSource = null;
                //Populate DataGridView with all revision info
                datatable_revisioninfo = ConvertArrayToDatatable(revisionInfo, numrows, numcols);
                bindSource_revisioninfo.DataSource = datatable_revisioninfo;
                dgv_revInfo.DataSource = bindSource_revisioninfo;
            }
        }

        //Formats each row header with the revision number
        private void dgv_revInfo_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
        #endregion

        #region SetupFunctions
        private void SetupComboBoxes()
        {
            cb_Stamp.DataSource = STAMPS;
        }

        private void SetupButtons()
        {
            btn_ApplytoAll.Enabled = false;
        }

        private void SetupDataGrids()
        {
            #region DataGridView_TitleBlocks
            dgv_TitleBlocks.AllowUserToAddRows = false;
            dgv_TitleBlocks.AllowUserToResizeRows = false;
            dgv_TitleBlocks.AllowUserToDeleteRows = false;

            blkRefIDcol.Name = "BlockReferenceID";
            blkRefIDcol.HeaderText = "BlockReferenceID";
            blkRefIDcol.DataPropertyName = "BlockReferenceID";
            blkRefIDcol.ReadOnly = true;
            blkRefIDcol.DisplayIndex = 0;
            dgv_TitleBlocks.Columns.Add(blkRefIDcol);

            pageNumbercol.Name = "PageNumber";
            pageNumbercol.HeaderText = "PageNumber";
            pageNumbercol.DataPropertyName = "PageNumber";
            pageNumbercol.ReadOnly = true;
            pageNumbercol.SortMode = DataGridViewColumnSortMode.Automatic;
            pageNumbercol.DisplayIndex = 1;
            dgv_TitleBlocks.Columns.Add(pageNumbercol);

            pageTitlecol.Name = "PageTitle";
            pageTitlecol.HeaderText = "PageTitle";
            pageTitlecol.DataPropertyName = "PageTitle";
            pageTitlecol.DisplayIndex = 2;
            dgv_TitleBlocks.Columns.Add(pageTitlecol);

            projNumcol.Name = "ProjectNumber";
            projNumcol.HeaderText = "ProjectNumber";
            projNumcol.DataPropertyName = "ProjectNumber";
            projNumcol.ReadOnly = true;
            projNumcol.DisplayIndex = 3;
            dgv_TitleBlocks.Columns.Add(projNumcol);

            drawnBycol.Name = "DrawnBy";
            drawnBycol.HeaderText = "DrawnBy";
            drawnBycol.DataPropertyName = "DrawnBy";
            drawnBycol.DisplayIndex = 4;
            dgv_TitleBlocks.Columns.Add(drawnBycol);

            checkedBycol.Name = "CheckedBy";
            checkedBycol.HeaderText = "CheckedBy";
            checkedBycol.DataPropertyName = "CheckedBy";
            checkedBycol.DisplayIndex = 5;
            dgv_TitleBlocks.Columns.Add(checkedBycol);

            stampcol.Name = "Stamp";
            stampcol.HeaderText = "Stamp";
            stampcol.DataPropertyName = "Stamp";
            stampcol.DataSource = STAMPS;
            stampcol.DisplayIndex = 6;
            dgv_TitleBlocks.Columns.Add(stampcol);

            revNumcol.Name = "RevisionNumber";
            revNumcol.HeaderText = "RevisionNumber";
            revNumcol.DataPropertyName = "RevisionNumber";
            revNumcol.DisplayIndex = 7;
            dgv_TitleBlocks.Columns.Add(revNumcol);
            #endregion

            #region DataGridView_RevisionInfo
            dgv_revInfo.AllowUserToAddRows = false;
            dgv_revInfo.AllowUserToResizeRows = false;
            dgv_revInfo.AllowUserToDeleteRows = false;
            dgv_revInfo.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dgv_revInfo.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            revInfoNumcol.Name = "RevNumber";
            revInfoNumcol.HeaderText = "RevNumber";
            revInfoNumcol.DataPropertyName = "RevNumber";
            revInfoNumcol.DisplayIndex = 0;
            revInfoNumcol.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgv_revInfo.Columns.Add(revInfoNumcol);

            revInfoTitlecol.Name = "RevTitle";
            revInfoTitlecol.HeaderText = "RevTitle";
            revInfoTitlecol.DataPropertyName = "RevTitle";
            revInfoTitlecol.DisplayIndex = 1;
            revInfoTitlecol.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgv_revInfo.Columns.Add(revInfoTitlecol);

            revInfoDatecol.Name = "RevDate";
            revInfoDatecol.HeaderText = "RevDate";
            revInfoDatecol.DataPropertyName = "RevDate";
            revInfoDatecol.DisplayIndex = 2;
            revInfoDatecol.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgv_revInfo.Columns.Add(revInfoDatecol);
            #endregion
        }
        #endregion

        #region HelperFunctions
        private System.Data.DataTable ConvertArrayToDatatable(string[,] revisionInfo, int numrows, int numcols)
        {
            System.Data.DataTable table = new System.Data.DataTable();

            for(int i = 0; i < numcols; i++)
            {
                string columnHeader = "";
                if(i == 0)
                {
                    columnHeader = "RevNumber";
                }
                else if(i == 1)
                {
                    columnHeader = "RevTitle";
                }
                else if(i == 2)
                {
                    columnHeader = "RevDate";
                }
                table.Columns.Add(columnHeader);
            }

            for(int i = 0; i < numrows; i++)
            {
                DataRow row = table.NewRow();

                for(int j = 0; j < numcols; j++)
                {
                    row[j] = revisionInfo[i, j];
                }

                table.Rows.Add(row);
            }

            return table;
        }
        #endregion        

        #region UnusedCode
        private static System.Data.DataTable ConvertListToDatatable<T>(List<T> data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            System.Data.DataTable table = new System.Data.DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    table.Columns.Add(prop.Name, prop.PropertyType.GetGenericArguments()[0]);
                else
                    table.Columns.Add(prop.Name, prop.PropertyType);
            }

            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }
            return table;
        }
        #endregion
    }

    public class Title_Block
    {
        public ObjectId BlockReferenceID { get; set; }
        public string PageNumber { get; set; }
        public string PageTitle { get; set; }
        public string ProjectNumber { get; set; }
        public string DrawnBy { get; set; }
        public string CheckedBy { get; set; }
        public string Stamp { get; set; }
        public string RevisionNumber { get; set; }
        public string[,] RevisionInfo { get; set; }        

        public Title_Block(ObjectId objectID)
        {
            BlockReferenceID = objectID;
            PageNumber = "";
            PageTitle = "";
            ProjectNumber = "";
            DrawnBy = "";
            CheckedBy = "";
            Stamp = "";
            RevisionNumber = "";
            RevisionInfo = new string[15, 3];
        }
    }
}
