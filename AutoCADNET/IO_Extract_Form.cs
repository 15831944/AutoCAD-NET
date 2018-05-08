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
    public partial class IO_Extract_Form : Form
    {
        Document doc;
        Database db;
        Editor ed;

        String[] dwgFiles;

        public ListView.ListViewItemCollection fileList { get; set; }

        public IO_Extract_Form()
        {
            InitializeComponent();
        }

        private void load_IOExtractionForm(object sender, EventArgs e)
        {
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            btn_Next.Enabled = false;
            btn_RemoveSelected.Enabled = false;

            ColumnHeader dummy = new ColumnHeader();
            dummy.Text = "Drawing Path:";
            dummy.Name = "col1";
            dummy.Width = -2;
            lv_Files.Columns.Add(dummy);

            rb_Full.Checked = false;
            rb_Smart.Checked = true;            
        } 

        private void click_AddDrawings(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select drawings:", null, "dwg", "AutoCADDrawings", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple | Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();

            if (dr != System.Windows.Forms.DialogResult.OK)
            {
                ed.WriteMessage("Invalid selection.");
                this.Close();
                return;
            }
            else
            {
                btn_Next.Enabled = true;
                dwgFiles = ofd.GetFilenames();
                foreach(string file in dwgFiles)
                {
                    lv_Files.Items.Add(file);
                }
                lv_Files.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void click_RemoveDrawings(object sender, EventArgs e)
        {
            foreach(ListViewItem item in lv_Files.SelectedItems)
            {
                lv_Files.Items.Remove(item);
            }

            if(!(lv_Files.Items.Count > 0) || !(lv_Files.SelectedItems.Count > 0))
            {
                btn_RemoveSelected.Enabled = false;
            }

            if(!(lv_Files.Items.Count > 0))
            {
                btn_Next.Enabled = false;
            }

            lv_Files.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
        }

        private void click_Next(object sender, EventArgs e)
        {
            if(lv_Files.Items.Count > 0)
            {
                List<string> filePaths = new List<string>();
                foreach(ListViewItem item in lv_Files.Items)
                {
                    filePaths.Add(item.Text);
                }
                IO_Extract_Dialogs.IO_Extract_Form_Progress IEFP = new IO_Extract_Dialogs.IO_Extract_Form_Progress();
                IEFP.doc = doc;
                IEFP.db = db;
                IEFP.ed = ed;
                IEFP.fileList = filePaths;
                IEFP.smartSearch = rb_Smart.Checked;
                this.Visible = false;
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(null, IEFP, false);
            }            
        }

        private void click_Cancel(object sender, EventArgs e)
        {
            this.Close();
        }

        private void selectedIndexChanged_Files(object sender, EventArgs e)
        {
            if(lv_Files.SelectedItems.Count > 0)
            {
                btn_RemoveSelected.Enabled = true;
            }
        } 
    }
}
