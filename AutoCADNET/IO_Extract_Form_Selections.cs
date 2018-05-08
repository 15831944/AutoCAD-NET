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
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
//using Microsoft.Office.Tools.Excel;

namespace IO_Extract_Dialogs
{
    public partial class IO_Extract_Form_Selections : Form
    {
        private ListViewColumnSorter lvColumnSorter = new ListViewColumnSorter();
        int[] slotCells = new int[] { 3, 37, 71, 105, 139, 173, 207, 241, 275, 309, 343, 377, 411, 445, 479, 513, 547, 581 };

        public List<IO_Block> IO_Blocks { get; set; }

        public IO_Extract_Form_Selections()
        {
            InitializeComponent();
        }

        private void load_IOExtractSelections(object sender, EventArgs e)
        {
            ColumnHeader rackCol = new ColumnHeader();
            rackCol.Text = "Rack";
            rackCol.Name = "col1";
            rackCol.Width = 50;

            ColumnHeader sortCol = new ColumnHeader();
            sortCol.Text = "Slot";
            sortCol.Name = "col2";
            sortCol.Width = 50;

            ColumnHeader partCol = new ColumnHeader();
            partCol.Text = "Part Number";
            partCol.Name = "col3";
            partCol.Width = 300;

            lv_IOCards.Columns.Add(rackCol);
            lv_IOCards.Columns.Add(sortCol);
            lv_IOCards.Columns.Add(partCol);

            foreach(IO_Block card in IO_Blocks)
            {
                string[] item = new string[3];
                item[0] = card.rack;
                item[1] = card.slot;
                item[2] = card.partNumber;

                ListViewItem itm = new ListViewItem(item);

                lv_IOCards.Items.Add(itm);
            }
            
            lv_IOCards.ListViewItemSorter = lvColumnSorter;

            lvColumnSorter.SortColumn = lv_IOCards.Columns.IndexOf(sortCol);
            lvColumnSorter.Order = SortOrder.Ascending;            
            lv_IOCards.Sort();

            lvColumnSorter.SortColumn = lv_IOCards.Columns.IndexOf(rackCol);
            lvColumnSorter.Order = SortOrder.Ascending;
            lv_IOCards.Sort();
        }

        private void click_Extract(object sender, EventArgs e)
        {
            int CardColumn = 2;
            int TagColumn = 4;
            int DescriptionColumn = 5;
            int SlotRowWidth = 34;

            List<IO_Block> Selected_Cards = new List<IO_Block>();
            ListView.CheckedIndexCollection holder = lv_IOCards.CheckedIndices;

            foreach(int index in holder)
            {
                Selected_Cards.Add(IO_Blocks[index]);
            }

            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            ExcelApp.ScreenUpdating = false;
            ExcelApp.EnableEvents = false;

            Excel.Workbook ExcelWB = ExcelApp.Workbooks.Open(tb_Path.Text);
            ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            Excel.Sheets ExcelSheets = ExcelWB.Sheets;

            foreach(IO_Block card in Selected_Cards)
            {
                int test = SlotRowWidth * Int32.Parse(card.slot) + 3;
                ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 3, CardColumn].Value = card.partNumber;

                int portCount = 0;
                for(int i = 0; i < card.portDescriptions.Length; i++)
                {                    
                    if(!String.IsNullOrEmpty(card.portDescriptions[i]))
                    {                        
                        ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 4 + portCount, DescriptionColumn].Value = card.portDescriptions[i];
                        portCount++;
                    }                     
                }
                ExcelApp.DisplayAlerts = false;

                Excel.Range rng = ExcelSheets[1].Range[ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 4, DescriptionColumn], ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 35, DescriptionColumn]];
                rng.Replace(@"\P", " ", Excel.XlLookAt.xlPart);
                rng.Replace(@" \P", " ", Excel.XlLookAt.xlPart);
                rng.Copy();

                ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 4, TagColumn].PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                rng = ExcelSheets[1].Range[ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 4, TagColumn], ExcelSheets[1].Cells[SlotRowWidth * Int32.Parse(card.slot) + 35, TagColumn]];
                rng.Replace("  ", "_", Excel.XlLookAt.xlPart);
                rng.Replace(" ", "_", Excel.XlLookAt.xlPart);
                rng.Replace("-", "_", Excel.XlLookAt.xlPart);
                rng.Replace("SPARE", "", Excel.XlLookAt.xlPart);

                ExcelApp.DisplayAlerts = true;
            }
            ExcelWB.Worksheets[1].Cells[3, 2].Select();

            Autodesk.AutoCAD.Windows.SaveFileDialog sfd = new Autodesk.AutoCAD.Windows.SaveFileDialog("Save IO Export Excel file as:", "IOExport", "xlsm", "ExcelFiles", Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.DoNotWarnIfFileExist);
            System.Windows.Forms.DialogResult dr = sfd.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                ExcelWB.SaveAs(sfd.Filename, Type.Missing, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                ExcelApp.EnableEvents = true;
                ExcelApp.Visible = true;
                ExcelApp.ScreenUpdating = true;
                ExcelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                ExcelWB.Unprotect();
                ExcelApp.Run("ExternalUpdate");
                ExcelWB.Protect();
                this.Close();            
            }
            else
            {
                ExcelWB.Close(false, Type.Missing, Type.Missing);
                this.Close();
            }
        }

        private void click_Cancel(object sender, EventArgs e)
        {
            this.Close();            
        }

        private void click_Browse(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Select Excel file:", null, "xlsm", "ExcelFiles", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple | Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                tb_Path.Text = ofd.Filename;
            }
        }

        private void columnclick_lv_IOCards(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvColumnSorter.Order == SortOrder.Ascending)
                {
                    lvColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvColumnSorter.SortColumn = e.Column;
                lvColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            lv_IOCards.Sort();
        }
    }

    public class ListViewColumnSorter : IComparer
    {
        /// <summary>
        /// Specifies the column to be sorted
        /// </summary>
        private int ColumnToSort;
        /// <summary>
        /// Specifies the order in which to sort (i.e. 'Ascending').
        /// </summary>
        private SortOrder OrderOfSort;
        /// <summary>
        /// Case insensitive comparer object
        /// </summary>
        private CaseInsensitiveComparer ObjectCompare;

        /// <summary>
        /// Class constructor.  Initializes various elements
        /// </summary>
        public ListViewColumnSorter()
        {
            // Initialize the column to '0'
            ColumnToSort = 0;

            // Initialize the sort order to 'none'
            OrderOfSort = SortOrder.None;

            // Initialize the CaseInsensitiveComparer object
            ObjectCompare = new CaseInsensitiveComparer();
        }

        /// <summary>
        /// This method is inherited from the IComparer interface.  It compares the two objects passed using a case insensitive comparison.
        /// </summary>
        /// <param name="x">First object to be compared</param>
        /// <param name="y">Second object to be compared</param>
        /// <returns>The result of the comparison. "0" if equal, negative if 'x' is less than 'y' and positive if 'x' is greater than 'y'</returns>
        public int Compare(object x, object y)
        {
            int compareResult;
            ListViewItem listviewX, listviewY;

            // Cast the objects to be compared to ListViewItem objects
            listviewX = (ListViewItem)x;
            listviewY = (ListViewItem)y;

            // Compare the two items
            compareResult = ObjectCompare.Compare(listviewX.SubItems[ColumnToSort].Text, listviewY.SubItems[ColumnToSort].Text);

            // Calculate correct return value based on object comparison
            if (OrderOfSort == SortOrder.Ascending)
            {
                // Ascending sort is selected, return normal result of compare operation
                return compareResult;
            }
            else if (OrderOfSort == SortOrder.Descending)
            {
                // Descending sort is selected, return negative result of compare operation
                return (-compareResult);
            }
            else
            {
                // Return '0' to indicate they are equal
                return 0;
            }
        }

        /// <summary>
        /// Gets or sets the number of the column to which to apply the sorting operation (Defaults to '0').
        /// </summary>
        public int SortColumn
        {
            set
            {
                ColumnToSort = value;
            }
            get
            {
                return ColumnToSort;
            }
        }

        /// <summary>
        /// Gets or sets the order of sorting to apply (for example, 'Ascending' or 'Descending').
        /// </summary>
        public SortOrder Order
        {
            set
            {
                OrderOfSort = value;
            }
            get
            {
                return OrderOfSort;
            }
        }

    }
}
