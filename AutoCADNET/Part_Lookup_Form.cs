using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.IO;

namespace Part_Lookup_Dialogs
{
    public partial class Part_Lookup_Form : Form
    {
        string CSV_network_Path = @"\\thehaskellco.net\Remote\AtlantaData\_Disciplines\Controls\4 Controls Programming\DVBs and DLLs\CSV Resources\";
        string CSV_FileName = @"PartLookup_PDFList.csv";
        string PDF_network_Path;
        //string network_Path = @"\\thehaskellco.net\Remote\AtlantaData\_Disciplines\Controls\4 Controls Programming\DVBs and DLLs\CSV Resources\";
        string navigating_to;

        List<PDF_Resource> PDFs = new List<PDF_Resource>();        

        Document doc;
        Database db;
        Editor ed;

        public Part_Lookup_Form()
        {
            InitializeComponent();

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            //Detects if connected to Haskell network
            //If not, closes form
            if (!System.IO.Directory.Exists(CSV_network_Path))
            {
                ed.WriteMessage("ERROR: Must be connected to Haskell network to retrieve files... \n");
                this.Close();
            }

            DefinePDFs();
            PopulateComboBox();

            //Loads the "default" PDF
            if(PDFs.Count == 0)
            {
                ed.WriteMessage("ERROR: No PDFs found in file list. \n");
                this.Close();
            }
        }        

        #region HelperFunctions
        private void DefinePDFs()
        {
            List<Tuple<string, string>> PDF_FileNames = new List<Tuple<string, string>>();

            if (System.IO.File.Exists(CSV_network_Path + CSV_FileName))
            {
                using (var reader = new StreamReader(CSV_network_Path + CSV_FileName))
                {
                    bool firstLine = true;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (firstLine)
                        {
                            firstLine = false;
                            PDF_network_Path = values[2];
                            continue;
                        }
                        else
                        {
                            PDF_FileNames.Add(new Tuple<string, string>(values[0], values[1]));
                        }
                    }
                }

                for (int i = 0; i < PDF_FileNames.Count; i++)
                {
                    PDFs.Add(new PDF_Resource(PDF_network_Path, PDF_FileNames[i].Item1, PDF_FileNames[i].Item2));
                }
            }
            else
            {
                ed.WriteMessage("ERROR: CSV File " + CSV_FileName + " not found in directory " + CSV_network_Path);
            }
        }

        private void PopulateComboBox()
        {
            foreach(PDF_Resource PDF in PDFs)
            {
                cb_PDFSelect.Items.Add(PDF.SeriesName);
            }
        }
        #endregion

        #region WebBrowserEvents
        private void Navigating_wb_PDF(object sender, WebBrowserNavigatingEventArgs e)
        {
            lbl_Displaying.Text = "Opening: " + System.IO.Path.GetFileName(navigating_to);
        }

        private void Navigated_wb_PDF(object sender, WebBrowserNavigatedEventArgs e)
        {
            lbl_Displaying.Text = "Viewing: " + System.IO.Path.GetFileName(navigating_to);
        }
        #endregion

        #region ComboBoxEvents      
        private void TextChanged_cb_PDFSelect(object sender, EventArgs e)
        {
            foreach (PDF_Resource PDF in PDFs)
            {
                if (cb_PDFSelect.Text.Equals(PDF.SeriesName))
                {
                    if (System.IO.File.Exists(PDF.FilePath))
                    {
                        wb_PDF.Navigate(PDF.FilePath);
                        navigating_to = PDF.FilePath;
                    }
                    else
                    {
                        ed.WriteMessage("ERROR: File not found. \n");
                    }
                }
            }
        }
        #endregion
    }

    public class PDF_Resource
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string SeriesName { get; set; }

        public PDF_Resource(string NetworkPath, string file_name, string series_name)
        {
            FileName = file_name;
            FilePath = NetworkPath + FileName;
            SeriesName = series_name;
        }
    }
}
