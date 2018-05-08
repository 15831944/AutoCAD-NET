using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace Tutorial_Dialogs
{
    public partial class Tutorials_by_Adam_Form : Form
    {
        string CSV_network_Path = @"\\thehaskellco.net\Remote\AtlantaData\_Disciplines\Controls\4 Controls Programming\DVBs and DLLs\CSV Resources\";
        string CSV_FileName = @"TutorialList.csv";
        string Tutorial_network_Path;
        string navigating_to;

        List<PDF_Resource> PDFs = new List<PDF_Resource>();

        Document doc;
        Database db;
        Editor ed;

        public Tutorials_by_Adam_Form()
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
            if (PDFs.Count == 0)
            {
                ed.WriteMessage("ERROR: No PDFs found in file list. \n");
                this.Close();
            }
        }

        #region ComboBoxEvents      
        private void TextChanged_cb_TutorialSelect(object sender, EventArgs e)
        {
            foreach (PDF_Resource PDF in PDFs)
            {
                if (cb_TutorialSelect.Text.Equals(PDF.SeriesName))
                {
                    if (System.IO.File.Exists(PDF.FilePath))
                    {
                        wb_Tutorial.Navigate(PDF.FilePath);
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
                            Tutorial_network_Path = values[2];
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
                    PDFs.Add(new PDF_Resource(Tutorial_network_Path, PDF_FileNames[i].Item1, PDF_FileNames[i].Item2));
                }
            }
            else
            {
                ed.WriteMessage("ERROR: CSV File " + CSV_FileName + " not found in directory " + CSV_network_Path);
            }
        }

        private void PopulateComboBox()
        {
            foreach (PDF_Resource PDF in PDFs)
            {
                cb_TutorialSelect.Items.Add(PDF.SeriesName);
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
