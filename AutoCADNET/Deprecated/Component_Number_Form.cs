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

namespace Component_Number_Dialogs
{
    public partial class Component_Number_Form : Form
    {
        Document doc;
        Database db;
        Editor ed;

        string[] prvw = new string[35];
        string preview;
        string lastValue1;
        string lastValue2;
        string lastValue3;
        string lastValue4;
        string lastValue5;

        public Component_Number_Form()
        {
            InitializeComponent();
        }

        private void Component_Number_Form_Load(object sender, EventArgs e)
        {
            ComboBoxSetup();

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            txtboxDWGName.Text = Path.GetFileName(doc.Name);

            try
            {
                lblCN1.Text = txtboxDWGName.Text[0].ToString();
                lblCN2.Text = txtboxDWGName.Text[1].ToString();
                lblCN3.Text = txtboxDWGName.Text[2].ToString();
                lblCN4.Text = txtboxDWGName.Text[3].ToString();
                lblCN5.Text = txtboxDWGName.Text[4].ToString();
                lblCN6.Text = txtboxDWGName.Text[5].ToString();
                lblCN7.Text = txtboxDWGName.Text[6].ToString();
                lblCN8.Text = txtboxDWGName.Text[7].ToString();
                lblCN9.Text = txtboxDWGName.Text[8].ToString();
                lblCN10.Text = txtboxDWGName.Text[9].ToString();
                lblCN11.Text = txtboxDWGName.Text[10].ToString();
                lblCN12.Text = txtboxDWGName.Text[11].ToString();
                lblCN13.Text = txtboxDWGName.Text[12].ToString();
                lblCN14.Text = txtboxDWGName.Text[13].ToString();
                lblCN15.Text = txtboxDWGName.Text[14].ToString();
                lblCN16.Text = txtboxDWGName.Text[15].ToString();
                lblCN17.Text = txtboxDWGName.Text[16].ToString();
                lblCN18.Text = txtboxDWGName.Text[17].ToString();
                lblCN19.Text = txtboxDWGName.Text[18].ToString();
                lblCN20.Text = txtboxDWGName.Text[19].ToString();
                lblCN21.Text = txtboxDWGName.Text[20].ToString();
                lblCN22.Text = txtboxDWGName.Text[21].ToString();
                lblCN23.Text = txtboxDWGName.Text[22].ToString();
                lblCN24.Text = txtboxDWGName.Text[23].ToString();
                lblCN25.Text = txtboxDWGName.Text[24].ToString();
                lblCN26.Text = txtboxDWGName.Text[25].ToString();
                lblCN27.Text = txtboxDWGName.Text[26].ToString();
                lblCN28.Text = txtboxDWGName.Text[27].ToString();
                lblCN29.Text = txtboxDWGName.Text[28].ToString();
                lblCN30.Text = txtboxDWGName.Text[29].ToString();

                cmbBoxCNL.Items.Add("1");
                cmbBoxCNL.Items.Add("2");
                cmbBoxCNL.Items.Add("3");
                cmbBoxCNL.Items.Add("4");
                cmbBoxCNL.Items.Add("5");
                cmbBoxCNL.Items.Add("6");
                cmbBoxCNL.Items.Add("7");
                cmbBoxCNL.Items.Add("8");

                cmbBoxCNSL.Items.Add("1");
                cmbBoxCNSL.Items.Add("2");
                cmbBoxCNSL.Items.Add("3");
                cmbBoxCNSL.Items.Add("4");
                cmbBoxCNSL.Items.Add("5");
                cmbBoxCNSL.Items.Add("6");
                cmbBoxCNSL.Items.Add("7");
                cmbBoxCNSL.Items.Add("8");
            }
            catch
            {
                ed.WriteMessage("Drawing name is less than 30 characters. ");
            }
        }

        private void btnRenumber_Click(object sender, EventArgs e)
        {
            string DieselEXP;

            DieselEXP = @"%<\AcDiesel ";

            for (int i = 0; i < prvw.Length; i++)
            {
                if (chckBxCN1.Checked)
                {
                    if (cmbBoxCN1.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "1,1)";
                    }
                }
                if (chckBxCN2.Checked)
                {
                    if (cmbBoxCN2.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "2,1)";
                    }
                }
                if (chckBxCN3.Checked)
                {
                    if (cmbBoxCN3.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "3,1)";
                    }
                }
                if (chckBxCN4.Checked)
                {
                    if (cmbBoxCN4.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "4,1)";
                    }
                }
                if (chckBxCN5.Checked)
                {
                    if (cmbBoxCN5.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "5,1)";
                    }
                }
                if (chckBxCN6.Checked)
                {
                    if (cmbBoxCN6.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "6,1)";
                    }
                }
                if (chckBxCN7.Checked)
                {
                    if (cmbBoxCN7.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "7,1)";
                    }
                }
                if (chckBxCN8.Checked)
                {
                    if (cmbBoxCN8.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "8,1)";
                    }
                }
                if (chckBxCN9.Checked)
                {
                    if (cmbBoxCN9.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "9,1)";
                    }
                }
                if (chckBxCN10.Checked)
                {
                    if (cmbBoxCN10.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "10,1)";
                    }
                }
                if (chckBxCN11.Checked)
                {
                    if (cmbBoxCN11.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "11,1)";
                    }
                }
                if (chckBxCN12.Checked)
                {
                    if (cmbBoxCN12.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "12,1)";
                    }
                }
                if (chckBxCN13.Checked)
                {
                    if (cmbBoxCN13.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "13,1)";
                    }
                }
                if (chckBxCN14.Checked)
                {
                    if (cmbBoxCN14.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "14,1)";
                    }
                }
                if (chckBxCN15.Checked)
                {
                    if (cmbBoxCN15.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "15,1)";
                    }
                }
                if (chckBxCN16.Checked)
                {
                    if (cmbBoxCN16.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "16,1)";
                    }
                }
                if (chckBxCN17.Checked)
                {
                    if (cmbBoxCN17.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "17,1)";
                    }
                }
                if (chckBxCN18.Checked)
                {
                    if (cmbBoxCN18.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "18,1)";
                    }
                }
                if (chckBxCN19.Checked)
                {
                    if (cmbBoxCN19.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "19,1)";
                    }
                }
                if (chckBxCN20.Checked)
                {
                    if (cmbBoxCN20.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "20,1)";
                    }
                }
                if (chckBxCN21.Checked)
                {
                    if (cmbBoxCN21.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "21,1)";
                    }
                }
                if (chckBxCN22.Checked)
                {
                    if (cmbBoxCN22.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "22,1)";
                    }
                }
                if (chckBxCN23.Checked)
                {
                    if (cmbBoxCN23.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "23,1)";
                    }
                }
                if (chckBxCN24.Checked)
                {
                    if (cmbBoxCN24.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "24,1)";
                    }
                }
                if (chckBxCN25.Checked)
                {
                    if (cmbBoxCN25.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "25,1)";
                    }
                }
                if (chckBxCN26.Checked)
                {
                    if (cmbBoxCN26.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "26,1)";
                    }
                }
                if (chckBxCN27.Checked)
                {
                    if (cmbBoxCN27.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "27,1)";
                    }
                }
                if (chckBxCN28.Checked)
                {
                    if (cmbBoxCN28.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "28,1)";
                    }
                }
                if (chckBxCN29.Checked)
                {
                    if (cmbBoxCN29.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "29,1)";
                    }
                }
                if (chckBxCN30.Checked)
                {
                    if (cmbBoxCN30.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "30,1)";
                    }
                }
                if (chckBxBreaks1.Checked)
                {
                    if (cmbBoxBreaks1.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + txtBoxBreaks1.Text;
                    }
                }
                if (chckBxBreaks2.Checked)
                {
                    if (cmbBoxBreaks2.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + txtBoxBreaks2.Text;
                    }
                }
                if (chckBxBreaks3.Checked)
                {
                    if (cmbBoxBreaks3.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + txtBoxBreaks3.Text;
                    }
                }
                if (chckBxBreaks4.Checked)
                {
                    if (cmbBoxBreaks4.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + txtBoxBreaks4.Text;
                    }
                }
                if (chckBxBreaks5.Checked)
                {
                    if (cmbBoxBreaks5.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + txtBoxBreaks5.Text;
                    }
                }
            }

            DieselEXP = DieselEXP + @">%";

            DocumentLock loc = doc.LockDocument();
            using (loc)
            {
                using (Transaction transaction = db.TransactionManager.StartTransaction())
                {
                    this.Hide();
                    PromptSelectionResult SPrompt = doc.Editor.GetSelection();

                    if (SPrompt.Status == PromptStatus.OK)
                    {
                        this.Show();
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
                                        BlockTableRecord bd = (BlockTableRecord)transaction.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);

                                        foreach (ObjectId attId in blkRef.AttributeCollection)
                                        {
                                            AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                            if (attRef.Tag == "TAG")
                                            {
                                                string currentText = attRef.TextString;
                                                string wire;
                                                string type;

                                                try
                                                {
                                                    wire = currentText.Substring(currentText.Length - Int32.Parse(cmbBoxCNL.Text), Int32.Parse(cmbBoxCNL.Text));
                                                }
                                                catch
                                                {
                                                    wire = "";
                                                }

                                                try
                                                {
                                                    type = currentText.Substring(0, Int32.Parse(cmbBoxCNSL.Text));
                                                }
                                                catch
                                                {
                                                    type = "";
                                                }

                                                attRef.TextString = "";
                                                attRef.TextString = type + DieselEXP + wire;                                                                                  
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

        private void UpdatePrvw(object sender, EventArgs e)
        {
            for (int i = 0; i < prvw.Length; i++)
            {
                prvw[i] = "";
            }

            preview = "";

            if (chckBxCN1.Checked && !cmbBoxCN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN1.Text)] = lblCN1.Text;
            }
            else if (!chckBxCN1.Checked && !cmbBoxCN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN1.Text)] = "";
            }

            if (chckBxCN2.Checked && !cmbBoxCN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN2.Text)] = lblCN2.Text;
            }
            else if (!chckBxCN2.Checked && !cmbBoxCN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN2.Text)] = "";
            }

            if (chckBxCN3.Checked && !cmbBoxCN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN3.Text)] = lblCN3.Text;
            }
            else if (!chckBxCN3.Checked && !cmbBoxCN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN3.Text)] = "";
            }

            if (chckBxCN4.Checked && !cmbBoxCN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN4.Text)] = lblCN4.Text;
            }
            else if (!chckBxCN4.Checked && !cmbBoxCN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN4.Text)] = "";
            }

            if (chckBxCN5.Checked && !cmbBoxCN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN5.Text)] = lblCN5.Text;
            }
            else if (!chckBxCN5.Checked && !cmbBoxCN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN5.Text)] = "";
            }

            if (chckBxCN6.Checked && !cmbBoxCN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN6.Text)] = lblCN6.Text;
            }
            else if (!chckBxCN6.Checked && !cmbBoxCN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN6.Text)] = "";
            }

            if (chckBxCN7.Checked && !cmbBoxCN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN7.Text)] = lblCN7.Text;
            }
            else if (!chckBxCN7.Checked && !cmbBoxCN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN7.Text)] = "";
            }

            if (chckBxCN8.Checked && !cmbBoxCN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN8.Text)] = lblCN8.Text;
            }
            else if (!chckBxCN8.Checked && !cmbBoxCN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN8.Text)] = "";
            }

            if (chckBxCN9.Checked && !cmbBoxCN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN9.Text)] = lblCN9.Text;
            }
            else if (!chckBxCN9.Checked && !cmbBoxCN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN9.Text)] = "";
            }

            if (chckBxCN10.Checked && !cmbBoxCN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN10.Text)] = lblCN10.Text;
            }
            else if (!chckBxCN10.Checked && !cmbBoxCN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN10.Text)] = "";
            }

            if (chckBxCN11.Checked && !cmbBoxCN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN11.Text)] = lblCN11.Text;
            }
            else if (!chckBxCN11.Checked && !cmbBoxCN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN11.Text)] = "";
            }

            if (chckBxCN12.Checked && !cmbBoxCN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN12.Text)] = lblCN12.Text;
            }
            else if (!chckBxCN12.Checked && !cmbBoxCN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN12.Text)] = "";
            }

            if (chckBxCN13.Checked && !cmbBoxCN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN13.Text)] = lblCN13.Text;
            }
            else if (!chckBxCN13.Checked && !cmbBoxCN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN13.Text)] = "";
            }

            if (chckBxCN14.Checked && !cmbBoxCN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN14.Text)] = lblCN14.Text;
            }
            else if (!chckBxCN14.Checked && !cmbBoxCN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN14.Text)] = "";
            }

            if (chckBxCN15.Checked && !cmbBoxCN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN15.Text)] = lblCN15.Text;
            }
            else if (!chckBxCN15.Checked && !cmbBoxCN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN15.Text)] = "";
            }

            if (chckBxCN16.Checked && !cmbBoxCN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN16.Text)] = lblCN16.Text;
            }
            else if (!chckBxCN16.Checked && !cmbBoxCN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN16.Text)] = "";
            }

            if (chckBxCN17.Checked && !cmbBoxCN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN17.Text)] = lblCN17.Text;
            }
            else if (!chckBxCN17.Checked && !cmbBoxCN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN17.Text)] = "";
            }

            if (chckBxCN18.Checked && !cmbBoxCN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN18.Text)] = lblCN18.Text;
            }
            else if (!chckBxCN18.Checked && !cmbBoxCN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN18.Text)] = "";
            }

            if (chckBxCN19.Checked && !cmbBoxCN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN19.Text)] = lblCN19.Text;
            }
            else if (!chckBxCN19.Checked && !cmbBoxCN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN19.Text)] = "";
            }

            if (chckBxCN20.Checked && !cmbBoxCN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN20.Text)] = lblCN20.Text;
            }
            else if (!chckBxCN20.Checked && !cmbBoxCN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN20.Text)] = "";
            }

            if (chckBxCN21.Checked && !cmbBoxCN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN21.Text)] = lblCN21.Text;
            }
            else if (!chckBxCN21.Checked && !cmbBoxCN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN21.Text)] = "";
            }

            if (chckBxCN22.Checked && !cmbBoxCN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN22.Text)] = lblCN22.Text;
            }
            else if (!chckBxCN22.Checked && !cmbBoxCN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN22.Text)] = "";
            }

            if (chckBxCN23.Checked && !cmbBoxCN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN23.Text)] = lblCN23.Text;
            }
            else if (!chckBxCN23.Checked && !cmbBoxCN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN23.Text)] = "";
            }

            if (chckBxCN24.Checked && !cmbBoxCN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN24.Text)] = lblCN24.Text;
            }
            else if (!chckBxCN24.Checked && !cmbBoxCN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN24.Text)] = "";
            }

            if (chckBxCN25.Checked && !cmbBoxCN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN25.Text)] = lblCN25.Text;
            }
            else if (!chckBxCN25.Checked && !cmbBoxCN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN25.Text)] = "";
            }

            if (chckBxCN26.Checked && !cmbBoxCN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN26.Text)] = lblCN26.Text;
            }
            else if (!chckBxCN26.Checked && !cmbBoxCN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN26.Text)] = "";
            }

            if (chckBxCN27.Checked && !cmbBoxCN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN27.Text)] = lblCN27.Text;
            }
            else if (!chckBxCN27.Checked && !cmbBoxCN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN27.Text)] = "";
            }

            if (chckBxCN28.Checked && !cmbBoxCN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN28.Text)] = lblCN28.Text;
            }
            else if (!chckBxCN28.Checked && !cmbBoxCN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN28.Text)] = "";
            }

            if (chckBxCN29.Checked && !cmbBoxCN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN29.Text)] = lblCN29.Text;
            }
            else if (!chckBxCN29.Checked && !cmbBoxCN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN29.Text)] = "";
            }

            if (chckBxCN30.Checked && !cmbBoxCN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN30.Text)] = lblCN30.Text;
            }
            else if (!chckBxCN30.Checked && !cmbBoxCN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN1.Text)] = "";
            }

            for (int i = 0; i < prvw.Length; i++)
            {
                preview = preview + prvw[i];

                if (prvw[i].Equals(""))
                {
                    preview = preview + " ";
                }
            }

            txtBoxPreview.Text = preview;

            UpdatePreview();
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            for (int i = 0; i < prvw.Length; i++)
            {
                prvw[i] = "";
            }

            preview = "";

            if (chckBxCN1.Checked && !cmbBoxCN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN1.Text)] = lblCN1.Text;
            }
            else if (!chckBxCN1.Checked && !cmbBoxCN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN1.Text)] = "";
            }

            if (chckBxCN2.Checked && !cmbBoxCN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN2.Text)] = lblCN2.Text;
            }
            else if (!chckBxCN2.Checked && !cmbBoxCN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN2.Text)] = "";
            }

            if (chckBxCN3.Checked && !cmbBoxCN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN3.Text)] = lblCN3.Text;
            }
            else if (!chckBxCN3.Checked && !cmbBoxCN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN3.Text)] = "";
            }

            if (chckBxCN4.Checked && !cmbBoxCN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN4.Text)] = lblCN4.Text;
            }
            else if (!chckBxCN4.Checked && !cmbBoxCN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN4.Text)] = "";
            }

            if (chckBxCN5.Checked && !cmbBoxCN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN5.Text)] = lblCN5.Text;
            }
            else if (!chckBxCN5.Checked && !cmbBoxCN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN5.Text)] = "";
            }

            if (chckBxCN6.Checked && !cmbBoxCN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN6.Text)] = lblCN6.Text;
            }
            else if (!chckBxCN6.Checked && !cmbBoxCN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN6.Text)] = "";
            }

            if (chckBxCN7.Checked && !cmbBoxCN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN7.Text)] = lblCN7.Text;
            }
            else if (!chckBxCN7.Checked && !cmbBoxCN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN7.Text)] = "";
            }

            if (chckBxCN8.Checked && !cmbBoxCN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN8.Text)] = lblCN8.Text;
            }
            else if (!chckBxCN8.Checked && !cmbBoxCN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN8.Text)] = "";
            }

            if (chckBxCN9.Checked && !cmbBoxCN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN9.Text)] = lblCN9.Text;
            }
            else if (!chckBxCN9.Checked && !cmbBoxCN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN9.Text)] = "";
            }

            if (chckBxCN10.Checked && !cmbBoxCN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN10.Text)] = lblCN10.Text;
            }
            else if (!chckBxCN10.Checked && !cmbBoxCN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN10.Text)] = "";
            }

            if (chckBxCN11.Checked && !cmbBoxCN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN11.Text)] = lblCN11.Text;
            }
            else if (!chckBxCN11.Checked && !cmbBoxCN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN11.Text)] = "";
            }

            if (chckBxCN12.Checked && !cmbBoxCN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN12.Text)] = lblCN12.Text;
            }
            else if (!chckBxCN12.Checked && !cmbBoxCN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN12.Text)] = "";
            }

            if (chckBxCN13.Checked && !cmbBoxCN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN13.Text)] = lblCN13.Text;
            }
            else if (!chckBxCN13.Checked && !cmbBoxCN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN13.Text)] = "";
            }

            if (chckBxCN14.Checked && !cmbBoxCN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN14.Text)] = lblCN14.Text;
            }
            else if (!chckBxCN14.Checked && !cmbBoxCN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN14.Text)] = "";
            }

            if (chckBxCN15.Checked && !cmbBoxCN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN15.Text)] = lblCN15.Text;
            }
            else if (!chckBxCN15.Checked && !cmbBoxCN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN15.Text)] = "";
            }

            if (chckBxCN16.Checked && !cmbBoxCN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN16.Text)] = lblCN16.Text;
            }
            else if (!chckBxCN16.Checked && !cmbBoxCN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN16.Text)] = "";
            }

            if (chckBxCN17.Checked && !cmbBoxCN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN17.Text)] = lblCN17.Text;
            }
            else if (!chckBxCN17.Checked && !cmbBoxCN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN17.Text)] = "";
            }

            if (chckBxCN18.Checked && !cmbBoxCN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN18.Text)] = lblCN18.Text;
            }
            else if (!chckBxCN18.Checked && !cmbBoxCN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN18.Text)] = "";
            }

            if (chckBxCN19.Checked && !cmbBoxCN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN19.Text)] = lblCN19.Text;
            }
            else if (!chckBxCN19.Checked && !cmbBoxCN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN19.Text)] = "";
            }

            if (chckBxCN20.Checked && !cmbBoxCN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN20.Text)] = lblCN20.Text;
            }
            else if (!chckBxCN20.Checked && !cmbBoxCN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN20.Text)] = "";
            }

            if (chckBxCN21.Checked && !cmbBoxCN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN21.Text)] = lblCN21.Text;
            }
            else if (!chckBxCN21.Checked && !cmbBoxCN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN21.Text)] = "";
            }

            if (chckBxCN22.Checked && !cmbBoxCN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN22.Text)] = lblCN22.Text;
            }
            else if (!chckBxCN22.Checked && !cmbBoxCN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN22.Text)] = "";
            }

            if (chckBxCN23.Checked && !cmbBoxCN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN23.Text)] = lblCN23.Text;
            }
            else if (!chckBxCN23.Checked && !cmbBoxCN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN23.Text)] = "";
            }

            if (chckBxCN24.Checked && !cmbBoxCN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN24.Text)] = lblCN24.Text;
            }
            else if (!chckBxCN24.Checked && !cmbBoxCN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN24.Text)] = "";
            }

            if (chckBxCN25.Checked && !cmbBoxCN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN25.Text)] = lblCN25.Text;
            }
            else if (!chckBxCN25.Checked && !cmbBoxCN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN25.Text)] = "";
            }

            if (chckBxCN26.Checked && !cmbBoxCN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN26.Text)] = lblCN26.Text;
            }
            else if (!chckBxCN26.Checked && !cmbBoxCN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN26.Text)] = "";
            }

            if (chckBxCN27.Checked && !cmbBoxCN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN27.Text)] = lblCN27.Text;
            }
            else if (!chckBxCN27.Checked && !cmbBoxCN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN27.Text)] = "";
            }

            if (chckBxCN28.Checked && !cmbBoxCN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN28.Text)] = lblCN28.Text;
            }
            else if (!chckBxCN28.Checked && !cmbBoxCN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN28.Text)] = "";
            }

            if (chckBxCN29.Checked && !cmbBoxCN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN29.Text)] = lblCN29.Text;
            }
            else if (!chckBxCN29.Checked && !cmbBoxCN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN29.Text)] = "";
            }

            if (chckBxCN30.Checked && !cmbBoxCN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN30.Text)] = lblCN30.Text;
            }
            else if (!chckBxCN30.Checked && !cmbBoxCN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxCN1.Text)] = "";
            }

            for (int i = 0; i < prvw.Length; i++)
            {
                preview = preview + prvw[i];

                if (prvw[i].Equals(""))
                {
                    preview = preview + " ";
                }
            }

            txtBoxPreview.Text = preview;

            UpdatePreview();
        }

        private void ComboBoxSetup()
        {
            cmbBoxBreaks1.Items.Add("1");
            cmbBoxBreaks1.Items.Add("2");
            cmbBoxBreaks1.Items.Add("3");
            cmbBoxBreaks1.Items.Add("4");
            cmbBoxBreaks1.Items.Add("5");
            cmbBoxBreaks1.Items.Add("6");
            cmbBoxBreaks1.Items.Add("7");
            cmbBoxBreaks1.Items.Add("8");
            cmbBoxBreaks1.Items.Add("9");
            cmbBoxBreaks1.Items.Add("10");
            cmbBoxBreaks1.Items.Add("11");
            cmbBoxBreaks1.Items.Add("12");
            cmbBoxBreaks1.Items.Add("13");
            cmbBoxBreaks1.Items.Add("14");
            cmbBoxBreaks1.Items.Add("15");
            cmbBoxBreaks1.Items.Add("16");
            cmbBoxBreaks1.Items.Add("17");
            cmbBoxBreaks1.Items.Add("18");
            cmbBoxBreaks1.Items.Add("19");
            cmbBoxBreaks1.Items.Add("20");
            cmbBoxBreaks1.Items.Add("21");
            cmbBoxBreaks1.Items.Add("22");
            cmbBoxBreaks1.Items.Add("23");
            cmbBoxBreaks1.Items.Add("24");
            cmbBoxBreaks1.Items.Add("25");
            cmbBoxBreaks1.Items.Add("26");
            cmbBoxBreaks1.Items.Add("27");
            cmbBoxBreaks1.Items.Add("28");
            cmbBoxBreaks1.Items.Add("29");
            cmbBoxBreaks1.Items.Add("30");

            cmbBoxBreaks2.Items.Add("1");
            cmbBoxBreaks2.Items.Add("2");
            cmbBoxBreaks2.Items.Add("3");
            cmbBoxBreaks2.Items.Add("4");
            cmbBoxBreaks2.Items.Add("5");
            cmbBoxBreaks2.Items.Add("6");
            cmbBoxBreaks2.Items.Add("7");
            cmbBoxBreaks2.Items.Add("8");
            cmbBoxBreaks2.Items.Add("9");
            cmbBoxBreaks2.Items.Add("10");
            cmbBoxBreaks2.Items.Add("11");
            cmbBoxBreaks2.Items.Add("12");
            cmbBoxBreaks2.Items.Add("13");
            cmbBoxBreaks2.Items.Add("14");
            cmbBoxBreaks2.Items.Add("15");
            cmbBoxBreaks2.Items.Add("16");
            cmbBoxBreaks2.Items.Add("17");
            cmbBoxBreaks2.Items.Add("18");
            cmbBoxBreaks2.Items.Add("19");
            cmbBoxBreaks2.Items.Add("20");
            cmbBoxBreaks2.Items.Add("21");
            cmbBoxBreaks2.Items.Add("22");
            cmbBoxBreaks2.Items.Add("23");
            cmbBoxBreaks2.Items.Add("24");
            cmbBoxBreaks2.Items.Add("25");
            cmbBoxBreaks2.Items.Add("26");
            cmbBoxBreaks2.Items.Add("27");
            cmbBoxBreaks2.Items.Add("28");
            cmbBoxBreaks2.Items.Add("29");
            cmbBoxBreaks2.Items.Add("30");

            cmbBoxBreaks3.Items.Add("1");
            cmbBoxBreaks3.Items.Add("2");
            cmbBoxBreaks3.Items.Add("3");
            cmbBoxBreaks3.Items.Add("4");
            cmbBoxBreaks3.Items.Add("5");
            cmbBoxBreaks3.Items.Add("6");
            cmbBoxBreaks3.Items.Add("7");
            cmbBoxBreaks3.Items.Add("8");
            cmbBoxBreaks3.Items.Add("9");
            cmbBoxBreaks3.Items.Add("10");
            cmbBoxBreaks3.Items.Add("11");
            cmbBoxBreaks3.Items.Add("12");
            cmbBoxBreaks3.Items.Add("13");
            cmbBoxBreaks3.Items.Add("14");
            cmbBoxBreaks3.Items.Add("15");
            cmbBoxBreaks3.Items.Add("16");
            cmbBoxBreaks3.Items.Add("17");
            cmbBoxBreaks3.Items.Add("18");
            cmbBoxBreaks3.Items.Add("19");
            cmbBoxBreaks3.Items.Add("20");
            cmbBoxBreaks3.Items.Add("21");
            cmbBoxBreaks3.Items.Add("22");
            cmbBoxBreaks3.Items.Add("23");
            cmbBoxBreaks3.Items.Add("24");
            cmbBoxBreaks3.Items.Add("25");
            cmbBoxBreaks3.Items.Add("26");
            cmbBoxBreaks3.Items.Add("27");
            cmbBoxBreaks3.Items.Add("28");
            cmbBoxBreaks3.Items.Add("29");
            cmbBoxBreaks3.Items.Add("30");

            cmbBoxBreaks4.Items.Add("1");
            cmbBoxBreaks4.Items.Add("2");
            cmbBoxBreaks4.Items.Add("3");
            cmbBoxBreaks4.Items.Add("4");
            cmbBoxBreaks4.Items.Add("5");
            cmbBoxBreaks4.Items.Add("6");
            cmbBoxBreaks4.Items.Add("7");
            cmbBoxBreaks4.Items.Add("8");
            cmbBoxBreaks4.Items.Add("9");
            cmbBoxBreaks4.Items.Add("10");
            cmbBoxBreaks4.Items.Add("11");
            cmbBoxBreaks4.Items.Add("12");
            cmbBoxBreaks4.Items.Add("13");
            cmbBoxBreaks4.Items.Add("14");
            cmbBoxBreaks4.Items.Add("15");
            cmbBoxBreaks4.Items.Add("16");
            cmbBoxBreaks4.Items.Add("17");
            cmbBoxBreaks4.Items.Add("18");
            cmbBoxBreaks4.Items.Add("19");
            cmbBoxBreaks4.Items.Add("20");
            cmbBoxBreaks4.Items.Add("21");
            cmbBoxBreaks4.Items.Add("22");
            cmbBoxBreaks4.Items.Add("23");
            cmbBoxBreaks4.Items.Add("24");
            cmbBoxBreaks4.Items.Add("25");
            cmbBoxBreaks4.Items.Add("26");
            cmbBoxBreaks4.Items.Add("27");
            cmbBoxBreaks4.Items.Add("28");
            cmbBoxBreaks4.Items.Add("29");
            cmbBoxBreaks4.Items.Add("30");

            cmbBoxBreaks5.Items.Add("1");
            cmbBoxBreaks5.Items.Add("2");
            cmbBoxBreaks5.Items.Add("3");
            cmbBoxBreaks5.Items.Add("4");
            cmbBoxBreaks5.Items.Add("5");
            cmbBoxBreaks5.Items.Add("6");
            cmbBoxBreaks5.Items.Add("7");
            cmbBoxBreaks5.Items.Add("8");
            cmbBoxBreaks5.Items.Add("9");
            cmbBoxBreaks5.Items.Add("10");
            cmbBoxBreaks5.Items.Add("11");
            cmbBoxBreaks5.Items.Add("12");
            cmbBoxBreaks5.Items.Add("13");
            cmbBoxBreaks5.Items.Add("14");
            cmbBoxBreaks5.Items.Add("15");
            cmbBoxBreaks5.Items.Add("16");
            cmbBoxBreaks5.Items.Add("17");
            cmbBoxBreaks5.Items.Add("18");
            cmbBoxBreaks5.Items.Add("19");
            cmbBoxBreaks5.Items.Add("20");
            cmbBoxBreaks5.Items.Add("21");
            cmbBoxBreaks5.Items.Add("22");
            cmbBoxBreaks5.Items.Add("23");
            cmbBoxBreaks5.Items.Add("24");
            cmbBoxBreaks5.Items.Add("25");
            cmbBoxBreaks5.Items.Add("26");
            cmbBoxBreaks5.Items.Add("27");
            cmbBoxBreaks5.Items.Add("28");
            cmbBoxBreaks5.Items.Add("29");
            cmbBoxBreaks5.Items.Add("30");

            cmbBoxCN1.Items.Add("1");
            cmbBoxCN1.Items.Add("2");
            cmbBoxCN1.Items.Add("3");
            cmbBoxCN1.Items.Add("4");
            cmbBoxCN1.Items.Add("5");
            cmbBoxCN1.Items.Add("6");
            cmbBoxCN1.Items.Add("7");
            cmbBoxCN1.Items.Add("8");
            cmbBoxCN1.Items.Add("9");
            cmbBoxCN1.Items.Add("10");
            cmbBoxCN1.Items.Add("11");
            cmbBoxCN1.Items.Add("12");
            cmbBoxCN1.Items.Add("13");
            cmbBoxCN1.Items.Add("14");
            cmbBoxCN1.Items.Add("15");
            cmbBoxCN1.Items.Add("16");
            cmbBoxCN1.Items.Add("17");
            cmbBoxCN1.Items.Add("18");
            cmbBoxCN1.Items.Add("19");
            cmbBoxCN1.Items.Add("20");
            cmbBoxCN1.Items.Add("21");
            cmbBoxCN1.Items.Add("22");
            cmbBoxCN1.Items.Add("23");
            cmbBoxCN1.Items.Add("24");
            cmbBoxCN1.Items.Add("25");
            cmbBoxCN1.Items.Add("26");
            cmbBoxCN1.Items.Add("27");
            cmbBoxCN1.Items.Add("28");
            cmbBoxCN1.Items.Add("29");
            cmbBoxCN1.Items.Add("30");

            cmbBoxCN2.Items.Add("1");
            cmbBoxCN2.Items.Add("2");
            cmbBoxCN2.Items.Add("3");
            cmbBoxCN2.Items.Add("4");
            cmbBoxCN2.Items.Add("5");
            cmbBoxCN2.Items.Add("6");
            cmbBoxCN2.Items.Add("7");
            cmbBoxCN2.Items.Add("8");
            cmbBoxCN2.Items.Add("9");
            cmbBoxCN2.Items.Add("10");
            cmbBoxCN2.Items.Add("11");
            cmbBoxCN2.Items.Add("12");
            cmbBoxCN2.Items.Add("13");
            cmbBoxCN2.Items.Add("14");
            cmbBoxCN2.Items.Add("15");
            cmbBoxCN2.Items.Add("16");
            cmbBoxCN2.Items.Add("17");
            cmbBoxCN2.Items.Add("18");
            cmbBoxCN2.Items.Add("19");
            cmbBoxCN2.Items.Add("20");
            cmbBoxCN2.Items.Add("21");
            cmbBoxCN2.Items.Add("22");
            cmbBoxCN2.Items.Add("23");
            cmbBoxCN2.Items.Add("24");
            cmbBoxCN2.Items.Add("25");
            cmbBoxCN2.Items.Add("26");
            cmbBoxCN2.Items.Add("27");
            cmbBoxCN2.Items.Add("28");
            cmbBoxCN2.Items.Add("29");
            cmbBoxCN2.Items.Add("30");

            cmbBoxCN3.Items.Add("1");
            cmbBoxCN3.Items.Add("2");
            cmbBoxCN3.Items.Add("3");
            cmbBoxCN3.Items.Add("4");
            cmbBoxCN3.Items.Add("5");
            cmbBoxCN3.Items.Add("6");
            cmbBoxCN3.Items.Add("7");
            cmbBoxCN3.Items.Add("8");
            cmbBoxCN3.Items.Add("9");
            cmbBoxCN3.Items.Add("10");
            cmbBoxCN3.Items.Add("11");
            cmbBoxCN3.Items.Add("12");
            cmbBoxCN3.Items.Add("13");
            cmbBoxCN3.Items.Add("14");
            cmbBoxCN3.Items.Add("15");
            cmbBoxCN3.Items.Add("16");
            cmbBoxCN3.Items.Add("17");
            cmbBoxCN3.Items.Add("18");
            cmbBoxCN3.Items.Add("19");
            cmbBoxCN3.Items.Add("20");
            cmbBoxCN3.Items.Add("21");
            cmbBoxCN3.Items.Add("22");
            cmbBoxCN3.Items.Add("23");
            cmbBoxCN3.Items.Add("24");
            cmbBoxCN3.Items.Add("25");
            cmbBoxCN3.Items.Add("26");
            cmbBoxCN3.Items.Add("27");
            cmbBoxCN3.Items.Add("28");
            cmbBoxCN3.Items.Add("29");
            cmbBoxCN3.Items.Add("30");
            cmbBoxCN4.Items.Add("1");
            cmbBoxCN4.Items.Add("2");
            cmbBoxCN4.Items.Add("3");
            cmbBoxCN4.Items.Add("4");
            cmbBoxCN4.Items.Add("5");
            cmbBoxCN4.Items.Add("6");
            cmbBoxCN4.Items.Add("7");
            cmbBoxCN4.Items.Add("8");
            cmbBoxCN4.Items.Add("9");
            cmbBoxCN4.Items.Add("10");
            cmbBoxCN4.Items.Add("11");
            cmbBoxCN4.Items.Add("12");
            cmbBoxCN4.Items.Add("13");
            cmbBoxCN4.Items.Add("14");
            cmbBoxCN4.Items.Add("15");
            cmbBoxCN4.Items.Add("16");
            cmbBoxCN4.Items.Add("17");
            cmbBoxCN4.Items.Add("18");
            cmbBoxCN4.Items.Add("19");
            cmbBoxCN4.Items.Add("20");
            cmbBoxCN4.Items.Add("21");
            cmbBoxCN4.Items.Add("22");
            cmbBoxCN4.Items.Add("23");
            cmbBoxCN4.Items.Add("24");
            cmbBoxCN4.Items.Add("25");
            cmbBoxCN4.Items.Add("26");
            cmbBoxCN4.Items.Add("27");
            cmbBoxCN4.Items.Add("28");
            cmbBoxCN4.Items.Add("29");
            cmbBoxCN4.Items.Add("30");

            cmbBoxCN5.Items.Add("1");
            cmbBoxCN5.Items.Add("2");
            cmbBoxCN5.Items.Add("3");
            cmbBoxCN5.Items.Add("4");
            cmbBoxCN5.Items.Add("5");
            cmbBoxCN5.Items.Add("6");
            cmbBoxCN5.Items.Add("7");
            cmbBoxCN5.Items.Add("8");
            cmbBoxCN5.Items.Add("9");
            cmbBoxCN5.Items.Add("10");
            cmbBoxCN5.Items.Add("11");
            cmbBoxCN5.Items.Add("12");
            cmbBoxCN5.Items.Add("13");
            cmbBoxCN5.Items.Add("14");
            cmbBoxCN5.Items.Add("15");
            cmbBoxCN5.Items.Add("16");
            cmbBoxCN5.Items.Add("17");
            cmbBoxCN5.Items.Add("18");
            cmbBoxCN5.Items.Add("19");
            cmbBoxCN5.Items.Add("20");
            cmbBoxCN5.Items.Add("21");
            cmbBoxCN5.Items.Add("22");
            cmbBoxCN5.Items.Add("23");
            cmbBoxCN5.Items.Add("24");
            cmbBoxCN5.Items.Add("25");
            cmbBoxCN5.Items.Add("26");
            cmbBoxCN5.Items.Add("27");
            cmbBoxCN5.Items.Add("28");
            cmbBoxCN5.Items.Add("29");
            cmbBoxCN5.Items.Add("30");

            cmbBoxCN6.Items.Add("1");
            cmbBoxCN6.Items.Add("2");
            cmbBoxCN6.Items.Add("3");
            cmbBoxCN6.Items.Add("4");
            cmbBoxCN6.Items.Add("5");
            cmbBoxCN6.Items.Add("6");
            cmbBoxCN6.Items.Add("7");
            cmbBoxCN6.Items.Add("8");
            cmbBoxCN6.Items.Add("9");
            cmbBoxCN6.Items.Add("10");
            cmbBoxCN6.Items.Add("11");
            cmbBoxCN6.Items.Add("12");
            cmbBoxCN6.Items.Add("13");
            cmbBoxCN6.Items.Add("14");
            cmbBoxCN6.Items.Add("15");
            cmbBoxCN6.Items.Add("16");
            cmbBoxCN6.Items.Add("17");
            cmbBoxCN6.Items.Add("18");
            cmbBoxCN6.Items.Add("19");
            cmbBoxCN6.Items.Add("20");
            cmbBoxCN6.Items.Add("21");
            cmbBoxCN6.Items.Add("22");
            cmbBoxCN6.Items.Add("23");
            cmbBoxCN6.Items.Add("24");
            cmbBoxCN6.Items.Add("25");
            cmbBoxCN6.Items.Add("26");
            cmbBoxCN6.Items.Add("27");
            cmbBoxCN6.Items.Add("28");
            cmbBoxCN6.Items.Add("29");
            cmbBoxCN6.Items.Add("30");

            cmbBoxCN7.Items.Add("1");
            cmbBoxCN7.Items.Add("2");
            cmbBoxCN7.Items.Add("3");
            cmbBoxCN7.Items.Add("4");
            cmbBoxCN7.Items.Add("5");
            cmbBoxCN7.Items.Add("6");
            cmbBoxCN7.Items.Add("7");
            cmbBoxCN7.Items.Add("8");
            cmbBoxCN7.Items.Add("9");
            cmbBoxCN7.Items.Add("10");
            cmbBoxCN7.Items.Add("11");
            cmbBoxCN7.Items.Add("12");
            cmbBoxCN7.Items.Add("13");
            cmbBoxCN7.Items.Add("14");
            cmbBoxCN7.Items.Add("15");
            cmbBoxCN7.Items.Add("16");
            cmbBoxCN7.Items.Add("17");
            cmbBoxCN7.Items.Add("18");
            cmbBoxCN7.Items.Add("19");
            cmbBoxCN7.Items.Add("20");
            cmbBoxCN7.Items.Add("21");
            cmbBoxCN7.Items.Add("22");
            cmbBoxCN7.Items.Add("23");
            cmbBoxCN7.Items.Add("24");
            cmbBoxCN7.Items.Add("25");
            cmbBoxCN7.Items.Add("26");
            cmbBoxCN7.Items.Add("27");
            cmbBoxCN7.Items.Add("28");
            cmbBoxCN7.Items.Add("29");
            cmbBoxCN7.Items.Add("30");

            cmbBoxCN8.Items.Add("1");
            cmbBoxCN8.Items.Add("2");
            cmbBoxCN8.Items.Add("3");
            cmbBoxCN8.Items.Add("4");
            cmbBoxCN8.Items.Add("5");
            cmbBoxCN8.Items.Add("6");
            cmbBoxCN8.Items.Add("7");
            cmbBoxCN8.Items.Add("8");
            cmbBoxCN8.Items.Add("9");
            cmbBoxCN8.Items.Add("10");
            cmbBoxCN8.Items.Add("11");
            cmbBoxCN8.Items.Add("12");
            cmbBoxCN8.Items.Add("13");
            cmbBoxCN8.Items.Add("14");
            cmbBoxCN8.Items.Add("15");
            cmbBoxCN8.Items.Add("16");
            cmbBoxCN8.Items.Add("17");
            cmbBoxCN8.Items.Add("18");
            cmbBoxCN8.Items.Add("19");
            cmbBoxCN8.Items.Add("20");
            cmbBoxCN8.Items.Add("21");
            cmbBoxCN8.Items.Add("22");
            cmbBoxCN8.Items.Add("23");
            cmbBoxCN8.Items.Add("24");
            cmbBoxCN8.Items.Add("25");
            cmbBoxCN8.Items.Add("26");
            cmbBoxCN8.Items.Add("27");
            cmbBoxCN8.Items.Add("28");
            cmbBoxCN8.Items.Add("29");
            cmbBoxCN8.Items.Add("30");

            cmbBoxCN9.Items.Add("1");
            cmbBoxCN9.Items.Add("2");
            cmbBoxCN9.Items.Add("3");
            cmbBoxCN9.Items.Add("4");
            cmbBoxCN9.Items.Add("5");
            cmbBoxCN9.Items.Add("6");
            cmbBoxCN9.Items.Add("7");
            cmbBoxCN9.Items.Add("8");
            cmbBoxCN9.Items.Add("9");
            cmbBoxCN9.Items.Add("10");
            cmbBoxCN9.Items.Add("11");
            cmbBoxCN9.Items.Add("12");
            cmbBoxCN9.Items.Add("13");
            cmbBoxCN9.Items.Add("14");
            cmbBoxCN9.Items.Add("15");
            cmbBoxCN9.Items.Add("16");
            cmbBoxCN9.Items.Add("17");
            cmbBoxCN9.Items.Add("18");
            cmbBoxCN9.Items.Add("19");
            cmbBoxCN9.Items.Add("20");
            cmbBoxCN9.Items.Add("21");
            cmbBoxCN9.Items.Add("22");
            cmbBoxCN9.Items.Add("23");
            cmbBoxCN9.Items.Add("24");
            cmbBoxCN9.Items.Add("25");
            cmbBoxCN9.Items.Add("26");
            cmbBoxCN9.Items.Add("27");
            cmbBoxCN9.Items.Add("28");
            cmbBoxCN9.Items.Add("29");
            cmbBoxCN9.Items.Add("30");

            cmbBoxCN10.Items.Add("1");
            cmbBoxCN10.Items.Add("2");
            cmbBoxCN10.Items.Add("3");
            cmbBoxCN10.Items.Add("4");
            cmbBoxCN10.Items.Add("5");
            cmbBoxCN10.Items.Add("6");
            cmbBoxCN10.Items.Add("7");
            cmbBoxCN10.Items.Add("8");
            cmbBoxCN10.Items.Add("9");
            cmbBoxCN10.Items.Add("10");
            cmbBoxCN10.Items.Add("11");
            cmbBoxCN10.Items.Add("12");
            cmbBoxCN10.Items.Add("13");
            cmbBoxCN10.Items.Add("14");
            cmbBoxCN10.Items.Add("15");
            cmbBoxCN10.Items.Add("16");
            cmbBoxCN10.Items.Add("17");
            cmbBoxCN10.Items.Add("18");
            cmbBoxCN10.Items.Add("19");
            cmbBoxCN10.Items.Add("20");
            cmbBoxCN10.Items.Add("21");
            cmbBoxCN10.Items.Add("22");
            cmbBoxCN10.Items.Add("23");
            cmbBoxCN10.Items.Add("24");
            cmbBoxCN10.Items.Add("25");
            cmbBoxCN10.Items.Add("26");
            cmbBoxCN10.Items.Add("27");
            cmbBoxCN10.Items.Add("28");
            cmbBoxCN10.Items.Add("29");
            cmbBoxCN10.Items.Add("30");

            cmbBoxCN11.Items.Add("1");
            cmbBoxCN11.Items.Add("2");
            cmbBoxCN11.Items.Add("3");
            cmbBoxCN11.Items.Add("4");
            cmbBoxCN11.Items.Add("5");
            cmbBoxCN11.Items.Add("6");
            cmbBoxCN11.Items.Add("7");
            cmbBoxCN11.Items.Add("8");
            cmbBoxCN11.Items.Add("9");
            cmbBoxCN11.Items.Add("10");
            cmbBoxCN11.Items.Add("11");
            cmbBoxCN11.Items.Add("12");
            cmbBoxCN11.Items.Add("13");
            cmbBoxCN11.Items.Add("14");
            cmbBoxCN11.Items.Add("15");
            cmbBoxCN11.Items.Add("16");
            cmbBoxCN11.Items.Add("17");
            cmbBoxCN11.Items.Add("18");
            cmbBoxCN11.Items.Add("19");
            cmbBoxCN11.Items.Add("20");
            cmbBoxCN11.Items.Add("21");
            cmbBoxCN11.Items.Add("22");
            cmbBoxCN11.Items.Add("23");
            cmbBoxCN11.Items.Add("24");
            cmbBoxCN11.Items.Add("25");
            cmbBoxCN11.Items.Add("26");
            cmbBoxCN11.Items.Add("27");
            cmbBoxCN11.Items.Add("28");
            cmbBoxCN11.Items.Add("29");
            cmbBoxCN11.Items.Add("30");

            cmbBoxCN12.Items.Add("1");
            cmbBoxCN12.Items.Add("2");
            cmbBoxCN12.Items.Add("3");
            cmbBoxCN12.Items.Add("4");
            cmbBoxCN12.Items.Add("5");
            cmbBoxCN12.Items.Add("6");
            cmbBoxCN12.Items.Add("7");
            cmbBoxCN12.Items.Add("8");
            cmbBoxCN12.Items.Add("9");
            cmbBoxCN12.Items.Add("10");
            cmbBoxCN12.Items.Add("11");
            cmbBoxCN12.Items.Add("12");
            cmbBoxCN12.Items.Add("13");
            cmbBoxCN12.Items.Add("14");
            cmbBoxCN12.Items.Add("15");
            cmbBoxCN12.Items.Add("16");
            cmbBoxCN12.Items.Add("17");
            cmbBoxCN12.Items.Add("18");
            cmbBoxCN12.Items.Add("19");
            cmbBoxCN12.Items.Add("20");
            cmbBoxCN12.Items.Add("21");
            cmbBoxCN12.Items.Add("22");
            cmbBoxCN12.Items.Add("23");
            cmbBoxCN12.Items.Add("24");
            cmbBoxCN12.Items.Add("25");
            cmbBoxCN12.Items.Add("26");
            cmbBoxCN12.Items.Add("27");
            cmbBoxCN12.Items.Add("28");
            cmbBoxCN12.Items.Add("29");
            cmbBoxCN12.Items.Add("30");

            cmbBoxCN13.Items.Add("1");
            cmbBoxCN13.Items.Add("2");
            cmbBoxCN13.Items.Add("3");
            cmbBoxCN13.Items.Add("4");
            cmbBoxCN13.Items.Add("5");
            cmbBoxCN13.Items.Add("6");
            cmbBoxCN13.Items.Add("7");
            cmbBoxCN13.Items.Add("8");
            cmbBoxCN13.Items.Add("9");
            cmbBoxCN13.Items.Add("10");
            cmbBoxCN13.Items.Add("11");
            cmbBoxCN13.Items.Add("12");
            cmbBoxCN13.Items.Add("13");
            cmbBoxCN13.Items.Add("14");
            cmbBoxCN13.Items.Add("15");
            cmbBoxCN13.Items.Add("16");
            cmbBoxCN13.Items.Add("17");
            cmbBoxCN13.Items.Add("18");
            cmbBoxCN13.Items.Add("19");
            cmbBoxCN13.Items.Add("20");
            cmbBoxCN13.Items.Add("21");
            cmbBoxCN13.Items.Add("22");
            cmbBoxCN13.Items.Add("23");
            cmbBoxCN13.Items.Add("24");
            cmbBoxCN13.Items.Add("25");
            cmbBoxCN13.Items.Add("26");
            cmbBoxCN13.Items.Add("27");
            cmbBoxCN13.Items.Add("28");
            cmbBoxCN13.Items.Add("29");
            cmbBoxCN13.Items.Add("30");

            cmbBoxCN14.Items.Add("1");
            cmbBoxCN14.Items.Add("2");
            cmbBoxCN14.Items.Add("3");
            cmbBoxCN14.Items.Add("4");
            cmbBoxCN14.Items.Add("5");
            cmbBoxCN14.Items.Add("6");
            cmbBoxCN14.Items.Add("7");
            cmbBoxCN14.Items.Add("8");
            cmbBoxCN14.Items.Add("9");
            cmbBoxCN14.Items.Add("10");
            cmbBoxCN14.Items.Add("11");
            cmbBoxCN14.Items.Add("12");
            cmbBoxCN14.Items.Add("13");
            cmbBoxCN14.Items.Add("14");
            cmbBoxCN14.Items.Add("15");
            cmbBoxCN14.Items.Add("16");
            cmbBoxCN14.Items.Add("17");
            cmbBoxCN14.Items.Add("18");
            cmbBoxCN14.Items.Add("19");
            cmbBoxCN14.Items.Add("20");
            cmbBoxCN14.Items.Add("21");
            cmbBoxCN14.Items.Add("22");
            cmbBoxCN14.Items.Add("23");
            cmbBoxCN14.Items.Add("24");
            cmbBoxCN14.Items.Add("25");
            cmbBoxCN14.Items.Add("26");
            cmbBoxCN14.Items.Add("27");
            cmbBoxCN14.Items.Add("28");
            cmbBoxCN14.Items.Add("29");
            cmbBoxCN14.Items.Add("30");

            cmbBoxCN15.Items.Add("1");
            cmbBoxCN15.Items.Add("2");
            cmbBoxCN15.Items.Add("3");
            cmbBoxCN15.Items.Add("4");
            cmbBoxCN15.Items.Add("5");
            cmbBoxCN15.Items.Add("6");
            cmbBoxCN15.Items.Add("7");
            cmbBoxCN15.Items.Add("8");
            cmbBoxCN15.Items.Add("9");
            cmbBoxCN15.Items.Add("10");
            cmbBoxCN15.Items.Add("11");
            cmbBoxCN15.Items.Add("12");
            cmbBoxCN15.Items.Add("13");
            cmbBoxCN15.Items.Add("14");
            cmbBoxCN15.Items.Add("15");
            cmbBoxCN15.Items.Add("16");
            cmbBoxCN15.Items.Add("17");
            cmbBoxCN15.Items.Add("18");
            cmbBoxCN15.Items.Add("19");
            cmbBoxCN15.Items.Add("20");
            cmbBoxCN15.Items.Add("21");
            cmbBoxCN15.Items.Add("22");
            cmbBoxCN15.Items.Add("23");
            cmbBoxCN15.Items.Add("24");
            cmbBoxCN15.Items.Add("25");
            cmbBoxCN15.Items.Add("26");
            cmbBoxCN15.Items.Add("27");
            cmbBoxCN15.Items.Add("28");
            cmbBoxCN15.Items.Add("29");
            cmbBoxCN15.Items.Add("30");

            cmbBoxCN16.Items.Add("1");
            cmbBoxCN16.Items.Add("2");
            cmbBoxCN16.Items.Add("3");
            cmbBoxCN16.Items.Add("4");
            cmbBoxCN16.Items.Add("5");
            cmbBoxCN16.Items.Add("6");
            cmbBoxCN16.Items.Add("7");
            cmbBoxCN16.Items.Add("8");
            cmbBoxCN16.Items.Add("9");
            cmbBoxCN16.Items.Add("10");
            cmbBoxCN16.Items.Add("11");
            cmbBoxCN16.Items.Add("12");
            cmbBoxCN16.Items.Add("13");
            cmbBoxCN16.Items.Add("14");
            cmbBoxCN16.Items.Add("15");
            cmbBoxCN16.Items.Add("16");
            cmbBoxCN16.Items.Add("17");
            cmbBoxCN16.Items.Add("18");
            cmbBoxCN16.Items.Add("19");
            cmbBoxCN16.Items.Add("20");
            cmbBoxCN16.Items.Add("21");
            cmbBoxCN16.Items.Add("22");
            cmbBoxCN16.Items.Add("23");
            cmbBoxCN16.Items.Add("24");
            cmbBoxCN16.Items.Add("25");
            cmbBoxCN16.Items.Add("26");
            cmbBoxCN16.Items.Add("27");
            cmbBoxCN16.Items.Add("28");
            cmbBoxCN16.Items.Add("29");
            cmbBoxCN16.Items.Add("30");

            cmbBoxCN17.Items.Add("1");
            cmbBoxCN17.Items.Add("2");
            cmbBoxCN17.Items.Add("3");
            cmbBoxCN17.Items.Add("4");
            cmbBoxCN17.Items.Add("5");
            cmbBoxCN17.Items.Add("6");
            cmbBoxCN17.Items.Add("7");
            cmbBoxCN17.Items.Add("8");
            cmbBoxCN17.Items.Add("9");
            cmbBoxCN17.Items.Add("10");
            cmbBoxCN17.Items.Add("11");
            cmbBoxCN17.Items.Add("12");
            cmbBoxCN17.Items.Add("13");
            cmbBoxCN17.Items.Add("14");
            cmbBoxCN17.Items.Add("15");
            cmbBoxCN17.Items.Add("16");
            cmbBoxCN17.Items.Add("17");
            cmbBoxCN17.Items.Add("18");
            cmbBoxCN17.Items.Add("19");
            cmbBoxCN17.Items.Add("20");
            cmbBoxCN17.Items.Add("21");
            cmbBoxCN17.Items.Add("22");
            cmbBoxCN17.Items.Add("23");
            cmbBoxCN17.Items.Add("24");
            cmbBoxCN17.Items.Add("25");
            cmbBoxCN17.Items.Add("26");
            cmbBoxCN17.Items.Add("27");
            cmbBoxCN17.Items.Add("28");
            cmbBoxCN17.Items.Add("29");
            cmbBoxCN17.Items.Add("30");

            cmbBoxCN18.Items.Add("1");
            cmbBoxCN18.Items.Add("2");
            cmbBoxCN18.Items.Add("3");
            cmbBoxCN18.Items.Add("4");
            cmbBoxCN18.Items.Add("5");
            cmbBoxCN18.Items.Add("6");
            cmbBoxCN18.Items.Add("7");
            cmbBoxCN18.Items.Add("8");
            cmbBoxCN18.Items.Add("9");
            cmbBoxCN18.Items.Add("10");
            cmbBoxCN18.Items.Add("11");
            cmbBoxCN18.Items.Add("12");
            cmbBoxCN18.Items.Add("13");
            cmbBoxCN18.Items.Add("14");
            cmbBoxCN18.Items.Add("15");
            cmbBoxCN18.Items.Add("16");
            cmbBoxCN18.Items.Add("17");
            cmbBoxCN18.Items.Add("18");
            cmbBoxCN18.Items.Add("19");
            cmbBoxCN18.Items.Add("20");
            cmbBoxCN18.Items.Add("21");
            cmbBoxCN18.Items.Add("22");
            cmbBoxCN18.Items.Add("23");
            cmbBoxCN18.Items.Add("24");
            cmbBoxCN18.Items.Add("25");
            cmbBoxCN18.Items.Add("26");
            cmbBoxCN18.Items.Add("27");
            cmbBoxCN18.Items.Add("28");
            cmbBoxCN18.Items.Add("29");
            cmbBoxCN18.Items.Add("30");

            cmbBoxCN19.Items.Add("1");
            cmbBoxCN19.Items.Add("2");
            cmbBoxCN19.Items.Add("3");
            cmbBoxCN19.Items.Add("4");
            cmbBoxCN19.Items.Add("5");
            cmbBoxCN19.Items.Add("6");
            cmbBoxCN19.Items.Add("7");
            cmbBoxCN19.Items.Add("8");
            cmbBoxCN19.Items.Add("9");
            cmbBoxCN19.Items.Add("10");
            cmbBoxCN19.Items.Add("11");
            cmbBoxCN19.Items.Add("12");
            cmbBoxCN19.Items.Add("13");
            cmbBoxCN19.Items.Add("14");
            cmbBoxCN19.Items.Add("15");
            cmbBoxCN19.Items.Add("16");
            cmbBoxCN19.Items.Add("17");
            cmbBoxCN19.Items.Add("18");
            cmbBoxCN19.Items.Add("19");
            cmbBoxCN19.Items.Add("20");
            cmbBoxCN19.Items.Add("21");
            cmbBoxCN19.Items.Add("22");
            cmbBoxCN19.Items.Add("23");
            cmbBoxCN19.Items.Add("24");
            cmbBoxCN19.Items.Add("25");
            cmbBoxCN19.Items.Add("26");
            cmbBoxCN19.Items.Add("27");
            cmbBoxCN19.Items.Add("28");
            cmbBoxCN19.Items.Add("29");
            cmbBoxCN19.Items.Add("30");

            cmbBoxCN20.Items.Add("1");
            cmbBoxCN20.Items.Add("2");
            cmbBoxCN20.Items.Add("3");
            cmbBoxCN20.Items.Add("4");
            cmbBoxCN20.Items.Add("5");
            cmbBoxCN20.Items.Add("6");
            cmbBoxCN20.Items.Add("7");
            cmbBoxCN20.Items.Add("8");
            cmbBoxCN20.Items.Add("9");
            cmbBoxCN20.Items.Add("10");
            cmbBoxCN20.Items.Add("11");
            cmbBoxCN20.Items.Add("12");
            cmbBoxCN20.Items.Add("13");
            cmbBoxCN20.Items.Add("14");
            cmbBoxCN20.Items.Add("15");
            cmbBoxCN20.Items.Add("16");
            cmbBoxCN20.Items.Add("17");
            cmbBoxCN20.Items.Add("18");
            cmbBoxCN20.Items.Add("19");
            cmbBoxCN20.Items.Add("20");
            cmbBoxCN20.Items.Add("21");
            cmbBoxCN20.Items.Add("22");
            cmbBoxCN20.Items.Add("23");
            cmbBoxCN20.Items.Add("24");
            cmbBoxCN20.Items.Add("25");
            cmbBoxCN20.Items.Add("26");
            cmbBoxCN20.Items.Add("27");
            cmbBoxCN20.Items.Add("28");
            cmbBoxCN20.Items.Add("29");
            cmbBoxCN20.Items.Add("30");

            cmbBoxCN21.Items.Add("1");
            cmbBoxCN21.Items.Add("2");
            cmbBoxCN21.Items.Add("3");
            cmbBoxCN21.Items.Add("4");
            cmbBoxCN21.Items.Add("5");
            cmbBoxCN21.Items.Add("6");
            cmbBoxCN21.Items.Add("7");
            cmbBoxCN21.Items.Add("8");
            cmbBoxCN21.Items.Add("9");
            cmbBoxCN21.Items.Add("10");
            cmbBoxCN21.Items.Add("11");
            cmbBoxCN21.Items.Add("12");
            cmbBoxCN21.Items.Add("13");
            cmbBoxCN21.Items.Add("14");
            cmbBoxCN21.Items.Add("15");
            cmbBoxCN21.Items.Add("16");
            cmbBoxCN21.Items.Add("17");
            cmbBoxCN21.Items.Add("18");
            cmbBoxCN21.Items.Add("19");
            cmbBoxCN21.Items.Add("20");
            cmbBoxCN21.Items.Add("21");
            cmbBoxCN21.Items.Add("22");
            cmbBoxCN21.Items.Add("23");
            cmbBoxCN21.Items.Add("24");
            cmbBoxCN21.Items.Add("25");
            cmbBoxCN21.Items.Add("26");
            cmbBoxCN21.Items.Add("27");
            cmbBoxCN21.Items.Add("28");
            cmbBoxCN21.Items.Add("29");
            cmbBoxCN21.Items.Add("30");

            cmbBoxCN22.Items.Add("1");
            cmbBoxCN22.Items.Add("2");
            cmbBoxCN22.Items.Add("3");
            cmbBoxCN22.Items.Add("4");
            cmbBoxCN22.Items.Add("5");
            cmbBoxCN22.Items.Add("6");
            cmbBoxCN22.Items.Add("7");
            cmbBoxCN22.Items.Add("8");
            cmbBoxCN22.Items.Add("9");
            cmbBoxCN22.Items.Add("10");
            cmbBoxCN22.Items.Add("11");
            cmbBoxCN22.Items.Add("12");
            cmbBoxCN22.Items.Add("13");
            cmbBoxCN22.Items.Add("14");
            cmbBoxCN22.Items.Add("15");
            cmbBoxCN22.Items.Add("16");
            cmbBoxCN22.Items.Add("17");
            cmbBoxCN22.Items.Add("18");
            cmbBoxCN22.Items.Add("19");
            cmbBoxCN22.Items.Add("20");
            cmbBoxCN22.Items.Add("21");
            cmbBoxCN22.Items.Add("22");
            cmbBoxCN22.Items.Add("23");
            cmbBoxCN22.Items.Add("24");
            cmbBoxCN22.Items.Add("25");
            cmbBoxCN22.Items.Add("26");
            cmbBoxCN22.Items.Add("27");
            cmbBoxCN22.Items.Add("28");
            cmbBoxCN22.Items.Add("29");
            cmbBoxCN22.Items.Add("30");

            cmbBoxCN23.Items.Add("1");
            cmbBoxCN23.Items.Add("2");
            cmbBoxCN23.Items.Add("3");
            cmbBoxCN23.Items.Add("4");
            cmbBoxCN23.Items.Add("5");
            cmbBoxCN23.Items.Add("6");
            cmbBoxCN23.Items.Add("7");
            cmbBoxCN23.Items.Add("8");
            cmbBoxCN23.Items.Add("9");
            cmbBoxCN23.Items.Add("10");
            cmbBoxCN23.Items.Add("11");
            cmbBoxCN23.Items.Add("12");
            cmbBoxCN23.Items.Add("13");
            cmbBoxCN23.Items.Add("14");
            cmbBoxCN23.Items.Add("15");
            cmbBoxCN23.Items.Add("16");
            cmbBoxCN23.Items.Add("17");
            cmbBoxCN23.Items.Add("18");
            cmbBoxCN23.Items.Add("19");
            cmbBoxCN23.Items.Add("20");
            cmbBoxCN23.Items.Add("21");
            cmbBoxCN23.Items.Add("22");
            cmbBoxCN23.Items.Add("23");
            cmbBoxCN23.Items.Add("24");
            cmbBoxCN23.Items.Add("25");
            cmbBoxCN23.Items.Add("26");
            cmbBoxCN23.Items.Add("27");
            cmbBoxCN23.Items.Add("28");
            cmbBoxCN23.Items.Add("29");
            cmbBoxCN23.Items.Add("30");

            cmbBoxCN24.Items.Add("1");
            cmbBoxCN24.Items.Add("2");
            cmbBoxCN24.Items.Add("3");
            cmbBoxCN24.Items.Add("4");
            cmbBoxCN24.Items.Add("5");
            cmbBoxCN24.Items.Add("6");
            cmbBoxCN24.Items.Add("7");
            cmbBoxCN24.Items.Add("8");
            cmbBoxCN24.Items.Add("9");
            cmbBoxCN24.Items.Add("10");
            cmbBoxCN24.Items.Add("11");
            cmbBoxCN24.Items.Add("12");
            cmbBoxCN24.Items.Add("13");
            cmbBoxCN24.Items.Add("14");
            cmbBoxCN24.Items.Add("15");
            cmbBoxCN24.Items.Add("16");
            cmbBoxCN24.Items.Add("17");
            cmbBoxCN24.Items.Add("18");
            cmbBoxCN24.Items.Add("19");
            cmbBoxCN24.Items.Add("20");
            cmbBoxCN24.Items.Add("21");
            cmbBoxCN24.Items.Add("22");
            cmbBoxCN24.Items.Add("23");
            cmbBoxCN24.Items.Add("24");
            cmbBoxCN24.Items.Add("25");
            cmbBoxCN24.Items.Add("26");
            cmbBoxCN24.Items.Add("27");
            cmbBoxCN24.Items.Add("28");
            cmbBoxCN24.Items.Add("29");
            cmbBoxCN24.Items.Add("30");

            cmbBoxCN25.Items.Add("1");
            cmbBoxCN25.Items.Add("2");
            cmbBoxCN25.Items.Add("3");
            cmbBoxCN25.Items.Add("4");
            cmbBoxCN25.Items.Add("5");
            cmbBoxCN25.Items.Add("6");
            cmbBoxCN25.Items.Add("7");
            cmbBoxCN25.Items.Add("8");
            cmbBoxCN25.Items.Add("9");
            cmbBoxCN25.Items.Add("10");
            cmbBoxCN25.Items.Add("11");
            cmbBoxCN25.Items.Add("12");
            cmbBoxCN25.Items.Add("13");
            cmbBoxCN25.Items.Add("14");
            cmbBoxCN25.Items.Add("15");
            cmbBoxCN25.Items.Add("16");
            cmbBoxCN25.Items.Add("17");
            cmbBoxCN25.Items.Add("18");
            cmbBoxCN25.Items.Add("19");
            cmbBoxCN25.Items.Add("20");
            cmbBoxCN25.Items.Add("21");
            cmbBoxCN25.Items.Add("22");
            cmbBoxCN25.Items.Add("23");
            cmbBoxCN25.Items.Add("24");
            cmbBoxCN25.Items.Add("25");
            cmbBoxCN25.Items.Add("26");
            cmbBoxCN25.Items.Add("27");
            cmbBoxCN25.Items.Add("28");
            cmbBoxCN25.Items.Add("29");
            cmbBoxCN25.Items.Add("30");

            cmbBoxCN26.Items.Add("1");
            cmbBoxCN26.Items.Add("2");
            cmbBoxCN26.Items.Add("3");
            cmbBoxCN26.Items.Add("4");
            cmbBoxCN26.Items.Add("5");
            cmbBoxCN26.Items.Add("6");
            cmbBoxCN26.Items.Add("7");
            cmbBoxCN26.Items.Add("8");
            cmbBoxCN26.Items.Add("9");
            cmbBoxCN26.Items.Add("10");
            cmbBoxCN26.Items.Add("11");
            cmbBoxCN26.Items.Add("12");
            cmbBoxCN26.Items.Add("13");
            cmbBoxCN26.Items.Add("14");
            cmbBoxCN26.Items.Add("15");
            cmbBoxCN26.Items.Add("16");
            cmbBoxCN26.Items.Add("17");
            cmbBoxCN26.Items.Add("18");
            cmbBoxCN26.Items.Add("19");
            cmbBoxCN26.Items.Add("20");
            cmbBoxCN26.Items.Add("21");
            cmbBoxCN26.Items.Add("22");
            cmbBoxCN26.Items.Add("23");
            cmbBoxCN26.Items.Add("24");
            cmbBoxCN26.Items.Add("25");
            cmbBoxCN26.Items.Add("26");
            cmbBoxCN26.Items.Add("27");
            cmbBoxCN26.Items.Add("28");
            cmbBoxCN26.Items.Add("29");
            cmbBoxCN26.Items.Add("30");

            cmbBoxCN27.Items.Add("1");
            cmbBoxCN27.Items.Add("2");
            cmbBoxCN27.Items.Add("3");
            cmbBoxCN27.Items.Add("4");
            cmbBoxCN27.Items.Add("5");
            cmbBoxCN27.Items.Add("6");
            cmbBoxCN27.Items.Add("7");
            cmbBoxCN27.Items.Add("8");
            cmbBoxCN27.Items.Add("9");
            cmbBoxCN27.Items.Add("10");
            cmbBoxCN27.Items.Add("11");
            cmbBoxCN27.Items.Add("12");
            cmbBoxCN27.Items.Add("13");
            cmbBoxCN27.Items.Add("14");
            cmbBoxCN27.Items.Add("15");
            cmbBoxCN27.Items.Add("16");
            cmbBoxCN27.Items.Add("17");
            cmbBoxCN27.Items.Add("18");
            cmbBoxCN27.Items.Add("19");
            cmbBoxCN27.Items.Add("20");
            cmbBoxCN27.Items.Add("21");
            cmbBoxCN27.Items.Add("22");
            cmbBoxCN27.Items.Add("23");
            cmbBoxCN27.Items.Add("24");
            cmbBoxCN27.Items.Add("25");
            cmbBoxCN27.Items.Add("26");
            cmbBoxCN27.Items.Add("27");
            cmbBoxCN27.Items.Add("28");
            cmbBoxCN27.Items.Add("29");
            cmbBoxCN27.Items.Add("30");

            cmbBoxCN28.Items.Add("1");
            cmbBoxCN28.Items.Add("2");
            cmbBoxCN28.Items.Add("3");
            cmbBoxCN28.Items.Add("4");
            cmbBoxCN28.Items.Add("5");
            cmbBoxCN28.Items.Add("6");
            cmbBoxCN28.Items.Add("7");
            cmbBoxCN28.Items.Add("8");
            cmbBoxCN28.Items.Add("9");
            cmbBoxCN28.Items.Add("10");
            cmbBoxCN28.Items.Add("11");
            cmbBoxCN28.Items.Add("12");
            cmbBoxCN28.Items.Add("13");
            cmbBoxCN28.Items.Add("14");
            cmbBoxCN28.Items.Add("15");
            cmbBoxCN28.Items.Add("16");
            cmbBoxCN28.Items.Add("17");
            cmbBoxCN28.Items.Add("18");
            cmbBoxCN28.Items.Add("19");
            cmbBoxCN28.Items.Add("20");
            cmbBoxCN28.Items.Add("21");
            cmbBoxCN28.Items.Add("22");
            cmbBoxCN28.Items.Add("23");
            cmbBoxCN28.Items.Add("24");
            cmbBoxCN28.Items.Add("25");
            cmbBoxCN28.Items.Add("26");
            cmbBoxCN28.Items.Add("27");
            cmbBoxCN28.Items.Add("28");
            cmbBoxCN28.Items.Add("29");
            cmbBoxCN28.Items.Add("30");

            cmbBoxCN29.Items.Add("1");
            cmbBoxCN29.Items.Add("2");
            cmbBoxCN29.Items.Add("3");
            cmbBoxCN29.Items.Add("4");
            cmbBoxCN29.Items.Add("5");
            cmbBoxCN29.Items.Add("6");
            cmbBoxCN29.Items.Add("7");
            cmbBoxCN29.Items.Add("8");
            cmbBoxCN29.Items.Add("9");
            cmbBoxCN29.Items.Add("10");
            cmbBoxCN29.Items.Add("11");
            cmbBoxCN29.Items.Add("12");
            cmbBoxCN29.Items.Add("13");
            cmbBoxCN29.Items.Add("14");
            cmbBoxCN29.Items.Add("15");
            cmbBoxCN29.Items.Add("16");
            cmbBoxCN29.Items.Add("17");
            cmbBoxCN29.Items.Add("18");
            cmbBoxCN29.Items.Add("19");
            cmbBoxCN29.Items.Add("20");
            cmbBoxCN29.Items.Add("21");
            cmbBoxCN29.Items.Add("22");
            cmbBoxCN29.Items.Add("23");
            cmbBoxCN29.Items.Add("24");
            cmbBoxCN29.Items.Add("25");
            cmbBoxCN29.Items.Add("26");
            cmbBoxCN29.Items.Add("27");
            cmbBoxCN29.Items.Add("28");
            cmbBoxCN29.Items.Add("29");
            cmbBoxCN29.Items.Add("30");

            cmbBoxCN30.Items.Add("1");
            cmbBoxCN30.Items.Add("2");
            cmbBoxCN30.Items.Add("3");
            cmbBoxCN30.Items.Add("4");
            cmbBoxCN30.Items.Add("5");
            cmbBoxCN30.Items.Add("6");
            cmbBoxCN30.Items.Add("7");
            cmbBoxCN30.Items.Add("8");
            cmbBoxCN30.Items.Add("9");
            cmbBoxCN30.Items.Add("10");
            cmbBoxCN30.Items.Add("11");
            cmbBoxCN30.Items.Add("12");
            cmbBoxCN30.Items.Add("13");
            cmbBoxCN30.Items.Add("14");
            cmbBoxCN30.Items.Add("15");
            cmbBoxCN30.Items.Add("16");
            cmbBoxCN30.Items.Add("17");
            cmbBoxCN30.Items.Add("18");
            cmbBoxCN30.Items.Add("19");
            cmbBoxCN30.Items.Add("20");
            cmbBoxCN30.Items.Add("21");
            cmbBoxCN30.Items.Add("22");
            cmbBoxCN30.Items.Add("23");
            cmbBoxCN30.Items.Add("24");
            cmbBoxCN30.Items.Add("25");
            cmbBoxCN30.Items.Add("26");
            cmbBoxCN30.Items.Add("27");
            cmbBoxCN30.Items.Add("28");
            cmbBoxCN30.Items.Add("29");
            cmbBoxCN30.Items.Add("30");
        }

        private void InsertBreak(object sender, EventArgs e)
        {
            try
            {
                prvw[Int32.Parse(lastValue1)] = "";
            }
            catch
            {

            }
            try
            {
                prvw[Int32.Parse(lastValue2)] = "";
            }
            catch
            {

            }
            try
            {
                prvw[Int32.Parse(lastValue3)] = "";
            }
            catch
            {

            }
            try
            {
                prvw[Int32.Parse(lastValue4)] = "";
            }
            catch
            {

            }
            try
            {
                prvw[Int32.Parse(lastValue5)] = "";
            }
            catch
            {

            }
            try
            {
                AddToComboBox();
            }
            catch
            {

            }

            try
            {
                RemoveFromComboBox();
            }
            catch
            {

            }
            try
            {
                UpdatePreview();
            }
            catch
            {

            }

            lastValue1 = cmbBoxBreaks1.Text;
            lastValue2 = cmbBoxBreaks2.Text;
            lastValue3 = cmbBoxBreaks3.Text;
            lastValue4 = cmbBoxBreaks4.Text;
            lastValue5 = cmbBoxBreaks5.Text;
        }

        private void UpdatePreview()
        {
            if (chckBxBreaks1.Checked && !cmbBoxBreaks1.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks1.Text)] = txtBoxBreaks1.Text;
            }
            else if (!chckBxBreaks1.Checked && !cmbBoxBreaks1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks1.Text)] = "";
            }

            if (chckBxBreaks2.Checked && !cmbBoxBreaks2.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks2.Text)] = txtBoxBreaks2.Text;
            }
            else if (!chckBxBreaks2.Checked && !cmbBoxBreaks2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks2.Text)] = "";
            }

            if (chckBxBreaks3.Checked && !cmbBoxBreaks3.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks3.Text)] = txtBoxBreaks3.Text;
            }
            else if (!chckBxBreaks3.Checked && !cmbBoxBreaks3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks3.Text)] = "";
            }

            if (chckBxBreaks4.Checked && !cmbBoxBreaks4.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks4.Text)] = txtBoxBreaks4.Text;
            }
            else if (!chckBxBreaks4.Checked && !cmbBoxBreaks4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks4.Text)] = "";
            }

            if (chckBxBreaks5.Checked && !cmbBoxBreaks5.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks5.Text)] = txtBoxBreaks5.Text;
            }
            else if (!chckBxBreaks5.Checked && !cmbBoxBreaks5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxBreaks5.Text)] = "";
            }

            preview = "";

            for (int i = 0; i < prvw.Length; i++)
            {
                preview = preview + prvw[i];

                if (prvw[i] == "")
                {
                    preview = preview + " ";
                }
            }

            txtBoxPreview.Text = preview;
        }

        private void AddToComboBox()
        {
            cmbBoxCN1.Items.Add(lastValue1);
            cmbBoxCN2.Items.Add(lastValue1);
            cmbBoxCN3.Items.Add(lastValue1);
            cmbBoxCN4.Items.Add(lastValue1);
            cmbBoxCN5.Items.Add(lastValue1);
            cmbBoxCN6.Items.Add(lastValue1);
            cmbBoxCN7.Items.Add(lastValue1);
            cmbBoxCN8.Items.Add(lastValue1);
            cmbBoxCN9.Items.Add(lastValue1);
            cmbBoxCN10.Items.Add(lastValue1);
            cmbBoxCN11.Items.Add(lastValue1);
            cmbBoxCN12.Items.Add(lastValue1);
            cmbBoxCN13.Items.Add(lastValue1);
            cmbBoxCN14.Items.Add(lastValue1);
            cmbBoxCN15.Items.Add(lastValue1);
            cmbBoxCN16.Items.Add(lastValue1);
            cmbBoxCN17.Items.Add(lastValue1);
            cmbBoxCN18.Items.Add(lastValue1);
            cmbBoxCN19.Items.Add(lastValue1);
            cmbBoxCN20.Items.Add(lastValue1);
            cmbBoxCN21.Items.Add(lastValue1);
            cmbBoxCN22.Items.Add(lastValue1);
            cmbBoxCN23.Items.Add(lastValue1);
            cmbBoxCN24.Items.Add(lastValue1);
            cmbBoxCN25.Items.Add(lastValue1);
            cmbBoxCN26.Items.Add(lastValue1);
            cmbBoxCN27.Items.Add(lastValue1);
            cmbBoxCN28.Items.Add(lastValue1);
            cmbBoxCN29.Items.Add(lastValue1);
            cmbBoxCN30.Items.Add(lastValue1);

            cmbBoxCN1.Items.Add(lastValue2);
            cmbBoxCN2.Items.Add(lastValue2);
            cmbBoxCN3.Items.Add(lastValue2);
            cmbBoxCN4.Items.Add(lastValue2);
            cmbBoxCN5.Items.Add(lastValue2);
            cmbBoxCN6.Items.Add(lastValue2);
            cmbBoxCN7.Items.Add(lastValue2);
            cmbBoxCN8.Items.Add(lastValue2);
            cmbBoxCN9.Items.Add(lastValue2);
            cmbBoxCN10.Items.Add(lastValue2);
            cmbBoxCN11.Items.Add(lastValue2);
            cmbBoxCN12.Items.Add(lastValue2);
            cmbBoxCN13.Items.Add(lastValue2);
            cmbBoxCN14.Items.Add(lastValue2);
            cmbBoxCN15.Items.Add(lastValue2);
            cmbBoxCN16.Items.Add(lastValue2);
            cmbBoxCN17.Items.Add(lastValue2);
            cmbBoxCN18.Items.Add(lastValue2);
            cmbBoxCN19.Items.Add(lastValue2);
            cmbBoxCN20.Items.Add(lastValue2);
            cmbBoxCN21.Items.Add(lastValue2);
            cmbBoxCN22.Items.Add(lastValue2);
            cmbBoxCN23.Items.Add(lastValue2);
            cmbBoxCN24.Items.Add(lastValue2);
            cmbBoxCN25.Items.Add(lastValue2);
            cmbBoxCN26.Items.Add(lastValue2);
            cmbBoxCN27.Items.Add(lastValue2);
            cmbBoxCN28.Items.Add(lastValue2);
            cmbBoxCN29.Items.Add(lastValue2);
            cmbBoxCN30.Items.Add(lastValue2);

            cmbBoxCN1.Items.Add(lastValue3);
            cmbBoxCN2.Items.Add(lastValue3);
            cmbBoxCN3.Items.Add(lastValue3);
            cmbBoxCN4.Items.Add(lastValue3);
            cmbBoxCN5.Items.Add(lastValue3);
            cmbBoxCN6.Items.Add(lastValue3);
            cmbBoxCN7.Items.Add(lastValue3);
            cmbBoxCN8.Items.Add(lastValue3);
            cmbBoxCN9.Items.Add(lastValue3);
            cmbBoxCN10.Items.Add(lastValue3);
            cmbBoxCN11.Items.Add(lastValue3);
            cmbBoxCN12.Items.Add(lastValue3);
            cmbBoxCN13.Items.Add(lastValue3);
            cmbBoxCN14.Items.Add(lastValue3);
            cmbBoxCN15.Items.Add(lastValue3);
            cmbBoxCN16.Items.Add(lastValue3);
            cmbBoxCN17.Items.Add(lastValue3);
            cmbBoxCN18.Items.Add(lastValue3);
            cmbBoxCN19.Items.Add(lastValue3);
            cmbBoxCN20.Items.Add(lastValue3);
            cmbBoxCN21.Items.Add(lastValue3);
            cmbBoxCN22.Items.Add(lastValue3);
            cmbBoxCN23.Items.Add(lastValue3);
            cmbBoxCN24.Items.Add(lastValue3);
            cmbBoxCN25.Items.Add(lastValue3);
            cmbBoxCN26.Items.Add(lastValue3);
            cmbBoxCN27.Items.Add(lastValue3);
            cmbBoxCN28.Items.Add(lastValue3);
            cmbBoxCN29.Items.Add(lastValue3);
            cmbBoxCN30.Items.Add(lastValue3);

            cmbBoxCN1.Items.Add(lastValue4);
            cmbBoxCN2.Items.Add(lastValue4);
            cmbBoxCN3.Items.Add(lastValue4);
            cmbBoxCN4.Items.Add(lastValue4);
            cmbBoxCN5.Items.Add(lastValue4);
            cmbBoxCN6.Items.Add(lastValue4);
            cmbBoxCN7.Items.Add(lastValue4);
            cmbBoxCN8.Items.Add(lastValue4);
            cmbBoxCN9.Items.Add(lastValue4);
            cmbBoxCN10.Items.Add(lastValue4);
            cmbBoxCN11.Items.Add(lastValue4);
            cmbBoxCN12.Items.Add(lastValue4);
            cmbBoxCN13.Items.Add(lastValue4);
            cmbBoxCN14.Items.Add(lastValue4);
            cmbBoxCN15.Items.Add(lastValue4);
            cmbBoxCN16.Items.Add(lastValue4);
            cmbBoxCN17.Items.Add(lastValue4);
            cmbBoxCN18.Items.Add(lastValue4);
            cmbBoxCN19.Items.Add(lastValue4);
            cmbBoxCN20.Items.Add(lastValue4);
            cmbBoxCN21.Items.Add(lastValue4);
            cmbBoxCN22.Items.Add(lastValue4);
            cmbBoxCN23.Items.Add(lastValue4);
            cmbBoxCN24.Items.Add(lastValue4);
            cmbBoxCN25.Items.Add(lastValue4);
            cmbBoxCN26.Items.Add(lastValue4);
            cmbBoxCN27.Items.Add(lastValue4);
            cmbBoxCN28.Items.Add(lastValue4);
            cmbBoxCN29.Items.Add(lastValue4);
            cmbBoxCN30.Items.Add(lastValue4);

            cmbBoxCN1.Items.Add(lastValue5);
            cmbBoxCN2.Items.Add(lastValue5);
            cmbBoxCN3.Items.Add(lastValue5);
            cmbBoxCN4.Items.Add(lastValue5);
            cmbBoxCN5.Items.Add(lastValue5);
            cmbBoxCN6.Items.Add(lastValue5);
            cmbBoxCN7.Items.Add(lastValue5);
            cmbBoxCN8.Items.Add(lastValue5);
            cmbBoxCN9.Items.Add(lastValue5);
            cmbBoxCN10.Items.Add(lastValue5);
            cmbBoxCN11.Items.Add(lastValue5);
            cmbBoxCN12.Items.Add(lastValue5);
            cmbBoxCN13.Items.Add(lastValue5);
            cmbBoxCN14.Items.Add(lastValue5);
            cmbBoxCN15.Items.Add(lastValue5);
            cmbBoxCN16.Items.Add(lastValue5);
            cmbBoxCN17.Items.Add(lastValue5);
            cmbBoxCN18.Items.Add(lastValue5);
            cmbBoxCN19.Items.Add(lastValue5);
            cmbBoxCN20.Items.Add(lastValue5);
            cmbBoxCN21.Items.Add(lastValue5);
            cmbBoxCN22.Items.Add(lastValue5);
            cmbBoxCN23.Items.Add(lastValue5);
            cmbBoxCN24.Items.Add(lastValue5);
            cmbBoxCN25.Items.Add(lastValue5);
            cmbBoxCN26.Items.Add(lastValue5);
            cmbBoxCN27.Items.Add(lastValue5);
            cmbBoxCN28.Items.Add(lastValue5);
            cmbBoxCN29.Items.Add(lastValue5);
            cmbBoxCN30.Items.Add(lastValue5);
        }

        private void RemoveFromComboBox()
        {
            cmbBoxCN1.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN2.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN3.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN4.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN5.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN6.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN7.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN8.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN9.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN10.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN11.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN12.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN13.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN14.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN15.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN16.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN17.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN18.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN19.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN20.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN21.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN22.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN23.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN24.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN25.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN26.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN27.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN28.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN29.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxCN30.Items.Remove(cmbBoxBreaks1.Text);

            cmbBoxCN1.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN2.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN3.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN4.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN5.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN6.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN7.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN8.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN9.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN10.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN11.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN12.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN13.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN14.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN15.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN16.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN17.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN18.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN19.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN20.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN21.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN22.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN23.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN24.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN25.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN26.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN27.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN28.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN29.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxCN30.Items.Remove(cmbBoxBreaks2.Text);

            cmbBoxCN1.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN2.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN3.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN4.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN5.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN6.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN7.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN8.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN9.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN10.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN11.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN12.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN13.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN14.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN15.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN16.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN17.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN18.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN19.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN20.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN21.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN22.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN23.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN24.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN25.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN26.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN27.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN28.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN29.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxCN30.Items.Remove(cmbBoxBreaks3.Text);

            cmbBoxCN1.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN2.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN3.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN4.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN5.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN6.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN7.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN8.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN9.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN10.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN11.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN12.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN13.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN14.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN15.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN16.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN17.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN18.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN19.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN20.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN21.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN22.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN23.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN24.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN25.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN26.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN27.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN28.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN29.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxCN30.Items.Remove(cmbBoxBreaks4.Text);

            cmbBoxCN1.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN2.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN3.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN4.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN5.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN6.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN7.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN8.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN9.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN10.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN11.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN12.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN13.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN14.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN15.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN16.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN17.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN18.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN19.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN20.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN21.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN22.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN23.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN24.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN25.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN26.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN27.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN28.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN29.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxCN30.Items.Remove(cmbBoxBreaks5.Text);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
