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

namespace Wire_Number_Dialogs
{
    public partial class Wire_Number_Form : Form
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

        public Wire_Number_Form()
        {
            InitializeComponent();
        }

        private void Wire_Number_Form_Load(object sender, EventArgs e)
        {
            ComboBoxSetup();

            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;

            txtboxDWGName.Text = Path.GetFileName(doc.Name);

            try
            {
                lblWN1.Text = txtboxDWGName.Text[0].ToString();
                lblWN2.Text = txtboxDWGName.Text[1].ToString();
                lblWN3.Text = txtboxDWGName.Text[2].ToString();
                lblWN4.Text = txtboxDWGName.Text[3].ToString();
                lblWN5.Text = txtboxDWGName.Text[4].ToString();
                lblWN6.Text = txtboxDWGName.Text[5].ToString();
                lblWN7.Text = txtboxDWGName.Text[6].ToString();
                lblWN8.Text = txtboxDWGName.Text[7].ToString();
                lblWN9.Text = txtboxDWGName.Text[8].ToString();
                lblWN10.Text = txtboxDWGName.Text[9].ToString();
                lblWN11.Text = txtboxDWGName.Text[10].ToString();
                lblWN12.Text = txtboxDWGName.Text[11].ToString();
                lblWN13.Text = txtboxDWGName.Text[12].ToString();
                lblWN14.Text = txtboxDWGName.Text[13].ToString();
                lblWN15.Text = txtboxDWGName.Text[14].ToString();
                lblWN16.Text = txtboxDWGName.Text[15].ToString();
                lblWN17.Text = txtboxDWGName.Text[16].ToString();
                lblWN18.Text = txtboxDWGName.Text[17].ToString();
                lblWN19.Text = txtboxDWGName.Text[18].ToString();
                lblWN20.Text = txtboxDWGName.Text[19].ToString();
                lblWN21.Text = txtboxDWGName.Text[20].ToString();
                lblWN22.Text = txtboxDWGName.Text[21].ToString();
                lblWN23.Text = txtboxDWGName.Text[22].ToString();
                lblWN24.Text = txtboxDWGName.Text[23].ToString();
                lblWN25.Text = txtboxDWGName.Text[24].ToString();
                lblWN26.Text = txtboxDWGName.Text[25].ToString();
                lblWN27.Text = txtboxDWGName.Text[26].ToString();
                lblWN28.Text = txtboxDWGName.Text[27].ToString();
                lblWN29.Text = txtboxDWGName.Text[28].ToString();
                lblWN30.Text = txtboxDWGName.Text[29].ToString();

                cmbBoxWNL.Items.Add("0");
                cmbBoxWNL.Items.Add("1");
                cmbBoxWNL.Items.Add("2");
                cmbBoxWNL.Items.Add("3");
                cmbBoxWNL.Items.Add("4");
                cmbBoxWNL.Items.Add("5");
                cmbBoxWNL.Items.Add("6");
                cmbBoxWNL.Items.Add("7");
                cmbBoxWNL.Items.Add("8");

                cmbBoxWNSL.Items.Add("0");
                cmbBoxWNSL.Items.Add("1");
                cmbBoxWNSL.Items.Add("2");
                cmbBoxWNSL.Items.Add("3");
                cmbBoxWNSL.Items.Add("4");
                cmbBoxWNSL.Items.Add("5");
                cmbBoxWNSL.Items.Add("6");
                cmbBoxWNSL.Items.Add("7");
                cmbBoxWNSL.Items.Add("8");
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
                if (chckBxWN1.Checked)
                {
                    if (cmbBoxWN1.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "1,1)";
                    }
                }
                if (chckBxWN2.Checked)
                {
                    if (cmbBoxWN2.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "2,1)";
                    }
                }
                if (chckBxWN3.Checked)
                {
                    if (cmbBoxWN3.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "3,1)";
                    }
                }
                if (chckBxWN4.Checked)
                {
                    if (cmbBoxWN4.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "4,1)";
                    }
                }
                if (chckBxWN5.Checked)
                {
                    if (cmbBoxWN5.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "5,1)";
                    }
                }
                if (chckBxWN6.Checked)
                {
                    if (cmbBoxWN6.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "6,1)";
                    }
                }
                if (chckBxWN7.Checked)
                {
                    if (cmbBoxWN7.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "7,1)";
                    }
                }
                if (chckBxWN8.Checked)
                {
                    if (cmbBoxWN8.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "8,1)";
                    }
                }
                if (chckBxWN9.Checked)
                {
                    if (cmbBoxWN9.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "9,1)";
                    }
                }
                if (chckBxWN10.Checked)
                {
                    if (cmbBoxWN10.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "10,1)";
                    }
                }
                if (chckBxWN11.Checked)
                {
                    if (cmbBoxWN11.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "11,1)";
                    }
                }
                if (chckBxWN12.Checked)
                {
                    if (cmbBoxWN12.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "12,1)";
                    }
                }
                if (chckBxWN13.Checked)
                {
                    if (cmbBoxWN13.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "13,1)";
                    }
                }
                if (chckBxWN14.Checked)
                {
                    if (cmbBoxWN14.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "14,1)";
                    }
                }
                if (chckBxWN15.Checked)
                {
                    if (cmbBoxWN15.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "15,1)";
                    }
                }
                if (chckBxWN16.Checked)
                {
                    if (cmbBoxWN16.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "16,1)";
                    }
                }
                if (chckBxWN17.Checked)
                {
                    if (cmbBoxWN17.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "17,1)";
                    }
                }
                if (chckBxWN18.Checked)
                {
                    if (cmbBoxWN18.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "18,1)";
                    }
                }
                if (chckBxWN19.Checked)
                {
                    if (cmbBoxWN19.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "19,1)";
                    }
                }
                if (chckBxWN20.Checked)
                {
                    if (cmbBoxWN20.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "20,1)";
                    }
                }
                if (chckBxWN21.Checked)
                {
                    if (cmbBoxWN21.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "21,1)";
                    }
                }
                if (chckBxWN22.Checked)
                {
                    if (cmbBoxWN22.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "22,1)";
                    }
                }
                if (chckBxWN23.Checked)
                {
                    if (cmbBoxWN23.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "23,1)";
                    }
                }
                if (chckBxWN24.Checked)
                {
                    if (cmbBoxWN24.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "24,1)";
                    }
                }
                if (chckBxWN25.Checked)
                {
                    if (cmbBoxWN25.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "25,1)";
                    }
                }
                if (chckBxWN26.Checked)
                {
                    if (cmbBoxWN26.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "26,1)";
                    }
                }
                if (chckBxWN27.Checked)
                {
                    if (cmbBoxWN27.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "27,1)";
                    }
                }
                if (chckBxWN28.Checked)
                {
                    if (cmbBoxWN28.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "28,1)";
                    }
                }
                if (chckBxWN29.Checked)
                {
                    if (cmbBoxWN29.Text.Equals(i.ToString()))
                    {
                        DieselEXP = DieselEXP + @"$(substr,$(getvar,dwgname)," + "29,1)";
                    }
                }
                if (chckBxWN30.Checked)
                {
                    if (cmbBoxWN30.Text.Equals(i.ToString()))
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

                    if(SPrompt.Status == PromptStatus.OK)
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

                                        if (bd.Name == "WD_WNH" || bd.Name == "WD_WNV")
                                        {
                                            foreach (ObjectId attId in blkRef.AttributeCollection)
                                            {
                                                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForWrite);

                                                if (attRef.Tag == "WIRENO")
                                                {
                                                    string currentText = attRef.TextString;
                                                    string wirePF;
                                                    string wireSF;
                                                    try
                                                    {
                                                        wireSF = currentText.Substring(currentText.Length - Int32.Parse(cmbBoxWNL.Text), Int32.Parse(cmbBoxWNL.Text));
                                                    }
                                                    catch
                                                    {
                                                        wireSF = "";
                                                    }

                                                    try
                                                    {
                                                        wirePF = currentText.Substring(0, Int32.Parse(cmbBoxWNSL.Text));
                                                    }
                                                    catch
                                                    {
                                                        wirePF = "";
                                                    }

                                                    attRef.TextString = "";
                                                    attRef.TextString = wirePF + DieselEXP + wireSF; 
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
            doc.SendStringToExecute("Regenall ", true, false, false);
        }        

        private void UpdatePrvw(object sender, EventArgs e)
        {
            for (int i = 0; i < prvw.Length; i++)
            {
                prvw[i] = "";
            }

            preview = "";
             
            if(chckBxWN1.Checked && !cmbBoxWN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN1.Text)] = lblWN1.Text;                 
            }
            else if (!chckBxWN1.Checked && !cmbBoxWN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN1.Text)] = "";
            }

            if(chckBxWN2.Checked && !cmbBoxWN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN2.Text)] = lblWN2.Text;
            }
            else if (!chckBxWN2.Checked && !cmbBoxWN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN2.Text)] = "";
            }

            if (chckBxWN3.Checked && !cmbBoxWN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN3.Text)] = lblWN3.Text;
            }
            else if (!chckBxWN3.Checked && !cmbBoxWN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN3.Text)] = "";
            }

            if (chckBxWN4.Checked && !cmbBoxWN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN4.Text)] = lblWN4.Text;
            }
            else if (!chckBxWN4.Checked && !cmbBoxWN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN4.Text)] = "";
            }

            if (chckBxWN5.Checked && !cmbBoxWN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN5.Text)] = lblWN5.Text;
            }
            else if (!chckBxWN5.Checked && !cmbBoxWN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN5.Text)] = "";
            }

             if (chckBxWN6.Checked && !cmbBoxWN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN6.Text)] = lblWN6.Text;
            }
             else if (!chckBxWN6.Checked && !cmbBoxWN6.Text.Equals(""))
             {
                 prvw[Int32.Parse(cmbBoxWN6.Text)] = "";
             }

             if (chckBxWN7.Checked && !cmbBoxWN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN7.Text)] = lblWN7.Text;
            }
             else if (!chckBxWN7.Checked && !cmbBoxWN7.Text.Equals(""))
             {
                 prvw[Int32.Parse(cmbBoxWN7.Text)] = "";
             }

            if (chckBxWN8.Checked && !cmbBoxWN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN8.Text)] = lblWN8.Text;
            }
            else if (!chckBxWN8.Checked && !cmbBoxWN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN8.Text)] = "";
            }

            if (chckBxWN9.Checked && !cmbBoxWN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN9.Text)] = lblWN9.Text;
            }
            else if (!chckBxWN9.Checked && !cmbBoxWN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN9.Text)] = "";
            }

            if (chckBxWN10.Checked && !cmbBoxWN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN10.Text)] = lblWN10.Text;
            }
            else if (!chckBxWN10.Checked && !cmbBoxWN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN10.Text)] = "";
            }

            if (chckBxWN11.Checked && !cmbBoxWN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN11.Text)] = lblWN11.Text;
            }
            else if (!chckBxWN11.Checked && !cmbBoxWN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN11.Text)] = "";
            }

            if (chckBxWN12.Checked && !cmbBoxWN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN12.Text)] = lblWN12.Text;
            }
            else if (!chckBxWN12.Checked && !cmbBoxWN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN12.Text)] = "";
            }

            if (chckBxWN13.Checked && !cmbBoxWN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN13.Text)] = lblWN13.Text;
            }
            else if (!chckBxWN13.Checked && !cmbBoxWN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN13.Text)] = "";
            }

            if (chckBxWN14.Checked && !cmbBoxWN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN14.Text)] = lblWN14.Text;
            }
            else if (!chckBxWN14.Checked && !cmbBoxWN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN14.Text)] = "";
            }

            if (chckBxWN15.Checked && !cmbBoxWN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN15.Text)] = lblWN15.Text;
            }
            else if (!chckBxWN15.Checked && !cmbBoxWN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN15.Text)] = "";
            }

            if (chckBxWN16.Checked && !cmbBoxWN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN16.Text)] = lblWN16.Text;
            }
            else if (!chckBxWN16.Checked && !cmbBoxWN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN16.Text)] = "";
            }

            if (chckBxWN17.Checked && !cmbBoxWN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN17.Text)] = lblWN17.Text;
            }
            else if (!chckBxWN17.Checked && !cmbBoxWN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN17.Text)] = "";
            }

            if (chckBxWN18.Checked && !cmbBoxWN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN18.Text)] = lblWN18.Text;
            }
            else if (!chckBxWN18.Checked && !cmbBoxWN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN18.Text)] = "";
            }

            if (chckBxWN19.Checked && !cmbBoxWN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN19.Text)] = lblWN19.Text;
            }
            else if (!chckBxWN19.Checked && !cmbBoxWN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN19.Text)] = "";
            }

            if (chckBxWN20.Checked && !cmbBoxWN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN20.Text)] = lblWN20.Text;
            }
            else if (!chckBxWN20.Checked && !cmbBoxWN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN20.Text)] = "";
            }

            if (chckBxWN21.Checked && !cmbBoxWN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN21.Text)] = lblWN21.Text;
            }
            else if (!chckBxWN21.Checked && !cmbBoxWN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN21.Text)] = "";
            }

            if (chckBxWN22.Checked && !cmbBoxWN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN22.Text)] = lblWN22.Text;
            }
            else if (!chckBxWN22.Checked && !cmbBoxWN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN22.Text)] = "";
            }

            if (chckBxWN23.Checked && !cmbBoxWN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN23.Text)] = lblWN23.Text;
            }
            else if (!chckBxWN23.Checked && !cmbBoxWN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN23.Text)] = "";
            }

            if (chckBxWN24.Checked && !cmbBoxWN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN24.Text)] = lblWN24.Text;
            }
            else if (!chckBxWN24.Checked && !cmbBoxWN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN24.Text)] = "";
            }

            if (chckBxWN25.Checked && !cmbBoxWN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN25.Text)] = lblWN25.Text;
            }
            else if (!chckBxWN25.Checked && !cmbBoxWN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN25.Text)] = "";
            }

            if (chckBxWN26.Checked && !cmbBoxWN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN26.Text)] = lblWN26.Text;
            }
            else if (!chckBxWN26.Checked && !cmbBoxWN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN26.Text)] = "";
            }

            if (chckBxWN27.Checked && !cmbBoxWN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN27.Text)] = lblWN27.Text;
            }
            else if (!chckBxWN27.Checked && !cmbBoxWN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN27.Text)] = "";
            }

            if (chckBxWN28.Checked && !cmbBoxWN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN28.Text)] = lblWN28.Text;
            }
            else if (!chckBxWN28.Checked && !cmbBoxWN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN28.Text)] = "";
            }

            if (chckBxWN29.Checked && !cmbBoxWN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN29.Text)] = lblWN29.Text;
            }
            else if (!chckBxWN29.Checked && !cmbBoxWN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN29.Text)] = "";
            }

            if (chckBxWN30.Checked && !cmbBoxWN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN30.Text)] = lblWN30.Text;
            }
            else if (!chckBxWN30.Checked && !cmbBoxWN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN1.Text)] = "";
            }

            for(int i = 0; i < prvw.Length; i++)
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

            if (chckBxWN1.Checked && !cmbBoxWN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN1.Text)] = lblWN1.Text;
            }
            else if (!chckBxWN1.Checked && !cmbBoxWN1.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN1.Text)] = "";
            }

            if (chckBxWN2.Checked && !cmbBoxWN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN2.Text)] = lblWN2.Text;
            }
            else if (!chckBxWN2.Checked && !cmbBoxWN2.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN2.Text)] = "";
            }

            if (chckBxWN3.Checked && !cmbBoxWN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN3.Text)] = lblWN3.Text;
            }
            else if (!chckBxWN3.Checked && !cmbBoxWN3.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN3.Text)] = "";
            }

            if (chckBxWN4.Checked && !cmbBoxWN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN4.Text)] = lblWN4.Text;
            }
            else if (!chckBxWN4.Checked && !cmbBoxWN4.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN4.Text)] = "";
            }

            if (chckBxWN5.Checked && !cmbBoxWN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN5.Text)] = lblWN5.Text;
            }
            else if (!chckBxWN5.Checked && !cmbBoxWN5.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN5.Text)] = "";
            }

            if (chckBxWN6.Checked && !cmbBoxWN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN6.Text)] = lblWN6.Text;
            }
            else if (!chckBxWN6.Checked && !cmbBoxWN6.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN6.Text)] = "";
            }

            if (chckBxWN7.Checked && !cmbBoxWN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN7.Text)] = lblWN7.Text;
            }
            else if (!chckBxWN7.Checked && !cmbBoxWN7.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN7.Text)] = "";
            }

            if (chckBxWN8.Checked && !cmbBoxWN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN8.Text)] = lblWN8.Text;
            }
            else if (!chckBxWN8.Checked && !cmbBoxWN8.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN8.Text)] = "";
            }

            if (chckBxWN9.Checked && !cmbBoxWN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN9.Text)] = lblWN9.Text;
            }
            else if (!chckBxWN9.Checked && !cmbBoxWN9.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN9.Text)] = "";
            }

            if (chckBxWN10.Checked && !cmbBoxWN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN10.Text)] = lblWN10.Text;
            }
            else if (!chckBxWN10.Checked && !cmbBoxWN10.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN10.Text)] = "";
            }

            if (chckBxWN11.Checked && !cmbBoxWN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN11.Text)] = lblWN11.Text;
            }
            else if (!chckBxWN11.Checked && !cmbBoxWN11.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN11.Text)] = "";
            }

            if (chckBxWN12.Checked && !cmbBoxWN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN12.Text)] = lblWN12.Text;
            }
            else if (!chckBxWN12.Checked && !cmbBoxWN12.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN12.Text)] = "";
            }

            if (chckBxWN13.Checked && !cmbBoxWN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN13.Text)] = lblWN13.Text;
            }
            else if (!chckBxWN13.Checked && !cmbBoxWN13.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN13.Text)] = "";
            }

            if (chckBxWN14.Checked && !cmbBoxWN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN14.Text)] = lblWN14.Text;
            }
            else if (!chckBxWN14.Checked && !cmbBoxWN14.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN14.Text)] = "";
            }

            if (chckBxWN15.Checked && !cmbBoxWN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN15.Text)] = lblWN15.Text;
            }
            else if (!chckBxWN15.Checked && !cmbBoxWN15.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN15.Text)] = "";
            }

            if (chckBxWN16.Checked && !cmbBoxWN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN16.Text)] = lblWN16.Text;
            }
            else if (!chckBxWN16.Checked && !cmbBoxWN16.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN16.Text)] = "";
            }

            if (chckBxWN17.Checked && !cmbBoxWN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN17.Text)] = lblWN17.Text;
            }
            else if (!chckBxWN17.Checked && !cmbBoxWN17.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN17.Text)] = "";
            }

            if (chckBxWN18.Checked && !cmbBoxWN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN18.Text)] = lblWN18.Text;
            }
            else if (!chckBxWN18.Checked && !cmbBoxWN18.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN18.Text)] = "";
            }

            if (chckBxWN19.Checked && !cmbBoxWN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN19.Text)] = lblWN19.Text;
            }
            else if (!chckBxWN19.Checked && !cmbBoxWN19.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN19.Text)] = "";
            }

            if (chckBxWN20.Checked && !cmbBoxWN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN20.Text)] = lblWN20.Text;
            }
            else if (!chckBxWN20.Checked && !cmbBoxWN20.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN20.Text)] = "";
            }

            if (chckBxWN21.Checked && !cmbBoxWN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN21.Text)] = lblWN21.Text;
            }
            else if (!chckBxWN21.Checked && !cmbBoxWN21.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN21.Text)] = "";
            }

            if (chckBxWN22.Checked && !cmbBoxWN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN22.Text)] = lblWN22.Text;
            }
            else if (!chckBxWN22.Checked && !cmbBoxWN22.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN22.Text)] = "";
            }

            if (chckBxWN23.Checked && !cmbBoxWN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN23.Text)] = lblWN23.Text;
            }
            else if (!chckBxWN23.Checked && !cmbBoxWN23.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN23.Text)] = "";
            }

            if (chckBxWN24.Checked && !cmbBoxWN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN24.Text)] = lblWN24.Text;
            }
            else if (!chckBxWN24.Checked && !cmbBoxWN24.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN24.Text)] = "";
            }

            if (chckBxWN25.Checked && !cmbBoxWN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN25.Text)] = lblWN25.Text;
            }
            else if (!chckBxWN25.Checked && !cmbBoxWN25.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN25.Text)] = "";
            }

            if (chckBxWN26.Checked && !cmbBoxWN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN26.Text)] = lblWN26.Text;
            }
            else if (!chckBxWN26.Checked && !cmbBoxWN26.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN26.Text)] = "";
            }

            if (chckBxWN27.Checked && !cmbBoxWN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN27.Text)] = lblWN27.Text;
            }
            else if (!chckBxWN27.Checked && !cmbBoxWN27.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN27.Text)] = "";
            }

            if (chckBxWN28.Checked && !cmbBoxWN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN28.Text)] = lblWN28.Text;
            }
            else if (!chckBxWN28.Checked && !cmbBoxWN28.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN28.Text)] = "";
            }

            if (chckBxWN29.Checked && !cmbBoxWN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN29.Text)] = lblWN29.Text;
            }
            else if (!chckBxWN29.Checked && !cmbBoxWN29.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN29.Text)] = "";
            }

            if (chckBxWN30.Checked && !cmbBoxWN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN30.Text)] = lblWN30.Text;
            }
            else if (!chckBxWN30.Checked && !cmbBoxWN30.Text.Equals(""))
            {
                prvw[Int32.Parse(cmbBoxWN1.Text)] = "";
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

            cmbBoxWN1.Items.Add("1");
            cmbBoxWN1.Items.Add("2");
            cmbBoxWN1.Items.Add("3");
            cmbBoxWN1.Items.Add("4");
            cmbBoxWN1.Items.Add("5");
            cmbBoxWN1.Items.Add("6");
            cmbBoxWN1.Items.Add("7");
            cmbBoxWN1.Items.Add("8");
            cmbBoxWN1.Items.Add("9");
            cmbBoxWN1.Items.Add("10");
            cmbBoxWN1.Items.Add("11");
            cmbBoxWN1.Items.Add("12");
            cmbBoxWN1.Items.Add("13");
            cmbBoxWN1.Items.Add("14");
            cmbBoxWN1.Items.Add("15");
            cmbBoxWN1.Items.Add("16");
            cmbBoxWN1.Items.Add("17");
            cmbBoxWN1.Items.Add("18");
            cmbBoxWN1.Items.Add("19");
            cmbBoxWN1.Items.Add("20");
            cmbBoxWN1.Items.Add("21");
            cmbBoxWN1.Items.Add("22");
            cmbBoxWN1.Items.Add("23");
            cmbBoxWN1.Items.Add("24");
            cmbBoxWN1.Items.Add("25");
            cmbBoxWN1.Items.Add("26");
            cmbBoxWN1.Items.Add("27");
            cmbBoxWN1.Items.Add("28");
            cmbBoxWN1.Items.Add("29");
            cmbBoxWN1.Items.Add("30");

            cmbBoxWN2.Items.Add("1");
            cmbBoxWN2.Items.Add("2");
            cmbBoxWN2.Items.Add("3");
            cmbBoxWN2.Items.Add("4");
            cmbBoxWN2.Items.Add("5");
            cmbBoxWN2.Items.Add("6");
            cmbBoxWN2.Items.Add("7");
            cmbBoxWN2.Items.Add("8");
            cmbBoxWN2.Items.Add("9");
            cmbBoxWN2.Items.Add("10");
            cmbBoxWN2.Items.Add("11");
            cmbBoxWN2.Items.Add("12");
            cmbBoxWN2.Items.Add("13");
            cmbBoxWN2.Items.Add("14");
            cmbBoxWN2.Items.Add("15");
            cmbBoxWN2.Items.Add("16");
            cmbBoxWN2.Items.Add("17");
            cmbBoxWN2.Items.Add("18");
            cmbBoxWN2.Items.Add("19");
            cmbBoxWN2.Items.Add("20");
            cmbBoxWN2.Items.Add("21");
            cmbBoxWN2.Items.Add("22");
            cmbBoxWN2.Items.Add("23");
            cmbBoxWN2.Items.Add("24");
            cmbBoxWN2.Items.Add("25");
            cmbBoxWN2.Items.Add("26");
            cmbBoxWN2.Items.Add("27");
            cmbBoxWN2.Items.Add("28");
            cmbBoxWN2.Items.Add("29");
            cmbBoxWN2.Items.Add("30");

            cmbBoxWN3.Items.Add("1");
            cmbBoxWN3.Items.Add("2");
            cmbBoxWN3.Items.Add("3");
            cmbBoxWN3.Items.Add("4");
            cmbBoxWN3.Items.Add("5");
            cmbBoxWN3.Items.Add("6");
            cmbBoxWN3.Items.Add("7");
            cmbBoxWN3.Items.Add("8");
            cmbBoxWN3.Items.Add("9");
            cmbBoxWN3.Items.Add("10");
            cmbBoxWN3.Items.Add("11");
            cmbBoxWN3.Items.Add("12");
            cmbBoxWN3.Items.Add("13");
            cmbBoxWN3.Items.Add("14");
            cmbBoxWN3.Items.Add("15");
            cmbBoxWN3.Items.Add("16");
            cmbBoxWN3.Items.Add("17");
            cmbBoxWN3.Items.Add("18");
            cmbBoxWN3.Items.Add("19");
            cmbBoxWN3.Items.Add("20");
            cmbBoxWN3.Items.Add("21");
            cmbBoxWN3.Items.Add("22");
            cmbBoxWN3.Items.Add("23");
            cmbBoxWN3.Items.Add("24");
            cmbBoxWN3.Items.Add("25");
            cmbBoxWN3.Items.Add("26");
            cmbBoxWN3.Items.Add("27");
            cmbBoxWN3.Items.Add("28");
            cmbBoxWN3.Items.Add("29");
            cmbBoxWN3.Items.Add("30");
            cmbBoxWN4.Items.Add("1");
            cmbBoxWN4.Items.Add("2");
            cmbBoxWN4.Items.Add("3");
            cmbBoxWN4.Items.Add("4");
            cmbBoxWN4.Items.Add("5");
            cmbBoxWN4.Items.Add("6");
            cmbBoxWN4.Items.Add("7");
            cmbBoxWN4.Items.Add("8");
            cmbBoxWN4.Items.Add("9");
            cmbBoxWN4.Items.Add("10");
            cmbBoxWN4.Items.Add("11");
            cmbBoxWN4.Items.Add("12");
            cmbBoxWN4.Items.Add("13");
            cmbBoxWN4.Items.Add("14");
            cmbBoxWN4.Items.Add("15");
            cmbBoxWN4.Items.Add("16");
            cmbBoxWN4.Items.Add("17");
            cmbBoxWN4.Items.Add("18");
            cmbBoxWN4.Items.Add("19");
            cmbBoxWN4.Items.Add("20");
            cmbBoxWN4.Items.Add("21");
            cmbBoxWN4.Items.Add("22");
            cmbBoxWN4.Items.Add("23");
            cmbBoxWN4.Items.Add("24");
            cmbBoxWN4.Items.Add("25");
            cmbBoxWN4.Items.Add("26");
            cmbBoxWN4.Items.Add("27");
            cmbBoxWN4.Items.Add("28");
            cmbBoxWN4.Items.Add("29");
            cmbBoxWN4.Items.Add("30");

            cmbBoxWN5.Items.Add("1");
            cmbBoxWN5.Items.Add("2");
            cmbBoxWN5.Items.Add("3");
            cmbBoxWN5.Items.Add("4");
            cmbBoxWN5.Items.Add("5");
            cmbBoxWN5.Items.Add("6");
            cmbBoxWN5.Items.Add("7");
            cmbBoxWN5.Items.Add("8");
            cmbBoxWN5.Items.Add("9");
            cmbBoxWN5.Items.Add("10");
            cmbBoxWN5.Items.Add("11");
            cmbBoxWN5.Items.Add("12");
            cmbBoxWN5.Items.Add("13");
            cmbBoxWN5.Items.Add("14");
            cmbBoxWN5.Items.Add("15");
            cmbBoxWN5.Items.Add("16");
            cmbBoxWN5.Items.Add("17");
            cmbBoxWN5.Items.Add("18");
            cmbBoxWN5.Items.Add("19");
            cmbBoxWN5.Items.Add("20");
            cmbBoxWN5.Items.Add("21");
            cmbBoxWN5.Items.Add("22");
            cmbBoxWN5.Items.Add("23");
            cmbBoxWN5.Items.Add("24");
            cmbBoxWN5.Items.Add("25");
            cmbBoxWN5.Items.Add("26");
            cmbBoxWN5.Items.Add("27");
            cmbBoxWN5.Items.Add("28");
            cmbBoxWN5.Items.Add("29");
            cmbBoxWN5.Items.Add("30");

            cmbBoxWN6.Items.Add("1");
            cmbBoxWN6.Items.Add("2");
            cmbBoxWN6.Items.Add("3");
            cmbBoxWN6.Items.Add("4");
            cmbBoxWN6.Items.Add("5");
            cmbBoxWN6.Items.Add("6");
            cmbBoxWN6.Items.Add("7");
            cmbBoxWN6.Items.Add("8");
            cmbBoxWN6.Items.Add("9");
            cmbBoxWN6.Items.Add("10");
            cmbBoxWN6.Items.Add("11");
            cmbBoxWN6.Items.Add("12");
            cmbBoxWN6.Items.Add("13");
            cmbBoxWN6.Items.Add("14");
            cmbBoxWN6.Items.Add("15");
            cmbBoxWN6.Items.Add("16");
            cmbBoxWN6.Items.Add("17");
            cmbBoxWN6.Items.Add("18");
            cmbBoxWN6.Items.Add("19");
            cmbBoxWN6.Items.Add("20");
            cmbBoxWN6.Items.Add("21");
            cmbBoxWN6.Items.Add("22");
            cmbBoxWN6.Items.Add("23");
            cmbBoxWN6.Items.Add("24");
            cmbBoxWN6.Items.Add("25");
            cmbBoxWN6.Items.Add("26");
            cmbBoxWN6.Items.Add("27");
            cmbBoxWN6.Items.Add("28");
            cmbBoxWN6.Items.Add("29");
            cmbBoxWN6.Items.Add("30");

            cmbBoxWN7.Items.Add("1");
            cmbBoxWN7.Items.Add("2");
            cmbBoxWN7.Items.Add("3");
            cmbBoxWN7.Items.Add("4");
            cmbBoxWN7.Items.Add("5");
            cmbBoxWN7.Items.Add("6");
            cmbBoxWN7.Items.Add("7");
            cmbBoxWN7.Items.Add("8");
            cmbBoxWN7.Items.Add("9");
            cmbBoxWN7.Items.Add("10");
            cmbBoxWN7.Items.Add("11");
            cmbBoxWN7.Items.Add("12");
            cmbBoxWN7.Items.Add("13");
            cmbBoxWN7.Items.Add("14");
            cmbBoxWN7.Items.Add("15");
            cmbBoxWN7.Items.Add("16");
            cmbBoxWN7.Items.Add("17");
            cmbBoxWN7.Items.Add("18");
            cmbBoxWN7.Items.Add("19");
            cmbBoxWN7.Items.Add("20");
            cmbBoxWN7.Items.Add("21");
            cmbBoxWN7.Items.Add("22");
            cmbBoxWN7.Items.Add("23");
            cmbBoxWN7.Items.Add("24");
            cmbBoxWN7.Items.Add("25");
            cmbBoxWN7.Items.Add("26");
            cmbBoxWN7.Items.Add("27");
            cmbBoxWN7.Items.Add("28");
            cmbBoxWN7.Items.Add("29");
            cmbBoxWN7.Items.Add("30");

            cmbBoxWN8.Items.Add("1");
            cmbBoxWN8.Items.Add("2");
            cmbBoxWN8.Items.Add("3");
            cmbBoxWN8.Items.Add("4");
            cmbBoxWN8.Items.Add("5");
            cmbBoxWN8.Items.Add("6");
            cmbBoxWN8.Items.Add("7");
            cmbBoxWN8.Items.Add("8");
            cmbBoxWN8.Items.Add("9");
            cmbBoxWN8.Items.Add("10");
            cmbBoxWN8.Items.Add("11");
            cmbBoxWN8.Items.Add("12");
            cmbBoxWN8.Items.Add("13");
            cmbBoxWN8.Items.Add("14");
            cmbBoxWN8.Items.Add("15");
            cmbBoxWN8.Items.Add("16");
            cmbBoxWN8.Items.Add("17");
            cmbBoxWN8.Items.Add("18");
            cmbBoxWN8.Items.Add("19");
            cmbBoxWN8.Items.Add("20");
            cmbBoxWN8.Items.Add("21");
            cmbBoxWN8.Items.Add("22");
            cmbBoxWN8.Items.Add("23");
            cmbBoxWN8.Items.Add("24");
            cmbBoxWN8.Items.Add("25");
            cmbBoxWN8.Items.Add("26");
            cmbBoxWN8.Items.Add("27");
            cmbBoxWN8.Items.Add("28");
            cmbBoxWN8.Items.Add("29");
            cmbBoxWN8.Items.Add("30");

            cmbBoxWN9.Items.Add("1");
            cmbBoxWN9.Items.Add("2");
            cmbBoxWN9.Items.Add("3");
            cmbBoxWN9.Items.Add("4");
            cmbBoxWN9.Items.Add("5");
            cmbBoxWN9.Items.Add("6");
            cmbBoxWN9.Items.Add("7");
            cmbBoxWN9.Items.Add("8");
            cmbBoxWN9.Items.Add("9");
            cmbBoxWN9.Items.Add("10");
            cmbBoxWN9.Items.Add("11");
            cmbBoxWN9.Items.Add("12");
            cmbBoxWN9.Items.Add("13");
            cmbBoxWN9.Items.Add("14");
            cmbBoxWN9.Items.Add("15");
            cmbBoxWN9.Items.Add("16");
            cmbBoxWN9.Items.Add("17");
            cmbBoxWN9.Items.Add("18");
            cmbBoxWN9.Items.Add("19");
            cmbBoxWN9.Items.Add("20");
            cmbBoxWN9.Items.Add("21");
            cmbBoxWN9.Items.Add("22");
            cmbBoxWN9.Items.Add("23");
            cmbBoxWN9.Items.Add("24");
            cmbBoxWN9.Items.Add("25");
            cmbBoxWN9.Items.Add("26");
            cmbBoxWN9.Items.Add("27");
            cmbBoxWN9.Items.Add("28");
            cmbBoxWN9.Items.Add("29");
            cmbBoxWN9.Items.Add("30");

            cmbBoxWN10.Items.Add("1");
            cmbBoxWN10.Items.Add("2");
            cmbBoxWN10.Items.Add("3");
            cmbBoxWN10.Items.Add("4");
            cmbBoxWN10.Items.Add("5");
            cmbBoxWN10.Items.Add("6");
            cmbBoxWN10.Items.Add("7");
            cmbBoxWN10.Items.Add("8");
            cmbBoxWN10.Items.Add("9");
            cmbBoxWN10.Items.Add("10");
            cmbBoxWN10.Items.Add("11");
            cmbBoxWN10.Items.Add("12");
            cmbBoxWN10.Items.Add("13");
            cmbBoxWN10.Items.Add("14");
            cmbBoxWN10.Items.Add("15");
            cmbBoxWN10.Items.Add("16");
            cmbBoxWN10.Items.Add("17");
            cmbBoxWN10.Items.Add("18");
            cmbBoxWN10.Items.Add("19");
            cmbBoxWN10.Items.Add("20");
            cmbBoxWN10.Items.Add("21");
            cmbBoxWN10.Items.Add("22");
            cmbBoxWN10.Items.Add("23");
            cmbBoxWN10.Items.Add("24");
            cmbBoxWN10.Items.Add("25");
            cmbBoxWN10.Items.Add("26");
            cmbBoxWN10.Items.Add("27");
            cmbBoxWN10.Items.Add("28");
            cmbBoxWN10.Items.Add("29");
            cmbBoxWN10.Items.Add("30");

            cmbBoxWN11.Items.Add("1");
            cmbBoxWN11.Items.Add("2");
            cmbBoxWN11.Items.Add("3");
            cmbBoxWN11.Items.Add("4");
            cmbBoxWN11.Items.Add("5");
            cmbBoxWN11.Items.Add("6");
            cmbBoxWN11.Items.Add("7");
            cmbBoxWN11.Items.Add("8");
            cmbBoxWN11.Items.Add("9");
            cmbBoxWN11.Items.Add("10");
            cmbBoxWN11.Items.Add("11");
            cmbBoxWN11.Items.Add("12");
            cmbBoxWN11.Items.Add("13");
            cmbBoxWN11.Items.Add("14");
            cmbBoxWN11.Items.Add("15");
            cmbBoxWN11.Items.Add("16");
            cmbBoxWN11.Items.Add("17");
            cmbBoxWN11.Items.Add("18");
            cmbBoxWN11.Items.Add("19");
            cmbBoxWN11.Items.Add("20");
            cmbBoxWN11.Items.Add("21");
            cmbBoxWN11.Items.Add("22");
            cmbBoxWN11.Items.Add("23");
            cmbBoxWN11.Items.Add("24");
            cmbBoxWN11.Items.Add("25");
            cmbBoxWN11.Items.Add("26");
            cmbBoxWN11.Items.Add("27");
            cmbBoxWN11.Items.Add("28");
            cmbBoxWN11.Items.Add("29");
            cmbBoxWN11.Items.Add("30");

            cmbBoxWN12.Items.Add("1");
            cmbBoxWN12.Items.Add("2");
            cmbBoxWN12.Items.Add("3");
            cmbBoxWN12.Items.Add("4");
            cmbBoxWN12.Items.Add("5");
            cmbBoxWN12.Items.Add("6");
            cmbBoxWN12.Items.Add("7");
            cmbBoxWN12.Items.Add("8");
            cmbBoxWN12.Items.Add("9");
            cmbBoxWN12.Items.Add("10");
            cmbBoxWN12.Items.Add("11");
            cmbBoxWN12.Items.Add("12");
            cmbBoxWN12.Items.Add("13");
            cmbBoxWN12.Items.Add("14");
            cmbBoxWN12.Items.Add("15");
            cmbBoxWN12.Items.Add("16");
            cmbBoxWN12.Items.Add("17");
            cmbBoxWN12.Items.Add("18");
            cmbBoxWN12.Items.Add("19");
            cmbBoxWN12.Items.Add("20");
            cmbBoxWN12.Items.Add("21");
            cmbBoxWN12.Items.Add("22");
            cmbBoxWN12.Items.Add("23");
            cmbBoxWN12.Items.Add("24");
            cmbBoxWN12.Items.Add("25");
            cmbBoxWN12.Items.Add("26");
            cmbBoxWN12.Items.Add("27");
            cmbBoxWN12.Items.Add("28");
            cmbBoxWN12.Items.Add("29");
            cmbBoxWN12.Items.Add("30");

            cmbBoxWN13.Items.Add("1");
            cmbBoxWN13.Items.Add("2");
            cmbBoxWN13.Items.Add("3");
            cmbBoxWN13.Items.Add("4");
            cmbBoxWN13.Items.Add("5");
            cmbBoxWN13.Items.Add("6");
            cmbBoxWN13.Items.Add("7");
            cmbBoxWN13.Items.Add("8");
            cmbBoxWN13.Items.Add("9");
            cmbBoxWN13.Items.Add("10");
            cmbBoxWN13.Items.Add("11");
            cmbBoxWN13.Items.Add("12");
            cmbBoxWN13.Items.Add("13");
            cmbBoxWN13.Items.Add("14");
            cmbBoxWN13.Items.Add("15");
            cmbBoxWN13.Items.Add("16");
            cmbBoxWN13.Items.Add("17");
            cmbBoxWN13.Items.Add("18");
            cmbBoxWN13.Items.Add("19");
            cmbBoxWN13.Items.Add("20");
            cmbBoxWN13.Items.Add("21");
            cmbBoxWN13.Items.Add("22");
            cmbBoxWN13.Items.Add("23");
            cmbBoxWN13.Items.Add("24");
            cmbBoxWN13.Items.Add("25");
            cmbBoxWN13.Items.Add("26");
            cmbBoxWN13.Items.Add("27");
            cmbBoxWN13.Items.Add("28");
            cmbBoxWN13.Items.Add("29");
            cmbBoxWN13.Items.Add("30");

            cmbBoxWN14.Items.Add("1");
            cmbBoxWN14.Items.Add("2");
            cmbBoxWN14.Items.Add("3");
            cmbBoxWN14.Items.Add("4");
            cmbBoxWN14.Items.Add("5");
            cmbBoxWN14.Items.Add("6");
            cmbBoxWN14.Items.Add("7");
            cmbBoxWN14.Items.Add("8");
            cmbBoxWN14.Items.Add("9");
            cmbBoxWN14.Items.Add("10");
            cmbBoxWN14.Items.Add("11");
            cmbBoxWN14.Items.Add("12");
            cmbBoxWN14.Items.Add("13");
            cmbBoxWN14.Items.Add("14");
            cmbBoxWN14.Items.Add("15");
            cmbBoxWN14.Items.Add("16");
            cmbBoxWN14.Items.Add("17");
            cmbBoxWN14.Items.Add("18");
            cmbBoxWN14.Items.Add("19");
            cmbBoxWN14.Items.Add("20");
            cmbBoxWN14.Items.Add("21");
            cmbBoxWN14.Items.Add("22");
            cmbBoxWN14.Items.Add("23");
            cmbBoxWN14.Items.Add("24");
            cmbBoxWN14.Items.Add("25");
            cmbBoxWN14.Items.Add("26");
            cmbBoxWN14.Items.Add("27");
            cmbBoxWN14.Items.Add("28");
            cmbBoxWN14.Items.Add("29");
            cmbBoxWN14.Items.Add("30");

            cmbBoxWN15.Items.Add("1");
            cmbBoxWN15.Items.Add("2");
            cmbBoxWN15.Items.Add("3");
            cmbBoxWN15.Items.Add("4");
            cmbBoxWN15.Items.Add("5");
            cmbBoxWN15.Items.Add("6");
            cmbBoxWN15.Items.Add("7");
            cmbBoxWN15.Items.Add("8");
            cmbBoxWN15.Items.Add("9");
            cmbBoxWN15.Items.Add("10");
            cmbBoxWN15.Items.Add("11");
            cmbBoxWN15.Items.Add("12");
            cmbBoxWN15.Items.Add("13");
            cmbBoxWN15.Items.Add("14");
            cmbBoxWN15.Items.Add("15");
            cmbBoxWN15.Items.Add("16");
            cmbBoxWN15.Items.Add("17");
            cmbBoxWN15.Items.Add("18");
            cmbBoxWN15.Items.Add("19");
            cmbBoxWN15.Items.Add("20");
            cmbBoxWN15.Items.Add("21");
            cmbBoxWN15.Items.Add("22");
            cmbBoxWN15.Items.Add("23");
            cmbBoxWN15.Items.Add("24");
            cmbBoxWN15.Items.Add("25");
            cmbBoxWN15.Items.Add("26");
            cmbBoxWN15.Items.Add("27");
            cmbBoxWN15.Items.Add("28");
            cmbBoxWN15.Items.Add("29");
            cmbBoxWN15.Items.Add("30");

            cmbBoxWN16.Items.Add("1");
            cmbBoxWN16.Items.Add("2");
            cmbBoxWN16.Items.Add("3");
            cmbBoxWN16.Items.Add("4");
            cmbBoxWN16.Items.Add("5");
            cmbBoxWN16.Items.Add("6");
            cmbBoxWN16.Items.Add("7");
            cmbBoxWN16.Items.Add("8");
            cmbBoxWN16.Items.Add("9");
            cmbBoxWN16.Items.Add("10");
            cmbBoxWN16.Items.Add("11");
            cmbBoxWN16.Items.Add("12");
            cmbBoxWN16.Items.Add("13");
            cmbBoxWN16.Items.Add("14");
            cmbBoxWN16.Items.Add("15");
            cmbBoxWN16.Items.Add("16");
            cmbBoxWN16.Items.Add("17");
            cmbBoxWN16.Items.Add("18");
            cmbBoxWN16.Items.Add("19");
            cmbBoxWN16.Items.Add("20");
            cmbBoxWN16.Items.Add("21");
            cmbBoxWN16.Items.Add("22");
            cmbBoxWN16.Items.Add("23");
            cmbBoxWN16.Items.Add("24");
            cmbBoxWN16.Items.Add("25");
            cmbBoxWN16.Items.Add("26");
            cmbBoxWN16.Items.Add("27");
            cmbBoxWN16.Items.Add("28");
            cmbBoxWN16.Items.Add("29");
            cmbBoxWN16.Items.Add("30");

            cmbBoxWN17.Items.Add("1");
            cmbBoxWN17.Items.Add("2");
            cmbBoxWN17.Items.Add("3");
            cmbBoxWN17.Items.Add("4");
            cmbBoxWN17.Items.Add("5");
            cmbBoxWN17.Items.Add("6");
            cmbBoxWN17.Items.Add("7");
            cmbBoxWN17.Items.Add("8");
            cmbBoxWN17.Items.Add("9");
            cmbBoxWN17.Items.Add("10");
            cmbBoxWN17.Items.Add("11");
            cmbBoxWN17.Items.Add("12");
            cmbBoxWN17.Items.Add("13");
            cmbBoxWN17.Items.Add("14");
            cmbBoxWN17.Items.Add("15");
            cmbBoxWN17.Items.Add("16");
            cmbBoxWN17.Items.Add("17");
            cmbBoxWN17.Items.Add("18");
            cmbBoxWN17.Items.Add("19");
            cmbBoxWN17.Items.Add("20");
            cmbBoxWN17.Items.Add("21");
            cmbBoxWN17.Items.Add("22");
            cmbBoxWN17.Items.Add("23");
            cmbBoxWN17.Items.Add("24");
            cmbBoxWN17.Items.Add("25");
            cmbBoxWN17.Items.Add("26");
            cmbBoxWN17.Items.Add("27");
            cmbBoxWN17.Items.Add("28");
            cmbBoxWN17.Items.Add("29");
            cmbBoxWN17.Items.Add("30");

            cmbBoxWN18.Items.Add("1");
            cmbBoxWN18.Items.Add("2");
            cmbBoxWN18.Items.Add("3");
            cmbBoxWN18.Items.Add("4");
            cmbBoxWN18.Items.Add("5");
            cmbBoxWN18.Items.Add("6");
            cmbBoxWN18.Items.Add("7");
            cmbBoxWN18.Items.Add("8");
            cmbBoxWN18.Items.Add("9");
            cmbBoxWN18.Items.Add("10");
            cmbBoxWN18.Items.Add("11");
            cmbBoxWN18.Items.Add("12");
            cmbBoxWN18.Items.Add("13");
            cmbBoxWN18.Items.Add("14");
            cmbBoxWN18.Items.Add("15");
            cmbBoxWN18.Items.Add("16");
            cmbBoxWN18.Items.Add("17");
            cmbBoxWN18.Items.Add("18");
            cmbBoxWN18.Items.Add("19");
            cmbBoxWN18.Items.Add("20");
            cmbBoxWN18.Items.Add("21");
            cmbBoxWN18.Items.Add("22");
            cmbBoxWN18.Items.Add("23");
            cmbBoxWN18.Items.Add("24");
            cmbBoxWN18.Items.Add("25");
            cmbBoxWN18.Items.Add("26");
            cmbBoxWN18.Items.Add("27");
            cmbBoxWN18.Items.Add("28");
            cmbBoxWN18.Items.Add("29");
            cmbBoxWN18.Items.Add("30");

            cmbBoxWN19.Items.Add("1");
            cmbBoxWN19.Items.Add("2");
            cmbBoxWN19.Items.Add("3");
            cmbBoxWN19.Items.Add("4");
            cmbBoxWN19.Items.Add("5");
            cmbBoxWN19.Items.Add("6");
            cmbBoxWN19.Items.Add("7");
            cmbBoxWN19.Items.Add("8");
            cmbBoxWN19.Items.Add("9");
            cmbBoxWN19.Items.Add("10");
            cmbBoxWN19.Items.Add("11");
            cmbBoxWN19.Items.Add("12");
            cmbBoxWN19.Items.Add("13");
            cmbBoxWN19.Items.Add("14");
            cmbBoxWN19.Items.Add("15");
            cmbBoxWN19.Items.Add("16");
            cmbBoxWN19.Items.Add("17");
            cmbBoxWN19.Items.Add("18");
            cmbBoxWN19.Items.Add("19");
            cmbBoxWN19.Items.Add("20");
            cmbBoxWN19.Items.Add("21");
            cmbBoxWN19.Items.Add("22");
            cmbBoxWN19.Items.Add("23");
            cmbBoxWN19.Items.Add("24");
            cmbBoxWN19.Items.Add("25");
            cmbBoxWN19.Items.Add("26");
            cmbBoxWN19.Items.Add("27");
            cmbBoxWN19.Items.Add("28");
            cmbBoxWN19.Items.Add("29");
            cmbBoxWN19.Items.Add("30");

            cmbBoxWN20.Items.Add("1");
            cmbBoxWN20.Items.Add("2");
            cmbBoxWN20.Items.Add("3");
            cmbBoxWN20.Items.Add("4");
            cmbBoxWN20.Items.Add("5");
            cmbBoxWN20.Items.Add("6");
            cmbBoxWN20.Items.Add("7");
            cmbBoxWN20.Items.Add("8");
            cmbBoxWN20.Items.Add("9");
            cmbBoxWN20.Items.Add("10");
            cmbBoxWN20.Items.Add("11");
            cmbBoxWN20.Items.Add("12");
            cmbBoxWN20.Items.Add("13");
            cmbBoxWN20.Items.Add("14");
            cmbBoxWN20.Items.Add("15");
            cmbBoxWN20.Items.Add("16");
            cmbBoxWN20.Items.Add("17");
            cmbBoxWN20.Items.Add("18");
            cmbBoxWN20.Items.Add("19");
            cmbBoxWN20.Items.Add("20");
            cmbBoxWN20.Items.Add("21");
            cmbBoxWN20.Items.Add("22");
            cmbBoxWN20.Items.Add("23");
            cmbBoxWN20.Items.Add("24");
            cmbBoxWN20.Items.Add("25");
            cmbBoxWN20.Items.Add("26");
            cmbBoxWN20.Items.Add("27");
            cmbBoxWN20.Items.Add("28");
            cmbBoxWN20.Items.Add("29");
            cmbBoxWN20.Items.Add("30");

            cmbBoxWN21.Items.Add("1");
            cmbBoxWN21.Items.Add("2");
            cmbBoxWN21.Items.Add("3");
            cmbBoxWN21.Items.Add("4");
            cmbBoxWN21.Items.Add("5");
            cmbBoxWN21.Items.Add("6");
            cmbBoxWN21.Items.Add("7");
            cmbBoxWN21.Items.Add("8");
            cmbBoxWN21.Items.Add("9");
            cmbBoxWN21.Items.Add("10");
            cmbBoxWN21.Items.Add("11");
            cmbBoxWN21.Items.Add("12");
            cmbBoxWN21.Items.Add("13");
            cmbBoxWN21.Items.Add("14");
            cmbBoxWN21.Items.Add("15");
            cmbBoxWN21.Items.Add("16");
            cmbBoxWN21.Items.Add("17");
            cmbBoxWN21.Items.Add("18");
            cmbBoxWN21.Items.Add("19");
            cmbBoxWN21.Items.Add("20");
            cmbBoxWN21.Items.Add("21");
            cmbBoxWN21.Items.Add("22");
            cmbBoxWN21.Items.Add("23");
            cmbBoxWN21.Items.Add("24");
            cmbBoxWN21.Items.Add("25");
            cmbBoxWN21.Items.Add("26");
            cmbBoxWN21.Items.Add("27");
            cmbBoxWN21.Items.Add("28");
            cmbBoxWN21.Items.Add("29");
            cmbBoxWN21.Items.Add("30");

            cmbBoxWN22.Items.Add("1");
            cmbBoxWN22.Items.Add("2");
            cmbBoxWN22.Items.Add("3");
            cmbBoxWN22.Items.Add("4");
            cmbBoxWN22.Items.Add("5");
            cmbBoxWN22.Items.Add("6");
            cmbBoxWN22.Items.Add("7");
            cmbBoxWN22.Items.Add("8");
            cmbBoxWN22.Items.Add("9");
            cmbBoxWN22.Items.Add("10");
            cmbBoxWN22.Items.Add("11");
            cmbBoxWN22.Items.Add("12");
            cmbBoxWN22.Items.Add("13");
            cmbBoxWN22.Items.Add("14");
            cmbBoxWN22.Items.Add("15");
            cmbBoxWN22.Items.Add("16");
            cmbBoxWN22.Items.Add("17");
            cmbBoxWN22.Items.Add("18");
            cmbBoxWN22.Items.Add("19");
            cmbBoxWN22.Items.Add("20");
            cmbBoxWN22.Items.Add("21");
            cmbBoxWN22.Items.Add("22");
            cmbBoxWN22.Items.Add("23");
            cmbBoxWN22.Items.Add("24");
            cmbBoxWN22.Items.Add("25");
            cmbBoxWN22.Items.Add("26");
            cmbBoxWN22.Items.Add("27");
            cmbBoxWN22.Items.Add("28");
            cmbBoxWN22.Items.Add("29");
            cmbBoxWN22.Items.Add("30");

            cmbBoxWN23.Items.Add("1");
            cmbBoxWN23.Items.Add("2");
            cmbBoxWN23.Items.Add("3");
            cmbBoxWN23.Items.Add("4");
            cmbBoxWN23.Items.Add("5");
            cmbBoxWN23.Items.Add("6");
            cmbBoxWN23.Items.Add("7");
            cmbBoxWN23.Items.Add("8");
            cmbBoxWN23.Items.Add("9");
            cmbBoxWN23.Items.Add("10");
            cmbBoxWN23.Items.Add("11");
            cmbBoxWN23.Items.Add("12");
            cmbBoxWN23.Items.Add("13");
            cmbBoxWN23.Items.Add("14");
            cmbBoxWN23.Items.Add("15");
            cmbBoxWN23.Items.Add("16");
            cmbBoxWN23.Items.Add("17");
            cmbBoxWN23.Items.Add("18");
            cmbBoxWN23.Items.Add("19");
            cmbBoxWN23.Items.Add("20");
            cmbBoxWN23.Items.Add("21");
            cmbBoxWN23.Items.Add("22");
            cmbBoxWN23.Items.Add("23");
            cmbBoxWN23.Items.Add("24");
            cmbBoxWN23.Items.Add("25");
            cmbBoxWN23.Items.Add("26");
            cmbBoxWN23.Items.Add("27");
            cmbBoxWN23.Items.Add("28");
            cmbBoxWN23.Items.Add("29");
            cmbBoxWN23.Items.Add("30");

            cmbBoxWN24.Items.Add("1");
            cmbBoxWN24.Items.Add("2");
            cmbBoxWN24.Items.Add("3");
            cmbBoxWN24.Items.Add("4");
            cmbBoxWN24.Items.Add("5");
            cmbBoxWN24.Items.Add("6");
            cmbBoxWN24.Items.Add("7");
            cmbBoxWN24.Items.Add("8");
            cmbBoxWN24.Items.Add("9");
            cmbBoxWN24.Items.Add("10");
            cmbBoxWN24.Items.Add("11");
            cmbBoxWN24.Items.Add("12");
            cmbBoxWN24.Items.Add("13");
            cmbBoxWN24.Items.Add("14");
            cmbBoxWN24.Items.Add("15");
            cmbBoxWN24.Items.Add("16");
            cmbBoxWN24.Items.Add("17");
            cmbBoxWN24.Items.Add("18");
            cmbBoxWN24.Items.Add("19");
            cmbBoxWN24.Items.Add("20");
            cmbBoxWN24.Items.Add("21");
            cmbBoxWN24.Items.Add("22");
            cmbBoxWN24.Items.Add("23");
            cmbBoxWN24.Items.Add("24");
            cmbBoxWN24.Items.Add("25");
            cmbBoxWN24.Items.Add("26");
            cmbBoxWN24.Items.Add("27");
            cmbBoxWN24.Items.Add("28");
            cmbBoxWN24.Items.Add("29");
            cmbBoxWN24.Items.Add("30");

            cmbBoxWN25.Items.Add("1");
            cmbBoxWN25.Items.Add("2");
            cmbBoxWN25.Items.Add("3");
            cmbBoxWN25.Items.Add("4");
            cmbBoxWN25.Items.Add("5");
            cmbBoxWN25.Items.Add("6");
            cmbBoxWN25.Items.Add("7");
            cmbBoxWN25.Items.Add("8");
            cmbBoxWN25.Items.Add("9");
            cmbBoxWN25.Items.Add("10");
            cmbBoxWN25.Items.Add("11");
            cmbBoxWN25.Items.Add("12");
            cmbBoxWN25.Items.Add("13");
            cmbBoxWN25.Items.Add("14");
            cmbBoxWN25.Items.Add("15");
            cmbBoxWN25.Items.Add("16");
            cmbBoxWN25.Items.Add("17");
            cmbBoxWN25.Items.Add("18");
            cmbBoxWN25.Items.Add("19");
            cmbBoxWN25.Items.Add("20");
            cmbBoxWN25.Items.Add("21");
            cmbBoxWN25.Items.Add("22");
            cmbBoxWN25.Items.Add("23");
            cmbBoxWN25.Items.Add("24");
            cmbBoxWN25.Items.Add("25");
            cmbBoxWN25.Items.Add("26");
            cmbBoxWN25.Items.Add("27");
            cmbBoxWN25.Items.Add("28");
            cmbBoxWN25.Items.Add("29");
            cmbBoxWN25.Items.Add("30");

            cmbBoxWN26.Items.Add("1");
            cmbBoxWN26.Items.Add("2");
            cmbBoxWN26.Items.Add("3");
            cmbBoxWN26.Items.Add("4");
            cmbBoxWN26.Items.Add("5");
            cmbBoxWN26.Items.Add("6");
            cmbBoxWN26.Items.Add("7");
            cmbBoxWN26.Items.Add("8");
            cmbBoxWN26.Items.Add("9");
            cmbBoxWN26.Items.Add("10");
            cmbBoxWN26.Items.Add("11");
            cmbBoxWN26.Items.Add("12");
            cmbBoxWN26.Items.Add("13");
            cmbBoxWN26.Items.Add("14");
            cmbBoxWN26.Items.Add("15");
            cmbBoxWN26.Items.Add("16");
            cmbBoxWN26.Items.Add("17");
            cmbBoxWN26.Items.Add("18");
            cmbBoxWN26.Items.Add("19");
            cmbBoxWN26.Items.Add("20");
            cmbBoxWN26.Items.Add("21");
            cmbBoxWN26.Items.Add("22");
            cmbBoxWN26.Items.Add("23");
            cmbBoxWN26.Items.Add("24");
            cmbBoxWN26.Items.Add("25");
            cmbBoxWN26.Items.Add("26");
            cmbBoxWN26.Items.Add("27");
            cmbBoxWN26.Items.Add("28");
            cmbBoxWN26.Items.Add("29");
            cmbBoxWN26.Items.Add("30");

            cmbBoxWN27.Items.Add("1");
            cmbBoxWN27.Items.Add("2");
            cmbBoxWN27.Items.Add("3");
            cmbBoxWN27.Items.Add("4");
            cmbBoxWN27.Items.Add("5");
            cmbBoxWN27.Items.Add("6");
            cmbBoxWN27.Items.Add("7");
            cmbBoxWN27.Items.Add("8");
            cmbBoxWN27.Items.Add("9");
            cmbBoxWN27.Items.Add("10");
            cmbBoxWN27.Items.Add("11");
            cmbBoxWN27.Items.Add("12");
            cmbBoxWN27.Items.Add("13");
            cmbBoxWN27.Items.Add("14");
            cmbBoxWN27.Items.Add("15");
            cmbBoxWN27.Items.Add("16");
            cmbBoxWN27.Items.Add("17");
            cmbBoxWN27.Items.Add("18");
            cmbBoxWN27.Items.Add("19");
            cmbBoxWN27.Items.Add("20");
            cmbBoxWN27.Items.Add("21");
            cmbBoxWN27.Items.Add("22");
            cmbBoxWN27.Items.Add("23");
            cmbBoxWN27.Items.Add("24");
            cmbBoxWN27.Items.Add("25");
            cmbBoxWN27.Items.Add("26");
            cmbBoxWN27.Items.Add("27");
            cmbBoxWN27.Items.Add("28");
            cmbBoxWN27.Items.Add("29");
            cmbBoxWN27.Items.Add("30");

            cmbBoxWN28.Items.Add("1");
            cmbBoxWN28.Items.Add("2");
            cmbBoxWN28.Items.Add("3");
            cmbBoxWN28.Items.Add("4");
            cmbBoxWN28.Items.Add("5");
            cmbBoxWN28.Items.Add("6");
            cmbBoxWN28.Items.Add("7");
            cmbBoxWN28.Items.Add("8");
            cmbBoxWN28.Items.Add("9");
            cmbBoxWN28.Items.Add("10");
            cmbBoxWN28.Items.Add("11");
            cmbBoxWN28.Items.Add("12");
            cmbBoxWN28.Items.Add("13");
            cmbBoxWN28.Items.Add("14");
            cmbBoxWN28.Items.Add("15");
            cmbBoxWN28.Items.Add("16");
            cmbBoxWN28.Items.Add("17");
            cmbBoxWN28.Items.Add("18");
            cmbBoxWN28.Items.Add("19");
            cmbBoxWN28.Items.Add("20");
            cmbBoxWN28.Items.Add("21");
            cmbBoxWN28.Items.Add("22");
            cmbBoxWN28.Items.Add("23");
            cmbBoxWN28.Items.Add("24");
            cmbBoxWN28.Items.Add("25");
            cmbBoxWN28.Items.Add("26");
            cmbBoxWN28.Items.Add("27");
            cmbBoxWN28.Items.Add("28");
            cmbBoxWN28.Items.Add("29");
            cmbBoxWN28.Items.Add("30");

            cmbBoxWN29.Items.Add("1");
            cmbBoxWN29.Items.Add("2");
            cmbBoxWN29.Items.Add("3");
            cmbBoxWN29.Items.Add("4");
            cmbBoxWN29.Items.Add("5");
            cmbBoxWN29.Items.Add("6");
            cmbBoxWN29.Items.Add("7");
            cmbBoxWN29.Items.Add("8");
            cmbBoxWN29.Items.Add("9");
            cmbBoxWN29.Items.Add("10");
            cmbBoxWN29.Items.Add("11");
            cmbBoxWN29.Items.Add("12");
            cmbBoxWN29.Items.Add("13");
            cmbBoxWN29.Items.Add("14");
            cmbBoxWN29.Items.Add("15");
            cmbBoxWN29.Items.Add("16");
            cmbBoxWN29.Items.Add("17");
            cmbBoxWN29.Items.Add("18");
            cmbBoxWN29.Items.Add("19");
            cmbBoxWN29.Items.Add("20");
            cmbBoxWN29.Items.Add("21");
            cmbBoxWN29.Items.Add("22");
            cmbBoxWN29.Items.Add("23");
            cmbBoxWN29.Items.Add("24");
            cmbBoxWN29.Items.Add("25");
            cmbBoxWN29.Items.Add("26");
            cmbBoxWN29.Items.Add("27");
            cmbBoxWN29.Items.Add("28");
            cmbBoxWN29.Items.Add("29");
            cmbBoxWN29.Items.Add("30");

            cmbBoxWN30.Items.Add("1");
            cmbBoxWN30.Items.Add("2");
            cmbBoxWN30.Items.Add("3");
            cmbBoxWN30.Items.Add("4");
            cmbBoxWN30.Items.Add("5");
            cmbBoxWN30.Items.Add("6");
            cmbBoxWN30.Items.Add("7");
            cmbBoxWN30.Items.Add("8");
            cmbBoxWN30.Items.Add("9");
            cmbBoxWN30.Items.Add("10");
            cmbBoxWN30.Items.Add("11");
            cmbBoxWN30.Items.Add("12");
            cmbBoxWN30.Items.Add("13");
            cmbBoxWN30.Items.Add("14");
            cmbBoxWN30.Items.Add("15");
            cmbBoxWN30.Items.Add("16");
            cmbBoxWN30.Items.Add("17");
            cmbBoxWN30.Items.Add("18");
            cmbBoxWN30.Items.Add("19");
            cmbBoxWN30.Items.Add("20");
            cmbBoxWN30.Items.Add("21");
            cmbBoxWN30.Items.Add("22");
            cmbBoxWN30.Items.Add("23");
            cmbBoxWN30.Items.Add("24");
            cmbBoxWN30.Items.Add("25");
            cmbBoxWN30.Items.Add("26");
            cmbBoxWN30.Items.Add("27");
            cmbBoxWN30.Items.Add("28");
            cmbBoxWN30.Items.Add("29");
            cmbBoxWN30.Items.Add("30");
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
            if(chckBxBreaks1.Checked && !cmbBoxBreaks1.Equals(""))
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
            cmbBoxWN1.Items.Add(lastValue1);
            cmbBoxWN2.Items.Add(lastValue1);
            cmbBoxWN3.Items.Add(lastValue1);
            cmbBoxWN4.Items.Add(lastValue1);
            cmbBoxWN5.Items.Add(lastValue1);
            cmbBoxWN6.Items.Add(lastValue1);
            cmbBoxWN7.Items.Add(lastValue1);
            cmbBoxWN8.Items.Add(lastValue1);
            cmbBoxWN9.Items.Add(lastValue1);
            cmbBoxWN10.Items.Add(lastValue1);
            cmbBoxWN11.Items.Add(lastValue1);
            cmbBoxWN12.Items.Add(lastValue1);
            cmbBoxWN13.Items.Add(lastValue1);
            cmbBoxWN14.Items.Add(lastValue1);
            cmbBoxWN15.Items.Add(lastValue1);
            cmbBoxWN16.Items.Add(lastValue1);
            cmbBoxWN17.Items.Add(lastValue1);
            cmbBoxWN18.Items.Add(lastValue1);
            cmbBoxWN19.Items.Add(lastValue1);
            cmbBoxWN20.Items.Add(lastValue1);
            cmbBoxWN21.Items.Add(lastValue1);
            cmbBoxWN22.Items.Add(lastValue1);
            cmbBoxWN23.Items.Add(lastValue1);
            cmbBoxWN24.Items.Add(lastValue1);
            cmbBoxWN25.Items.Add(lastValue1);
            cmbBoxWN26.Items.Add(lastValue1);
            cmbBoxWN27.Items.Add(lastValue1);
            cmbBoxWN28.Items.Add(lastValue1);
            cmbBoxWN29.Items.Add(lastValue1);
            cmbBoxWN30.Items.Add(lastValue1);

            cmbBoxWN1.Items.Add(lastValue2);
            cmbBoxWN2.Items.Add(lastValue2);
            cmbBoxWN3.Items.Add(lastValue2);
            cmbBoxWN4.Items.Add(lastValue2);
            cmbBoxWN5.Items.Add(lastValue2);
            cmbBoxWN6.Items.Add(lastValue2);
            cmbBoxWN7.Items.Add(lastValue2);
            cmbBoxWN8.Items.Add(lastValue2);
            cmbBoxWN9.Items.Add(lastValue2);
            cmbBoxWN10.Items.Add(lastValue2);
            cmbBoxWN11.Items.Add(lastValue2);
            cmbBoxWN12.Items.Add(lastValue2);
            cmbBoxWN13.Items.Add(lastValue2);
            cmbBoxWN14.Items.Add(lastValue2);
            cmbBoxWN15.Items.Add(lastValue2);
            cmbBoxWN16.Items.Add(lastValue2);
            cmbBoxWN17.Items.Add(lastValue2);
            cmbBoxWN18.Items.Add(lastValue2);
            cmbBoxWN19.Items.Add(lastValue2);
            cmbBoxWN20.Items.Add(lastValue2);
            cmbBoxWN21.Items.Add(lastValue2);
            cmbBoxWN22.Items.Add(lastValue2);
            cmbBoxWN23.Items.Add(lastValue2);
            cmbBoxWN24.Items.Add(lastValue2);
            cmbBoxWN25.Items.Add(lastValue2);
            cmbBoxWN26.Items.Add(lastValue2);
            cmbBoxWN27.Items.Add(lastValue2);
            cmbBoxWN28.Items.Add(lastValue2);
            cmbBoxWN29.Items.Add(lastValue2);
            cmbBoxWN30.Items.Add(lastValue2);

            cmbBoxWN1.Items.Add(lastValue3);
            cmbBoxWN2.Items.Add(lastValue3);
            cmbBoxWN3.Items.Add(lastValue3);
            cmbBoxWN4.Items.Add(lastValue3);
            cmbBoxWN5.Items.Add(lastValue3);
            cmbBoxWN6.Items.Add(lastValue3);
            cmbBoxWN7.Items.Add(lastValue3);
            cmbBoxWN8.Items.Add(lastValue3);
            cmbBoxWN9.Items.Add(lastValue3);
            cmbBoxWN10.Items.Add(lastValue3);
            cmbBoxWN11.Items.Add(lastValue3);
            cmbBoxWN12.Items.Add(lastValue3);
            cmbBoxWN13.Items.Add(lastValue3);
            cmbBoxWN14.Items.Add(lastValue3);
            cmbBoxWN15.Items.Add(lastValue3);
            cmbBoxWN16.Items.Add(lastValue3);
            cmbBoxWN17.Items.Add(lastValue3);
            cmbBoxWN18.Items.Add(lastValue3);
            cmbBoxWN19.Items.Add(lastValue3);
            cmbBoxWN20.Items.Add(lastValue3);
            cmbBoxWN21.Items.Add(lastValue3);
            cmbBoxWN22.Items.Add(lastValue3);
            cmbBoxWN23.Items.Add(lastValue3);
            cmbBoxWN24.Items.Add(lastValue3);
            cmbBoxWN25.Items.Add(lastValue3);
            cmbBoxWN26.Items.Add(lastValue3);
            cmbBoxWN27.Items.Add(lastValue3);
            cmbBoxWN28.Items.Add(lastValue3);
            cmbBoxWN29.Items.Add(lastValue3);
            cmbBoxWN30.Items.Add(lastValue3);

            cmbBoxWN1.Items.Add(lastValue4);
            cmbBoxWN2.Items.Add(lastValue4);
            cmbBoxWN3.Items.Add(lastValue4);
            cmbBoxWN4.Items.Add(lastValue4);
            cmbBoxWN5.Items.Add(lastValue4);
            cmbBoxWN6.Items.Add(lastValue4);
            cmbBoxWN7.Items.Add(lastValue4);
            cmbBoxWN8.Items.Add(lastValue4);
            cmbBoxWN9.Items.Add(lastValue4);
            cmbBoxWN10.Items.Add(lastValue4);
            cmbBoxWN11.Items.Add(lastValue4);
            cmbBoxWN12.Items.Add(lastValue4);
            cmbBoxWN13.Items.Add(lastValue4);
            cmbBoxWN14.Items.Add(lastValue4);
            cmbBoxWN15.Items.Add(lastValue4);
            cmbBoxWN16.Items.Add(lastValue4);
            cmbBoxWN17.Items.Add(lastValue4);
            cmbBoxWN18.Items.Add(lastValue4);
            cmbBoxWN19.Items.Add(lastValue4);
            cmbBoxWN20.Items.Add(lastValue4);
            cmbBoxWN21.Items.Add(lastValue4);
            cmbBoxWN22.Items.Add(lastValue4);
            cmbBoxWN23.Items.Add(lastValue4);
            cmbBoxWN24.Items.Add(lastValue4);
            cmbBoxWN25.Items.Add(lastValue4);
            cmbBoxWN26.Items.Add(lastValue4);
            cmbBoxWN27.Items.Add(lastValue4);
            cmbBoxWN28.Items.Add(lastValue4);
            cmbBoxWN29.Items.Add(lastValue4);
            cmbBoxWN30.Items.Add(lastValue4);

            cmbBoxWN1.Items.Add(lastValue5);
            cmbBoxWN2.Items.Add(lastValue5);
            cmbBoxWN3.Items.Add(lastValue5);
            cmbBoxWN4.Items.Add(lastValue5);
            cmbBoxWN5.Items.Add(lastValue5);
            cmbBoxWN6.Items.Add(lastValue5);
            cmbBoxWN7.Items.Add(lastValue5);
            cmbBoxWN8.Items.Add(lastValue5);
            cmbBoxWN9.Items.Add(lastValue5);
            cmbBoxWN10.Items.Add(lastValue5);
            cmbBoxWN11.Items.Add(lastValue5);
            cmbBoxWN12.Items.Add(lastValue5);
            cmbBoxWN13.Items.Add(lastValue5);
            cmbBoxWN14.Items.Add(lastValue5);
            cmbBoxWN15.Items.Add(lastValue5);
            cmbBoxWN16.Items.Add(lastValue5);
            cmbBoxWN17.Items.Add(lastValue5);
            cmbBoxWN18.Items.Add(lastValue5);
            cmbBoxWN19.Items.Add(lastValue5);
            cmbBoxWN20.Items.Add(lastValue5);
            cmbBoxWN21.Items.Add(lastValue5);
            cmbBoxWN22.Items.Add(lastValue5);
            cmbBoxWN23.Items.Add(lastValue5);
            cmbBoxWN24.Items.Add(lastValue5);
            cmbBoxWN25.Items.Add(lastValue5);
            cmbBoxWN26.Items.Add(lastValue5);
            cmbBoxWN27.Items.Add(lastValue5);
            cmbBoxWN28.Items.Add(lastValue5);
            cmbBoxWN29.Items.Add(lastValue5);
            cmbBoxWN30.Items.Add(lastValue5);
        }

        private void RemoveFromComboBox()
        {
            cmbBoxWN1.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN2.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN3.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN4.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN5.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN6.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN7.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN8.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN9.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN10.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN11.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN12.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN13.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN14.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN15.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN16.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN17.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN18.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN19.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN20.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN21.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN22.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN23.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN24.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN25.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN26.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN27.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN28.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN29.Items.Remove(cmbBoxBreaks1.Text);
            cmbBoxWN30.Items.Remove(cmbBoxBreaks1.Text);

            cmbBoxWN1.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN2.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN3.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN4.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN5.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN6.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN7.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN8.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN9.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN10.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN11.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN12.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN13.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN14.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN15.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN16.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN17.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN18.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN19.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN20.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN21.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN22.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN23.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN24.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN25.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN26.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN27.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN28.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN29.Items.Remove(cmbBoxBreaks2.Text);
            cmbBoxWN30.Items.Remove(cmbBoxBreaks2.Text);

            cmbBoxWN1.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN2.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN3.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN4.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN5.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN6.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN7.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN8.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN9.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN10.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN11.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN12.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN13.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN14.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN15.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN16.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN17.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN18.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN19.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN20.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN21.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN22.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN23.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN24.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN25.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN26.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN27.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN28.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN29.Items.Remove(cmbBoxBreaks3.Text);
            cmbBoxWN30.Items.Remove(cmbBoxBreaks3.Text);

            cmbBoxWN1.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN2.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN3.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN4.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN5.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN6.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN7.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN8.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN9.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN10.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN11.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN12.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN13.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN14.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN15.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN16.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN17.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN18.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN19.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN20.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN21.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN22.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN23.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN24.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN25.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN26.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN27.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN28.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN29.Items.Remove(cmbBoxBreaks4.Text);
            cmbBoxWN30.Items.Remove(cmbBoxBreaks4.Text);

            cmbBoxWN1.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN2.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN3.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN4.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN5.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN6.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN7.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN8.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN9.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN10.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN11.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN12.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN13.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN14.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN15.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN16.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN17.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN18.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN19.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN20.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN21.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN22.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN23.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN24.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN25.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN26.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN27.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN28.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN29.Items.Remove(cmbBoxBreaks5.Text);
            cmbBoxWN30.Items.Remove(cmbBoxBreaks5.Text);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
