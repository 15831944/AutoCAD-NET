namespace Tutorial_Dialogs
{
    partial class Tutorials_by_Adam_Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.cb_TutorialSelect = new System.Windows.Forms.ComboBox();
            this.lbl_SelectInstr = new System.Windows.Forms.Label();
            this.wb_Tutorial = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // cb_TutorialSelect
            // 
            this.cb_TutorialSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cb_TutorialSelect.FormattingEnabled = true;
            this.cb_TutorialSelect.Location = new System.Drawing.Point(230, 13);
            this.cb_TutorialSelect.Name = "cb_TutorialSelect";
            this.cb_TutorialSelect.Size = new System.Drawing.Size(273, 24);
            this.cb_TutorialSelect.TabIndex = 4;
            this.cb_TutorialSelect.TextChanged += new System.EventHandler(this.TextChanged_cb_TutorialSelect);
            // 
            // lbl_SelectInstr
            // 
            this.lbl_SelectInstr.AutoSize = true;
            this.lbl_SelectInstr.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_SelectInstr.Location = new System.Drawing.Point(12, 9);
            this.lbl_SelectInstr.Name = "lbl_SelectInstr";
            this.lbl_SelectInstr.Size = new System.Drawing.Size(200, 25);
            this.lbl_SelectInstr.TabIndex = 3;
            this.lbl_SelectInstr.Text = "Select tutorial to view:";
            // 
            // wb_Tutorial
            // 
            this.wb_Tutorial.Location = new System.Drawing.Point(241, 14);
            this.wb_Tutorial.MinimumSize = new System.Drawing.Size(20, 20);
            this.wb_Tutorial.Name = "wb_Tutorial";
            this.wb_Tutorial.Size = new System.Drawing.Size(148, 20);
            this.wb_Tutorial.TabIndex = 5;
            // 
            // Tutorials_by_Adam_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(515, 50);
            this.Controls.Add(this.cb_TutorialSelect);
            this.Controls.Add(this.wb_Tutorial);
            this.Controls.Add(this.lbl_SelectInstr);
            this.Name = "Tutorials_by_Adam_Form";
            this.Text = "Command Tutorials";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cb_TutorialSelect;
        private System.Windows.Forms.Label lbl_SelectInstr;
        private System.Windows.Forms.WebBrowser wb_Tutorial;
    }
}