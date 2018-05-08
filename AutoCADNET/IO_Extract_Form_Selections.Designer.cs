namespace IO_Extract_Dialogs
{
    partial class IO_Extract_Form_Selections
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
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Extract = new System.Windows.Forms.Button();
            this.label_1 = new System.Windows.Forms.Label();
            this.lv_IOCards = new System.Windows.Forms.ListView();
            this.btn_Browse = new System.Windows.Forms.Button();
            this.tb_Path = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(441, 367);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(96, 29);
            this.btn_Cancel.TabIndex = 10;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.click_Cancel);
            // 
            // btn_Extract
            // 
            this.btn_Extract.Location = new System.Drawing.Point(339, 367);
            this.btn_Extract.Name = "btn_Extract";
            this.btn_Extract.Size = new System.Drawing.Size(96, 29);
            this.btn_Extract.TabIndex = 9;
            this.btn_Extract.Text = "Extract";
            this.btn_Extract.UseVisualStyleBackColor = true;
            this.btn_Extract.Click += new System.EventHandler(this.click_Extract);
            // 
            // label_1
            // 
            this.label_1.AutoSize = true;
            this.label_1.Location = new System.Drawing.Point(12, 9);
            this.label_1.Name = "label_1";
            this.label_1.Size = new System.Drawing.Size(303, 17);
            this.label_1.TabIndex = 8;
            this.label_1.Text = "Select the IO Cards to include in the extraction.";
            // 
            // lv_IOCards
            // 
            this.lv_IOCards.CheckBoxes = true;
            this.lv_IOCards.Location = new System.Drawing.Point(12, 30);
            this.lv_IOCards.Name = "lv_IOCards";
            this.lv_IOCards.Size = new System.Drawing.Size(533, 275);
            this.lv_IOCards.TabIndex = 6;
            this.lv_IOCards.UseCompatibleStateImageBehavior = false;
            this.lv_IOCards.View = System.Windows.Forms.View.Details;
            this.lv_IOCards.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.columnclick_lv_IOCards);
            // 
            // btn_Browse
            // 
            this.btn_Browse.Location = new System.Drawing.Point(12, 339);
            this.btn_Browse.Name = "btn_Browse";
            this.btn_Browse.Size = new System.Drawing.Size(32, 23);
            this.btn_Browse.TabIndex = 11;
            this.btn_Browse.Text = "...";
            this.btn_Browse.UseVisualStyleBackColor = true;
            this.btn_Browse.Click += new System.EventHandler(this.click_Browse);
            // 
            // tb_Path
            // 
            this.tb_Path.Location = new System.Drawing.Point(51, 339);
            this.tb_Path.Name = "tb_Path";
            this.tb_Path.Size = new System.Drawing.Size(494, 22);
            this.tb_Path.TabIndex = 12;
            this.tb_Path.Text = @"\\thehaskellco.net\Remote\AtlantaData\_Disciplines\Controls\3 Controls Design\Engineering Tools\Tag Import Tools\IO_TAGS " + "rev2.1.xlsm";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(50, 319);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(200, 17);
            this.label1.TabIndex = 13;
            this.label1.Text = "IO Tag Export Utility Excel File:";
            // 
            // IO_Extract_Form_Selections
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 413);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tb_Path);
            this.Controls.Add(this.btn_Browse);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_Extract);
            this.Controls.Add(this.label_1);
            this.Controls.Add(this.lv_IOCards);
            this.Name = "IO_Extract_Form_Selections";
            this.Text = "IO Card Selection";
            this.Load += new System.EventHandler(this.load_IOExtractSelections);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Extract;
        private System.Windows.Forms.Label label_1;
        private System.Windows.Forms.ListView lv_IOCards;
        private System.Windows.Forms.Button btn_Browse;
        private System.Windows.Forms.TextBox tb_Path;
        private System.Windows.Forms.Label label1;
    }
}