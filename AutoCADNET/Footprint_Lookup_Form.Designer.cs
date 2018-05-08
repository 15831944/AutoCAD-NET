namespace Footprint_Lookup_Dialogs
{
    partial class Footprint_Lookup_Form
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
            this.txtbx_partNumSearch = new System.Windows.Forms.TextBox();
            this.lbl_partNumSearch = new System.Windows.Forms.Label();
            this.cmbbx_partNumbers = new System.Windows.Forms.ComboBox();
            this.lbl_partNumbers = new System.Windows.Forms.Label();
            this.lbl_Description = new System.Windows.Forms.Label();
            this.txtBx_Description = new System.Windows.Forms.TextBox();
            this.btn_Search = new System.Windows.Forms.Button();
            this.btn_Add = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtbx_partNumSearch
            // 
            this.txtbx_partNumSearch.Location = new System.Drawing.Point(12, 33);
            this.txtbx_partNumSearch.Name = "txtbx_partNumSearch";
            this.txtbx_partNumSearch.Size = new System.Drawing.Size(225, 22);
            this.txtbx_partNumSearch.TabIndex = 0;
            this.txtbx_partNumSearch.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.previewKeyDown_txtbx_partNumSearch);
            // 
            // lbl_partNumSearch
            // 
            this.lbl_partNumSearch.AutoSize = true;
            this.lbl_partNumSearch.Location = new System.Drawing.Point(12, 13);
            this.lbl_partNumSearch.Name = "lbl_partNumSearch";
            this.lbl_partNumSearch.Size = new System.Drawing.Size(137, 17);
            this.lbl_partNumSearch.TabIndex = 1;
            this.lbl_partNumSearch.Text = "Part Number Search";
            // 
            // cmbbx_partNumbers
            // 
            this.cmbbx_partNumbers.FormattingEnabled = true;
            this.cmbbx_partNumbers.Location = new System.Drawing.Point(12, 96);
            this.cmbbx_partNumbers.Name = "cmbbx_partNumbers";
            this.cmbbx_partNumbers.Size = new System.Drawing.Size(225, 24);
            this.cmbbx_partNumbers.TabIndex = 2;
            this.cmbbx_partNumbers.TextChanged += new System.EventHandler(this.txtChanged_cmbBx_PartNum);
            // 
            // lbl_partNumbers
            // 
            this.lbl_partNumbers.AutoSize = true;
            this.lbl_partNumbers.Location = new System.Drawing.Point(12, 76);
            this.lbl_partNumbers.Name = "lbl_partNumbers";
            this.lbl_partNumbers.Size = new System.Drawing.Size(95, 17);
            this.lbl_partNumbers.TabIndex = 3;
            this.lbl_partNumbers.Text = "Part Numbers";
            // 
            // lbl_Description
            // 
            this.lbl_Description.AutoSize = true;
            this.lbl_Description.Location = new System.Drawing.Point(261, 78);
            this.lbl_Description.Name = "lbl_Description";
            this.lbl_Description.Size = new System.Drawing.Size(79, 17);
            this.lbl_Description.TabIndex = 4;
            this.lbl_Description.Text = "Description";
            // 
            // txtBx_Description
            // 
            this.txtBx_Description.Location = new System.Drawing.Point(264, 98);
            this.txtBx_Description.Name = "txtBx_Description";
            this.txtBx_Description.Size = new System.Drawing.Size(380, 22);
            this.txtBx_Description.TabIndex = 5;
            // 
            // btn_Search
            // 
            this.btn_Search.Location = new System.Drawing.Point(264, 27);
            this.btn_Search.Name = "btn_Search";
            this.btn_Search.Size = new System.Drawing.Size(118, 34);
            this.btn_Search.TabIndex = 6;
            this.btn_Search.Text = "Search";
            this.btn_Search.UseVisualStyleBackColor = true;
            this.btn_Search.Click += new System.EventHandler(this.click_btn_Search);
            // 
            // btn_Add
            // 
            this.btn_Add.Location = new System.Drawing.Point(264, 139);
            this.btn_Add.Name = "btn_Add";
            this.btn_Add.Size = new System.Drawing.Size(118, 34);
            this.btn_Add.TabIndex = 7;
            this.btn_Add.Text = "Add Part";
            this.btn_Add.UseVisualStyleBackColor = true;
            this.btn_Add.Click += new System.EventHandler(this.click_btn_AddPart);
            // 
            // Footprint_Lookup_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(656, 186);
            this.Controls.Add(this.btn_Add);
            this.Controls.Add(this.btn_Search);
            this.Controls.Add(this.txtBx_Description);
            this.Controls.Add(this.lbl_Description);
            this.Controls.Add(this.lbl_partNumbers);
            this.Controls.Add(this.cmbbx_partNumbers);
            this.Controls.Add(this.lbl_partNumSearch);
            this.Controls.Add(this.txtbx_partNumSearch);
            this.Name = "Footprint_Lookup_Form";
            this.Text = "Footprint Lookup";
            this.Load += new System.EventHandler(this.Footprint_Lookup_Form_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtbx_partNumSearch;
        private System.Windows.Forms.Label lbl_partNumSearch;
        private System.Windows.Forms.ComboBox cmbbx_partNumbers;
        private System.Windows.Forms.Label lbl_partNumbers;
        private System.Windows.Forms.Label lbl_Description;
        private System.Windows.Forms.TextBox txtBx_Description;
        private System.Windows.Forms.Button btn_Search;
        private System.Windows.Forms.Button btn_Add;
    }
}