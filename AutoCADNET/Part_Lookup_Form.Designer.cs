namespace Part_Lookup_Dialogs
{
    partial class Part_Lookup_Form
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
            this.cb_PDFSelect = new System.Windows.Forms.ComboBox();
            this.lbl_Displaying = new System.Windows.Forms.Label();
            this.wb_PDF = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // cb_PDFSelect
            // 
            this.cb_PDFSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cb_PDFSelect.FormattingEnabled = true;
            this.cb_PDFSelect.Location = new System.Drawing.Point(420, 14);
            this.cb_PDFSelect.Name = "cb_PDFSelect";
            this.cb_PDFSelect.Size = new System.Drawing.Size(273, 24);
            this.cb_PDFSelect.TabIndex = 2;
            this.cb_PDFSelect.TextChanged += new System.EventHandler(this.TextChanged_cb_PDFSelect);
            // 
            // lbl_Displaying
            // 
            this.lbl_Displaying.AutoSize = true;
            this.lbl_Displaying.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Displaying.Location = new System.Drawing.Point(12, 12);
            this.lbl_Displaying.Name = "lbl_Displaying";
            this.lbl_Displaying.Size = new System.Drawing.Size(182, 25);
            this.lbl_Displaying.TabIndex = 1;
            this.lbl_Displaying.Text = "Select PDF to view:";
            // 
            // wb_PDF
            // 
            this.wb_PDF.Location = new System.Drawing.Point(465, 14);
            this.wb_PDF.MinimumSize = new System.Drawing.Size(20, 20);
            this.wb_PDF.Name = "wb_PDF";
            this.wb_PDF.Size = new System.Drawing.Size(174, 20);
            this.wb_PDF.TabIndex = 3;
            this.wb_PDF.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.Navigated_wb_PDF);
            this.wb_PDF.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.Navigating_wb_PDF);
            // 
            // Part_Lookup_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 51);
            this.Controls.Add(this.cb_PDFSelect);
            this.Controls.Add(this.wb_PDF);
            this.Controls.Add(this.lbl_Displaying);
            this.Name = "Part_Lookup_Form";
            this.Text = "VFD Lookup";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox cb_PDFSelect;
        private System.Windows.Forms.Label lbl_Displaying;
        private System.Windows.Forms.WebBrowser wb_PDF;
    }
}