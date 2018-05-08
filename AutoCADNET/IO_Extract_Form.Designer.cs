namespace IO_Extract_Dialogs
{
    partial class IO_Extract_Form
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
            this.lv_Files = new System.Windows.Forms.ListView();
            this.btn_SelectDrawings = new System.Windows.Forms.Button();
            this.btn_RemoveSelected = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_Next = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.rb_Full = new System.Windows.Forms.RadioButton();
            this.rb_Smart = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // lv_Files
            // 
            this.lv_Files.Location = new System.Drawing.Point(12, 29);
            this.lv_Files.Name = "lv_Files";
            this.lv_Files.Size = new System.Drawing.Size(533, 275);
            this.lv_Files.TabIndex = 0;
            this.lv_Files.UseCompatibleStateImageBehavior = false;
            this.lv_Files.View = System.Windows.Forms.View.Details;
            this.lv_Files.SelectedIndexChanged += new System.EventHandler(this.selectedIndexChanged_Files);
            // 
            // btn_SelectDrawings
            // 
            this.btn_SelectDrawings.Location = new System.Drawing.Point(551, 29);
            this.btn_SelectDrawings.Name = "btn_SelectDrawings";
            this.btn_SelectDrawings.Size = new System.Drawing.Size(139, 29);
            this.btn_SelectDrawings.TabIndex = 1;
            this.btn_SelectDrawings.Text = "Add Drawings:";
            this.btn_SelectDrawings.UseVisualStyleBackColor = true;
            this.btn_SelectDrawings.Click += new System.EventHandler(this.click_AddDrawings);
            // 
            // btn_RemoveSelected
            // 
            this.btn_RemoveSelected.Location = new System.Drawing.Point(551, 97);
            this.btn_RemoveSelected.Name = "btn_RemoveSelected";
            this.btn_RemoveSelected.Size = new System.Drawing.Size(139, 29);
            this.btn_RemoveSelected.TabIndex = 2;
            this.btn_RemoveSelected.Text = "Remove Selected:";
            this.btn_RemoveSelected.UseVisualStyleBackColor = true;
            this.btn_RemoveSelected.Click += new System.EventHandler(this.click_RemoveDrawings);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(320, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Select the drawing files to include in IO extraction.";
            // 
            // btn_Next
            // 
            this.btn_Next.Location = new System.Drawing.Point(347, 310);
            this.btn_Next.Name = "btn_Next";
            this.btn_Next.Size = new System.Drawing.Size(96, 29);
            this.btn_Next.TabIndex = 4;
            this.btn_Next.Text = "Next >>";
            this.btn_Next.UseVisualStyleBackColor = true;
            this.btn_Next.Click += new System.EventHandler(this.click_Next);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(449, 310);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(96, 29);
            this.btn_Cancel.TabIndex = 5;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.click_Cancel);
            // 
            // rb_Full
            // 
            this.rb_Full.AutoSize = true;
            this.rb_Full.Location = new System.Drawing.Point(166, 310);
            this.rb_Full.Name = "rb_Full";
            this.rb_Full.Size = new System.Drawing.Size(143, 21);
            this.rb_Full.TabIndex = 6;
            this.rb_Full.TabStop = true;
            this.rb_Full.Text = "Full Search (Slow)";
            this.rb_Full.UseVisualStyleBackColor = true;
            // 
            // rb_Smart
            // 
            this.rb_Smart.AutoSize = true;
            this.rb_Smart.Location = new System.Drawing.Point(166, 333);
            this.rb_Smart.Name = "rb_Smart";
            this.rb_Smart.Size = new System.Drawing.Size(115, 21);
            this.rb_Smart.TabIndex = 7;
            this.rb_Smart.TabStop = true;
            this.rb_Smart.Text = "Smart Search";
            this.rb_Smart.UseVisualStyleBackColor = true;
            // 
            // IO_Extract_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(697, 370);
            this.Controls.Add(this.rb_Smart);
            this.Controls.Add(this.rb_Full);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_Next);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_RemoveSelected);
            this.Controls.Add(this.btn_SelectDrawings);
            this.Controls.Add(this.lv_Files);
            this.Name = "IO_Extract_Form";
            this.Text = "IO Extraction Utility";
            this.Load += new System.EventHandler(this.load_IOExtractionForm);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView lv_Files;
        private System.Windows.Forms.Button btn_SelectDrawings;
        private System.Windows.Forms.Button btn_RemoveSelected;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_Next;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.RadioButton rb_Full;
        private System.Windows.Forms.RadioButton rb_Smart;
    }
}