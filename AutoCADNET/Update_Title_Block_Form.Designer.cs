namespace Update_Title_Block_Dialogs
{
    partial class Update_Title_Block_Form
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
            this.btn_Execute = new System.Windows.Forms.Button();
            this.dgv_TitleBlocks = new System.Windows.Forms.DataGridView();
            this.dgv_revInfo = new System.Windows.Forms.DataGridView();
            this.lbl_dgv_TitleBlocks_Title = new System.Windows.Forms.Label();
            this.lbl_dgv_revInfo_Title = new System.Windows.Forms.Label();
            this.tb_DrawnBy = new System.Windows.Forms.TextBox();
            this.lbl_DrawnBy = new System.Windows.Forms.Label();
            this.lbl_CheckedBy = new System.Windows.Forms.Label();
            this.tb_CheckedBy = new System.Windows.Forms.TextBox();
            this.lbl_Stamp = new System.Windows.Forms.Label();
            this.cb_Stamp = new System.Windows.Forms.ComboBox();
            this.lbl_revNum = new System.Windows.Forms.Label();
            this.tb_revNum = new System.Windows.Forms.TextBox();
            this.chckB_DrawnBy = new System.Windows.Forms.CheckBox();
            this.chckB_CheckedBy = new System.Windows.Forms.CheckBox();
            this.chckB_revNum = new System.Windows.Forms.CheckBox();
            this.chckB_Stamp = new System.Windows.Forms.CheckBox();
            this.btn_ApplytoAll = new System.Windows.Forms.Button();
            this.lbl_Use = new System.Windows.Forms.Label();
            this.lbl_effect1 = new System.Windows.Forms.Label();
            this.tb_revTitle = new System.Windows.Forms.TextBox();
            this.lbl_revTitle = new System.Windows.Forms.Label();
            this.lbl_revDate = new System.Windows.Forms.Label();
            this.tb_revDate = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.rb_first = new System.Windows.Forms.RadioButton();
            this.rb_last = new System.Windows.Forms.RadioButton();
            this.rb_firstEmpty = new System.Windows.Forms.RadioButton();
            this.btn_ApplyToAll_revInfo = new System.Windows.Forms.Button();
            this.btn_Clear_ALLrevInfo = new System.Windows.Forms.Button();
            this.btn_Clear_SelectedRevInfo = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_TitleBlocks)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_revInfo)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_Execute
            // 
            this.btn_Execute.Location = new System.Drawing.Point(1101, 375);
            this.btn_Execute.Name = "btn_Execute";
            this.btn_Execute.Size = new System.Drawing.Size(156, 48);
            this.btn_Execute.TabIndex = 2;
            this.btn_Execute.Text = "Execute";
            this.btn_Execute.UseVisualStyleBackColor = true;
            this.btn_Execute.Click += new System.EventHandler(this.btn_Execute_Click);
            // 
            // dgv_TitleBlocks
            // 
            this.dgv_TitleBlocks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_TitleBlocks.Location = new System.Drawing.Point(12, 29);
            this.dgv_TitleBlocks.Name = "dgv_TitleBlocks";
            this.dgv_TitleBlocks.RowTemplate.Height = 24;
            this.dgv_TitleBlocks.Size = new System.Drawing.Size(829, 285);
            this.dgv_TitleBlocks.TabIndex = 15;
            this.dgv_TitleBlocks.SelectionChanged += new System.EventHandler(this.selectionChanged_dgv_TitleBlocks);
            // 
            // dgv_revInfo
            // 
            this.dgv_revInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_revInfo.Location = new System.Drawing.Point(847, 29);
            this.dgv_revInfo.Name = "dgv_revInfo";
            this.dgv_revInfo.RowTemplate.Height = 24;
            this.dgv_revInfo.Size = new System.Drawing.Size(409, 285);
            this.dgv_revInfo.TabIndex = 16;
            this.dgv_revInfo.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.cellValueChanged_dgv_revInfo);
            this.dgv_revInfo.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dgv_revInfo_RowPostPaint);
            this.dgv_revInfo.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.dgv_revInfo_PreviewKeyDown);
            // 
            // lbl_dgv_TitleBlocks_Title
            // 
            this.lbl_dgv_TitleBlocks_Title.AutoSize = true;
            this.lbl_dgv_TitleBlocks_Title.Location = new System.Drawing.Point(12, 9);
            this.lbl_dgv_TitleBlocks_Title.Name = "lbl_dgv_TitleBlocks_Title";
            this.lbl_dgv_TitleBlocks_Title.Size = new System.Drawing.Size(107, 17);
            this.lbl_dgv_TitleBlocks_Title.TabIndex = 17;
            this.lbl_dgv_TitleBlocks_Title.Text = "Title Block Data";
            // 
            // lbl_dgv_revInfo_Title
            // 
            this.lbl_dgv_revInfo_Title.AutoSize = true;
            this.lbl_dgv_revInfo_Title.Location = new System.Drawing.Point(844, 9);
            this.lbl_dgv_revInfo_Title.Name = "lbl_dgv_revInfo_Title";
            this.lbl_dgv_revInfo_Title.Size = new System.Drawing.Size(288, 17);
            this.lbl_dgv_revInfo_Title.TabIndex = 18;
            this.lbl_dgv_revInfo_Title.Text = "Select title block to view revision information:";
            // 
            // tb_DrawnBy
            // 
            this.tb_DrawnBy.Location = new System.Drawing.Point(12, 342);
            this.tb_DrawnBy.Name = "tb_DrawnBy";
            this.tb_DrawnBy.Size = new System.Drawing.Size(102, 22);
            this.tb_DrawnBy.TabIndex = 19;
            // 
            // lbl_DrawnBy
            // 
            this.lbl_DrawnBy.AutoSize = true;
            this.lbl_DrawnBy.Location = new System.Drawing.Point(12, 322);
            this.lbl_DrawnBy.Name = "lbl_DrawnBy";
            this.lbl_DrawnBy.Size = new System.Drawing.Size(71, 17);
            this.lbl_DrawnBy.TabIndex = 20;
            this.lbl_DrawnBy.Text = "Drawn by:";
            // 
            // lbl_CheckedBy
            // 
            this.lbl_CheckedBy.AutoSize = true;
            this.lbl_CheckedBy.Location = new System.Drawing.Point(139, 322);
            this.lbl_CheckedBy.Name = "lbl_CheckedBy";
            this.lbl_CheckedBy.Size = new System.Drawing.Size(86, 17);
            this.lbl_CheckedBy.TabIndex = 22;
            this.lbl_CheckedBy.Text = "Checked by:";
            // 
            // tb_CheckedBy
            // 
            this.tb_CheckedBy.Location = new System.Drawing.Point(139, 342);
            this.tb_CheckedBy.Name = "tb_CheckedBy";
            this.tb_CheckedBy.Size = new System.Drawing.Size(102, 22);
            this.tb_CheckedBy.TabIndex = 21;
            // 
            // lbl_Stamp
            // 
            this.lbl_Stamp.AutoSize = true;
            this.lbl_Stamp.Location = new System.Drawing.Point(266, 322);
            this.lbl_Stamp.Name = "lbl_Stamp";
            this.lbl_Stamp.Size = new System.Drawing.Size(52, 17);
            this.lbl_Stamp.TabIndex = 24;
            this.lbl_Stamp.Text = "Stamp:";
            // 
            // cb_Stamp
            // 
            this.cb_Stamp.FormattingEnabled = true;
            this.cb_Stamp.Location = new System.Drawing.Point(269, 340);
            this.cb_Stamp.Name = "cb_Stamp";
            this.cb_Stamp.Size = new System.Drawing.Size(225, 24);
            this.cb_Stamp.TabIndex = 25;
            // 
            // lbl_revNum
            // 
            this.lbl_revNum.AutoSize = true;
            this.lbl_revNum.Location = new System.Drawing.Point(521, 320);
            this.lbl_revNum.Name = "lbl_revNum";
            this.lbl_revNum.Size = new System.Drawing.Size(118, 17);
            this.lbl_revNum.TabIndex = 27;
            this.lbl_revNum.Text = "Revision number:";
            // 
            // tb_revNum
            // 
            this.tb_revNum.Location = new System.Drawing.Point(521, 340);
            this.tb_revNum.Name = "tb_revNum";
            this.tb_revNum.Size = new System.Drawing.Size(102, 22);
            this.tb_revNum.TabIndex = 26;
            // 
            // chckB_DrawnBy
            // 
            this.chckB_DrawnBy.AutoSize = true;
            this.chckB_DrawnBy.Location = new System.Drawing.Point(108, 375);
            this.chckB_DrawnBy.Name = "chckB_DrawnBy";
            this.chckB_DrawnBy.Size = new System.Drawing.Size(90, 21);
            this.chckB_DrawnBy.TabIndex = 28;
            this.chckB_DrawnBy.Text = "Drawn By";
            this.chckB_DrawnBy.UseVisualStyleBackColor = true;
            this.chckB_DrawnBy.CheckedChanged += new System.EventHandler(this.chckB_DrawnBy_CheckedChanged);
            // 
            // chckB_CheckedBy
            // 
            this.chckB_CheckedBy.AutoSize = true;
            this.chckB_CheckedBy.Location = new System.Drawing.Point(108, 402);
            this.chckB_CheckedBy.Name = "chckB_CheckedBy";
            this.chckB_CheckedBy.Size = new System.Drawing.Size(105, 21);
            this.chckB_CheckedBy.TabIndex = 29;
            this.chckB_CheckedBy.Text = "Checked By";
            this.chckB_CheckedBy.UseVisualStyleBackColor = true;
            this.chckB_CheckedBy.CheckedChanged += new System.EventHandler(this.chckB_CheckedBy_CheckedChanged);
            // 
            // chckB_revNum
            // 
            this.chckB_revNum.AutoSize = true;
            this.chckB_revNum.Location = new System.Drawing.Point(238, 402);
            this.chckB_revNum.Name = "chckB_revNum";
            this.chckB_revNum.Size = new System.Drawing.Size(138, 21);
            this.chckB_revNum.TabIndex = 30;
            this.chckB_revNum.Text = "Revision Number";
            this.chckB_revNum.UseVisualStyleBackColor = true;
            this.chckB_revNum.CheckStateChanged += new System.EventHandler(this.chckB_revNum_CheckedChanged);
            // 
            // chckB_Stamp
            // 
            this.chckB_Stamp.AutoSize = true;
            this.chckB_Stamp.Location = new System.Drawing.Point(238, 375);
            this.chckB_Stamp.Name = "chckB_Stamp";
            this.chckB_Stamp.Size = new System.Drawing.Size(70, 21);
            this.chckB_Stamp.TabIndex = 31;
            this.chckB_Stamp.Text = "Stamp";
            this.chckB_Stamp.UseVisualStyleBackColor = true;
            this.chckB_Stamp.CheckedChanged += new System.EventHandler(this.chckB_Stamp_CheckedChanged);
            // 
            // btn_ApplytoAll
            // 
            this.btn_ApplytoAll.Location = new System.Drawing.Point(400, 390);
            this.btn_ApplytoAll.Name = "btn_ApplytoAll";
            this.btn_ApplytoAll.Size = new System.Drawing.Size(109, 27);
            this.btn_ApplytoAll.TabIndex = 32;
            this.btn_ApplytoAll.Text = "Apply to All";
            this.btn_ApplytoAll.UseVisualStyleBackColor = true;
            this.btn_ApplytoAll.Click += new System.EventHandler(this.btn_ApplytoAll_Click);
            // 
            // lbl_Use
            // 
            this.lbl_Use.AutoSize = true;
            this.lbl_Use.Location = new System.Drawing.Point(42, 390);
            this.lbl_Use.Name = "lbl_Use";
            this.lbl_Use.Size = new System.Drawing.Size(41, 17);
            this.lbl_Use.TabIndex = 33;
            this.lbl_Use.Text = "Use?";
            // 
            // lbl_effect1
            // 
            this.lbl_effect1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lbl_effect1.Location = new System.Drawing.Point(648, 321);
            this.lbl_effect1.Name = "lbl_effect1";
            this.lbl_effect1.Size = new System.Drawing.Size(2, 108);
            this.lbl_effect1.TabIndex = 34;
            // 
            // tb_revTitle
            // 
            this.tb_revTitle.Location = new System.Drawing.Point(672, 340);
            this.tb_revTitle.Name = "tb_revTitle";
            this.tb_revTitle.Size = new System.Drawing.Size(102, 22);
            this.tb_revTitle.TabIndex = 35;
            // 
            // lbl_revTitle
            // 
            this.lbl_revTitle.AutoSize = true;
            this.lbl_revTitle.Location = new System.Drawing.Point(669, 320);
            this.lbl_revTitle.Name = "lbl_revTitle";
            this.lbl_revTitle.Size = new System.Drawing.Size(92, 17);
            this.lbl_revTitle.TabIndex = 36;
            this.lbl_revTitle.Text = "Revision title:";
            // 
            // lbl_revDate
            // 
            this.lbl_revDate.AutoSize = true;
            this.lbl_revDate.Location = new System.Drawing.Point(793, 320);
            this.lbl_revDate.Name = "lbl_revDate";
            this.lbl_revDate.Size = new System.Drawing.Size(98, 17);
            this.lbl_revDate.TabIndex = 38;
            this.lbl_revDate.Text = "Revision date:";
            // 
            // tb_revDate
            // 
            this.tb_revDate.Location = new System.Drawing.Point(796, 340);
            this.tb_revDate.Name = "tb_revDate";
            this.tb_revDate.Size = new System.Drawing.Size(102, 22);
            this.tb_revDate.TabIndex = 37;
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(922, 320);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(2, 108);
            this.label1.TabIndex = 39;
            // 
            // rb_first
            // 
            this.rb_first.AutoSize = true;
            this.rb_first.Location = new System.Drawing.Point(707, 369);
            this.rb_first.Name = "rb_first";
            this.rb_first.Size = new System.Drawing.Size(56, 21);
            this.rb_first.TabIndex = 40;
            this.rb_first.Text = "First";
            this.rb_first.UseVisualStyleBackColor = true;
            this.rb_first.CheckedChanged += new System.EventHandler(this.rb_first_CheckedChanged);
            // 
            // rb_last
            // 
            this.rb_last.AutoSize = true;
            this.rb_last.Location = new System.Drawing.Point(707, 393);
            this.rb_last.Name = "rb_last";
            this.rb_last.Size = new System.Drawing.Size(56, 21);
            this.rb_last.TabIndex = 41;
            this.rb_last.Text = "Last";
            this.rb_last.UseVisualStyleBackColor = true;
            this.rb_last.CheckedChanged += new System.EventHandler(this.rb_last_CheckedChanged);
            // 
            // rb_firstEmpty
            // 
            this.rb_firstEmpty.AutoSize = true;
            this.rb_firstEmpty.Checked = true;
            this.rb_firstEmpty.Location = new System.Drawing.Point(780, 369);
            this.rb_firstEmpty.Name = "rb_firstEmpty";
            this.rb_firstEmpty.Size = new System.Drawing.Size(99, 21);
            this.rb_firstEmpty.TabIndex = 42;
            this.rb_firstEmpty.TabStop = true;
            this.rb_firstEmpty.Text = "First Empty";
            this.rb_firstEmpty.UseVisualStyleBackColor = true;
            this.rb_firstEmpty.CheckedChanged += new System.EventHandler(this.rb_firstEmpty_CheckedChanged);
            // 
            // btn_ApplyToAll_revInfo
            // 
            this.btn_ApplyToAll_revInfo.Location = new System.Drawing.Point(780, 390);
            this.btn_ApplyToAll_revInfo.Name = "btn_ApplyToAll_revInfo";
            this.btn_ApplyToAll_revInfo.Size = new System.Drawing.Size(107, 27);
            this.btn_ApplyToAll_revInfo.TabIndex = 43;
            this.btn_ApplyToAll_revInfo.Text = "Apply to All";
            this.btn_ApplyToAll_revInfo.UseVisualStyleBackColor = true;
            this.btn_ApplyToAll_revInfo.Click += new System.EventHandler(this.btn_ApplyToAll_revInfo_Click);
            // 
            // btn_Clear_ALLrevInfo
            // 
            this.btn_Clear_ALLrevInfo.Location = new System.Drawing.Point(930, 346);
            this.btn_Clear_ALLrevInfo.Name = "btn_Clear_ALLrevInfo";
            this.btn_Clear_ALLrevInfo.Size = new System.Drawing.Size(169, 27);
            this.btn_Clear_ALLrevInfo.TabIndex = 44;
            this.btn_Clear_ALLrevInfo.Text = "Clear ALL Rev Info";
            this.btn_Clear_ALLrevInfo.UseVisualStyleBackColor = true;
            this.btn_Clear_ALLrevInfo.Click += new System.EventHandler(this.btn_Clear_ALLrevInfo_Click);
            // 
            // btn_Clear_SelectedRevInfo
            // 
            this.btn_Clear_SelectedRevInfo.Location = new System.Drawing.Point(930, 317);
            this.btn_Clear_SelectedRevInfo.Name = "btn_Clear_SelectedRevInfo";
            this.btn_Clear_SelectedRevInfo.Size = new System.Drawing.Size(169, 27);
            this.btn_Clear_SelectedRevInfo.TabIndex = 45;
            this.btn_Clear_SelectedRevInfo.Text = "Clear Selected Rev Info";
            this.btn_Clear_SelectedRevInfo.UseVisualStyleBackColor = true;
            this.btn_Clear_SelectedRevInfo.Click += new System.EventHandler(this.Click_btn_ClearSelectedRevInfo);
            // 
            // Update_Title_Block_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1269, 434);
            this.Controls.Add(this.btn_Clear_SelectedRevInfo);
            this.Controls.Add(this.btn_Clear_ALLrevInfo);
            this.Controls.Add(this.btn_ApplyToAll_revInfo);
            this.Controls.Add(this.rb_firstEmpty);
            this.Controls.Add(this.rb_last);
            this.Controls.Add(this.rb_first);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbl_revDate);
            this.Controls.Add(this.tb_revDate);
            this.Controls.Add(this.lbl_revTitle);
            this.Controls.Add(this.tb_revTitle);
            this.Controls.Add(this.lbl_effect1);
            this.Controls.Add(this.lbl_Use);
            this.Controls.Add(this.btn_ApplytoAll);
            this.Controls.Add(this.chckB_Stamp);
            this.Controls.Add(this.chckB_revNum);
            this.Controls.Add(this.chckB_CheckedBy);
            this.Controls.Add(this.chckB_DrawnBy);
            this.Controls.Add(this.lbl_revNum);
            this.Controls.Add(this.tb_revNum);
            this.Controls.Add(this.cb_Stamp);
            this.Controls.Add(this.lbl_Stamp);
            this.Controls.Add(this.lbl_CheckedBy);
            this.Controls.Add(this.tb_CheckedBy);
            this.Controls.Add(this.lbl_DrawnBy);
            this.Controls.Add(this.tb_DrawnBy);
            this.Controls.Add(this.lbl_dgv_revInfo_Title);
            this.Controls.Add(this.lbl_dgv_TitleBlocks_Title);
            this.Controls.Add(this.dgv_TitleBlocks);
            this.Controls.Add(this.btn_Execute);
            this.Controls.Add(this.dgv_revInfo);
            this.Name = "Update_Title_Block_Form";
            this.Text = "Update Title Blocks";
            this.Load += new System.EventHandler(this.Update_Title_Block_Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_TitleBlocks)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_revInfo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Execute;
        private System.Windows.Forms.DataGridView dgv_TitleBlocks;
        private System.Windows.Forms.DataGridView dgv_revInfo;
        private System.Windows.Forms.Label lbl_dgv_TitleBlocks_Title;
        private System.Windows.Forms.Label lbl_dgv_revInfo_Title;
        private System.Windows.Forms.TextBox tb_DrawnBy;
        private System.Windows.Forms.Label lbl_DrawnBy;
        private System.Windows.Forms.Label lbl_CheckedBy;
        private System.Windows.Forms.TextBox tb_CheckedBy;
        private System.Windows.Forms.Label lbl_Stamp;
        private System.Windows.Forms.ComboBox cb_Stamp;
        private System.Windows.Forms.Label lbl_revNum;
        private System.Windows.Forms.TextBox tb_revNum;
        private System.Windows.Forms.CheckBox chckB_DrawnBy;
        private System.Windows.Forms.CheckBox chckB_CheckedBy;
        private System.Windows.Forms.CheckBox chckB_revNum;
        private System.Windows.Forms.CheckBox chckB_Stamp;
        private System.Windows.Forms.Button btn_ApplytoAll;
        private System.Windows.Forms.Label lbl_Use;
        private System.Windows.Forms.Label lbl_effect1;
        private System.Windows.Forms.TextBox tb_revTitle;
        private System.Windows.Forms.Label lbl_revTitle;
        private System.Windows.Forms.Label lbl_revDate;
        private System.Windows.Forms.TextBox tb_revDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rb_first;
        private System.Windows.Forms.RadioButton rb_last;
        private System.Windows.Forms.RadioButton rb_firstEmpty;
        private System.Windows.Forms.Button btn_ApplyToAll_revInfo;
        private System.Windows.Forms.Button btn_Clear_ALLrevInfo;
        private System.Windows.Forms.Button btn_Clear_SelectedRevInfo;
    }
}