namespace XRDataProcess
{
    partial class WinRoadDisList
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinRoadDisList));
            this.dataGridView_Dislist = new System.Windows.Forms.DataGridView();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Mile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Mark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_update = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView_curdis = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dataGridView_MLDisImgList = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dataGridView_MLNoDisImgList = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox_DisName = new System.Windows.Forms.TextBox();
            this.comboBox_BSType = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Dislist)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_curdis)).BeginInit();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_MLDisImgList)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_MLNoDisImgList)).BeginInit();
            this.tableLayoutPanel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView_Dislist
            // 
            this.dataGridView_Dislist.AllowUserToAddRows = false;
            this.dataGridView_Dislist.AllowUserToDeleteRows = false;
            this.dataGridView_Dislist.AllowUserToResizeRows = false;
            this.dataGridView_Dislist.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_Dislist.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_Dislist.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column3,
            this.Column_Mile,
            this.Column_Type,
            this.Column_Mark,
            this.Column1,
            this.Column2});
            this.dataGridView_Dislist.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_Dislist.Location = new System.Drawing.Point(3, 32);
            this.dataGridView_Dislist.MultiSelect = false;
            this.dataGridView_Dislist.Name = "dataGridView_Dislist";
            this.dataGridView_Dislist.ReadOnly = true;
            this.dataGridView_Dislist.RowHeadersVisible = false;
            this.dataGridView_Dislist.RowTemplate.Height = 23;
            this.dataGridView_Dislist.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_Dislist.Size = new System.Drawing.Size(520, 165);
            this.dataGridView_Dislist.TabIndex = 1;
            this.dataGridView_Dislist.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_Dislist_CellDoubleClick);
            // 
            // Column3
            // 
            this.Column3.HeaderText = "编号";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            // 
            // Column_Mile
            // 
            this.Column_Mile.HeaderText = "桩号";
            this.Column_Mile.Name = "Column_Mile";
            this.Column_Mile.ReadOnly = true;
            // 
            // Column_Type
            // 
            this.Column_Type.HeaderText = "类型";
            this.Column_Type.Name = "Column_Type";
            this.Column_Type.ReadOnly = true;
            // 
            // Column_Mark
            // 
            this.Column_Mark.HeaderText = "长度(m)";
            this.Column_Mark.Name = "Column_Mark";
            this.Column_Mark.ReadOnly = true;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "宽度(m)";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "面积(m2)";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView_Dislist, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_update, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(526, 200);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // button_update
            // 
            this.button_update.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_update.Location = new System.Drawing.Point(3, 3);
            this.button_update.Name = "button_update";
            this.button_update.Size = new System.Drawing.Size(520, 23);
            this.button_update.TabIndex = 2;
            this.button_update.Text = "刷新病害列表";
            this.button_update.UseVisualStyleBackColor = true;
            this.button_update.Click += new System.EventHandler(this.button_update_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 30);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(540, 232);
            this.tabControl1.TabIndex = 3;
            //this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.tableLayoutPanel3);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(532, 206);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "当前病害列表";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Controls.Add(this.button1, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.dataGridView_curdis, 0, 1);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(532, 206);
            this.tableLayoutPanel3.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(3, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(526, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "更新病害";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView_curdis
            // 
            this.dataGridView_curdis.AllowUserToAddRows = false;
            this.dataGridView_curdis.AllowUserToDeleteRows = false;
            this.dataGridView_curdis.AllowUserToResizeRows = false;
            this.dataGridView_curdis.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_curdis.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_curdis.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9});
            this.dataGridView_curdis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_curdis.Location = new System.Drawing.Point(3, 32);
            this.dataGridView_curdis.MultiSelect = false;
            this.dataGridView_curdis.Name = "dataGridView_curdis";
            this.dataGridView_curdis.ReadOnly = true;
            this.dataGridView_curdis.RowHeadersVisible = false;
            this.dataGridView_curdis.RowTemplate.Height = 23;
            this.dataGridView_curdis.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_curdis.Size = new System.Drawing.Size(526, 171);
            this.dataGridView_curdis.TabIndex = 3;
            this.dataGridView_curdis.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_curdis_CellDoubleClick);
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.HeaderText = "桩号";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.HeaderText = "类型";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.HeaderText = "长度(m)";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.HeaderText = "宽度(m)";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            this.dataGridViewTextBoxColumn8.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.HeaderText = "面积(m2)";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.ReadOnly = true;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tableLayoutPanel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(532, 206);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "工程病害列表";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dataGridView_MLDisImgList);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(532, 206);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "自动识别有病害图像列表";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataGridView_MLDisImgList
            // 
            this.dataGridView_MLDisImgList.AllowUserToAddRows = false;
            this.dataGridView_MLDisImgList.AllowUserToDeleteRows = false;
            this.dataGridView_MLDisImgList.AllowUserToResizeRows = false;
            this.dataGridView_MLDisImgList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_MLDisImgList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_MLDisImgList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.dataGridView_MLDisImgList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_MLDisImgList.Location = new System.Drawing.Point(3, 3);
            this.dataGridView_MLDisImgList.MultiSelect = false;
            this.dataGridView_MLDisImgList.Name = "dataGridView_MLDisImgList";
            this.dataGridView_MLDisImgList.ReadOnly = true;
            this.dataGridView_MLDisImgList.RowHeadersVisible = false;
            this.dataGridView_MLDisImgList.RowTemplate.Height = 23;
            this.dataGridView_MLDisImgList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_MLDisImgList.Size = new System.Drawing.Size(526, 200);
            this.dataGridView_MLDisImgList.TabIndex = 1;
            this.dataGridView_MLDisImgList.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_Dislist_CellDoubleClick);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.FillWeight = 25.3807F;
            this.dataGridViewTextBoxColumn1.HeaderText = "桩号";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.FillWeight = 174.6192F;
            this.dataGridViewTextBoxColumn2.HeaderText = "图像名";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dataGridView_MLNoDisImgList);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(532, 206);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "自动识别无病害图像列表";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dataGridView_MLNoDisImgList
            // 
            this.dataGridView_MLNoDisImgList.AllowUserToAddRows = false;
            this.dataGridView_MLNoDisImgList.AllowUserToDeleteRows = false;
            this.dataGridView_MLNoDisImgList.AllowUserToResizeRows = false;
            this.dataGridView_MLNoDisImgList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_MLNoDisImgList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_MLNoDisImgList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7});
            this.dataGridView_MLNoDisImgList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_MLNoDisImgList.Location = new System.Drawing.Point(3, 3);
            this.dataGridView_MLNoDisImgList.MultiSelect = false;
            this.dataGridView_MLNoDisImgList.Name = "dataGridView_MLNoDisImgList";
            this.dataGridView_MLNoDisImgList.ReadOnly = true;
            this.dataGridView_MLNoDisImgList.RowHeadersVisible = false;
            this.dataGridView_MLNoDisImgList.RowTemplate.Height = 23;
            this.dataGridView_MLNoDisImgList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_MLNoDisImgList.Size = new System.Drawing.Size(526, 200);
            this.dataGridView_MLNoDisImgList.TabIndex = 1;
            this.dataGridView_MLNoDisImgList.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_Dislist_CellDoubleClick);
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.FillWeight = 25.38071F;
            this.dataGridViewTextBoxColumn6.HeaderText = "桩号";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.FillWeight = 174.6193F;
            this.dataGridViewTextBoxColumn7.HeaderText = "图像名";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.ReadOnly = true;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.tabControl1, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(540, 262);
            this.tableLayoutPanel2.TabIndex = 4;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textBox_DisName);
            this.panel1.Controls.Add(this.comboBox_BSType);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(540, 30);
            this.panel1.TabIndex = 4;
            // 
            // textBox_DisName
            // 
            this.textBox_DisName.Location = new System.Drawing.Point(151, 3);
            this.textBox_DisName.Name = "textBox_DisName";
            this.textBox_DisName.Size = new System.Drawing.Size(100, 21);
            this.textBox_DisName.TabIndex = 2;
            this.textBox_DisName.Text = "路框差";
            // 
            // comboBox_BSType
            // 
            this.comboBox_BSType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_BSType.FormattingEnabled = true;
            this.comboBox_BSType.Items.AddRange(new object[] {
            "连续浏览",
            "仅浏览自动识别有病害图像",
            "仅浏览自动识别无病害图像",
            "仅浏览指定病害图像"});
            this.comboBox_BSType.Location = new System.Drawing.Point(3, 3);
            this.comboBox_BSType.Name = "comboBox_BSType";
            this.comboBox_BSType.Size = new System.Drawing.Size(142, 20);
            this.comboBox_BSType.TabIndex = 1;
            this.comboBox_BSType.SelectedIndexChanged += new System.EventHandler(this.comboBox_BSType_SelectedIndexChanged);
            // 
            // WinRoadDisList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(540, 262);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "WinRoadDisList";
            this.Text = "WinRoadDisList";
            this.Load += new System.EventHandler(this.WinRoadDisList_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Dislist)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_curdis)).EndInit();
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_MLDisImgList)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_MLNoDisImgList)).EndInit();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView_Dislist;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button_update;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comboBox_BSType;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dataGridView_MLDisImgList;
        private System.Windows.Forms.DataGridView dataGridView_MLNoDisImgList;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView dataGridView_curdis;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.TextBox textBox_DisName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Mile;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Mark;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
    }
}