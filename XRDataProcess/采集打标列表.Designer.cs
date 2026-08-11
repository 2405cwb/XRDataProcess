namespace XRDataProcess
{
    partial class 采集打标列表
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(采集打标列表));
            this.dataGridView_Mark = new System.Windows.Forms.DataGridView();
            this.Column_Mile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Mark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tbox_dmi = new System.Windows.Forms.TextBox();
            this.tbox_mile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_add = new System.Windows.Forms.Button();
            this.button_deleteMark = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.tbox_mark = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button_deleteCali = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Mark)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView_Mark
            // 
            this.dataGridView_Mark.AllowUserToAddRows = false;
            this.dataGridView_Mark.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_Mark.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_Mark.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column_Mile,
            this.Column_Type,
            this.Column_Mark});
            this.dataGridView_Mark.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_Mark.Location = new System.Drawing.Point(3, 107);
            this.dataGridView_Mark.MultiSelect = true;
            this.dataGridView_Mark.Name = "dataGridView_Mark";
            this.dataGridView_Mark.ReadOnly = true;
            this.dataGridView_Mark.RowHeadersVisible = false;
            this.dataGridView_Mark.RowHeadersWidth = 51;
            this.dataGridView_Mark.RowTemplate.Height = 23;
            this.dataGridView_Mark.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_Mark.Size = new System.Drawing.Size(348, 145);
            this.dataGridView_Mark.TabIndex = 0;
            this.dataGridView_Mark.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_Mark_CellContentClick);
            this.dataGridView_Mark.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView_Mark_CellMouseDoubleClick);
            this.dataGridView_Mark.Sorted += new System.EventHandler(this.dataGridView_Mark_Sorted);
            // 
            // Column_Mile
            // 
            this.Column_Mile.HeaderText = "桩号";
            this.Column_Mile.MinimumWidth = 6;
            this.Column_Mile.Name = "Column_Mile";
            this.Column_Mile.ReadOnly = true;
            // 
            // Column_Type
            // 
            this.Column_Type.HeaderText = "类型";
            this.Column_Type.MinimumWidth = 6;
            this.Column_Type.Name = "Column_Type";
            this.Column_Type.ReadOnly = true;
            // 
            // Column_Mark
            // 
            this.Column_Mark.HeaderText = "标记";
            this.Column_Mark.MinimumWidth = 6;
            this.Column_Mark.Name = "Column_Mark";
            this.Column_Mark.ReadOnly = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView_Mark, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.groupBox1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 104F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(354, 255);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tableLayoutPanel2);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(0);
            this.groupBox1.Size = new System.Drawing.Size(354, 104);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "划分报表单元";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 4;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayoutPanel2.Controls.Add(this.tbox_dmi, 3, 0);
            this.tableLayoutPanel2.Controls.Add(this.tbox_mile, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.button_add, 2, 1);
            this.tableLayoutPanel2.Controls.Add(this.label3, 2, 0);
            this.tableLayoutPanel2.Controls.Add(this.button1, 2, 2);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.tbox_mark, 1, 2);
            this.tableLayoutPanel2.Controls.Add(this.label6, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.comboBox1, 1, 1);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 14);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 3;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.34F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(354, 90);
            this.tableLayoutPanel2.TabIndex = 113;
            // 
            // tbox_dmi
            // 
            this.tbox_dmi.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbox_dmi.Location = new System.Drawing.Point(248, 6);
            this.tbox_dmi.Margin = new System.Windows.Forms.Padding(2, 6, 2, 6);
            this.tbox_dmi.Name = "tbox_dmi";
            this.tbox_dmi.Size = new System.Drawing.Size(104, 21);
            this.tbox_dmi.TabIndex = 113;
            // 
            // tbox_mile
            // 
            this.tbox_mile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbox_mile.Location = new System.Drawing.Point(72, 6);
            this.tbox_mile.Margin = new System.Windows.Forms.Padding(2, 6, 2, 6);
            this.tbox_mile.Name = "tbox_mile";
            this.tbox_mile.Size = new System.Drawing.Size(102, 21);
            this.tbox_mile.TabIndex = 112;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(2, 8);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 8, 2, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "桩号：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // button_add
            // 
            this.button_add.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_add.Location = new System.Drawing.Point(176, 29);
            this.button_add.Margin = new System.Windows.Forms.Padding(0);
            this.button_add.Name = "button_add";
            this.button_add.Size = new System.Drawing.Size(88, 29);
            this.button_add.TabIndex = 110;
            this.button_add.Text = "添加";
            this.button_add.UseVisualStyleBackColor = true;
            this.button_add.Click += new System.EventHandler(this.button_add_Click);
            //
            // button_deleteMark
            //
            this.button_deleteMark.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_deleteMark.Location = new System.Drawing.Point(265, 29);
            this.button_deleteMark.Margin = new System.Windows.Forms.Padding(0);
            this.button_deleteMark.Name = "button_deleteMark";
            this.button_deleteMark.Size = new System.Drawing.Size(89, 29);
            this.button_deleteMark.TabIndex = 117;
            this.button_deleteMark.Text = "删除所选";
            this.button_deleteMark.UseVisualStyleBackColor = true;
            this.button_deleteMark.Click += new System.EventHandler(this.button_deleteMark_Click);
            this.tableLayoutPanel2.Controls.Add(this.button_deleteMark, 3, 1);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Location = new System.Drawing.Point(178, 8);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 8, 2, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 111;
            this.label3.Text = "里程：";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // button1
            // 
            this.tableLayoutPanel2.SetColumnSpan(this.button1, 2);
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(176, 58);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(178, 32);
            this.button1.TabIndex = 114;
            this.button1.Text = "批量导入打标";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(2, 66);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 8, 2, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 16);
            this.label2.TabIndex = 108;
            this.label2.Text = "标记：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // tbox_mark
            // 
            this.tbox_mark.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbox_mark.Location = new System.Drawing.Point(72, 64);
            this.tbox_mark.Margin = new System.Windows.Forms.Padding(2, 6, 2, 6);
            this.tbox_mark.Name = "tbox_mark";
            this.tbox_mark.Size = new System.Drawing.Size(102, 21);
            this.tbox_mark.TabIndex = 109;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Location = new System.Drawing.Point(3, 29);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 29);
            this.label6.TabIndex = 115;
            this.label6.Text = "类型：";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBox1
            // 
            this.comboBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "路面单元",
            "路面情况"});
            this.comboBox1.Location = new System.Drawing.Point(73, 32);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(100, 20);
            this.comboBox1.TabIndex = 116;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(368, 287);
            this.tabControl1.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tableLayoutPanel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(360, 261);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "打标";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tableLayoutPanel3);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(360, 261);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "校桩";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 4;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.Controls.Add(this.dataGridView1, 0, 3);
            this.tableLayoutPanel3.Controls.Add(this.label4, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.label5, 2, 0);
            this.tableLayoutPanel3.Controls.Add(this.textBox1, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.textBox2, 3, 0);
            this.tableLayoutPanel3.Controls.Add(this.button3, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.button4, 2, 1);
            this.tableLayoutPanel3.Controls.Add(this.button_deleteCali, 0, 2);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 4;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(354, 255);
            this.tableLayoutPanel3.TabIndex = 0;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.tableLayoutPanel3.SetColumnSpan(this.dataGridView1, 4);
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 103);
            this.dataGridView1.MultiSelect = true;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(348, 149);
            this.dataGridView1.TabIndex = 6;
            this.dataGridView1.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_CellMouseDoubleClick);
            this.dataGridView1.Sorted += new System.EventHandler(this.dataGridView1_Sorted);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "桩号";
            this.dataGridViewTextBoxColumn1.MinimumWidth = 6;
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "里程";
            this.dataGridViewTextBoxColumn2.MinimumWidth = 6;
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Location = new System.Drawing.Point(3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 36);
            this.label4.TabIndex = 0;
            this.label4.Text = "桩号：";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Location = new System.Drawing.Point(180, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 36);
            this.label5.TabIndex = 1;
            this.label5.Text = "里程：";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // textBox1
            // 
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Location = new System.Drawing.Point(52, 6);
            this.textBox1.Margin = new System.Windows.Forms.Padding(2, 6, 2, 6);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(123, 21);
            this.textBox1.TabIndex = 2;
            // 
            // textBox2
            // 
            this.textBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox2.Location = new System.Drawing.Point(229, 6);
            this.textBox2.Margin = new System.Windows.Forms.Padding(2, 6, 2, 6);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(123, 21);
            this.textBox2.TabIndex = 3;
            // 
            // button3
            // 
            this.tableLayoutPanel3.SetColumnSpan(this.button3, 2);
            this.button3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button3.Location = new System.Drawing.Point(0, 36);
            this.button3.Margin = new System.Windows.Forms.Padding(0);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(177, 36);
            this.button3.TabIndex = 4;
            this.button3.Text = "添加";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.tableLayoutPanel3.SetColumnSpan(this.button4, 2);
            this.button4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button4.Location = new System.Drawing.Point(177, 36);
            this.button4.Margin = new System.Windows.Forms.Padding(0);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(177, 36);
            this.button4.TabIndex = 5;
            this.button4.Text = "重载工程";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            //
            // button_deleteCali
            //
            this.tableLayoutPanel3.SetColumnSpan(this.button_deleteCali, 4);
            this.button_deleteCali.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_deleteCali.Location = new System.Drawing.Point(0, 72);
            this.button_deleteCali.Margin = new System.Windows.Forms.Padding(0);
            this.button_deleteCali.Name = "button_deleteCali";
            this.button_deleteCali.Size = new System.Drawing.Size(354, 28);
            this.button_deleteCali.TabIndex = 7;
            this.button_deleteCali.Text = "删除所选校桩";
            this.button_deleteCali.UseVisualStyleBackColor = true;
            this.button_deleteCali.Click += new System.EventHandler(this.button_deleteCali_Click);
            // 
            // 采集打标列表
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 287);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "采集打标列表";
            this.Text = "WinRoadMark";
            this.Load += new System.EventHandler(this.采集打标列表_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Mark)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView_Mark;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Mile;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Mark;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button_add;
        private System.Windows.Forms.Button button_deleteMark;
        private System.Windows.Forms.TextBox tbox_mark;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TextBox tbox_dmi;
        private System.Windows.Forms.TextBox tbox_mile;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button_deleteCali;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}
