namespace XRDataProcess
{
    partial class 统计单元长度
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_Yes = new System.Windows.Forms.Button();
            this.button_No = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox_DetectYear = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox_DetectNum = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox_DistrictCode = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.button_Yes, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.button_No, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.comboBox1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.comboBox2, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.comboBox_DetectYear, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.label6, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.comboBox_DetectNum, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.label7, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.textBox_DistrictCode, 1, 4);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 6;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(341, 214);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // button_Yes
            // 
            this.button_Yes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Yes.Location = new System.Drawing.Point(3, 163);
            this.button_Yes.Name = "button_Yes";
            this.button_Yes.Size = new System.Drawing.Size(164, 48);
            this.button_Yes.TabIndex = 0;
            this.button_Yes.Text = "确定";
            this.button_Yes.UseVisualStyleBackColor = true;
            this.button_Yes.Click += new System.EventHandler(this.button_Yes_Click);
            // 
            // button_No
            // 
            this.button_No.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_No.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_No.Location = new System.Drawing.Point(173, 163);
            this.button_No.Name = "button_No";
            this.button_No.Size = new System.Drawing.Size(165, 48);
            this.button_No.TabIndex = 1;
            this.button_No.Text = "取消";
            this.button_No.UseVisualStyleBackColor = true;
            this.button_No.Click += new System.EventHandler(this.button_No_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(164, 32);
            this.label1.TabIndex = 2;
            this.label1.Text = "病害统计单元长度(m)";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBox1
            // 
            this.comboBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "100",
            "1000"});
            this.comboBox1.Location = new System.Drawing.Point(173, 5);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(165, 20);
            this.comboBox1.TabIndex = 3;
            this.comboBox1.Text = "100";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(3, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 32);
            this.label2.TabIndex = 4;
            this.label2.Text = "评价指标单元长度(m)";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBox2
            // 
            this.comboBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "100",
            "1000"});
            this.comboBox2.Location = new System.Drawing.Point(173, 37);
            this.comboBox2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(165, 20);
            this.comboBox2.TabIndex = 5;
            this.comboBox2.Text = "100";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Location = new System.Drawing.Point(3, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(164, 32);
            this.label3.TabIndex = 41;
            this.label3.Tag = "-1";
            this.label3.Text = "检测任务年份";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBox_DetectYear
            // 
            this.comboBox_DetectYear.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox_DetectYear.FormattingEnabled = true;
            this.comboBox_DetectYear.Items.AddRange(new object[] {
            "2018",
            "2019",
            "2020",
            "2021",
            "2022",
            "2023",
            "2024",
            "2025",
            "2026",
            "2027",
            "2028",
            "2029"});
            this.comboBox_DetectYear.Location = new System.Drawing.Point(173, 67);
            this.comboBox_DetectYear.Name = "comboBox_DetectYear";
            this.comboBox_DetectYear.Size = new System.Drawing.Size(165, 20);
            this.comboBox_DetectYear.TabIndex = 42;
            this.comboBox_DetectYear.Tag = "-1";
            this.comboBox_DetectYear.Text = "2021";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Location = new System.Drawing.Point(3, 96);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(164, 32);
            this.label6.TabIndex = 43;
            this.label6.Tag = "-1";
            this.label6.Text = "检测次数";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBox_DetectNum
            // 
            this.comboBox_DetectNum.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox_DetectNum.FormattingEnabled = true;
            this.comboBox_DetectNum.Items.AddRange(new object[] {
            "1-一次",
            "2-二次",
            "3-三次",
            "4-四次",
            "5-五次"});
            this.comboBox_DetectNum.Location = new System.Drawing.Point(173, 99);
            this.comboBox_DetectNum.Name = "comboBox_DetectNum";
            this.comboBox_DetectNum.Size = new System.Drawing.Size(165, 20);
            this.comboBox_DetectNum.TabIndex = 44;
            this.comboBox_DetectNum.Tag = "-1";
            this.comboBox_DetectNum.Text = "1-一次";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label7.Location = new System.Drawing.Point(3, 128);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(164, 32);
            this.label7.TabIndex = 45;
            this.label7.Tag = "-1";
            this.label7.Text = "行政区代码";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBox_DistrictCode
            // 
            this.textBox_DistrictCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_DistrictCode.Location = new System.Drawing.Point(173, 131);
            this.textBox_DistrictCode.Name = "textBox_DistrictCode";
            this.textBox_DistrictCode.Size = new System.Drawing.Size(165, 21);
            this.textBox_DistrictCode.TabIndex = 46;
            this.textBox_DistrictCode.Tag = "-1";
            this.textBox_DistrictCode.Text = "620123";
            // 
            // 统计单元长度
            // 
            this.AcceptButton = this.button_Yes;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button_No;
            this.ClientSize = new System.Drawing.Size(341, 214);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "统计单元长度";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "统计单元长度";
            this.Load += new System.EventHandler(this.统计单元长度_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button_Yes;
        private System.Windows.Forms.Button button_No;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox_DetectYear;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox_DetectNum;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox_DistrictCode;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox comboBox1;
    }
}