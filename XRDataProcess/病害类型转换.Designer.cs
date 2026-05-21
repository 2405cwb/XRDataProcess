namespace XRDataProcess
{
    partial class 病害类型转换
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(病害类型转换));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button_TranDis = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.textBox_log = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.comboBox_Src = new System.Windows.Forms.ComboBox();
            this.comboBox_Dest = new System.Windows.Forms.ComboBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 30);
            this.label1.TabIndex = 0;
            this.label1.Text = "病害框图时规范";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(3, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 30);
            this.label2.TabIndex = 2;
            this.label2.Text = "要转换出表规范";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button_TranDis
            // 
            this.button_TranDis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_TranDis.Location = new System.Drawing.Point(255, 3);
            this.button_TranDis.Name = "button_TranDis";
            this.tableLayoutPanel1.SetRowSpan(this.button_TranDis, 2);
            this.button_TranDis.Size = new System.Drawing.Size(163, 54);
            this.button_TranDis.TabIndex = 4;
            this.button_TranDis.Text = "病害类型转换";
            this.button_TranDis.UseVisualStyleBackColor = true;
            this.button_TranDis.Click += new System.EventHandler(this.button_TranDis_Click);
            // 
            // progressBar1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.progressBar1, 3);
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar1.Location = new System.Drawing.Point(3, 63);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(415, 24);
            this.progressBar1.TabIndex = 5;
            // 
            // textBox_log
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.textBox_log, 3);
            this.textBox_log.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox_log.Location = new System.Drawing.Point(3, 93);
            this.textBox_log.Multiline = true;
            this.textBox_log.Name = "textBox_log";
            this.textBox_log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_log.Size = new System.Drawing.Size(415, 198);
            this.textBox_log.TabIndex = 6;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40F));
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBox_log, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.progressBar1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_TranDis, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.comboBox_Src, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.comboBox_Dest, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(421, 294);
            this.tableLayoutPanel1.TabIndex = 7;
            // 
            // comboBox_Src
            // 
            this.comboBox_Src.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox_Src.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Src.FormattingEnabled = true;
            this.comboBox_Src.Items.AddRange(new object[] {
            "等级公路2018"});
            this.comboBox_Src.Location = new System.Drawing.Point(108, 5);
            this.comboBox_Src.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.comboBox_Src.Name = "comboBox_Src";
            this.comboBox_Src.Size = new System.Drawing.Size(141, 20);
            this.comboBox_Src.TabIndex = 7;
            // 
            // comboBox_Dest
            // 
            this.comboBox_Dest.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox_Dest.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Dest.FormattingEnabled = true;
            this.comboBox_Dest.Items.AddRange(new object[] {
            "北京农村公路",
            "辽宁农村公路"});
            this.comboBox_Dest.Location = new System.Drawing.Point(108, 35);
            this.comboBox_Dest.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.comboBox_Dest.Name = "comboBox_Dest";
            this.comboBox_Dest.Size = new System.Drawing.Size(141, 20);
            this.comboBox_Dest.TabIndex = 8;
            // 
            // 病害类型转换
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(421, 294);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "病害类型转换";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "病害类型转换";
            this.Load += new System.EventHandler(this.病害类型转换_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_TranDis;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox textBox_log;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ComboBox comboBox_Src;
        private System.Windows.Forms.ComboBox comboBox_Dest;
    }
}