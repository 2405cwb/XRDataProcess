namespace XRDataProcess
{
    partial class YGView
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.m_n_dbd = new System.Windows.Forms.NumericUpDown();
            this.m_b_zm = new System.Windows.Forms.CheckBox();
            this.m_n_ld = new System.Windows.Forms.NumericUpDown();
            this.m_b_YG = new System.Windows.Forms.CheckBox();
            this.标签_1 = new System.Windows.Forms.Label();
            this.标签_2 = new System.Windows.Forms.Label();
            this.按钮_照明分析 = new System.Windows.Forms.Button();
            this.标签_3 = new System.Windows.Forms.Label();
            this.标签_4 = new System.Windows.Forms.Label();
            this.锐化半径 = new System.Windows.Forms.NumericUpDown();
            this.锐化强度 = new System.Windows.Forms.NumericUpDown();
            this.锐化 = new System.Windows.Forms.CheckBox();
            this.m_pic = new System.Windows.Forms.PictureBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.m_n_dbd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_n_ld)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.锐化半径)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.锐化强度)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_pic)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // m_n_dbd
            // 
            this.m_n_dbd.Dock = System.Windows.Forms.DockStyle.Fill;
            this.m_n_dbd.Location = new System.Drawing.Point(117, 360);
            this.m_n_dbd.Margin = new System.Windows.Forms.Padding(0);
            this.m_n_dbd.Minimum = new decimal(new int[] {
            100,
            0,
            0,
            -2147483648});
            this.m_n_dbd.Name = "m_n_dbd";
            this.m_n_dbd.Size = new System.Drawing.Size(118, 21);
            this.m_n_dbd.TabIndex = 1;
            this.m_n_dbd.ValueChanged += new System.EventHandler(this.ParaValueChanged);
            // 
            // m_b_zm
            // 
            this.m_b_zm.AutoSize = true;
            this.m_b_zm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.m_b_zm.Location = new System.Drawing.Point(120, 303);
            this.m_b_zm.Name = "m_b_zm";
            this.m_b_zm.Padding = new System.Windows.Forms.Padding(3);
            this.m_b_zm.Size = new System.Drawing.Size(112, 24);
            this.m_b_zm.TabIndex = 0;
            this.m_b_zm.Text = "纵向灰度校正";
            this.m_b_zm.UseVisualStyleBackColor = true;
            this.m_b_zm.CheckedChanged += new System.EventHandler(this.ModeChanged);
            // 
            // m_n_ld
            // 
            this.m_n_ld.Dock = System.Windows.Forms.DockStyle.Fill;
            this.m_n_ld.Location = new System.Drawing.Point(0, 360);
            this.m_n_ld.Margin = new System.Windows.Forms.Padding(0);
            this.m_n_ld.Minimum = new decimal(new int[] {
            100,
            0,
            0,
            -2147483648});
            this.m_n_ld.Name = "m_n_ld";
            this.m_n_ld.Size = new System.Drawing.Size(117, 21);
            this.m_n_ld.TabIndex = 0;
            this.m_n_ld.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.m_n_ld.ValueChanged += new System.EventHandler(this.ParaValueChanged);
            // 
            // m_b_YG
            // 
            this.m_b_YG.AutoSize = true;
            this.m_b_YG.Dock = System.Windows.Forms.DockStyle.Fill;
            this.m_b_YG.Location = new System.Drawing.Point(3, 303);
            this.m_b_YG.Name = "m_b_YG";
            this.m_b_YG.Padding = new System.Windows.Forms.Padding(3);
            this.m_b_YG.Size = new System.Drawing.Size(111, 24);
            this.m_b_YG.TabIndex = 0;
            this.m_b_YG.Text = "全局灰度校正";
            this.m_b_YG.UseVisualStyleBackColor = true;
            this.m_b_YG.CheckedChanged += new System.EventHandler(this.ModeChanged);
            // 
            // 标签_1
            // 
            this.标签_1.AutoSize = true;
            this.标签_1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.标签_1.Location = new System.Drawing.Point(3, 333);
            this.标签_1.Margin = new System.Windows.Forms.Padding(3);
            this.标签_1.Name = "标签_1";
            this.标签_1.Size = new System.Drawing.Size(111, 24);
            this.标签_1.TabIndex = 2;
            this.标签_1.Text = "亮度";
            this.标签_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // 标签_2
            // 
            this.标签_2.AutoSize = true;
            this.标签_2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.标签_2.Location = new System.Drawing.Point(120, 333);
            this.标签_2.Margin = new System.Windows.Forms.Padding(3);
            this.标签_2.Name = "标签_2";
            this.标签_2.Size = new System.Drawing.Size(112, 24);
            this.标签_2.TabIndex = 3;
            this.标签_2.Text = "对比度";
            this.标签_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // 按钮_照明分析
            // 
            this.按钮_照明分析.Dock = System.Windows.Forms.DockStyle.Fill;
            this.按钮_照明分析.Location = new System.Drawing.Point(117, 450);
            this.按钮_照明分析.Margin = new System.Windows.Forms.Padding(0);
            this.按钮_照明分析.Name = "按钮_照明分析";
            this.按钮_照明分析.Size = new System.Drawing.Size(118, 30);
            this.按钮_照明分析.TabIndex = 4;
            this.按钮_照明分析.Text = "统计原始灰度";
            this.按钮_照明分析.UseVisualStyleBackColor = true;
            this.按钮_照明分析.Click += new System.EventHandler(this.按钮_照明分析_Click);
            // 
            // 标签_3
            // 
            this.标签_3.AutoSize = true;
            this.标签_3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.标签_3.Location = new System.Drawing.Point(3, 393);
            this.标签_3.Margin = new System.Windows.Forms.Padding(3);
            this.标签_3.Name = "标签_3";
            this.标签_3.Size = new System.Drawing.Size(111, 24);
            this.标签_3.TabIndex = 5;
            this.标签_3.Text = "锐化半径";
            this.标签_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // 标签_4
            // 
            this.标签_4.AutoSize = true;
            this.标签_4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.标签_4.Location = new System.Drawing.Point(120, 393);
            this.标签_4.Margin = new System.Windows.Forms.Padding(3);
            this.标签_4.Name = "标签_4";
            this.标签_4.Size = new System.Drawing.Size(112, 24);
            this.标签_4.TabIndex = 6;
            this.标签_4.Text = "锐化强度";
            this.标签_4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // 锐化半径
            // 
            this.锐化半径.Dock = System.Windows.Forms.DockStyle.Fill;
            this.锐化半径.Location = new System.Drawing.Point(0, 420);
            this.锐化半径.Margin = new System.Windows.Forms.Padding(0);
            this.锐化半径.Name = "锐化半径";
            this.锐化半径.Size = new System.Drawing.Size(117, 21);
            this.锐化半径.TabIndex = 7;
            this.锐化半径.ValueChanged += new System.EventHandler(this.锐化_ValueChanged);
            // 
            // 锐化强度
            // 
            this.锐化强度.Dock = System.Windows.Forms.DockStyle.Fill;
            this.锐化强度.Location = new System.Drawing.Point(117, 420);
            this.锐化强度.Margin = new System.Windows.Forms.Padding(0);
            this.锐化强度.Name = "锐化强度";
            this.锐化强度.Size = new System.Drawing.Size(118, 21);
            this.锐化强度.TabIndex = 8;
            this.锐化强度.ValueChanged += new System.EventHandler(this.锐化_ValueChanged);
            // 
            // 锐化
            // 
            this.锐化.AutoSize = true;
            this.锐化.Dock = System.Windows.Forms.DockStyle.Fill;
            this.锐化.Location = new System.Drawing.Point(0, 450);
            this.锐化.Margin = new System.Windows.Forms.Padding(0);
            this.锐化.Name = "锐化";
            this.锐化.Size = new System.Drawing.Size(117, 30);
            this.锐化.TabIndex = 9;
            this.锐化.Text = "蜕化";
            this.锐化.UseVisualStyleBackColor = true;
            this.锐化.CheckedChanged += new System.EventHandler(this.锐化_ValueChanged);
            // 
            // m_pic
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.m_pic, 2);
            this.m_pic.Dock = System.Windows.Forms.DockStyle.Fill;
            this.m_pic.Location = new System.Drawing.Point(3, 3);
            this.m_pic.Name = "m_pic";
            this.m_pic.Padding = new System.Windows.Forms.Padding(3);
            this.m_pic.Size = new System.Drawing.Size(229, 294);
            this.m_pic.TabIndex = 0;
            this.m_pic.TabStop = false;
            this.m_pic.Paint += new System.Windows.Forms.PaintEventHandler(this.m_pic_Paint);
            this.m_pic.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_pic_MouseDown);
            this.m_pic.MouseUp += new System.Windows.Forms.MouseEventHandler(this.m_pic_MouseUp);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.m_n_dbd, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.锐化强度, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.锐化半径, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.标签_4, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.标签_3, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.m_b_YG, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.m_n_ld, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.m_b_zm, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.标签_2, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.标签_1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.锐化, 0, 6);
            this.tableLayoutPanel1.Controls.Add(this.按钮_照明分析, 1, 6);
            this.tableLayoutPanel1.Controls.Add(this.m_pic, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 7;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(235, 480);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // YGView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "YGView";
            this.Size = new System.Drawing.Size(235, 480);
            ((System.ComponentModel.ISupportInitialize)(this.m_n_dbd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_n_ld)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.锐化半径)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.锐化强度)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_pic)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NumericUpDown m_n_dbd;
        private System.Windows.Forms.CheckBox m_b_zm;
        private System.Windows.Forms.NumericUpDown m_n_ld;
        private System.Windows.Forms.CheckBox m_b_YG;
        private System.Windows.Forms.Label 标签_1;
        private System.Windows.Forms.Label 标签_2;
        private System.Windows.Forms.Button 按钮_照明分析;
        private System.Windows.Forms.PictureBox m_pic;
        private System.Windows.Forms.Label 标签_3;
        private System.Windows.Forms.Label 标签_4;
        private System.Windows.Forms.NumericUpDown 锐化半径;
        private System.Windows.Forms.NumericUpDown 锐化强度;
        private System.Windows.Forms.CheckBox 锐化;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
    }
}
