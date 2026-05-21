namespace XRDataProcess
{
    partial class SelectIRM
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SelectIRM));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.chk_MPD = new System.Windows.Forms.CheckBox();
            this.chk_IRI = new System.Windows.Forms.CheckBox();
            this.chk_MTD = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox_Thresh1 = new System.Windows.Forms.TextBox();
            this.textBox_Thresh0 = new System.Windows.Forms.TextBox();
            this.chk_LasFilter = new System.Windows.Forms.CheckBox();
            this.chk_Geoalig = new System.Windows.Forms.CheckBox();
            this.chk_RUT = new System.Windows.Forms.CheckBox();
            this.button_confirm = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.checkBox_mohao = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.tableLayoutPanel1);
            this.groupBox1.Location = new System.Drawing.Point(18, 18);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Size = new System.Drawing.Size(548, 362);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "选择项";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 39.83287F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.16713F));
            this.tableLayoutPanel1.Controls.Add(this.chk_MPD, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.chk_IRI, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.chk_MTD, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.chk_Geoalig, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.chk_RUT, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.checkBox_mohao, 1, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(4, 25);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(540, 333);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // chk_MPD
            // 
            this.chk_MPD.AutoSize = true;
            this.chk_MPD.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chk_MPD.Location = new System.Drawing.Point(4, 202);
            this.chk_MPD.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_MPD.Name = "chk_MPD";
            this.chk_MPD.Size = new System.Drawing.Size(207, 58);
            this.chk_MPD.TabIndex = 6;
            this.chk_MPD.Text = "计算构造深度MPD";
            this.chk_MPD.UseVisualStyleBackColor = true;
            // 
            // chk_IRI
            // 
            this.chk_IRI.AutoSize = true;
            this.chk_IRI.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chk_IRI.Location = new System.Drawing.Point(4, 4);
            this.chk_IRI.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_IRI.Name = "chk_IRI";
            this.chk_IRI.Size = new System.Drawing.Size(207, 58);
            this.chk_IRI.TabIndex = 3;
            this.chk_IRI.Text = "计算/清除平整度IRI";
            this.chk_IRI.UseVisualStyleBackColor = true;
            this.chk_IRI.CheckedChanged += new System.EventHandler(this.chk_IRI_CheckedChanged);
            // 
            // chk_MTD
            // 
            this.chk_MTD.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.chk_MTD, 2);
            this.chk_MTD.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chk_MTD.Location = new System.Drawing.Point(4, 136);
            this.chk_MTD.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_MTD.Name = "chk_MTD";
            this.chk_MTD.Size = new System.Drawing.Size(532, 58);
            this.chk_MTD.TabIndex = 5;
            this.chk_MTD.Text = "计算构造深度SMTD";
            this.chk_MTD.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textBox_Thresh1);
            this.panel1.Controls.Add(this.textBox_Thresh0);
            this.panel1.Controls.Add(this.chk_LasFilter);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(219, 4);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(317, 58);
            this.panel1.TabIndex = 7;
            // 
            // textBox_Thresh1
            // 
            this.textBox_Thresh1.Location = new System.Drawing.Point(225, 16);
            this.textBox_Thresh1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBox_Thresh1.Name = "textBox_Thresh1";
            this.textBox_Thresh1.Size = new System.Drawing.Size(61, 28);
            this.textBox_Thresh1.TabIndex = 2;
            this.textBox_Thresh1.Text = "20";
            // 
            // textBox_Thresh0
            // 
            this.textBox_Thresh0.Location = new System.Drawing.Point(156, 16);
            this.textBox_Thresh0.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBox_Thresh0.Name = "textBox_Thresh0";
            this.textBox_Thresh0.Size = new System.Drawing.Size(61, 28);
            this.textBox_Thresh0.TabIndex = 1;
            this.textBox_Thresh0.Text = "5";
            // 
            // chk_LasFilter
            // 
            this.chk_LasFilter.AutoSize = true;
            this.chk_LasFilter.Location = new System.Drawing.Point(4, 20);
            this.chk_LasFilter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_LasFilter.Name = "chk_LasFilter";
            this.chk_LasFilter.Size = new System.Drawing.Size(133, 22);
            this.chk_LasFilter.TabIndex = 0;
            this.chk_LasFilter.Text = "IRI激光去噪";
            this.chk_LasFilter.UseVisualStyleBackColor = true;
            // 
            // chk_Geoalig
            // 
            this.chk_Geoalig.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.chk_Geoalig, 2);
            this.chk_Geoalig.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chk_Geoalig.Location = new System.Drawing.Point(4, 268);
            this.chk_Geoalig.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_Geoalig.Name = "chk_Geoalig";
            this.chk_Geoalig.Size = new System.Drawing.Size(532, 61);
            this.chk_Geoalig.TabIndex = 8;
            this.chk_Geoalig.Text = "计算几何线形";
            this.chk_Geoalig.UseVisualStyleBackColor = true;
            // 
            // chk_RUT
            // 
            this.chk_RUT.AutoSize = true;
            this.chk_RUT.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chk_RUT.Location = new System.Drawing.Point(4, 70);
            this.chk_RUT.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chk_RUT.Name = "chk_RUT";
            this.chk_RUT.Size = new System.Drawing.Size(207, 58);
            this.chk_RUT.TabIndex = 9;
            this.chk_RUT.Text = "计算车辙RUT";
            this.chk_RUT.UseVisualStyleBackColor = true;
            // 
            // button_confirm
            // 
            this.button_confirm.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_confirm.Location = new System.Drawing.Point(58, 400);
            this.button_confirm.Margin = new System.Windows.Forms.Padding(0);
            this.button_confirm.Name = "button_confirm";
            this.button_confirm.Size = new System.Drawing.Size(159, 54);
            this.button_confirm.TabIndex = 13;
            this.button_confirm.Text = "确定";
            this.button_confirm.UseVisualStyleBackColor = true;
            this.button_confirm.Click += new System.EventHandler(this.button_confirm_Click);
            // 
            // button_cancel
            // 
            this.button_cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_cancel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_cancel.Location = new System.Drawing.Point(322, 400);
            this.button_cancel.Margin = new System.Windows.Forms.Padding(0);
            this.button_cancel.Name = "button_cancel";
            this.button_cancel.Size = new System.Drawing.Size(159, 54);
            this.button_cancel.TabIndex = 14;
            this.button_cancel.Text = "取消";
            this.button_cancel.UseVisualStyleBackColor = true;
            this.button_cancel.Click += new System.EventHandler(this.button_cancel_Click);
            // 
            // checkBox_mohao
            // 
            this.checkBox_mohao.AutoSize = true;
            this.checkBox_mohao.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkBox_mohao.Location = new System.Drawing.Point(218, 201);
            this.checkBox_mohao.Name = "checkBox_mohao";
            this.checkBox_mohao.Size = new System.Drawing.Size(319, 60);
            this.checkBox_mohao.TabIndex = 10;
            this.checkBox_mohao.Text = "导出磨耗原始数据";
            this.checkBox_mohao.UseVisualStyleBackColor = true;
            // 
            // SelectIRM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(572, 468);
            this.Controls.Add(this.button_cancel);
            this.Controls.Add(this.button_confirm);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SelectIRM";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "选择IRM";
            this.Load += new System.EventHandler(this.SelectIRM_Load);
            this.groupBox1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button_confirm;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.CheckBox chk_IRI;
        private System.Windows.Forms.CheckBox chk_MTD;
        private System.Windows.Forms.Button button_cancel;
        private System.Windows.Forms.CheckBox chk_MPD;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox_Thresh1;
        private System.Windows.Forms.TextBox textBox_Thresh0;
        private System.Windows.Forms.CheckBox chk_LasFilter;
        private System.Windows.Forms.CheckBox chk_Geoalig;
        private System.Windows.Forms.CheckBox chk_RUT;
        private System.Windows.Forms.CheckBox checkBox_mohao;
    }
}