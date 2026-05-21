namespace XRDataProcess
{
    partial class WinProcessBar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinProcessBar));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.loginfo = new System.Windows.Forms.TextBox();
            this.bar_main = new System.Windows.Forms.ProgressBar();
            this.bar_iri = new System.Windows.Forms.ProgressBar();
            this.bar_rut = new System.Windows.Forms.ProgressBar();
            this.bar_mtd = new System.Windows.Forms.ProgressBar();
            this.textcnt = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.bar_mpd = new System.Windows.Forms.ProgressBar();
            this.label6 = new System.Windows.Forms.Label();
            this.bar_geoalig = new System.Windows.Forms.ProgressBar();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.label5, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.loginfo, 0, 6);
            this.tableLayoutPanel1.Controls.Add(this.bar_main, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.bar_iri, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.bar_rut, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.bar_mtd, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.textcnt, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.bar_mpd, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.label6, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.bar_geoalig, 1, 5);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 7;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 34F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(569, 328);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 36);
            this.label1.TabIndex = 0;
            this.label1.Text = "总进度";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Location = new System.Drawing.Point(3, 36);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 36);
            this.label3.TabIndex = 2;
            this.label3.Text = "计算IRI进度";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Location = new System.Drawing.Point(3, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(107, 36);
            this.label4.TabIndex = 3;
            this.label4.Text = "计算Rut进度";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Location = new System.Drawing.Point(3, 108);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 36);
            this.label5.TabIndex = 4;
            this.label5.Text = "计算MTD进度";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // loginfo
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.loginfo, 3);
            this.loginfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.loginfo.Location = new System.Drawing.Point(3, 219);
            this.loginfo.Multiline = true;
            this.loginfo.Name = "loginfo";
            this.loginfo.ReadOnly = true;
            this.loginfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.loginfo.Size = new System.Drawing.Size(563, 107);
            this.loginfo.TabIndex = 5;
            // 
            // bar_main
            // 
            this.bar_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_main.Location = new System.Drawing.Point(116, 3);
            this.bar_main.Maximum = 10000;
            this.bar_main.Name = "bar_main";
            this.bar_main.Size = new System.Drawing.Size(392, 30);
            this.bar_main.TabIndex = 6;
            // 
            // bar_iri
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.bar_iri, 2);
            this.bar_iri.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_iri.Location = new System.Drawing.Point(116, 39);
            this.bar_iri.Maximum = 10000;
            this.bar_iri.Name = "bar_iri";
            this.bar_iri.Size = new System.Drawing.Size(450, 30);
            this.bar_iri.TabIndex = 8;
            // 
            // bar_rut
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.bar_rut, 2);
            this.bar_rut.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_rut.Location = new System.Drawing.Point(116, 75);
            this.bar_rut.Maximum = 10000;
            this.bar_rut.Name = "bar_rut";
            this.bar_rut.Size = new System.Drawing.Size(450, 30);
            this.bar_rut.TabIndex = 9;
            // 
            // bar_mtd
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.bar_mtd, 2);
            this.bar_mtd.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_mtd.Location = new System.Drawing.Point(116, 111);
            this.bar_mtd.Maximum = 10000;
            this.bar_mtd.Name = "bar_mtd";
            this.bar_mtd.Size = new System.Drawing.Size(450, 30);
            this.bar_mtd.TabIndex = 10;
            // 
            // textcnt
            // 
            this.textcnt.AutoSize = true;
            this.textcnt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textcnt.Location = new System.Drawing.Point(514, 0);
            this.textcnt.Name = "textcnt";
            this.textcnt.Size = new System.Drawing.Size(52, 36);
            this.textcnt.TabIndex = 11;
            this.textcnt.Text = "(0/0)";
            this.textcnt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(3, 144);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 36);
            this.label2.TabIndex = 12;
            this.label2.Text = "计算MPD进度";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // bar_mpd
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.bar_mpd, 2);
            this.bar_mpd.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_mpd.Location = new System.Drawing.Point(116, 147);
            this.bar_mpd.Maximum = 10000;
            this.bar_mpd.Name = "bar_mpd";
            this.bar_mpd.Size = new System.Drawing.Size(450, 30);
            this.bar_mpd.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Location = new System.Drawing.Point(3, 180);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(107, 36);
            this.label6.TabIndex = 14;
            this.label6.Text = "计算几何线形进度";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // bar_geoalig
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.bar_geoalig, 2);
            this.bar_geoalig.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_geoalig.Location = new System.Drawing.Point(116, 183);
            this.bar_geoalig.Maximum = 10000;
            this.bar_geoalig.Name = "bar_geoalig";
            this.bar_geoalig.Size = new System.Drawing.Size(450, 30);
            this.bar_geoalig.TabIndex = 15;
            // 
            // WinProcessBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 328);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "WinProcessBar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "数据处理进度";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox loginfo;
        private System.Windows.Forms.ProgressBar bar_main;
        private System.Windows.Forms.ProgressBar bar_iri;
        private System.Windows.Forms.ProgressBar bar_rut;
        private System.Windows.Forms.ProgressBar bar_mtd;
        private System.Windows.Forms.Label textcnt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar bar_mpd;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ProgressBar bar_geoalig;
    }
}