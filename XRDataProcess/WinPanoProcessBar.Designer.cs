namespace XRDataProcess
{
    partial class WinPanoProcessBar
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinPanoProcessBar));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.loginfo = new System.Windows.Forms.TextBox();
            this.bar_main = new System.Windows.Forms.ProgressBar();
            this.textcnt = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.bar_pano = new System.Windows.Forms.ProgressBar();
            this.timer_start = new System.Windows.Forms.Timer(this.components);
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
            this.tableLayoutPanel1.Controls.Add(this.loginfo, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.bar_main, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.textcnt, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.bar_pano, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(485, 319);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 47);
            this.label1.TabIndex = 0;
            this.label1.Text = "总进度";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // loginfo
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.loginfo, 3);
            this.loginfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.loginfo.Location = new System.Drawing.Point(3, 97);
            this.loginfo.Multiline = true;
            this.loginfo.Name = "loginfo";
            this.loginfo.ReadOnly = true;
            this.loginfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.loginfo.Size = new System.Drawing.Size(479, 219);
            this.loginfo.TabIndex = 5;
            // 
            // bar_main
            // 
            this.bar_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_main.Location = new System.Drawing.Point(100, 3);
            this.bar_main.Maximum = 10000;
            this.bar_main.Name = "bar_main";
            this.bar_main.Size = new System.Drawing.Size(333, 41);
            this.bar_main.TabIndex = 6;
            // 
            // textcnt
            // 
            this.textcnt.AutoSize = true;
            this.textcnt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textcnt.Location = new System.Drawing.Point(439, 0);
            this.textcnt.Name = "textcnt";
            this.textcnt.Size = new System.Drawing.Size(43, 47);
            this.textcnt.TabIndex = 11;
            this.textcnt.Text = "(0/0)";
            this.textcnt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(3, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 47);
            this.label2.TabIndex = 12;
            this.label2.Text = "全景拼接进度";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // bar_pano
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.bar_pano, 2);
            this.bar_pano.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bar_pano.Location = new System.Drawing.Point(100, 50);
            this.bar_pano.MarqueeAnimationSpeed = 10000;
            this.bar_pano.Name = "bar_pano";
            this.bar_pano.Size = new System.Drawing.Size(382, 41);
            this.bar_pano.TabIndex = 13;
            // 
            // WinPanoProcessBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(485, 319);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "WinPanoProcessBar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "数据处理进度";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox loginfo;
        private System.Windows.Forms.ProgressBar bar_main;
        private System.Windows.Forms.Timer timer_start;
        private System.Windows.Forms.Label textcnt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar bar_pano;
    }
}