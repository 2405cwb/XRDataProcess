namespace XRDataProcess
{
    partial class WinPanoImg
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinPanoImg));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_mile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_dmi = new System.Windows.Forms.TextBox();
            this.button_jump = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.button_last = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.button_play = new System.Windows.Forms.Button();
            this.button_next = new System.Windows.Forms.Button();
            this.button_speedsub = new System.Windows.Forms.Button();
            this.button_speedadd = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.progressBar_per = new System.Windows.Forms.ProgressBar();
            this.label6 = new System.Windows.Forms.Label();
            this.label_ImgPath = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.toolTip_label = new System.Windows.Forms.ToolTip(this.components);
            this.timer_roadplay = new System.Windows.Forms.Timer(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.label_ImgPath, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1000, 577);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // flowLayoutPanel3
            // 
            this.flowLayoutPanel3.BackColor = System.Drawing.Color.Transparent;
            this.flowLayoutPanel3.Controls.Add(this.label1);
            this.flowLayoutPanel3.Controls.Add(this.textBox_mile);
            this.flowLayoutPanel3.Controls.Add(this.label2);
            this.flowLayoutPanel3.Controls.Add(this.textBox_dmi);
            this.flowLayoutPanel3.Controls.Add(this.button_jump);
            this.flowLayoutPanel3.Controls.Add(this.label3);
            this.flowLayoutPanel3.Controls.Add(this.button_last);
            this.flowLayoutPanel3.Controls.Add(this.button_play);
            this.flowLayoutPanel3.Controls.Add(this.button_next);
            this.flowLayoutPanel3.Controls.Add(this.button_speedsub);
            this.flowLayoutPanel3.Controls.Add(this.button_speedadd);
            this.flowLayoutPanel3.Controls.Add(this.label5);
            this.flowLayoutPanel3.Controls.Add(this.progressBar_per);
            this.flowLayoutPanel3.Controls.Add(this.label6);
            this.flowLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.flowLayoutPanel3.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(994, 24);
            this.flowLayoutPanel3.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 4);
            this.label1.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 12);
            this.label1.TabIndex = 16;
            this.label1.Text = " 桩号";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // textBox_mile
            // 
            this.textBox_mile.Location = new System.Drawing.Point(35, 2);
            this.textBox_mile.Margin = new System.Windows.Forms.Padding(0, 2, 0, 0);
            this.textBox_mile.Name = "textBox_mile";
            this.textBox_mile.Size = new System.Drawing.Size(50, 21);
            this.textBox_mile.TabIndex = 17;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(85, 4);
            this.label2.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 18;
            this.label2.Text = "  里程";
            // 
            // textBox_dmi
            // 
            this.textBox_dmi.Location = new System.Drawing.Point(126, 2);
            this.textBox_dmi.Margin = new System.Windows.Forms.Padding(0, 2, 0, 0);
            this.textBox_dmi.Name = "textBox_dmi";
            this.textBox_dmi.Size = new System.Drawing.Size(50, 21);
            this.textBox_dmi.TabIndex = 19;
            // 
            // button_jump
            // 
            this.button_jump.Location = new System.Drawing.Point(176, 0);
            this.button_jump.Margin = new System.Windows.Forms.Padding(0);
            this.button_jump.Name = "button_jump";
            this.button_jump.Size = new System.Drawing.Size(60, 24);
            this.button_jump.TabIndex = 20;
            this.button_jump.Text = "跳转";
            this.button_jump.UseVisualStyleBackColor = true;
            this.button_jump.Click += new System.EventHandler(this.button_jump_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(236, 4);
            this.label3.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 21;
            this.label3.Text = "  ";
            // 
            // button_last
            // 
            this.button_last.ImageIndex = 4;
            this.button_last.ImageList = this.imageList1;
            this.button_last.Location = new System.Drawing.Point(253, 0);
            this.button_last.Margin = new System.Windows.Forms.Padding(0);
            this.button_last.Name = "button_last";
            this.button_last.Size = new System.Drawing.Size(30, 24);
            this.button_last.TabIndex = 22;
            this.button_last.UseVisualStyleBackColor = true;
            this.button_last.Click += new System.EventHandler(this.button_last_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Add_16x16.png");
            this.imageList1.Images.SetKeyName(1, "First_16x16.png");
            this.imageList1.Images.SetKeyName(2, "Last_16x16.png");
            this.imageList1.Images.SetKeyName(3, "Next_16x16.png");
            this.imageList1.Images.SetKeyName(4, "Prev_16x16.png");
            this.imageList1.Images.SetKeyName(5, "Remove_16x16.png");
            this.imageList1.Images.SetKeyName(6, "SelectAll_16x16.png");
            this.imageList1.Images.SetKeyName(7, "Media_16x16.png");
            // 
            // button_play
            // 
            this.button_play.ImageIndex = 7;
            this.button_play.ImageList = this.imageList1;
            this.button_play.Location = new System.Drawing.Point(283, 0);
            this.button_play.Margin = new System.Windows.Forms.Padding(0);
            this.button_play.Name = "button_play";
            this.button_play.Size = new System.Drawing.Size(30, 24);
            this.button_play.TabIndex = 23;
            this.button_play.UseVisualStyleBackColor = true;
            this.button_play.Click += new System.EventHandler(this.button_play_Click);
            // 
            // button_next
            // 
            this.button_next.ImageIndex = 3;
            this.button_next.ImageList = this.imageList1;
            this.button_next.Location = new System.Drawing.Point(313, 0);
            this.button_next.Margin = new System.Windows.Forms.Padding(0);
            this.button_next.Name = "button_next";
            this.button_next.Size = new System.Drawing.Size(30, 24);
            this.button_next.TabIndex = 24;
            this.button_next.UseVisualStyleBackColor = true;
            this.button_next.Click += new System.EventHandler(this.button_next_Click);
            // 
            // button_speedsub
            // 
            this.button_speedsub.ImageIndex = 5;
            this.button_speedsub.ImageList = this.imageList1;
            this.button_speedsub.Location = new System.Drawing.Point(343, 0);
            this.button_speedsub.Margin = new System.Windows.Forms.Padding(0);
            this.button_speedsub.Name = "button_speedsub";
            this.button_speedsub.Size = new System.Drawing.Size(30, 24);
            this.button_speedsub.TabIndex = 25;
            this.button_speedsub.UseVisualStyleBackColor = true;
            this.button_speedsub.Click += new System.EventHandler(this.button_speedsub_Click);
            // 
            // button_speedadd
            // 
            this.button_speedadd.ImageIndex = 0;
            this.button_speedadd.ImageList = this.imageList1;
            this.button_speedadd.Location = new System.Drawing.Point(373, 0);
            this.button_speedadd.Margin = new System.Windows.Forms.Padding(0);
            this.button_speedadd.Name = "button_speedadd";
            this.button_speedadd.Size = new System.Drawing.Size(30, 24);
            this.button_speedadd.TabIndex = 26;
            this.button_speedadd.UseVisualStyleBackColor = true;
            this.button_speedadd.Click += new System.EventHandler(this.button_speedadd_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(403, 4);
            this.label5.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(17, 12);
            this.label5.TabIndex = 27;
            this.label5.Text = "  ";
            // 
            // progressBar_per
            // 
            this.progressBar_per.Location = new System.Drawing.Point(420, 0);
            this.progressBar_per.Margin = new System.Windows.Forms.Padding(0);
            this.progressBar_per.Name = "progressBar_per";
            this.progressBar_per.Size = new System.Drawing.Size(100, 23);
            this.progressBar_per.TabIndex = 28;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(520, 4);
            this.label6.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 29;
            this.label6.Text = "  ";
            // 
            // label_ImgPath
            // 
            this.label_ImgPath.AutoSize = true;
            this.label_ImgPath.Location = new System.Drawing.Point(3, 550);
            this.label_ImgPath.Name = "label_ImgPath";
            this.label_ImgPath.Size = new System.Drawing.Size(53, 12);
            this.label_ImgPath.TabIndex = 1;
            this.label_ImgPath.Text = "图像路径";
            this.label_ImgPath.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.label_ImgPath_MouseDoubleClick);
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 33);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(994, 514);
            this.panel1.TabIndex = 4;
            // 
            // timer_roadplay
            // 
            this.timer_roadplay.Interval = 1024;
            // 
            // WinPanoImg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1000, 577);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "WinPanoImg";
            this.Text = "WinPanoImg";
            this.Load += new System.EventHandler(this.WinPanoImg_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel3.ResumeLayout(false);
            this.flowLayoutPanel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label_ImgPath;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_mile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_dmi;
        private System.Windows.Forms.Button button_jump;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button_last;
        private System.Windows.Forms.Button button_play;
        private System.Windows.Forms.Button button_next;
        private System.Windows.Forms.Button button_speedsub;
        private System.Windows.Forms.Button button_speedadd;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ProgressBar progressBar_per;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ToolTip toolTip_label;
        private System.Windows.Forms.Timer timer_roadplay;
        private System.Windows.Forms.Panel panel1;
    }
}