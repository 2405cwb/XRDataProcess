namespace XRDataProcess
{
    partial class WinRoadNew
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinRoadNew));
            this.panel_Img = new System.Windows.Forms.Panel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox_road = new System.Windows.Forms.PictureBox();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.label_imgpath = new System.Windows.Forms.Label();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_mile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_dmi = new System.Windows.Forms.TextBox();
            this.button_jump = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.button_lq = new System.Windows.Forms.Button();
            this.button_sn = new System.Windows.Forms.Button();
            this.button_SS = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.button_last = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.button_play = new System.Windows.Forms.Button();
            this.button_next = new System.Windows.Forms.Button();
            this.button_speedsub = new System.Windows.Forms.Button();
            this.button_speedadd = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.progressBar_per = new System.Windows.Forms.ProgressBar();
            this.label6 = new System.Windows.Forms.Label();
            this.drawModel_Combox = new System.Windows.Forms.ComboBox();
            this.timer_roadplay = new System.Windows.Forms.Timer(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panel_Img.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_road)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_Img
            // 
            this.panel_Img.Controls.Add(this.tableLayoutPanel1);
            this.panel_Img.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel_Img.Location = new System.Drawing.Point(0, 0);
            this.panel_Img.Name = "panel_Img";
            this.panel_Img.Size = new System.Drawing.Size(904, 314);
            this.panel_Img.TabIndex = 2;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(904, 314);
            this.tableLayoutPanel1.TabIndex = 20;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.pictureBox_road);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 27);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(904, 260);
            this.panel1.TabIndex = 0;
            // 
            // pictureBox_road
            // 
            this.pictureBox_road.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox_road.Location = new System.Drawing.Point(0, 0);
            this.pictureBox_road.Margin = new System.Windows.Forms.Padding(0);
            this.pictureBox_road.Name = "pictureBox_road";
            this.pictureBox_road.Size = new System.Drawing.Size(904, 260);
            this.pictureBox_road.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox_road.TabIndex = 15;
            this.pictureBox_road.TabStop = false;
            this.pictureBox_road.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBox_road_Paint);
            this.pictureBox_road.MouseClick += new System.Windows.Forms.MouseEventHandler(this.pictureBox_road_MouseClick);
            this.pictureBox_road.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.pictureBox_road_MouseDoubleClick);
            this.pictureBox_road.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBox_road_MouseDown);
            this.pictureBox_road.MouseEnter += new System.EventHandler(this.pictureBox_road_MouseEnter);
            this.pictureBox_road.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBox_road_MouseMove);
            this.pictureBox_road.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pictureBox_road_MouseUp);
            this.pictureBox_road.Resize += new System.EventHandler(this.pictureBox_road_Resize);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.label_imgpath);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 287);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(904, 27);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // label_imgpath
            // 
            this.label_imgpath.AutoSize = true;
            this.label_imgpath.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label_imgpath.Location = new System.Drawing.Point(3, 3);
            this.label_imgpath.Margin = new System.Windows.Forms.Padding(3, 3, 0, 0);
            this.label_imgpath.Name = "label_imgpath";
            this.label_imgpath.Size = new System.Drawing.Size(29, 12);
            this.label_imgpath.TabIndex = 3;
            this.label_imgpath.Text = "null";
            this.label_imgpath.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.label_imgpath_MouseDoubleClick);
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.flowLayoutPanel2.Controls.Add(this.label1);
            this.flowLayoutPanel2.Controls.Add(this.textBox_mile);
            this.flowLayoutPanel2.Controls.Add(this.label2);
            this.flowLayoutPanel2.Controls.Add(this.textBox_dmi);
            this.flowLayoutPanel2.Controls.Add(this.button_jump);
            this.flowLayoutPanel2.Controls.Add(this.label3);
            this.flowLayoutPanel2.Controls.Add(this.button_lq);
            this.flowLayoutPanel2.Controls.Add(this.button_sn);
            this.flowLayoutPanel2.Controls.Add(this.button_SS);
            this.flowLayoutPanel2.Controls.Add(this.label4);
            this.flowLayoutPanel2.Controls.Add(this.button_last);
            this.flowLayoutPanel2.Controls.Add(this.button_play);
            this.flowLayoutPanel2.Controls.Add(this.button_next);
            this.flowLayoutPanel2.Controls.Add(this.button_speedsub);
            this.flowLayoutPanel2.Controls.Add(this.button_speedadd);
            this.flowLayoutPanel2.Controls.Add(this.label5);
            this.flowLayoutPanel2.Controls.Add(this.progressBar_per);
            this.flowLayoutPanel2.Controls.Add(this.label6);
            this.flowLayoutPanel2.Controls.Add(this.drawModel_Combox);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(904, 27);
            this.flowLayoutPanel2.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(0, 4);
            this.label1.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 12);
            this.label1.TabIndex = 0;
            this.label1.Tag = "-1";
            this.label1.Text = " 桩号";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // textBox_mile
            // 
            this.textBox_mile.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_mile.Location = new System.Drawing.Point(35, 2);
            this.textBox_mile.Margin = new System.Windows.Forms.Padding(0, 2, 0, 0);
            this.textBox_mile.Name = "textBox_mile";
            this.textBox_mile.Size = new System.Drawing.Size(50, 21);
            this.textBox_mile.TabIndex = 1;
            this.textBox_mile.Tag = "-1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(85, 4);
            this.label2.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 2;
            this.label2.Tag = "-1";
            this.label2.Text = "  里程";
            // 
            // textBox_dmi
            // 
            this.textBox_dmi.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_dmi.Location = new System.Drawing.Point(126, 2);
            this.textBox_dmi.Margin = new System.Windows.Forms.Padding(0, 2, 0, 0);
            this.textBox_dmi.Name = "textBox_dmi";
            this.textBox_dmi.Size = new System.Drawing.Size(50, 21);
            this.textBox_dmi.TabIndex = 3;
            this.textBox_dmi.Tag = "-1";
            // 
            // button_jump
            // 
            this.button_jump.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_jump.Location = new System.Drawing.Point(176, 0);
            this.button_jump.Margin = new System.Windows.Forms.Padding(0);
            this.button_jump.Name = "button_jump";
            this.button_jump.Size = new System.Drawing.Size(40, 24);
            this.button_jump.TabIndex = 4;
            this.button_jump.Tag = "-1";
            this.button_jump.Text = "跳转";
            this.button_jump.UseVisualStyleBackColor = true;
            this.button_jump.Click += new System.EventHandler(this.button_jump_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(216, 4);
            this.label3.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(8, 7);
            this.label3.TabIndex = 5;
            this.label3.Text = " ";
            // 
            // button_lq
            // 
            this.button_lq.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_lq.Location = new System.Drawing.Point(224, 0);
            this.button_lq.Margin = new System.Windows.Forms.Padding(0);
            this.button_lq.Name = "button_lq";
            this.button_lq.Size = new System.Drawing.Size(40, 24);
            this.button_lq.TabIndex = 15;
            this.button_lq.Tag = "0";
            this.button_lq.Text = "沥青";
            this.button_lq.UseVisualStyleBackColor = true;
            this.button_lq.Click += new System.EventHandler(this.button_RoadType_Click);
            // 
            // button_sn
            // 
            this.button_sn.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_sn.Location = new System.Drawing.Point(264, 0);
            this.button_sn.Margin = new System.Windows.Forms.Padding(0);
            this.button_sn.Name = "button_sn";
            this.button_sn.Size = new System.Drawing.Size(40, 24);
            this.button_sn.TabIndex = 16;
            this.button_sn.Tag = "1";
            this.button_sn.Text = "水泥";
            this.button_sn.UseVisualStyleBackColor = true;
            this.button_sn.Click += new System.EventHandler(this.button_RoadType_Click);
            // 
            // button_SS
            // 
            this.button_SS.Location = new System.Drawing.Point(304, 0);
            this.button_SS.Margin = new System.Windows.Forms.Padding(0);
            this.button_SS.Name = "button_SS";
            this.button_SS.Size = new System.Drawing.Size(40, 24);
            this.button_SS.TabIndex = 18;
            this.button_SS.Tag = "2";
            this.button_SS.Text = "砂石";
            this.button_SS.UseVisualStyleBackColor = true;
            this.button_SS.Click += new System.EventHandler(this.button_RoadType_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(344, 4);
            this.label4.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(12, 7);
            this.label4.TabIndex = 17;
            this.label4.Text = "  ";
            // 
            // button_last
            // 
            this.button_last.ImageIndex = 4;
            this.button_last.ImageList = this.imageList1;
            this.button_last.Location = new System.Drawing.Point(356, 0);
            this.button_last.Margin = new System.Windows.Forms.Padding(0);
            this.button_last.Name = "button_last";
            this.button_last.Size = new System.Drawing.Size(30, 24);
            this.button_last.TabIndex = 6;
            this.button_last.Tag = "-1";
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
            this.button_play.Location = new System.Drawing.Point(386, 0);
            this.button_play.Margin = new System.Windows.Forms.Padding(0);
            this.button_play.Name = "button_play";
            this.button_play.Size = new System.Drawing.Size(30, 24);
            this.button_play.TabIndex = 7;
            this.button_play.Tag = "-1";
            this.button_play.UseVisualStyleBackColor = true;
            this.button_play.Click += new System.EventHandler(this.button_play_Click);
            // 
            // button_next
            // 
            this.button_next.ImageIndex = 3;
            this.button_next.ImageList = this.imageList1;
            this.button_next.Location = new System.Drawing.Point(416, 0);
            this.button_next.Margin = new System.Windows.Forms.Padding(0);
            this.button_next.Name = "button_next";
            this.button_next.Size = new System.Drawing.Size(30, 24);
            this.button_next.TabIndex = 8;
            this.button_next.Tag = "-1";
            this.button_next.UseVisualStyleBackColor = true;
            this.button_next.Click += new System.EventHandler(this.button_next_Click);
            // 
            // button_speedsub
            // 
            this.button_speedsub.ImageIndex = 5;
            this.button_speedsub.ImageList = this.imageList1;
            this.button_speedsub.Location = new System.Drawing.Point(446, 0);
            this.button_speedsub.Margin = new System.Windows.Forms.Padding(0);
            this.button_speedsub.Name = "button_speedsub";
            this.button_speedsub.Size = new System.Drawing.Size(30, 24);
            this.button_speedsub.TabIndex = 10;
            this.button_speedsub.Tag = "-1";
            this.button_speedsub.UseVisualStyleBackColor = true;
            this.button_speedsub.Click += new System.EventHandler(this.button_speedsub_Click);
            // 
            // button_speedadd
            // 
            this.button_speedadd.ImageIndex = 0;
            this.button_speedadd.ImageList = this.imageList1;
            this.button_speedadd.Location = new System.Drawing.Point(476, 0);
            this.button_speedadd.Margin = new System.Windows.Forms.Padding(0);
            this.button_speedadd.Name = "button_speedadd";
            this.button_speedadd.Size = new System.Drawing.Size(30, 24);
            this.button_speedadd.TabIndex = 11;
            this.button_speedadd.Tag = "-1";
            this.button_speedadd.UseVisualStyleBackColor = true;
            this.button_speedadd.Click += new System.EventHandler(this.button_speedadd_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(506, 4);
            this.label5.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(12, 7);
            this.label5.TabIndex = 12;
            this.label5.Text = "  ";
            // 
            // progressBar_per
            // 
            this.progressBar_per.Location = new System.Drawing.Point(518, 0);
            this.progressBar_per.Margin = new System.Windows.Forms.Padding(0);
            this.progressBar_per.Name = "progressBar_per";
            this.progressBar_per.Size = new System.Drawing.Size(100, 23);
            this.progressBar_per.TabIndex = 13;
            this.progressBar_per.Tag = "-1";
            this.toolTip1.SetToolTip(this.progressBar_per, "当前进度");
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(618, 4);
            this.label6.Margin = new System.Windows.Forms.Padding(0, 4, 0, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 14;
            this.label6.Text = "  ";
            // 
            // drawModel_Combox
            // 
            this.drawModel_Combox.FormattingEnabled = true;
            this.drawModel_Combox.Items.AddRange(new object[] {
            "常规绘制",
            "片状绘制(D)",
            "线状绘制(B)(N键结束绘制)"});
            this.drawModel_Combox.Location = new System.Drawing.Point(635, 0);
            this.drawModel_Combox.Margin = new System.Windows.Forms.Padding(0);
            this.drawModel_Combox.Name = "drawModel_Combox";
            this.drawModel_Combox.Size = new System.Drawing.Size(120, 20);
            this.drawModel_Combox.TabIndex = 19;
            this.drawModel_Combox.SelectedIndexChanged += new System.EventHandler(this.drawModel_Combox_SelectedIndexChanged);
            // 
            // timer_roadplay
            // 
            this.timer_roadplay.Interval = 1024;
            this.timer_roadplay.Tick += new System.EventHandler(this.timer_roadplay_Tick);
            // 
            // WinRoadNew
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(904, 314);
            this.Controls.Add(this.panel_Img);
            this.KeyPreview = true;
            this.Name = "WinRoadNew";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.WinRoadNew_FormClosed);
            this.Load += new System.EventHandler(this.WinRoadNew_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.WinRoadNew_KeyDown);
            this.KeyUp += new System.Windows.Forms.KeyEventHandler(this.WinRoadNew_KeyUp);
            this.panel_Img.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_road)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

       

        #endregion

        private System.Windows.Forms.Panel panel_Img;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox_road;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Label label_imgpath;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_mile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_dmi;
        private System.Windows.Forms.Button button_jump;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button_lq;
        private System.Windows.Forms.Button button_sn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button_last;
        private System.Windows.Forms.Button button_play;
        private System.Windows.Forms.Button button_next;
        private System.Windows.Forms.Button button_speedsub;
        private System.Windows.Forms.Button button_speedadd;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ProgressBar progressBar_per;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Timer timer_roadplay;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button_SS;
        private System.Windows.Forms.ComboBox drawModel_Combox;
    }
}