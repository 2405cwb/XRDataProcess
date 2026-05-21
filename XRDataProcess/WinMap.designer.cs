namespace XRDataProcess
{
    partial class WinMap
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinMap));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_delete = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button_add = new System.Windows.Forms.Button();
            this.button_Clear = new System.Windows.Forms.Button();
            this.button_ShowStart = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button_ShowMile = new System.Windows.Forms.Button();
            this.button_ShowRoad = new System.Windows.Forms.Button();
            this.webView2_map = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.timer_update = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.webView2_map)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(4);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.BackColor = System.Drawing.Color.Transparent;
            this.splitContainer1.Panel1.Controls.Add(this.panel1);
            this.splitContainer1.Panel2Collapsed = true;
            this.splitContainer1.Size = new System.Drawing.Size(1070, 564);
            this.splitContainer1.SplitterDistance = 204;
            this.splitContainer1.SplitterWidth = 6;
            this.splitContainer1.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.button_Clear);
            this.panel1.Controls.Add(this.button_ShowStart);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.button_ShowMile);
            this.panel1.Controls.Add(this.button_ShowRoad);
            this.panel1.Controls.Add(this.webView2_map);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1070, 564);
            this.panel1.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.button_delete);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.button_add);
            this.groupBox1.Location = new System.Drawing.Point(678, 50);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(260, 124);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "桩号标签";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 44);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 18);
            this.label1.TabIndex = 4;
            this.label1.Text = "标签内容";
            // 
            // button_delete
            // 
            this.button_delete.Location = new System.Drawing.Point(134, 80);
            this.button_delete.Margin = new System.Windows.Forms.Padding(4);
            this.button_delete.Name = "button_delete";
            this.button_delete.Size = new System.Drawing.Size(112, 34);
            this.button_delete.TabIndex = 7;
            this.button_delete.Text = "删除";
            this.button_delete.UseVisualStyleBackColor = true;
            this.button_delete.Click += new System.EventHandler(this.button_delete_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(112, 39);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(132, 28);
            this.textBox2.TabIndex = 5;
            // 
            // button_add
            // 
            this.button_add.Location = new System.Drawing.Point(12, 80);
            this.button_add.Margin = new System.Windows.Forms.Padding(4);
            this.button_add.Name = "button_add";
            this.button_add.Size = new System.Drawing.Size(112, 34);
            this.button_add.TabIndex = 6;
            this.button_add.Text = "添加";
            this.button_add.UseVisualStyleBackColor = true;
            this.button_add.Click += new System.EventHandler(this.button_add_Click);
            // 
            // button_Clear
            // 
            this.button_Clear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Clear.Enabled = false;
            this.button_Clear.Location = new System.Drawing.Point(947, 174);
            this.button_Clear.Margin = new System.Windows.Forms.Padding(4);
            this.button_Clear.Name = "button_Clear";
            this.button_Clear.Size = new System.Drawing.Size(112, 45);
            this.button_Clear.TabIndex = 11;
            this.button_Clear.Text = "清除地图";
            this.button_Clear.UseVisualStyleBackColor = true;
            this.button_Clear.Click += new System.EventHandler(this.button_Clear_Click);
            // 
            // button_ShowStart
            // 
            this.button_ShowStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_ShowStart.Location = new System.Drawing.Point(944, 10);
            this.button_ShowStart.Margin = new System.Windows.Forms.Padding(4);
            this.button_ShowStart.Name = "button_ShowStart";
            this.button_ShowStart.Size = new System.Drawing.Size(112, 45);
            this.button_ShowStart.TabIndex = 8;
            this.button_ShowStart.Text = "定位到起点";
            this.button_ShowStart.UseVisualStyleBackColor = true;
            this.button_ShowStart.Click += new System.EventHandler(this.button_ShowStart_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(947, 227);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 45);
            this.button1.TabIndex = 12;
            this.button1.Text = "截图按钮";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(678, 10);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(258, 28);
            this.textBox1.TabIndex = 2;
            // 
            // button_ShowMile
            // 
            this.button_ShowMile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_ShowMile.Location = new System.Drawing.Point(947, 120);
            this.button_ShowMile.Margin = new System.Windows.Forms.Padding(4);
            this.button_ShowMile.Name = "button_ShowMile";
            this.button_ShowMile.Size = new System.Drawing.Size(112, 45);
            this.button_ShowMile.TabIndex = 10;
            this.button_ShowMile.Text = "显示里程桩";
            this.button_ShowMile.UseVisualStyleBackColor = true;
            this.button_ShowMile.Click += new System.EventHandler(this.button_ShowMile_Click);
            // 
            // button_ShowRoad
            // 
            this.button_ShowRoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_ShowRoad.Location = new System.Drawing.Point(947, 64);
            this.button_ShowRoad.Margin = new System.Windows.Forms.Padding(4);
            this.button_ShowRoad.Name = "button_ShowRoad";
            this.button_ShowRoad.Size = new System.Drawing.Size(112, 45);
            this.button_ShowRoad.TabIndex = 9;
            this.button_ShowRoad.Text = "显示轨迹";
            this.button_ShowRoad.UseVisualStyleBackColor = true;
            this.button_ShowRoad.Click += new System.EventHandler(this.button_ShowRoad_Click);
            // 
            // webView2_map
            // 
            this.webView2_map.AllowExternalDrop = true;
            this.webView2_map.CreationProperties = null;
            this.webView2_map.DefaultBackgroundColor = System.Drawing.Color.White;
            this.webView2_map.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webView2_map.Location = new System.Drawing.Point(0, 0);
            this.webView2_map.Margin = new System.Windows.Forms.Padding(4);
            this.webView2_map.MinimumSize = new System.Drawing.Size(30, 30);
            this.webView2_map.Name = "webView2_map";
            this.webView2_map.Size = new System.Drawing.Size(1070, 564);
            this.webView2_map.TabIndex = 0;
            this.webView2_map.ZoomFactor = 1D;
            // 
            // timer_update
            // 
            this.timer_update.Interval = 10;
            this.timer_update.Tick += new System.EventHandler(this.timer_update_Tick);
            // 
            // WinMap
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1070, 564);
            this.ControlBox = false;
            this.Controls.Add(this.splitContainer1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "WinMap";
            this.Opacity = 0.5D;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Tag = "5";
            this.TopMost = true;
            this.Shown += new System.EventHandler(this.WinMap_Shown);
            this.splitContainer1.Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.webView2_map)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Timer timer_update;
        private System.Windows.Forms.Button button_Clear;
        private System.Windows.Forms.Button button_ShowMile;
        private System.Windows.Forms.Button button_ShowRoad;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_delete;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button button_add;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button_ShowStart;
        private System.Windows.Forms.Button button1;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView2_map;
        private System.Windows.Forms.Panel panel1;
    }
}