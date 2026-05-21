namespace XRDataProcess
{
    partial class 选择导出报告
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(选择导出报告));
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button_Yes = new System.Windows.Forms.Button();
            this.button_No = new System.Windows.Forms.Button();
            this.label_log = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.treeView1, 2);
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.treeView1.Location = new System.Drawing.Point(3, 3);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(644, 559);
            this.treeView1.StateImageList = this.imageList1;
            this.treeView1.TabIndex = 0;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "unchecked.png");
            this.imageList1.Images.SetKeyName(1, "checked.png");
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.treeView1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.button_Yes, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_No, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label_log, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(650, 635);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // button_Yes
            // 
            this.button_Yes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_Yes.Location = new System.Drawing.Point(3, 568);
            this.button_Yes.Name = "button_Yes";
            this.button_Yes.Size = new System.Drawing.Size(319, 34);
            this.button_Yes.TabIndex = 1;
            this.button_Yes.Text = "确定";
            this.button_Yes.UseVisualStyleBackColor = true;
            this.button_Yes.Click += new System.EventHandler(this.button_Yes_Click);
            // 
            // button_No
            // 
            this.button_No.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_No.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button_No.Location = new System.Drawing.Point(328, 568);
            this.button_No.Name = "button_No";
            this.button_No.Size = new System.Drawing.Size(319, 34);
            this.button_No.TabIndex = 2;
            this.button_No.Text = "取消";
            this.button_No.UseVisualStyleBackColor = true;
            this.button_No.Click += new System.EventHandler(this.button_No_Click);
            // 
            // label_log
            // 
            this.label_log.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.label_log, 2);
            this.label_log.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label_log.Location = new System.Drawing.Point(3, 605);
            this.label_log.Name = "label_log";
            this.label_log.Size = new System.Drawing.Size(644, 30);
            this.label_log.TabIndex = 3;
            this.label_log.Text = "选择导出报告";
            this.label_log.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // 选择导出报告
            // 
            this.AcceptButton = this.button_Yes;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button_No;
            this.ClientSize = new System.Drawing.Size(650, 635);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "选择导出报告";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "选择导出报告";
            this.Load += new System.EventHandler(this.选择导出报告_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button_Yes;
        private System.Windows.Forms.Button button_No;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label label_log;
    }
}