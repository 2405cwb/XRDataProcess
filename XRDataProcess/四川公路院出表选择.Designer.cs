namespace XRDataProcess
{
    partial class 四川公路院出表选择
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.checkedListBoxControl1 = new DevExpress.XtraEditors.CheckedListBoxControl();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.checkedListBoxControl1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(171, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "全选";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(171, 118);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "确定";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(171, 147);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "取消";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // checkedListBoxControl1
            // 
            this.checkedListBoxControl1.Items.AddRange(new DevExpress.XtraEditors.Controls.CheckedListBoxItem[] {
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "PCI", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "病害列表", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "RQI", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "RDI", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "PBI", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "Δh", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "PWI", System.Windows.Forms.CheckState.Checked),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(null, "SMTD", System.Windows.Forms.CheckState.Checked)});
            this.checkedListBoxControl1.Location = new System.Drawing.Point(12, 12);
            this.checkedListBoxControl1.Name = "checkedListBoxControl1";
            this.checkedListBoxControl1.Size = new System.Drawing.Size(140, 158);
            this.checkedListBoxControl1.TabIndex = 6;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(171, 41);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 7;
            this.button4.Text = "全不选";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // 四川公路院出表选择
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(282, 183);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.checkedListBoxControl1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "四川公路院出表选择";
            this.Text = "四川公路院出表选择";
            ((System.ComponentModel.ISupportInitialize)(this.checkedListBoxControl1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private DevExpress.XtraEditors.CheckedListBoxControl checkedListBoxControl1;
        private System.Windows.Forms.Button button4;
    }
}