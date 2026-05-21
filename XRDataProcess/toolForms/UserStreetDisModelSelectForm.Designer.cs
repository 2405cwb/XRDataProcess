namespace XRDataProcess.toolForms
{
    partial class UserStreetDisModelSelectForm
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
            this.tablePanel1 = new DevExpress.Utils.Layout.TablePanel();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.tablePanel1)).BeginInit();
            this.tablePanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tablePanel1
            // 
            this.tablePanel1.Columns.AddRange(new DevExpress.Utils.Layout.TablePanelColumn[] {
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 55F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 50F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 50F)});
            this.tablePanel1.Controls.Add(this.simpleButton3);
            this.tablePanel1.Controls.Add(this.simpleButton2);
            this.tablePanel1.Controls.Add(this.simpleButton1);
            this.tablePanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tablePanel1.Location = new System.Drawing.Point(0, 0);
            this.tablePanel1.Margin = new System.Windows.Forms.Padding(5);
            this.tablePanel1.Name = "tablePanel1";
            this.tablePanel1.Padding = new System.Windows.Forms.Padding(10);
            this.tablePanel1.Rows.AddRange(new DevExpress.Utils.Layout.TablePanelRow[] {
            new DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26F)});
            this.tablePanel1.Size = new System.Drawing.Size(684, 92);
            this.tablePanel1.TabIndex = 0;
            // 
            // simpleButton1
            // 
            this.tablePanel1.SetColumn(this.simpleButton1, 0);
            this.simpleButton1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.simpleButton1.Location = new System.Drawing.Point(20, 20);
            this.simpleButton1.Margin = new System.Windows.Forms.Padding(10);
            this.simpleButton1.Name = "simpleButton1";
            this.tablePanel1.SetRow(this.simpleButton1, 0);
            this.simpleButton1.Size = new System.Drawing.Size(216, 52);
            this.simpleButton1.TabIndex = 0;
            this.simpleButton1.Text = "沿线设施病害(1)";
            this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // simpleButton2
            // 
            this.tablePanel1.SetColumn(this.simpleButton2, 1);
            this.simpleButton2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.simpleButton2.Location = new System.Drawing.Point(256, 20);
            this.simpleButton2.Margin = new System.Windows.Forms.Padding(10);
            this.simpleButton2.Name = "simpleButton2";
            this.tablePanel1.SetRow(this.simpleButton2, 0);
            this.simpleButton2.Size = new System.Drawing.Size(194, 52);
            this.simpleButton2.TabIndex = 1;
            this.simpleButton2.Text = "路基损坏病害(2)";
            this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
            // 
            // simpleButton3
            // 
            this.tablePanel1.SetColumn(this.simpleButton3, 2);
            this.simpleButton3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.simpleButton3.Location = new System.Drawing.Point(470, 20);
            this.simpleButton3.Margin = new System.Windows.Forms.Padding(10);
            this.simpleButton3.Name = "simpleButton3";
            this.tablePanel1.SetRow(this.simpleButton3, 0);
            this.simpleButton3.Size = new System.Drawing.Size(194, 52);
            this.simpleButton3.TabIndex = 2;
            this.simpleButton3.Text = "人工自定义病害(3)";
            this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
            // 
            // UserStreetDisModelSelectForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 22F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 92);
            this.Controls.Add(this.tablePanel1);
            this.KeyPreview = true;
            this.Name = "UserStreetDisModelSelectForm";
            this.Text = "景观病害模式选择";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.UserStreetDisModelSelectForm_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.tablePanel1)).EndInit();
            this.tablePanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.Utils.Layout.TablePanel tablePanel1;
        private DevExpress.XtraEditors.SimpleButton simpleButton3;
        private DevExpress.XtraEditors.SimpleButton simpleButton2;
        private DevExpress.XtraEditors.SimpleButton simpleButton1;
    }
}