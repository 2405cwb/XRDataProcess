namespace XRDataProcess.toolForms
{
    partial class CustomStreetInfoForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CustomStreetInfoForm));
            this.fluentDesignFormControl1 = new DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl();
            this.modifyDisListFormBtn = new DevExpress.XtraBars.BarButtonItem();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.添加ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gridSplitContainer1 = new DevExpress.XtraGrid.GridSplitContainer();
            this.tablePanel1 = new DevExpress.Utils.Layout.TablePanel();
            this.disCnt_txt = new DevExpress.XtraEditors.TextEdit();
            this.labelcontrol5 = new DevExpress.XtraEditors.LabelControl();
            this.disName_cb = new DevExpress.XtraEditors.ComboBoxEdit();
            this.disRemark = new DevExpress.XtraEditors.TextEdit();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.fluentDesignFormControl1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridSplitContainer1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridSplitContainer1.Panel1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridSplitContainer1.Panel2)).BeginInit();
            this.gridSplitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tablePanel1)).BeginInit();
            this.tablePanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.disCnt_txt.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.disName_cb.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.disRemark.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // fluentDesignFormControl1
            // 
            this.fluentDesignFormControl1.FluentDesignForm = this;
            this.fluentDesignFormControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.modifyDisListFormBtn});
            this.fluentDesignFormControl1.Location = new System.Drawing.Point(0, 0);
            this.fluentDesignFormControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.fluentDesignFormControl1.Name = "fluentDesignFormControl1";
            this.fluentDesignFormControl1.Size = new System.Drawing.Size(939, 46);
            this.fluentDesignFormControl1.TabIndex = 2;
            this.fluentDesignFormControl1.TabStop = false;
            this.fluentDesignFormControl1.TitleItemLinks.Add(this.modifyDisListFormBtn);
            // 
            // modifyDisListFormBtn
            // 
            this.modifyDisListFormBtn.Border = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.modifyDisListFormBtn.Caption = "编辑自定义病害列表";
            this.modifyDisListFormBtn.Description = "自定义景观病害类型设置";
            this.modifyDisListFormBtn.Id = 6;
            this.modifyDisListFormBtn.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("modifyDisListFormBtn.ImageOptions.Image")));
            this.modifyDisListFormBtn.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("modifyDisListFormBtn.ImageOptions.LargeImage")));
            this.modifyDisListFormBtn.Name = "modifyDisListFormBtn";
            this.modifyDisListFormBtn.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.modifyDisListFormBtn.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.modifyDisListFormBtn_ItemClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.添加ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(117, 34);
            // 
            // 添加ToolStripMenuItem
            // 
            this.添加ToolStripMenuItem.Name = "添加ToolStripMenuItem";
            this.添加ToolStripMenuItem.Size = new System.Drawing.Size(116, 30);
            this.添加ToolStripMenuItem.Text = "添加";
            // 
            // gridSplitContainer1
            // 
            this.gridSplitContainer1.Dock = System.Windows.Forms.DockStyle.Top;
            this.gridSplitContainer1.Grid = null;
            this.gridSplitContainer1.Location = new System.Drawing.Point(0, 31);
            this.gridSplitContainer1.Name = "gridSplitContainer1";
            this.gridSplitContainer1.Size = new System.Drawing.Size(567, 311);
            this.gridSplitContainer1.TabIndex = 7;
            // 
            // tablePanel1
            // 
            this.tablePanel1.Columns.AddRange(new DevExpress.Utils.Layout.TablePanelColumn[] {
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 29.52F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 49.44F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 50F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 76.04F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 50F)});
            this.tablePanel1.Controls.Add(this.disCnt_txt);
            this.tablePanel1.Controls.Add(this.labelcontrol5);
            this.tablePanel1.Controls.Add(this.disName_cb);
            this.tablePanel1.Controls.Add(this.disRemark);
            this.tablePanel1.Controls.Add(this.labelControl4);
            this.tablePanel1.Controls.Add(this.labelControl3);
            this.tablePanel1.Controls.Add(this.simpleButton3);
            this.tablePanel1.Controls.Add(this.labelControl2);
            this.tablePanel1.Controls.Add(this.labelControl1);
            this.tablePanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tablePanel1.Location = new System.Drawing.Point(0, 46);
            this.tablePanel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tablePanel1.Name = "tablePanel1";
            this.tablePanel1.Rows.AddRange(new DevExpress.Utils.Layout.TablePanelRow[] {
            new DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26F),
            new DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26F)});
            this.tablePanel1.Size = new System.Drawing.Size(939, 101);
            this.tablePanel1.TabIndex = 8;
            // 
            // disCnt_txt
            // 
            this.tablePanel1.SetColumn(this.disCnt_txt, 2);
            this.disCnt_txt.EditValue = "1";
            this.disCnt_txt.Location = new System.Drawing.Point(294, 49);
            this.disCnt_txt.Name = "disCnt_txt";
            this.disCnt_txt.Properties.MaskSettings.Set("MaskManagerType", typeof(DevExpress.Data.Mask.NumericMaskManager));
            this.disCnt_txt.Properties.MaskSettings.Set("MaskManagerSignature", "allowNull=False");
            this.disCnt_txt.Properties.MaskSettings.Set("mask", "d");
            this.tablePanel1.SetRow(this.disCnt_txt, 1);
            this.disCnt_txt.Size = new System.Drawing.Size(178, 28);
            this.disCnt_txt.TabIndex = 7;
            // 
            // labelcontrol5
            // 
            this.tablePanel1.SetColumn(this.labelcontrol5, 2);
            this.labelcontrol5.Location = new System.Drawing.Point(294, 3);
            this.labelcontrol5.Name = "labelcontrol5";
            this.tablePanel1.SetRow(this.labelcontrol5, 0);
            this.labelcontrol5.Size = new System.Drawing.Size(77, 20);
            this.labelcontrol5.TabIndex = 6;
            this.labelcontrol5.Text = "个数\\长度";
            // 
            // disName_cb
            // 
            this.tablePanel1.SetColumn(this.disName_cb, 1);
            this.disName_cb.Location = new System.Drawing.Point(113, 49);
            this.disName_cb.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.disName_cb.Name = "disName_cb";
            this.disName_cb.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.tablePanel1.SetRow(this.disName_cb, 1);
            this.disName_cb.Size = new System.Drawing.Size(174, 28);
            this.disName_cb.TabIndex = 5;
            // 
            // disRemark
            // 
            this.tablePanel1.SetColumn(this.disRemark, 3);
            this.disRemark.Location = new System.Drawing.Point(479, 49);
            this.disRemark.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.disRemark.Name = "disRemark";
            this.tablePanel1.SetRow(this.disRemark, 1);
            this.disRemark.Size = new System.Drawing.Size(272, 28);
            this.disRemark.TabIndex = 4;
            // 
            // labelControl4
            // 
            this.labelControl4.Location = new System.Drawing.Point(4, 52);
            this.labelControl4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(63, 22);
            this.labelControl4.TabIndex = 3;
            this.labelControl4.Text = "K0+000";
            // 
            // labelControl3
            // 
            this.tablePanel1.SetColumn(this.labelControl3, 3);
            this.labelControl3.Location = new System.Drawing.Point(479, 5);
            this.labelControl3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.labelControl3.Name = "labelControl3";
            this.tablePanel1.SetRow(this.labelControl3, 0);
            this.labelControl3.Size = new System.Drawing.Size(36, 16);
            this.labelControl3.TabIndex = 2;
            this.labelControl3.Text = "备注";
            // 
            // simpleButton3
            // 
            this.tablePanel1.SetColumn(this.simpleButton3, 4);
            this.simpleButton3.Location = new System.Drawing.Point(759, 45);
            this.simpleButton3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.simpleButton3.Name = "simpleButton3";
            this.tablePanel1.SetRow(this.simpleButton3, 1);
            this.simpleButton3.Size = new System.Drawing.Size(176, 36);
            this.simpleButton3.TabIndex = 5;
            this.simpleButton3.Text = "添加";
            this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
            // 
            // labelControl2
            // 
            this.tablePanel1.SetColumn(this.labelControl2, 1);
            this.labelControl2.Location = new System.Drawing.Point(113, 5);
            this.labelControl2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.labelControl2.Name = "labelControl2";
            this.tablePanel1.SetRow(this.labelControl2, 0);
            this.labelControl2.Size = new System.Drawing.Size(72, 16);
            this.labelControl2.TabIndex = 1;
            this.labelControl2.Text = "病害名称";
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(4, 5);
            this.labelControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(36, 22);
            this.labelControl1.TabIndex = 0;
            this.labelControl1.Text = "桩号";
            // 
            // CustomStreetInfoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 22F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(939, 168);
            this.Controls.Add(this.tablePanel1);
            this.Controls.Add(this.fluentDesignFormControl1);
            this.DoubleBuffered = true;
            this.FluentDesignFormControl = this.fluentDesignFormControl1;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "CustomStreetInfoForm";
            this.Text = "景观自定义病害";
            ((System.ComponentModel.ISupportInitialize)(this.fluentDesignFormControl1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridSplitContainer1.Panel1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridSplitContainer1.Panel2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridSplitContainer1)).EndInit();
            this.gridSplitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tablePanel1)).EndInit();
            this.tablePanel1.ResumeLayout(false);
            this.tablePanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.disCnt_txt.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.disName_cb.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.disRemark.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl fluentDesignFormControl1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 添加ToolStripMenuItem;
        private DevExpress.XtraBars.BarButtonItem modifyDisListFormBtn;
        private DevExpress.XtraGrid.GridSplitContainer gridSplitContainer1;
        private DevExpress.Utils.Layout.TablePanel tablePanel1;
        private DevExpress.XtraEditors.ComboBoxEdit disName_cb;
        private DevExpress.XtraEditors.TextEdit disRemark;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.SimpleButton simpleButton3;
        private DevExpress.XtraEditors.TextEdit disCnt_txt;
        private DevExpress.XtraEditors.LabelControl labelcontrol5;
    }
}