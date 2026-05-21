namespace XRDataProcess.toolForms
{
    partial class modifyStreetDisListForm
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
            this.listBoxControl1 = new DevExpress.XtraEditors.ListBoxControl();
            this.okBtn = new DevExpress.XtraEditors.SimpleButton();
            this.cancelBtn = new DevExpress.XtraEditors.SimpleButton();
            this.txtInput = new DevExpress.XtraEditors.TextEdit();
            this.btnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.btnDelete = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.listBoxControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtInput.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // listBoxControl1
            // 
            this.listBoxControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.listBoxControl1.Location = new System.Drawing.Point(0, 0);
            this.listBoxControl1.Name = "listBoxControl1";
            this.listBoxControl1.Size = new System.Drawing.Size(656, 240); // 高度减小，留出空间
            this.listBoxControl1.TabIndex = 0;
            // 
            // txtInput
            // 
            this.txtInput.Location = new System.Drawing.Point(12, 255);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(350, 28); // 输入框
            this.txtInput.TabIndex = 2;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(375, 252);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(90, 34);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click); // 绑定事件
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(480, 252);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(90, 34);
            this.btnDelete.TabIndex = 4;
            this.btnDelete.Text = "删除选中";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click); // 绑定事件
            // 
            // okBtn
            // 
            this.okBtn.Location = new System.Drawing.Point(143, 312);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(112, 34);
            this.okBtn.TabIndex = 1;
            this.okBtn.Text = "确定";
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click); // 绑定事件
            // 
            // cancelBtn
            // 
            this.cancelBtn.Location = new System.Drawing.Point(404, 312);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(112, 34);
            this.cancelBtn.TabIndex = 1;
            this.cancelBtn.Text = "取消";
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click); // 绑定事件
            // 
            // modifyStreetDisListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 22F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(656, 366);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.txtInput);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.okBtn);
            this.Controls.Add(this.listBoxControl1);
            this.Name = "modifyStreetDisListForm";
            this.Text = "自定义景观病害列表";
            ((System.ComponentModel.ISupportInitialize)(this.listBoxControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtInput.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.ListBoxControl listBoxControl1;
        private DevExpress.XtraEditors.SimpleButton okBtn;
        private DevExpress.XtraEditors.SimpleButton cancelBtn;
        private DevExpress.XtraEditors.TextEdit txtInput; // 新增输入框
        private DevExpress.XtraEditors.SimpleButton btnAdd; // 新增按钮
        private DevExpress.XtraEditors.SimpleButton btnDelete; // 删除按钮
    }
}