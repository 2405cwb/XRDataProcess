namespace XRDataProcess
{
    partial class SnbkSetForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SnbkSetForm));
            this.btn_cancel = new System.Windows.Forms.Button();
            this.btn_confirm = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txb_bknum = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btn_cancel
            // 
            this.btn_cancel.Location = new System.Drawing.Point(207, 137);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(75, 30);
            this.btn_cancel.TabIndex = 0;
            this.btn_cancel.Text = "取  消";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // btn_confirm
            // 
            this.btn_confirm.Location = new System.Drawing.Point(63, 137);
            this.btn_confirm.Name = "btn_confirm";
            this.btn_confirm.Size = new System.Drawing.Size(75, 30);
            this.btn_confirm.TabIndex = 1;
            this.btn_confirm.Text = "确  定";
            this.btn_confirm.UseVisualStyleBackColor = true;
            this.btn_confirm.Click += new System.EventHandler(this.btn_confirm_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(61, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "请输入板块编号：";
            // 
            // txb_bknum
            // 
            this.txb_bknum.Location = new System.Drawing.Point(168, 50);
            this.txb_bknum.Name = "txb_bknum";
            this.txb_bknum.Size = new System.Drawing.Size(114, 21);
            this.txb_bknum.TabIndex = 3;
            // 
            // SnbkSetForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(360, 190);
            this.Controls.Add(this.txb_bknum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_confirm);
            this.Controls.Add(this.btn_cancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SnbkSetForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "水泥板块编号设置";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_cancel;
        private System.Windows.Forms.Button btn_confirm;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txb_bknum;
    }
}