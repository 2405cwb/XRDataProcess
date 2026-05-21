namespace XRDataProcess
{
    partial class 设置单元编号
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(设置单元编号));
            this.btn_confirm = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txb_unitnum = new System.Windows.Forms.TextBox();
            this.btn_cancel = new System.Windows.Forms.Button();
            this.txb_RoadNum = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_confirm
            // 
            this.btn_confirm.Location = new System.Drawing.Point(62, 172);
            this.btn_confirm.Name = "btn_confirm";
            this.btn_confirm.Size = new System.Drawing.Size(75, 23);
            this.btn_confirm.TabIndex = 0;
            this.btn_confirm.Text = "确  定";
            this.btn_confirm.UseVisualStyleBackColor = true;
            this.btn_confirm.Click += new System.EventHandler(this.btn_confirm_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 70);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(125, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "请输入起点单元编号：";
            // 
            // txb_unitnum
            // 
            this.txb_unitnum.Location = new System.Drawing.Point(162, 67);
            this.txb_unitnum.Name = "txb_unitnum";
            this.txb_unitnum.Size = new System.Drawing.Size(141, 21);
            this.txb_unitnum.TabIndex = 2;
            this.txb_unitnum.Text = "0";
            // 
            // btn_cancel
            // 
            this.btn_cancel.Location = new System.Drawing.Point(228, 172);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_cancel.TabIndex = 3;
            this.btn_cancel.Text = "退  出";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // txb_RoadNum
            // 
            this.txb_RoadNum.Location = new System.Drawing.Point(161, 108);
            this.txb_RoadNum.Name = "txb_RoadNum";
            this.txb_RoadNum.Size = new System.Drawing.Size(141, 21);
            this.txb_RoadNum.TabIndex = 5;
            this.txb_RoadNum.Text = "1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "请输入检测车道数(个)：";
            // 
            // 设置单元编号
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(335, 237);
            this.Controls.Add(this.txb_RoadNum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_cancel);
            this.Controls.Add(this.txb_unitnum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_confirm);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "设置单元编号";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "设置单元编号";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_confirm;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txb_unitnum;
        private System.Windows.Forms.Button btn_cancel;
        private System.Windows.Forms.TextBox txb_RoadNum;
        private System.Windows.Forms.Label label2;
    }
}