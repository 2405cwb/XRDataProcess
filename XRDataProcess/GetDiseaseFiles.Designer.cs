namespace XRDataProcess
{
    partial class GetDiseaseFiles
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GetDiseaseFiles));
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_Path = new System.Windows.Forms.TextBox();
            this.button_Select = new System.Windows.Forms.Button();
            this.button_Start = new System.Windows.Forms.Button();
            this.button_Close = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "目的文件夹";
            // 
            // textBox_Path
            // 
            this.textBox_Path.Location = new System.Drawing.Point(84, 35);
            this.textBox_Path.Name = "textBox_Path";
            this.textBox_Path.Size = new System.Drawing.Size(322, 21);
            this.textBox_Path.TabIndex = 1;
            // 
            // button_Select
            // 
            this.button_Select.Location = new System.Drawing.Point(13, 78);
            this.button_Select.Name = "button_Select";
            this.button_Select.Size = new System.Drawing.Size(106, 35);
            this.button_Select.TabIndex = 2;
            this.button_Select.Text = "选择目的文件夹";
            this.button_Select.UseVisualStyleBackColor = true;
            this.button_Select.Click += new System.EventHandler(this.button_Select_Click);
            // 
            // button_Start
            // 
            this.button_Start.Location = new System.Drawing.Point(151, 77);
            this.button_Start.Name = "button_Start";
            this.button_Start.Size = new System.Drawing.Size(106, 35);
            this.button_Start.TabIndex = 3;
            this.button_Start.Text = "开始提取";
            this.button_Start.UseVisualStyleBackColor = true;
            this.button_Start.Click += new System.EventHandler(this.button_Start_Click);
            // 
            // button_Close
            // 
            this.button_Close.Location = new System.Drawing.Point(300, 77);
            this.button_Close.Name = "button_Close";
            this.button_Close.Size = new System.Drawing.Size(106, 35);
            this.button_Close.TabIndex = 4;
            this.button_Close.Text = "关闭";
            this.button_Close.UseVisualStyleBackColor = true;
            this.button_Close.Click += new System.EventHandler(this.button_Close_Click);
            // 
            // GetDiseaseFiles
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 134);
            this.Controls.Add(this.button_Close);
            this.Controls.Add(this.button_Start);
            this.Controls.Add(this.button_Select);
            this.Controls.Add(this.textBox_Path);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "GetDiseaseFiles";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "提取工程中的病害文件";
            this.Load += new System.EventHandler(this.GetDiseaseFiles_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_Path;
        private System.Windows.Forms.Button button_Select;
        private System.Windows.Forms.Button button_Start;
        private System.Windows.Forms.Button button_Close;
    }
}