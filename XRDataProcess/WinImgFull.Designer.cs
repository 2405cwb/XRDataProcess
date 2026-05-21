namespace XRDataProcess
{
    partial class WinImgFull
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinImgFull));
            this.pictureBox_Img = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_Img)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox_Img
            // 
            this.pictureBox_Img.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox_Img.Location = new System.Drawing.Point(0, 0);
            this.pictureBox_Img.Name = "pictureBox_Img";
            this.pictureBox_Img.Size = new System.Drawing.Size(284, 262);
            this.pictureBox_Img.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox_Img.TabIndex = 0;
            this.pictureBox_Img.TabStop = false;
            this.pictureBox_Img.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBox_Img_Paint);
            // 
            // WinImgFull
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.pictureBox_Img);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "WinImgFull";
            this.Text = "WinImgFull";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_Img)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox_Img;
    }
}