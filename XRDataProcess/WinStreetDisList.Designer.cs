namespace XRDataProcess
{
    partial class WinStreetDisList
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WinStreetDisList));
            this.panel1 = new System.Windows.Forms.Panel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridView_Dislist = new System.Windows.Forms.DataGridView();
            this.Mile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Score = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Dislist)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tableLayoutPanel1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(347, 333);
            this.panel1.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView_Dislist, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.button1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.110629F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 90.88937F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(347, 333);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // dataGridView_Dislist
            // 
            this.dataGridView_Dislist.AllowUserToAddRows = false;
            this.dataGridView_Dislist.AllowUserToDeleteRows = false;
            this.dataGridView_Dislist.AllowUserToResizeRows = false;
            this.dataGridView_Dislist.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_Dislist.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_Dislist.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Mile,
            this.Type,
            this.Score});
            this.dataGridView_Dislist.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_Dislist.Location = new System.Drawing.Point(3, 33);
            this.dataGridView_Dislist.MultiSelect = false;
            this.dataGridView_Dislist.Name = "dataGridView_Dislist";
            this.dataGridView_Dislist.ReadOnly = true;
            this.dataGridView_Dislist.RowHeadersVisible = false;
            this.dataGridView_Dislist.RowTemplate.Height = 23;
            this.dataGridView_Dislist.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_Dislist.Size = new System.Drawing.Size(341, 297);
            this.dataGridView_Dislist.TabIndex = 3;
            // 
            // Mile
            // 
            this.Mile.HeaderText = "桩号";
            this.Mile.Name = "Mile";
            this.Mile.ReadOnly = true;
            // 
            // Type
            // 
            this.Type.HeaderText = "损坏类型";
            this.Type.Name = "Type";
            this.Type.ReadOnly = true;
            // 
            // Score
            // 
            this.Score.HeaderText = "扣分值";
            this.Score.Name = "Score";
            this.Score.ReadOnly = true;
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(3, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(341, 24);
            this.button1.TabIndex = 0;
            this.button1.Text = "更新病害";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // WinStreetDisList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(347, 333);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "WinStreetDisList";
            this.Text = "WinStreetDis";
            this.Load += new System.EventHandler(this.WinStreetDisList_Load_1);
            this.panel1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Dislist)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dataGridView_Dislist;
        private System.Windows.Forms.DataGridViewTextBoxColumn Mile;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Score;
        private System.Windows.Forms.Button button1;

    }
}