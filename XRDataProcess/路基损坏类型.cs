using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace XRDataProcess
{
    public partial class 路基损坏类型 : Form
    {
        public List<StreetDisRecord> _DisRecord = null;

        private int _direction = 0;
        private int _smile = 0, _emile = 0;
        private int _side = -1;
        private Rectangle _rect = default;
        public 路基损坏类型(int mile, int dir, int smile, int emile,int side = 0 , Rectangle rect = default)
        {
            InitializeComponent();
            
            _direction = dir;
            _smile = Math.Min(smile, emile);
            _emile = Math.Max(smile, emile);

            mask_mile.Text = mile.ToString("K0000+000");
            AddDiseaseControls(tableLayoutPanel1);

            _DisRecord = new List<StreetDisRecord>();
            _side = side;
            _rect = rect;
        }

        private void AddDiseaseControls(TableLayoutPanel panel)
        {
            int i = 0;
            int rcnt = 1;
            foreach (StreetDiseaseType dis in DiseaseTypes.roadbeddislist)
            {
                CheckBox type = new CheckBox()
                {
                    Text = string.Format("{0}(&{1})", dis.disname, dis.shortcut),
                    Top = 22 + i++ * 20,
                    Left = 20,
                    Width = 200,
                    Tag = rcnt
                };
                panel.Controls.Add(type, 0, rcnt);
                toolTip1.SetToolTip(type, dis.description);

                TextBox dislen = new TextBox() { Tag = rcnt, Text = "0", TextAlign = HorizontalAlignment.Center };
                panel.Controls.Add(dislen, 1, rcnt);

                TextBox discnt = new TextBox() { Tag = rcnt, Text = "0", TextAlign = HorizontalAlignment.Center };
                panel.Controls.Add(discnt, 2, rcnt);

                TextBox dissocre = new TextBox() { Tag = rcnt, Text = "0", TextAlign = HorizontalAlignment.Center, ReadOnly = true };
                panel.Controls.Add(dissocre, 3, rcnt);

                MaskedTextBox dismile = new MaskedTextBox() { Tag = rcnt, Text = mask_mile.Text, Mask = "K0000+000", TextAlign = HorizontalAlignment.Center };
                panel.Controls.Add(dismile, 4, rcnt);

                if (dis.unitval == 0)
                {
                    dislen.Text = dis.unitval.ToString();
                    discnt.Enabled = false;
                }
                else
                {
                    dislen.Enabled = false;
                    discnt.Text = dis.unitval.ToString();
                }

                dislen.TextChanged += new EventHandler(ckb_CheckedChanged);
                discnt.TextChanged += new EventHandler(ckb_CheckedChanged);
                type.CheckedChanged += new EventHandler(ckb_CheckedChanged);
                dismile.TextChanged += new EventHandler(dismile_TextChanged);

                rcnt++;
            }
        }

        void dismile_TextChanged(object sender, EventArgs e)
        {
            Control rbt = (Control)sender;
            int rcnt = Convert.ToInt32(rbt.Tag);

            CheckBox cbx = (CheckBox)tableLayoutPanel1.GetControlFromPosition(0, rcnt);
            if (!cbx.Checked)
            {
                return;
            }

            if (rbt.Text.Contains(' '))
            {
                MessageBox.Show("输入的桩号中包含空格，请检查！");
                return;
            }

            int dislen = Convert.ToInt32(tableLayoutPanel1.GetControlFromPosition(1, rcnt).Text);
            if (DiseaseTypes.roadbeddislist[rcnt - 1].unitval == 0)
            {
                int smile = int.Parse(tableLayoutPanel1.GetControlFromPosition(4, rcnt).Text.Replace("K", "").Replace("+", ""));
                int emile = smile + dislen * _direction;
                if (smile < _smile || smile > _emile || emile < _smile || emile > _emile)
                {
                    MessageBox.Show("输入的损坏桩号及长度区间跨越了工程区间范围，请重新输入桩号或损坏长度！");
                    return;
                }

                if (Math.Ceiling(Math.Max(smile, emile) * 0.01) != (Math.Floor(Math.Min(smile, emile) * 0.01) + 1))
                {
                    MessageBox.Show("输入的损坏桩号及长度区间跨越了百米桩，请重新输入桩号或损坏长度！");
                    return;
                }
            }
        }
        void ckb_CheckedChanged(object sender, EventArgs e)
        {
            Control rbt = (Control)sender;
            int rcnt = Convert.ToInt32(rbt.Tag);

            CheckBox cbx = (CheckBox)tableLayoutPanel1.GetControlFromPosition(0, rcnt);
            TextBox disscore = (TextBox)tableLayoutPanel1.GetControlFromPosition(3, rcnt);
            disscore.Text = "0";

            if (!cbx.Checked)
            {
                return;
            }

            int dislen = 0;
            try
            {
                dislen = Convert.ToInt32(tableLayoutPanel1.GetControlFromPosition(1, rcnt).Text);
            }
            catch
            {
                MessageBox.Show("输入的损坏长度不合法，请重新输入整数！");
                return;
            }
            int discnt = 0;
            try
            {
                discnt = Convert.ToInt32(tableLayoutPanel1.GetControlFromPosition(2, rcnt).Text);
            }
            catch
            {
                MessageBox.Show("输入的损坏处/个数不合法，请重新输入整数！");
                return;
            }

            if (DiseaseTypes.roadbeddislist[rcnt - 1].unitval == 0)
            {
                int smile = int.Parse(tableLayoutPanel1.GetControlFromPosition(4, rcnt).Text.Replace("K","").Replace("+",""));
                int emile = smile + dislen * _direction;
                if (Math.Ceiling(Math.Max(smile, emile) * 0.01) != (Math.Floor(Math.Min(smile, emile) * 0.01)+1))
                {
                    MessageBox.Show("输入的损坏桩号及长度区间跨越了百米桩，请重新输入桩号或损坏长度！");
                    return;
                }
            }

            disscore.Text = ComputeScore(rcnt - 1, dislen, discnt).ToString();
        }

        int ComputeScore(int idx, double dislen, double discnt)
        {
            int score = 0;
            if (DiseaseTypes.roadbeddislist[idx].unitval == 0)
            {
                score = (int)Math.Ceiling(dislen * DiseaseTypes.roadbeddislist[idx].unitscore);
            }
            else
            {
                score = (int)((int)Math.Ceiling(discnt) * DiseaseTypes.roadbeddislist[idx].unitscore);
            }
            return score;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < DiseaseTypes.roadbeddislist.Count; ++i)
            {
                int score = Convert.ToInt32(tableLayoutPanel1.GetControlFromPosition(3, i + 1).Text);
                if (score > 0)
                {
                    string len = tableLayoutPanel1.GetControlFromPosition(1, i + 1).Text;
                    string num = tableLayoutPanel1.GetControlFromPosition(2, i + 1).Text;
                    string mile = tableLayoutPanel1.GetControlFromPosition(4, i + 1).Text;
                    _DisRecord.Add(new StreetDisRecord(mile, DiseaseTypes.roadbeddislist[i].disname, score, num, len,_side,_rect));
                }
            }

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
