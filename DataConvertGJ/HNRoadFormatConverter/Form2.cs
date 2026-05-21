using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HNRoadFormatConverter
{
    public partial class Form2 : Form
    {
      
        public Form2(string num1,string num2)
        {
            InitializeComponent();
            this.maskedTextBox1.Text = num1 + "_" + num2 +"_";
            this.num1 = num1;
            this.num2 = num2;

        }
        /// <summary>
        /// 周长
        /// </summary>
        public string num1{ get; set; }
        //脉冲
        public string num2{ get; set; }
        //k
        public string k{ get; set; }
        //b
        public string b{ get; set; }
        public string result = "";

        public bool isOk=false;
        private void button1_Click(object sender, EventArgs e)
        {
            result = this.maskedTextBox1.Text;
            isOk = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            isOk = false;
            this.Close();
        }
    }
}
