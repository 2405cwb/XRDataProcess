using HNRoadFormatConverter.MyEntitys;
using Sunny.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HNRoadFormatConverter
{
    public partial class LpFileCalculateIri : Form
    {
        bool isInFile = false;
        public LpFileCalculateIri(GJProject gJProject)
        {
           
            InitializeComponent(); 
            _gJProject = gJProject;
            isInFile = false;
           setLpDatas();
            button2.Visible = false;
        }
        public LpFileCalculateIri(string lpfile)
        {

            InitializeComponent();
            _lpFile = lpfile;
            setLpDatas(_lpFile);
            isInFile = true;
        }

        List<string> lpDatas = new List<string>();

        public void setLpDatas()
        {
            this.lpDatas = _gJProject.getLpFileText();
            // 使用换行符连接所有项
            memoEdit1.Text = string.Join(Environment.NewLine, lpDatas); 
        }

        public void setLpDatas(string lpFile)
        {
            this.lpDatas = File.ReadAllLines(lpFile).ToList();
            // 使用换行符连接所有项
            memoEdit1.Text = string.Join(Environment.NewLine, lpDatas);
        }


        private GJProject _gJProject;
        private string _lpFile;
   

        private void button1_Click_1(object sender, EventArgs e)
        {
            int space = int.Parse(comboBox1.Text);
            List<string> list = new List<string>(
    memoEdit1.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None)
);
            if (isInFile)
            {
               // List<double> iriLeft = IRM_Algorithm.WorkBankIRIAlgo_withSpeed(datas, 0, space, 0.1);
            }
            else
            {
                List<string> result = _gJProject.calculateIriValue(list, space);

                textBox2.Text = string.Join(Environment.NewLine, result);
            }
              
        }

     

        private void LpFileCalculateIri_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (isInFile)
            {


                //文本输入的
                File.WriteAllLines(_lpFile, textBox2.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            }
            else
            {
                
            }
        }
    }
}
