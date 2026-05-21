using System.Diagnostics;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //方法1
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(@"C:\Users\Administrator\Desktop\新建文件夹\\新建 Microsoft Word 文档.docx")
            {
                UseShellExecute = true
            };
            p.Start();
         

        }
    }
}