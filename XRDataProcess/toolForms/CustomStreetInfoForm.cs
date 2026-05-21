using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using DevExpress.XtraEditors.Filtering;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors.Repository;

namespace XRDataProcess.toolForms
{

  
    public partial class CustomStreetInfoForm : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
    { 
        
        
        /// <summary>
       /// 获取用户专属的布局文件完整路径（%LocalAppData%）
       /// </summary>
        private string GetUserLayoutPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");
            Directory.CreateDirectory(appFolder);               // 确保目录存在
            return Path.Combine(appFolder, UserSignDisConfigFileName);
        } 
        private static int comboxIndex = 0;
        private static readonly string UserSignDisConfigFileName = "userStreetDis.txt";
        private static readonly string UserSignDisConfigFilePath = AppDomain.CurrentDomain.BaseDirectory + UserSignDisConfigFileName;
        private static  string m_filePath;
         
        private  UserSignMsg m_streetDis;
       
        public CustomStreetInfoForm(int curMile, string filePath) : this(0,curMile, filePath, Rectangle.Empty)
        {

        }

        private int m_side = 0;
         
        private Rectangle m_rectangle =default;
        public CustomStreetInfoForm(int side,int curMile, string filePath, Rectangle rectangle)
        {
            m_side = side;
            InitializeComponent();
            m_rectangle = rectangle;
            if (!File.Exists(GetUserLayoutPath()))
            {
                File.Copy(UserSignDisConfigFilePath, GetUserLayoutPath());
            }
            labelControl4.Text = curMile.ToString("K0+000"); 
            this.StartPosition = FormStartPosition.CenterParent;
             
            modifyStreetDisList();


           
            // 启用工具栏
            // EnableGridViewToolbar();
        }

        private void modifyStreetDisList(  )
        {
            disName_cb.Properties.Items.Clear();
            string[] txts = File.ReadAllLines(GetUserLayoutPath());
            for (int i = 0; i < txts.Length; i++)
            {
                disName_cb.Properties.Items.Add(txts[i]);

            }
            if (comboxIndex < disName_cb.Properties.Items.Count)
            {
              

            }
            else
            {
                comboxIndex = 0;
            }
            disName_cb.SelectedIndex = comboxIndex;
        }
        
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        // 配置 GridView 列
       
        private void simpleButton1_Click(object sender, EventArgs e)
        { 
           
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            comboxIndex = disName_cb.SelectedIndex;
            string curMsg = string.Join(" ", labelControl4.Text,  disName_cb.Text, disCnt_txt.Text, disRemark.Text);

            UserSignMsg userSignMsg  ;
            if (m_rectangle!=Rectangle.Empty)
            {
                userSignMsg = new UserSignMsg(curMsg, m_rectangle,m_side);
            }
            else
            {
                 userSignMsg = new UserSignMsg(curMsg);
            }
            m_streetDis = userSignMsg; 
           this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public UserSignMsg getUser()
        {
            return m_streetDis;
        }



        private void modifyDisListFormBtn_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            var oldList = File.ReadAllLines(GetUserLayoutPath()).ToList();

            // 1. 实例化窗口，并把旧数据传进去
            // 使用 using 语句是个好习惯，窗口关闭后自动释放资源
            using (var frm = new XRDataProcess.toolForms.modifyStreetDisListForm(oldList))
            {
                // 2. 弹出模态窗口 (ShowDialog)
                // 代码会暂停在这里，直到用户关闭那个窗口
                frm.StartPosition = FormStartPosition.CenterParent;
                DialogResult result = frm.ShowDialog();

                // 3. 判断用户点的是“确定”还是“取消”
                if (result == DialogResult.OK)
                {
                    // 4. 【关键步骤】调用子窗体的公开方法获取新数据
                    List<string> newList = frm.GetResultList();

                    File.WriteAllLines(GetUserLayoutPath(), newList);

                    modifyStreetDisList();

                    XtraMessageBox.Show($"修改成功！现在有 {newList.Count} 条数据。");
                }
                else
                {
                    // 用户点了取消，或者直接点了右上角的X，什么都不做
                    Console.WriteLine("用户取消了操作");
                }
            }

         

        }

     
    }
}