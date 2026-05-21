using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Mask;
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
    public partial class ChangeRoadPictureSizeForm : DevExpress.XtraEditors.XtraForm
    {
        // 用来把用户选择的结果返回给调用方
        public string SelectedSize { get; private set; }
        // 允许的最小/最大尺寸
        private const int MinSize = 10;
        private const int MaxSize = 10000;
        public ChangeRoadPictureSizeForm(int width, int height)
        {
          
            InitializeComponent();

            textEdit1.Text = width.ToString();
            textEdit2.Text = height.ToString();
            // 设置输入验证（只能输入数字）
            textEdit1.Properties.MaskSettings.Configure<MaskSettings.Regular>(settings =>
            {
                settings.MaskExpression = @"\d*"; // 仅允许输入数字（0-9）
            });
            textEdit2.Properties.MaskSettings.Configure<MaskSettings.Regular>(settings =>
            {
                settings.MaskExpression = @"\d*"; // 仅允许输入数字（0-9）
            });

        }

        public (int,int) getUserSize()
        {

            if (!int.TryParse(textEdit1.Text, out int width) ||
                !int.TryParse(textEdit2.Text, out int height))
            {
                throw new ArgumentException("请输入有效的数字！");
            }
            // 验证尺寸范围
            if (width < MinSize || width > MaxSize ||
                height < MinSize || height > MaxSize)
            {
                throw new ArgumentException($"尺寸必须在 {MinSize}-{MaxSize} 之间！");
            }
            return (width, height);
        }

        private void ChangeRoadPictureSizeForm_Load(object sender, EventArgs e)
        {

        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                // 触发验证
                var _ = getUserSize();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (ArgumentException ex)
            {
                XtraMessageBox.Show(ex.Message, "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}