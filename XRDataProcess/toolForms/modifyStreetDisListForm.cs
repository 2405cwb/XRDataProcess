using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace XRDataProcess.toolForms
{
    public partial class modifyStreetDisListForm : DevExpress.XtraEditors.XtraForm
    {
        // 使用 BindingList 作为数据源，这样列表增删时界面会自动更新
        private BindingList<string> _dataList;

        public modifyStreetDisListForm(List<string> currentStreetDis)
        {
            InitializeComponent();

            // 1. 初始化数据
            if (currentStreetDis != null)
            {
                // 创建一个新的 BindingList，避免直接修改传入的原始引用（直到点击确定）
                _dataList = new BindingList<string>(new List<string>(currentStreetDis));
            }
            else
            {
                _dataList = new BindingList<string>();
            }

            // 2. 绑定到 ListBoxControl
            listBoxControl1.DataSource = _dataList;
        }

        /// <summary>
        /// 获取最终修改后的列表
        /// </summary>
        public List<string> GetResultList()
        {
            return _dataList.ToList();
        }

        /// <summary>
        /// 新增按钮点击事件
        /// </summary>
        private void btnAdd_Click(object sender, EventArgs e)
        {
            string newItem = txtInput.Text.Trim();

            if (string.IsNullOrEmpty(newItem))
            {
                XtraMessageBox.Show("请输入内容！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_dataList.Contains(newItem))
            {
                XtraMessageBox.Show("该项已存在，请勿重复添加！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _dataList.Add(newItem); // 添加数据，UI自动刷新
            txtInput.Text = ""; // 清空输入框
            txtInput.Focus(); // 聚焦方便继续输入
        }

        /// <summary>
        /// 删除按钮点击事件
        /// </summary>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            // 获取选中项
            string selectedItem = listBoxControl1.SelectedItem as string;

            if (selectedItem == null)
            {
                XtraMessageBox.Show("请先选择要删除的行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 确认删除（可选）
            // if (XtraMessageBox.Show($"确定要删除“{selectedItem}”吗？", "询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _dataList.Remove(selectedItem); // 移除数据，UI自动刷新
            }
        }

        /// <summary>
        /// 确定按钮
        /// </summary>
        private void okBtn_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// 取消按钮
        /// </summary>
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}