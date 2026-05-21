using DevExpress.XtraEditors;
using Framework.Other;
using NPOI.SS.Formula.Functions;
using OperateIniFile;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace XRDataProcess
{


    public partial class HighAccuracySettingForm : DevExpress.XtraEditors.XtraForm
    {
        public HighAccuracySettingForm()
        {
            InitializeComponent();
            readParam();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }
        public POS_CONVERT_INFO info = new POS_CONVERT_INFO();
        private void simpleButton1_Click(object sender, EventArgs e)
        { 
            setParam();
            saveParam(info);
            this.DialogResult = DialogResult.OK;
        }


        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
        public POS_CONVERT_INFO getUserSelectInfo()
        {
            
            //根据用户选择进行设置 
            return info;
        }
        
        private void setFormParam(POS_CONVERT_INFO value)
        {
            textEdit1.Text = value.dCenterL.ToString();
            comboBoxEdit1.SelectedIndex = value.nSphereType ;
            comboBoxEdit2.SelectedIndex = value.nProjectType;
            textEdit3.Text = value.dProjectHeight.ToString();
            textEdit2.Text = value.dEastAdd.ToString();
            textEdit4.Text = value.dAverageLat.ToString();
            textEdit5.Text = value.dProjectScale.ToString();

          
            switch (value.nUseConvertModel)
            {
                case 0:
                    checkEdit1.Checked = false;
                    break;
                case 1:
                    {
                        checkEdit1.Checked = true;
                        comboBoxEdit3.SelectedIndex = 1;
                    
                    }
                    break;
                case 2:
                    {
                        checkEdit1.Checked = true;
                        comboBoxEdit3.SelectedIndex = 0;
                    }
                    break;
                default:
                    break;
            }

            textEdit6.Text=value.dFourX.ToString();
            textEdit7.Text=value.dFourY.ToString();
            textEdit8.Text=value.dFourR.ToString();
            textEdit9.Text = value.dFourK.ToString();

            textEdit19.Text =  value.dOffsetX.ToString();
            textEdit18.Text =  value.dOffsetY.ToString();
            textEdit17.Text =  value.dOffsetZ.ToString();
            textEdit16.Text =  value.dRotateX.ToString();
            textEdit15.Text =  value.dRotateY.ToString();
            textEdit14.Text = value.dRotateZ.ToString();
            textEdit13.Text = value.dK.ToString(); 
        }
        private void readParam()
        {
            //获取本地参数
            string iniPath = AppDomain.CurrentDomain.BaseDirectory + "ConfigFile\\HighAccuracySettingFormParams.ini";
            IniFiles inisetting = new IniFiles(iniPath);
            POS_CONVERT_INFO localInfo = new POS_CONVERT_INFO();
            localInfo.dCenterL = inisetting.ReadDouble("SETTING", "CenterL", 0);
            localInfo.nSphereType = inisetting.ReadInteger("SETTING", "nSphereType", 0);
            localInfo.nProjectType = inisetting.ReadInteger("SETTING", "nProjectType", 0);
            localInfo.dProjectHeight = inisetting.ReadDouble("SETTING", "dProjectHeight", 0);
            localInfo.dEastAdd = inisetting.ReadDouble("SETTING", "dEastAdd", 0);
            localInfo.dAverageLat = inisetting.ReadDouble("SETTING", "dAverageLat", 0);
            localInfo.dProjectScale = inisetting.ReadDouble("SETTING", "dProjectScale", 0);
            localInfo.nUseConvertModel = inisetting.ReadInteger("SETTING", "nUseConvertModel", 0);
            localInfo.dFourX = inisetting.ReadDouble("SETTING", "dFourX", 0);
            localInfo.dFourY = inisetting.ReadDouble("SETTING", "dFourY", 0);
            localInfo.dFourR = inisetting.ReadDouble("SETTING", "dFourR", 0);
            localInfo.dFourK = inisetting.ReadDouble("SETTING", "dFourK", 0);
            localInfo.dOffsetX = inisetting.ReadDouble("SETTING", "dOffsetX", 0);
            localInfo.dOffsetY = inisetting.ReadDouble("SETTING", "dOffsetY", 0);
            localInfo.dOffsetZ = inisetting.ReadDouble("SETTING", "dOffsetZ", 0);
            localInfo.dRotateX = inisetting.ReadDouble("SETTING", "dRotateX", 0);
            localInfo.dRotateY = inisetting.ReadDouble("SETTING", "dRotateY", 0);
            localInfo.dRotateZ = inisetting.ReadDouble("SETTING", "dRotateZ", 0);
            localInfo.dK = inisetting.ReadDouble("SETTING", "dK", 0);
            setFormParam(localInfo);
           
        } 
        private void setParam()
        {
            this.info.dCenterL = Convert.ToDouble(textEdit1.EditValue);
            this.info.nSphereType = comboBoxEdit1.SelectedIndex;
            this.info.nProjectType = comboBoxEdit2.SelectedIndex;  
            this.info.dProjectHeight = Convert.ToDouble(textEdit3.EditValue);
            this.info.dEastAdd = Convert.ToDouble(textEdit2.EditValue);
            this.info.dAverageLat = Convert.ToDouble(textEdit4.EditValue);
            this.info.dProjectScale =Convert.ToDouble(textEdit5.EditValue);
            if (!checkEdit1.Checked)
            {
                this.info.nUseConvertModel = 0;
            }
            else
            {
                switch (comboBoxEdit3.SelectedIndex)
                {
                    case 0:
                        this.info.nUseConvertModel = 2; break;
                    case 1:
                        this.info.nUseConvertModel = 1; break;
                    default:
                        break;
                }
            }
            //这里注意顺序是否正确 
            this.info.dFourX = Convert.ToDouble(textEdit6.EditValue);
            this.info.dFourY = Convert.ToDouble(textEdit7.EditValue);
            this.info.dFourR = Convert.ToDouble(textEdit8.EditValue);
            this.info.dFourK = Convert.ToDouble(textEdit9.EditValue);
             //9 七参数值设置;
             this.info.dOffsetX = Convert.ToDouble(textEdit19.EditValue);
             this.info.dOffsetY = Convert.ToDouble(textEdit18.EditValue);
             this.info.dOffsetZ = Convert.ToDouble(textEdit17.EditValue);
             this.info.dRotateX = Convert.ToDouble(textEdit16.EditValue);
             this.info.dRotateY = Convert.ToDouble(textEdit15.EditValue);
             this.info.dRotateZ = Convert.ToDouble(textEdit14.EditValue);
             this.info.dK =       Convert.ToDouble(textEdit13.EditValue);
           
        }
        private bool  saveParam(POS_CONVERT_INFO info)
        {
            string iniPath = AppDomain.CurrentDomain.BaseDirectory + "ConfigFile\\HighAccuracySettingFormParams.ini";

            IniFiles inisetting = new IniFiles(iniPath);
            inisetting.WriteDouble("SETTING", "CenterL", info.dCenterL);
            inisetting.WriteInteger("SETTING", "nSphereType", info.nSphereType);
            inisetting.WriteInteger("SETTING", "nProjectType", info.nProjectType);
            inisetting.WriteDouble("SETTING", "dProjectHeight", info.dProjectHeight);
            inisetting.WriteDouble("SETTING", "dEastAdd", info.dEastAdd);
            inisetting.WriteDouble("SETTING", "dAverageLat", info.dAverageLat);
            inisetting.WriteDouble("SETTING", "dProjectScale", info.dProjectScale);
            inisetting.WriteInteger("SETTING", "nUseConvertModel", info.nUseConvertModel);
            inisetting.WriteDouble("SETTING", "dFourX", info.dFourX);
            inisetting.WriteDouble("SETTING", "dFourY", info.dFourY);
            inisetting.WriteDouble("SETTING", "dFourR", info.dFourR);
            inisetting.WriteDouble("SETTING", "dFourK", info.dFourK);

            inisetting.WriteDouble("SETTING", "dOffsetX", info.dOffsetX);
            inisetting.WriteDouble("SETTING", "dOffsetY", info.dOffsetY);
            inisetting.WriteDouble("SETTING", "dOffsetZ", info.dOffsetZ);
            inisetting.WriteDouble("SETTING", "dRotateX", info.dRotateX); 
            inisetting.WriteDouble("SETTING", "dRotateY", info.dRotateY); 
            inisetting.WriteDouble("SETTING", "dRotateZ", info.dRotateZ); 
            inisetting.WriteDouble("SETTING", "dK", info.dK);

            return true;
        }
        // 遍历 GroupControl 中的所有控件，并将它们设置为只读
        private void SetControlsReadOnly(GroupControl groupControl,bool isReadOnly)
        {
            foreach (var control in groupControl.Controls)
            {
                if (control is BaseEdit)
                {
                    ((BaseEdit)control).Properties.ReadOnly = isReadOnly;
                }
                else if (control is GroupControl)
                {
                    SetControlsReadOnly((GroupControl)control,isReadOnly);
                }
            }
        }
        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkEdit1.Checked)
            {
                SetControlsReadOnly(groupControl3,true);
                SetControlsReadOnly(groupControl4,true);
            }
            else
            {
                SetControlsReadOnly(groupControl3, false);
                SetControlsReadOnly(groupControl4, false);
            }
        }

        private void comboBoxEdit3_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxEdit3.SelectedIndex)
            {
                case 0:
                    SetControlsReadOnly(groupControl3, true); 
                    SetControlsReadOnly(groupControl4, false);
                    break;
                case 1:
                    SetControlsReadOnly(groupControl4, true);
                    SetControlsReadOnly(groupControl3, false);
                    break;
                default:
                    break;
            }
        }

        private void HighAccuracySettingForm_Load(object sender, EventArgs e)
        {
            SetControlsReadOnly(groupControl3, true);
        }

        private void groupControl1_Paint(object sender, PaintEventArgs e)
        {

        }
    }

    public struct POS_CONVERT_INFO
    { 
        // 1 设置中央经线,单位度点度;
      public  double dCenterL;

        // 2 设置椭球体，0表示北京54,1表示西安80,2表示WGS84,3表示CGCS2000;
        public int nSphereType;

        //3 投影方法设置，0表示高斯三度带投影，1表示高斯6度带，2表示墨卡托投影，3表示横轴墨卡托投影;
        public int nProjectType;

        //4 设置投影面高程;
        public double dProjectHeight;

        // 5 东向加常数;
        public double dEastAdd;

        //6 平均纬度;
        public double dAverageLat;

        //7 尺度因子;
        public double dProjectScale;

        //8 是否使用七参数或四参数转换，1表示使用四参数，2表示使用七参数，0表示不使用;
        public int nUseConvertModel;

        // 四参数设置,不考虑高程;
        public   double dFourX;
        public   double dFourY;
        public   double dFourR;
        public   double dFourK;

        //9 七参数值设置;
        public   double dOffsetX;
        public   double dOffsetY;
        public   double dOffsetZ;
        public   double dRotateX;
        public   double dRotateY;
        public   double dRotateZ;
        public   double dK;
    };
}