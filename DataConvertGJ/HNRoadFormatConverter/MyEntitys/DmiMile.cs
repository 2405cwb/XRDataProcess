using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HNRoadFormatConverter.MyEntitys
{
    public class DmiMile
    {
        /// <summary>
        /// 里程
        /// </summary>
        public int _Dmi;

        /// <summary>
        /// 桩号
        /// </summary>
        public int _Mile;

        public DmiMile(int d, int m)
        {
            _Dmi = d;
            _Mile = m;
        }
        public DmiMile(DataGridViewRow row)
        {
            _Mile = Convert.ToInt32(row.Cells[0].Value);
            _Dmi = Convert.ToInt32(row.Cells[1].Value);
        }

        public override string ToString()
        {
            return string.Format("{0} {1}", _Dmi, _Mile);
        }
    }

    public class MarkInfo
    {
        /// <summary>
        /// 桩号
        /// </summary>
        public int _Mile;

        /// <summary>
        /// 打标类型
        /// </summary>
        public string _Type;

        /// <summary>
        /// 打标信息
        /// </summary>
        public string _Info;

        public MarkInfo()
        {
            _Mile = 0;
            _Type = null;
            _Info = null;
        }
        public MarkInfo(string info)
        {
            string[] str = info.Split(' ');
            _Mile = int.Parse(str[0].Replace("K", "").Replace("+", ""));

            str = str[str.Length - 1].Split(':');
            _Type = str[0];
            _Info = str[1];
        }
        public MarkInfo(DataGridViewRow row)
        {
            _Mile = Convert.ToInt32(row.Cells[0].Value);
            _Type = Convert.ToString(row.Cells[1].Value);
            _Info = Convert.ToString(row.Cells[2].Value);
        }
    }
}
