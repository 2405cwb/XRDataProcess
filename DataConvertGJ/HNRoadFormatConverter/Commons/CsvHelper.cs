using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Commons
{
    public class CsvHelper
    {
        /// <summary>
        /// 把csv读到dt中
        /// </summary>
        /// <param name="file">csv文件路径</param>
        /// <returns>dt</returns>
        public static DataTable ReadDataFromCsv(string file)
        {
            DataTable dt = null;

            if (File.Exists(file))
            {
                dt = new DataTable();
                FileStream fs = new FileStream(file, FileMode.OpenOrCreate, FileAccess.Read);

                StreamReader sr = new StreamReader(fs, Encoding.Default);

                string head = sr.ReadLine();
                string[] headNames = head.Split(',');
                for (int i = 0; i < headNames.Length; i++)
                {
                    //dt.Columns.Add(headNames[i], typeof(string));
                    dt.Columns.Add("Columns" + i.ToString(), typeof(string));
                }

                #region ==添加行数据==
                //string[] values = lineStr.Split(',');
                DataRow drHead = dt.NewRow();
                for (int n = 0; n < headNames.Length; n++)
                {
                    drHead[n] = headNames[n];
                }
                dt.Rows.Add(drHead);

                while (!sr.EndOfStream)
                {
                    #region ==循环读取文件==
                    string lineStr = sr.ReadLine();
                    if (lineStr == null || lineStr.Length == 0)
                        continue;
                    string[] values = lineStr.Split(',');
                    #region ==添加行数据==
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < values.Length; i++)
                    {
                        dr[i] = values[i];
                    }
                    dt.Rows.Add(dr);
                    #endregion
                    #endregion
                }
                fs.Close();
                sr.Close();
                #endregion
            }
            return dt;
        }

        static public void WriteDataToCsv(DataTable dt, string fullPath)
        {
            FileInfo fi = new FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            string data = "";

           // 写出列名称
            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    data += dt.Columns[i].ColumnName.ToString();
            //    if (i < dt.Columns.Count - 1)
            //    {
            //        data += ",";
            //    }
            //}
            //sw.WriteLine(data);

            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string str = dt.Rows[i][j].ToString();
                    str = str.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                    {
                        str = string.Format("\"{0}\"", str);
                    }

                    data += str;
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
        }


        public static void ConvertCsvToTxt(string filePath,bool deleteFile = true)
        {
            // 读取CSV文件
            string[] lines = File.ReadAllLines(filePath);

            // 获取同名的TXT文件路径
            string txtFilePath = Path.ChangeExtension(filePath, ".txt");

            // 创建同名的TXT文件并将CSV内容写入其中
            using (StreamWriter writer = new StreamWriter(txtFilePath))
            {
                foreach (string line in lines)
                {
                    writer.WriteLine(line);
                }
            }
            if (deleteFile)
            {
                File.Delete(filePath);
            }
           
        }
    }
}
