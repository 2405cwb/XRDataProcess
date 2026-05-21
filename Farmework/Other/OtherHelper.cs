using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Framework.Other
{
    public class OtherHelper
    {

     public   static Dictionary<string, string> ParseIniFile(string filePath)
        {
            var iniSettings = new Dictionary<string, string>();

            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);

                foreach (string line in lines)
                {
                    // 忽略空行和注释
                    if (string.IsNullOrWhiteSpace(line) || line.Trim().StartsWith(";"))
                        continue;

                    // 解析键值对
                    int equalsIndex = line.IndexOf('=');
                    if (equalsIndex > 0)
                    {
                        string key = line.Substring(0, equalsIndex).Trim();
                        string value = line.Substring(equalsIndex + 1).Trim();
                        iniSettings[key] = value;
                    }
                }
            }

            return iniSettings;
        }

        /// 给定一个字符串，从头到位遍历 遇到中文返回
        /// </summary>
        /// <param name="testStr"></param>
        /// <returns></returns>
        public static string removeChineseLetter(string input)
        {
            int index = -1;
            int code = 0;
            int chfrom = Convert.ToInt32("4e00", 16);    //范围（0x4e00～0x9fff）转换成int（chfrom～chend）
            int chend = Convert.ToInt32("9fff", 16);
            if (input != "")
            {

                for (int i = 0; i < input.Length; ++i)
                {
                    code = Char.ConvertToUtf32(input, i);    //获得字符串input中指定索引index处字符unicode编码

                    if (code >= chfrom && code <= chend)
                    {
                        index = i;
                        break;//当code在中文范围内返回true

                    }
                    else
                    {
                        continue;   //当code不在中文范围内返回false
                    }
                }

            }
            if (index == -1)
            {
                //纯中文 或者无中文
                return input;
            }
            else
            {
                return input.Substring(0, index);
            }

        }


        /// <summary>
        /// 获取指定文件的编码
        /// 以防止在不知道文件编码格式的情况下处理文件而造成的乱码问题
        /// </summary>
        /// <param name="filename">文件路径</param>
        /// <returns></returns>
        public static System.Text.Encoding GetFileEncodeType(string filePath)
        {

            if (filePath == null)
            {
                throw new ArgumentNullException("filePath");
            }
            Encoding encoding1 = Encoding.Default;
            if (File.Exists(filePath))
            {
                try
                {
                    using (FileStream stream1 = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        if (stream1.Length > 0)
                        {
                            using (StreamReader reader1 = new StreamReader(stream1, true))
                            {
                                char[] chArray1 = new char[1];
                                reader1.Read(chArray1, 0, 1);
                                encoding1 = reader1.CurrentEncoding;
                                reader1.BaseStream.Position = 0;
                                if (encoding1 == Encoding.UTF8)
                                {
                                    byte[] buffer1 = encoding1.GetPreamble();
                                    if (stream1.Length >= buffer1.Length)
                                    {
                                        byte[] buffer2 = new byte[buffer1.Length];
                                        stream1.Read(buffer2, 0, buffer2.Length);
                                        for (int num1 = 0; num1 < buffer2.Length; num1++)
                                        {
                                            if (buffer2[num1] != buffer1[num1])
                                            {
                                                encoding1 = Encoding.Default;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        encoding1 = Encoding.Default;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception )
                {
                    throw;
                }
                if (encoding1 == null)
                {
                    encoding1 = Encoding.UTF8;
                }
            }
            return encoding1;
        }
    }
}

