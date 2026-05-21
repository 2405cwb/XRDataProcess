using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace OperateIniFile
{
    public class IniFileOpr
    {
        private static StringBuilder _StringBuilder = new StringBuilder(1024);
        [DllImport("kernel32")]//返回0表示失败，非0为成功
        private static extern long WritePrivateProfileString(string section, string key,
            string val, string filePath);

        [DllImport("kernel32")]//返回取得字符串缓冲区的长度
        private static extern long GetPrivateProfileString(string section, string key,
            string def, StringBuilder retVal, int size, string filePath);

        public static string ReadIniData(string Section, string Key, string NoText, string iniFilePath)
        {
            if (File.Exists(iniFilePath))
            {
                _StringBuilder.Clear();
                GetPrivateProfileString(Section, Key, NoText, _StringBuilder, 1024, iniFilePath);
                return _StringBuilder.ToString();
            }
            else
            {
                return String.Empty;
            }
        }

        public static bool WriteIniData(string Section, string Key, string Value, string iniFilePath)
        {
            if (!File.Exists(iniFilePath))
                File.WriteAllText(iniFilePath, "");
            long OpStation = WritePrivateProfileString(Section, Key, Value, iniFilePath);
            if (OpStation == 0)
                return false;
            else
                return true;
        }
    }
}
