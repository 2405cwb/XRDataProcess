using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Farmework.Other
{
    public static class EncodingDetector
    {
        /// <summary>
        /// 给定文件的路径，读取文件的二进制数据，判断文件的编码类型
        /// </summary>
        /// <param name="FILE_NAME">文件路径</param>
        /// <returns>文件的编码类型</returns>
        public static System.Text.Encoding GetType(string FILE_NAME)
        {
            using (FileStream fs = new FileStream(FILE_NAME, FileMode.Open, FileAccess.Read))
            {
                return GetType(fs);
            }
        }

        /// <summary>
        /// 通过给定的文件流，判断文件的编码类型
        /// </summary>
        /// <param name="fs">文件流</param>
        /// <returns>文件的编码类型</returns>
        public static System.Text.Encoding GetType(FileStream fs)
        {
            // BOM 字节序列（不变）
            byte[] Unicode = new byte[] { 0xFF, 0xFE, 0x41 };
            byte[] UnicodeBIG = new byte[] { 0xFE, 0xFF, 0x00 };
            byte[] UTF8 = new byte[] { 0xEF, 0xBB, 0xBF }; // 带 BOM

            Encoding reVal = Encoding.Default;  // 默认 GBK

            // 采样前 1024 字节（够检测，避免大文件）
            byte[] buffer = new byte[1024];
            int bytesRead = fs.Read(buffer, 0, buffer.Length);
            if (bytesRead == 0) return Encoding.Default;

            // 先检查 BOM
            if (bytesRead >= 3 && buffer[0] == UTF8[0] && buffer[1] == UTF8[1] && buffer[2] == UTF8[2])
            {
                return Encoding.UTF8;
            }
            else if (bytesRead >= 3 && buffer[0] == UnicodeBIG[0] && buffer[1] == UnicodeBIG[1] && buffer[2] == UnicodeBIG[2])
            {
                return Encoding.BigEndianUnicode;
            }
            else if (bytesRead >= 3 && buffer[0] == Unicode[0] && buffer[1] == Unicode[1] && buffer[2] == Unicode[2])
            {
                return Encoding.Unicode;
            }

            // 无 BOM：检查无 BOM UTF-8
            if (IsUTF8Bytes(buffer, bytesRead))
            {
                reVal = Encoding.UTF8;
            }
            // 新增：检查 GB2312
            else if (IsGB2312Bytes(buffer, bytesRead))
            {
                reVal = Encoding.GetEncoding("GB2312");
            }
            // 可选：其他编码（如 GBK/Default 已作为 fallback）

            return reVal;
        }

        /// <summary>
        /// 判断是否是不带 BOM 的 UTF8 格式（原函数，微调为 buffer）
        /// </summary>
        /// <param name="data">字节数组</param>
        /// <param name="length">有效长度</param>
        /// <returns></returns>
        private static bool IsUTF8Bytes(byte[] data, int length)
        {
            int charByteCounter = 1;
            byte curByte;
            for (int i = 0; i < length; i++)
            {
                curByte = data[i];
                if (charByteCounter == 1)
                {
                    if (curByte >= 0x80)
                    {
                        while (((curByte <<= 1) & 0x80) != 0)
                        {
                            charByteCounter++;
                        }
                        if (charByteCounter == 1 || charByteCounter > 6)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    if ((curByte & 0xC0) != 0x80)
                    {
                        return false;
                    }
                    charByteCounter--;
                }
            }
            return charByteCounter == 1;
        }

        /// <summary>
        /// 新增：判断是否是 GB2312 格式（双字节检查 + 乱码验证）
        /// </summary>
        /// <param name="data">字节数组</param>
        /// <param name="length">有效长度</param>
        /// <returns></returns>
        private static bool IsGB2312Bytes(byte[] data, int length)
        {
            try
            {
                // 尝试用 GB2312 解码
                Encoding gbEnc = Encoding.GetEncoding("GB2312");
                string sampleText = gbEnc.GetString(data, 0, length);

                // 检查替换字符 �（乱码标志）
                if (sampleText.Contains('\uFFFD'))
                {
                    return false;  // 有乱码，非 GB2312
                }

                // 检查是否包含有效中文（CJK 汉字范围）
                if (!HasValidChinese(sampleText))
                {
                    return false;  // 无中文，疑似非 GB2312
                }

                // 简单字节范围检查（可选增强：双字节模式）
                int doubleByteCount = 0;
                for (int i = 0; i < length - 1; i += 2)  // 假设双字节
                {
                    if (data[i] >= 0xA1 && data[i] <= 0xF7 && data[i + 1] >= 0xA1 && data[i + 1] <= 0xFE)
                    {
                        doubleByteCount++;
                    }
                }
                // 如果 >10% 是双字节汉字，视为 GB2312
                return doubleByteCount > (length / 20);  // 阈值可调
            }
            catch (DecoderFallbackException)
            {
                return false;  // 解码失败
            }
        }

        /// <summary>
        /// 辅助：检查是否包含有效中文汉字
        /// </summary>
        private static bool HasValidChinese(string text)
        {
            foreach (char c in text)
            {
                if (c >= 0x4E00 && c <= 0x9FFF)  // CJK 统一汉字
                {
                    return true;
                }
            }
            return false;
        }
    }
}
