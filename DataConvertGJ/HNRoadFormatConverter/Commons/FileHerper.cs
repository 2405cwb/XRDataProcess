using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Commons
{
    public class FileHerper
    {
        //private static readonly LogHelper _log = new LogHelper(typeof(FileHerper)) ; 
        
        #region 把字符串转换成 byte 的二进制流。
        /// <summary>
        /// 把字符串转换成 byte 的二进制流。
        /// </summary>
        /// <param name="strmemo"></param>
        /// <param name="enc">字节流的编码方式，例如 Encoding.UTF7，Encoding.UTF8; Encoding.Unicode;</param>
        /// <returns></returns>
        static public MemoryStream getMemoryStream(string strmemo, Encoding enc)
        {
            if (strmemo == string.Empty)
            {
                return null;
            }
            Byte[] btMemo = enc.GetBytes(strmemo);
            MemoryStream ms = new MemoryStream(btMemo);
            return ms;
        }
        static public MemoryStream getMemoryStream(string strmemo)
        {
            return getMemoryStream(strmemo, Encoding.UTF8);
        }
        #endregion
        #region  将字节流写入文件
        /// <summary>
        /// 将字节流写入文件
        /// </summary>
        /// <param name="mem">字节流</param>
        /// <param name="filepath">文件全路径</param>
        /// <returns></returns>
        static public bool WriteFileMemory(MemoryStream ms, string filepath, string businessType)
        {
            if (ms == null) return false;
            try
            {
                FileStream fs = new FileStream(filepath, FileMode.Create);
                Byte[] btBlob = new Byte[ms.Length];
                ms.Read(btBlob, 0, btBlob.Length);
                fs.Write(btBlob, 0, btBlob.Length);
                fs.Close();
                return true;
            }
            catch
            {
                //_log.Error("字节流写入文件的过程中出错！"+ businessType);
                return false;
            }
        }
        #endregion

    }
}
