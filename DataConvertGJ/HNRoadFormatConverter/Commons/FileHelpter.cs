using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Commons
{
    public static class FileHelpter
    {  /// <summary>
       /// 复制文件夹下的所有文件、目录到指定的文件夹
       /// </summary>
       /// <param name="dir">源文件夹地址</param>
       /// <param name="desDir">指定的文件夹地址</param>
        public static void CopyFileAndDir(string dir, string desDir)
        {



            if (!System.IO.Directory.Exists(desDir))
            {
                System.IO.Directory.CreateDirectory(desDir);
            }
            else
            {
               // Directory.Delete(desDir, true);
                //System.IO.Directory.CreateDirectory(desDir);
            }
                IEnumerable<string> files = System.IO.Directory.EnumerateFileSystemEntries(dir);
            if (files != null && files.Count() > 0)
            {
                foreach (var item in files)
                {
                    string desPath = System.IO.Path.Combine(desDir, System.IO.Path.GetFileName(item));

                    //如果是文件
                    var fileExist = System.IO.File.Exists(item);
                    if (fileExist)
                    {
                        //复制文件到指定目录下                     
                        System.IO.File.Copy(item, desPath, true);
                        continue;
                    }

                    //如果是文件夹                   
                    CopyFileAndDir(item, desPath);

                }
            }
        }
    }
}
