using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using LadybugAPI;
using System.Windows.Forms;

namespace XRDataProcess
{
    class MyPanoStitch
    {
        //1、检查数据路径中有没有中文
        //2、生成 图像序号、GPS时间、桩号关联文件
        //3、全景图像拼接
        public static bool StitchImg(string prj, WinPanoProcessBar bar)
        {
            string[] pgrpath = null;
            if (GetAllPgrData(prj, ref pgrpath))
            {
                initStream(prj);
                uint firstcnt = 0;
                uint startnum = 0;
                bool IsFirst = true;
                double sval = 0;
                bar.SetPanoVal(0.05);
                double pval = 0.95 / pgrpath.Length;
                for (int i = 0; i < pgrpath.Length; ++i )
                {
                    sval = 0.05 + i * pval;
                    bar.SetPanoVal(sval);
                    StitchPgr(prj, pgrpath[i], bar, sval, pval, ref IsFirst, ref firstcnt, ref startnum);
                }
                cleanup();
            }
            bar.SetPanoVal(1.0);
            return true;
        }

        // 获取原始pgr文件路径
        private static bool GetAllPgrData(string prj, ref string[] pgrpath)
        {
            DirectoryInfo tdir = new DirectoryInfo(prj + @"\PanoImg\Camera0\OriData");
            FileInfo[] pgrfs = tdir.GetFiles("*.pgr");
            if (pgrfs.Length < 1)
            {
                return false;
            }

            List<string> pgrlist = new List<string>();
            pgrlist.Add(pgrfs[0].FullName);
            string oldstr = pgrfs[0].FullName, curstr;
            foreach (FileInfo pgr in pgrfs)
            {
                curstr = pgr.FullName;
                if (oldstr.Substring(0, oldstr.Length - 7) != curstr.Substring(0, curstr.Length - 7))
                {
                    pgrlist.Add(pgr.FullName);
                    oldstr = curstr;
                }
            }
            pgrlist.Sort(delegate(string x, string y) { return x.CompareTo(y); });
            pgrpath = pgrlist.ToArray();
            return true;
        }

        //ladybug相关
        // contexts and some key components
        private static IntPtr m_ladybugContext;
        private static IntPtr m_streamContext;
        private static LadybugImage m_currentImage;
        private static LadybugProcessedImage m_curpanoImage;
        private static byte[] m_textureBuffer = null;
        private static uint m_oriImgWidth = 2646;
        private static uint m_oriImgHeight = 2048;

        //拼接图像
        unsafe private static void StitchPgr(string prj, string pgrpath, WinPanoProcessBar bar, double sval, double pval, ref bool IsFirst, ref uint firstcnt, ref uint startnum)
        {
            LadybugError error = Ladybug.InitializeStreamForReading(m_streamContext, pgrpath, false);
            handleError(error);

            string imgname, imgdirname;
            uint imgnum = 0;
            uint imgcnt = 0;
            error = Ladybug.GetStreamNumOfImages(m_streamContext, out imgnum);
            handleError(error);

            DateTime startTime = new DateTime(1970, 1, 1);
            DateTime TranslateDate;

            fixed (byte* texBufPtr = m_textureBuffer)
            {
                byte** texBufPtrArray = stackalloc byte*[Ladybug.LADYBUG_NUM_CAMERAS];
                for (int cami = 0; cami < Ladybug.LADYBUG_NUM_CAMERAS; ++cami)
                {
                    texBufPtrArray[cami] = texBufPtr + cami * m_oriImgWidth * m_oriImgHeight * sizeof(ushort) * 4;
                }

                error = Ladybug.GoToImage(m_streamContext, 0);
                handleError(error);
                for (uint i = 0; i < imgnum; ++i)
                {
                    error = Ladybug.ReadImageFromStream(m_streamContext, out m_currentImage);
                    handleError(error);

                    if (IsFirst)
                    {
                        firstcnt = m_currentImage.imageInfo.ulSequenceId;
                        IsFirst = false;
                    }
                    imgcnt = m_currentImage.imageInfo.ulSequenceId - firstcnt + startnum;

                    imgdirname = string.Format("{0}\\PanoImg\\Camera0\\Image_{1:0000}", prj, imgcnt / 1000);
                    if (!Directory.Exists(imgdirname))
                    {
                        CreatFolder(imgdirname);
                    }

                    TranslateDate = startTime.AddSeconds(m_currentImage.imageInfo.ulTimeSeconds);
                    TranslateDate = TranslateDate.AddHours(8.0);
                    imgname = string.Format("{0}\\{1:000}_{2:HHmmss}{3:000}.jpeg", imgdirname, imgcnt % 1000, TranslateDate, m_currentImage.imageInfo.ulTimeMicroSeconds / 1000);

                    if (!File.Exists(imgname))
                    {
                        error = Ladybug.ConvertImage(m_ladybugContext, ref m_currentImage, texBufPtrArray, LadybugPixelFormat.LADYBUG_BGRU16);
                        handleError(error);

                        error = Ladybug.UpdateTextures(m_ladybugContext, Ladybug.LADYBUG_NUM_CAMERAS, texBufPtrArray, LadybugPixelFormat.LADYBUG_BGRU16);
                        handleError(error);

                        error = Ladybug.RenderOffScreenImage(m_ladybugContext, LadybugOutputImage.LADYBUG_PANORAMIC, LadybugPixelFormat.LADYBUG_BGR16, out m_curpanoImage);
                        handleError(error);

                        error = Ladybug.SaveImage(m_ladybugContext, ref m_curpanoImage, imgname, LadybugSaveFileFormat.LADYBUG_FILEFORMAT_JPG, true);
                        handleError(error);
                    }

                    sval = 0.05 + i * pval / imgnum;
                    bar.SetPanoVal(sval);
                }

                startnum = startnum + imgnum;
            }

            error = Ladybug.StopStream(m_streamContext);
            handleError(error);
        }

        //创建文件夹
        private static void CreatFolder(string imgdirname)
        {
            string[] strs = imgdirname.Split('\\');
            int len = strs.Length;
            string dirname = strs[0];
            for (int i = 1; i < len; ++i)
            {
                dirname = string.Format("{0}\\{1}", dirname, strs[i]);
                if (!Directory.Exists(dirname))
                {
                    Directory.CreateDirectory(dirname);
                }
            }
        }

        //异常句柄
        private static void handleError(LadybugError errorCode)
        {
            if (errorCode != LadybugError.LADYBUG_OK)
            {
                //MessageBox.Show(System.Runtime.InteropServices.Marshal.PtrToStringAnsi(Ladybug.ErrorToString(errorCode)));
            }
        }

        //初始化ladybug库
        private static void initStream(string prj)
        {
            LadybugError error = Ladybug.CreateContext(out m_ladybugContext);
            handleError(error);

            error = Ladybug.LoadConfig(m_ladybugContext, prj + "\\PanoImg\\Camera0\\CameraConfig.cal");
            handleError(error);

            uint gpunum = 0;
            error = Ladybug.GetNumGPUs(m_ladybugContext, out gpunum);
            handleError(error);
            if (gpunum > 0)
                error = Ladybug.SetColorProcessingMethod(m_ladybugContext, LadybugColorProcessingMethod.LADYBUG_HQLINEAR_GPU);
            else
                error = Ladybug.SetColorProcessingMethod(m_ladybugContext, LadybugColorProcessingMethod.LADYBUG_HQLINEAR);
            handleError(error);

            // set falloff correction value and flag
            error = Ladybug.SetFalloffCorrectionFlag(m_ladybugContext, true);
            handleError(error);
            error = Ladybug.SetFalloffCorrectionAttenuation(m_ladybugContext, 1.0f);
            handleError(error);
            
            // Set blending width
            error = Ladybug.SetBlendingParams(m_ladybugContext, 100);
            handleError(error);

            string[] tstrs = File.ReadAllLines(prj + "\\PanoImg\\Camera0\\CameraConfig.cal");
            foreach (string str in tstrs)
            {
                string[] ttstr = str.Split(' ');
                if (ttstr[0] == "Resolution")
                {
                    m_oriImgWidth = uint.Parse(ttstr[1]);
                    m_oriImgHeight = uint.Parse(ttstr[2]);
                    break;
                }
            }
            error = Ladybug.InitializeAlphaMasks(m_ladybugContext, m_oriImgWidth, m_oriImgHeight, false);
            handleError(error);

            error = Ladybug.SetAlphaMasking(m_ladybugContext, true);
            handleError(error);

            error = Ladybug.ConfigureOutputImages(m_ladybugContext, (0x1 << 12) /*LadybugOutputImage.LADYBUG_PANORAMIC*/);
            handleError(error);

            error = Ladybug.SetOffScreenImageSize(m_ladybugContext, LadybugOutputImage.LADYBUG_PANORAMIC, 8192, 4096);
            handleError(error);

            error = Ladybug.CreateStreamContext(out m_streamContext);
            handleError(error);

            m_textureBuffer = new byte[Ladybug.LADYBUG_NUM_CAMERAS * m_oriImgWidth * m_oriImgHeight * sizeof(ushort) * 4];
        }

        //销毁ladybug库
        private static void cleanup()
        {
            LadybugError error;
            if (m_streamContext != IntPtr.Zero)
            {
                error = Ladybug.StopStream(m_streamContext);
                handleError(error);

                error = Ladybug.DestroyStreamContext(ref m_streamContext);
                handleError(error);

                m_streamContext = IntPtr.Zero;
            }

            if (m_ladybugContext != IntPtr.Zero)
            {
                error = Ladybug.ReleaseOffScreenImage(m_ladybugContext, LadybugOutputImage.LADYBUG_PANORAMIC);
                handleError(error);

                error = Ladybug.DestroyContext(ref m_ladybugContext);
                handleError(error);

                m_ladybugContext = IntPtr.Zero;
            }

            if (m_textureBuffer != null)
            {
                m_textureBuffer = null;
            }

            GC.Collect();
        }
    }
}
