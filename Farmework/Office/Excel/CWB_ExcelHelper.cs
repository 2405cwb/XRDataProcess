using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;

namespace Framework.Office.Excel
{
    public class CWB_ExcelHelper
    {

        public static void disposeExcel<T>(ref T app) where T : MSExcel._Application
        {
            if (app != null)
            {
                try
                {
                    app.DisplayAlerts = false; // 禁用保存提示框
                    app.Quit();
                    int generation = System.GC.GetGeneration(app);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(app);
                    app = default(T); // 防止再次使用已经被释放的 COM 对象 
                }
                catch (Exception)
                {

                    
                }
               
            }

        }  /// <summary>
           /// 
           /// </summary>
           /// <param name="sheet"></param>
           /// <param name="col">计算列数</param>
           /// <param name="row">计算行数  </param>
           /// <param name="rangStr">想要获得的区域 如  "A2:O" </param>
           /// <param name="srcobj"></param>
           /// <param name="userownum">计算得到的数据域行数</param>
        public static void getExcelRangeData(MSExcel.Worksheet sheet, int col, int row, string rangStr, ref object[,] srcobj, ref int userownum)
        {
            userownum = judegeusedrow(sheet, col, row);
            MSExcel.Range srcrange = sheet.get_Range(rangStr + userownum.ToString());
            srcobj = (object[,])srcrange.Value2; //3.1之后为正式数据

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="col">计算行数 开始的行数</param>
        /// <param name="row">计算行数  开始的列数</param>
        /// <param name="rangStr">想要获得的区域 如  "A2:O"</param>
        /// <returns></returns>
        public static MSExcel.Range getExcelRange(MSExcel.Worksheet sheet, int col, int row, string rangStr)
        {
            int userownum = judegeusedrow(sheet, col, row);
            MSExcel.Range srcrange = sheet.get_Range(rangStr + userownum.ToString());
            return srcrange;
        }
        public static string GetCol(char instr)
        {
            string outstr;
            if (instr > 'Z')
            {
                instr = (char)(instr - 'A');
                outstr = ((char)(instr / 26 + 'A' - 1)).ToString() + ((char)(instr % 26 + 'A')).ToString();
            }
            else
            {
                outstr = instr.ToString();
            }
            return outstr;
        }
        public static void SetBorderLine(MSExcel.Range range, int Border)
        {
            if ((Border & 1) == 1)//边缘上
            {
                range.Borders[MSExcel.XlBordersIndex.xlEdgeTop].Weight = MSExcel.XlBorderWeight.xlThick;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeTop].LineStyle = MSExcel.XlLineStyle.xlContinuous;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeTop].Weight = 2;
            }
            if ((Border & 2) == 2)//边缘右
            {
                range.Borders[MSExcel.XlBordersIndex.xlEdgeRight].Weight = MSExcel.XlBorderWeight.xlThick;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeRight].LineStyle = MSExcel.XlLineStyle.xlContinuous;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeRight].Weight = 2;
            }
            if ((Border & 4) == 4)//边缘下
            {
                range.Borders[MSExcel.XlBordersIndex.xlEdgeBottom].Weight = MSExcel.XlBorderWeight.xlThick;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeBottom].LineStyle = MSExcel.XlLineStyle.xlContinuous;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeBottom].Weight = 2;
            }
            if ((Border & 8) == 8)//边缘左
            {
                range.Borders[MSExcel.XlBordersIndex.xlEdgeLeft].Weight = MSExcel.XlBorderWeight.xlThick;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeLeft].LineStyle = MSExcel.XlLineStyle.xlContinuous;
                range.Borders[MSExcel.XlBordersIndex.xlEdgeLeft].Weight = 2;
            }
            if ((Border & 16) == 16)//内部水平
            {
                range.Borders[MSExcel.XlBordersIndex.xlInsideHorizontal].Weight = MSExcel.XlBorderWeight.xlThick;
                range.Borders[MSExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = MSExcel.XlLineStyle.xlContinuous;
                range.Borders[MSExcel.XlBordersIndex.xlInsideHorizontal].Weight = 2;
            }
            if ((Border & 32) == 32)//内部竖直
            {
                range.Borders[MSExcel.XlBordersIndex.xlInsideVertical].Weight = MSExcel.XlBorderWeight.xlThick;
                range.Borders[MSExcel.XlBordersIndex.xlInsideVertical].LineStyle = MSExcel.XlLineStyle.xlContinuous;
                range.Borders[MSExcel.XlBordersIndex.xlInsideVertical].Weight = 2;
            }
        }

        public static void WriteExcel(int excelRow, int excelColumn, int cellCountRow, int cellCountColumn, string excelValue, MSExcel._Worksheet _Worksheet, int Border)
        {
            MSExcel.Range _Range = null;
            string point1 = "", point2 = "";
            //如列数超过26,请在这里做一些修改....
            if (excelColumn < 27)
            {
                point1 = ((char)(excelColumn + 64)).ToString()
                    + excelRow.ToString();//单元格起始点 如:A1
            }
            else
            {
                point1 = ((char)(excelColumn / 26 + 64)).ToString()
                    + ((char)(excelColumn % 26 + 64)).ToString()
                    + excelRow.ToString();//单元格起始点 如:A1
            }
            if (excelColumn < 27)
            {
                point2 = ((char)(excelColumn + cellCountColumn + 64 - 1)).ToString()
                    + (excelRow + cellCountRow - 1).ToString();//单元格结束点 如:B4
            }
            else
            {
                point2 = ((char)((excelColumn) / 26 + 64)).ToString()
                    + ((char)((excelColumn) % 26 + cellCountColumn + 64 - 1)).ToString()
                    + (excelRow + cellCountRow - 1).ToString();//单元格结束点 如:B4
            }
            _Range = _Worksheet.get_Range(point1, point2);//获取单元格
            if (cellCountColumn > 0)
            {
                _Range.MergeCells = true; //合并单元格
            }
            if (Border > 0)
            {
                SetBorderLine(_Range, Border);
            }
            _Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//把内容写入单元格
        }


        //根据指定列倒序
        public static void ReflectionColnum(MSExcel.Worksheet _Worksheet, MSExcel.Range srcrange, MSExcel.Range sortrange)
        {
            srcrange.Sort(sortrange, MSExcel.XlSortOrder.xlAscending,
                Type.Missing, Type.Missing,
                MSExcel.XlSortOrder.xlAscending,
                Type.Missing,
                MSExcel.XlSortOrder.xlAscending,
                MSExcel.XlYesNoGuess.xlNo,
                Type.Missing,
                Type.Missing,
                MSExcel.XlSortOrientation.xlSortColumns,
                MSExcel.XlSortMethod.xlPinYin,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal);
        }
        public static void ReflectionColnumDescending(MSExcel.Worksheet _Worksheet, MSExcel.Range srcrange, MSExcel.Range sortrange)
        {
            srcrange.Sort(sortrange, MSExcel.XlSortOrder.xlDescending,
                Type.Missing, Type.Missing,
                MSExcel.XlSortOrder.xlDescending,
                Type.Missing,
                MSExcel.XlSortOrder.xlDescending,
                MSExcel.XlYesNoGuess.xlNo,
                Type.Missing,
                Type.Missing,
                MSExcel.XlSortOrientation.xlSortColumns,
                MSExcel.XlSortMethod.xlPinYin,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal);
        }
        //判断某列有数值的行数, Column从1开始
        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="Column">开始列数 从1开始</param>
        /// <param name="startrow">默认从第二行开始</param>
        /// <returns></returns>
        public static int judegeusedrow(MSExcel.Worksheet worksheet, int Column, int startrow = 2)
        {
            int usedrowcnt = startrow - 1;
            MSExcel.Range trange = worksheet.get_Range(String.Format("{0}:{0}", GetCol((char)('A' + Column - 1))));
            object[,] tobj = (object[,])trange.Value2;
            for (int i = startrow; ; i++)
            {
                if (tobj[i, 1] != null)
                {
                    ++usedrowcnt;
                }
                else
                {
                    break;
                }
            }
            return usedrowcnt;
        }

        //判断某行有数值的列数
        public static int judegeusedcol(MSExcel.Worksheet worksheet, int Rownum, int startcol = 2)
        {
            int usedcolcnt = startcol - 1;
            MSExcel.Range trange = worksheet.get_Range(String.Format("{0}:{0}", Rownum));
            object[,] tobj = (object[,])trange.Value2;
            for (int i = startcol; ; i++)
            {
                if (tobj[1, i] != null)
                {
                    ++usedcolcnt;
                }
                else
                {
                    break;
                }
            }
            return usedcolcnt;
        }

        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。
        /// 注意：图片必须是绝对物理路径
        /// </summary>
        /// <param name="RangeName">单元格名称，例如：B4</param>
        /// <param name="PicturePath">要插入图片的绝对路径。</param>
        public static void InsertPicture(string RangeName, MSExcel._Worksheet sheet, string PicturePath)
        {
            MSExcel.Range rng = (MSExcel.Range)sheet.get_Range(RangeName, Type.Missing);
            rng.Select();
            float PicLeft, PicTop, PicWidth, PicHeight;    //距离左边距离，顶部距离，图片宽度、高度
            PicTop = Convert.ToSingle(rng.Top);
            PicWidth = Convert.ToSingle(rng.MergeArea.Width);
            PicHeight = Convert.ToSingle(rng.Height);
            PicWidth = Convert.ToSingle(rng.Width);
            PicLeft = Convert.ToSingle(rng.Left);//+ (Convert.ToSingle(rng.MergeArea.Width) - PicWidth) / 2;
            try
            {
                MSExcel.Pictures pics = (MSExcel.Pictures)sheet.Pictures(Type.Missing);
                pics.Insert(PicturePath, Type.Missing);
                pics.Left = (double)rng.Left;
                pics.Top = (double)rng.Top;
                pics.Width = (double)rng.Width;
                pics.Height = (double)rng.Height;

            }
            catch
            {
            }

            //            如果是要在某个区域插入，改区域没有命名的话，直接传入选中区域

            //Cell1 = SourceSheet.Cells[第几行, 第几列];
            //            Cell2 = SourceSheet.Cells[Row, Column];

            //            SourceRange = SourceSheet.get_Range(Cell1, Cell2);

            //            然后把上面的这句去掉 Excel.Range rng = (Excel.Range)sheet.get_Range(RangeName, Type.Missing);

            //            把rng换成SourceRange
            //sheet.Shapes.AddPicture(PicturePath, Microsoft.Office.Core.MsoTriState.msoFalse,
            // Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, PicWidth, PicHeight);
        }


        /// <summary>
        /// 检查图片是否为灰度图
        /// </summary>
        private static bool IsGrayscale(Bitmap image)
        {
            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    Color pixelColor = image.GetPixel(x, y);
                    if (pixelColor.R != pixelColor.G || pixelColor.G != pixelColor.B)
                    {
                        return false; // 发现非灰度像素
                    }
                }
            }
            return true;
        }

        /// 将图片插入到指定的单元格位置，并设置图片的宽度和高度。
        /// 注意：图片必须是绝对物理路径
        /// </summary>
        /// <param name="rng">Excel单元格选中的区域</param>
        /// <param name="PicturePath">要插入图片的绝对路径。</param>
        public static void InsertPicture(MSExcel.Range rng, MSExcel._Worksheet sheet, string picturePath, double ratio)
        { 
            //rng.Select(); 
            try
            {
                // 获取单元格的尺寸
                float picLeft = Convert.ToSingle(rng.Left);
                float picTop = Convert.ToSingle(rng.Top);
                float picWidth = Convert.ToSingle(rng.Width);
                float picHeight = Convert.ToSingle(rng.Width) * (float)ratio;

                //参数含义：
                //图片路径
                //是否链接到文件
                //图片插入时是否随文档一起保存
                //图片在文档中的坐标位置 坐标
                //图片显示的宽度和高度msoCTruemsoFalse

                // 限制图片尺寸不超过单元格
                float cellHeight = Convert.ToSingle(rng.Height);
                if (picHeight > cellHeight)
                {
                    picHeight = cellHeight;
                    picWidth = picHeight / (float)ratio; // 按比例调整宽度
                }

                string sheetPath = sheet.Application.ActiveWorkbook.Path;
             
              //  sheet.Shapes.AddPicture(tempPath, Microsoft.Office.Core.MsoTriState.msoFalse,  Microsoft.Office.Core.MsoTriState.msoTrue, picLeft, picTop, picWidth, picHeight);
                sheet.Shapes.AddPicture(picturePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, picLeft, picTop, picWidth, picHeight);

            }
            catch (Exception )
            {
                //MessageBox.Show("错误：" + ex.Message);
            }
        }


        public static void InsertPicture_Compress(MSExcel.Range rng, MSExcel._Worksheet sheet, string picturePath, double ratio)
        {
            //rng.Select(); 
            try
            {
                // 获取单元格的尺寸
                float picLeft = Convert.ToSingle(rng.Left);
                float picTop = Convert.ToSingle(rng.Top);
                float picWidth = Convert.ToSingle(rng.Width);
                float picHeight = Convert.ToSingle(rng.Width) * (float)ratio;

                //参数含义：
                //图片路径
                //是否链接到文件
                //图片插入时是否随文档一起保存
                //图片在文档中的坐标位置 坐标
                //图片显示的宽度和高度msoCTruemsoFalse

                // 限制图片尺寸不超过单元格
                float cellHeight = Convert.ToSingle(rng.Height);
                if (picHeight > cellHeight)
                {
                    picHeight = cellHeight;
                    picWidth = picHeight / (float)ratio; // 按比例调整宽度
                }



                // 临时文件路径用于保存处理后的图片
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".jpg");
                // 检查图片是否为灰度图
                // 加载原始图片
                using (Bitmap originalImage = new Bitmap(picturePath))
                {
                    Bitmap processedImage;

                    // 检查图片是否为灰度图
                    if (!IsGrayscale(originalImage))
                    {
                        // 对彩色图进行灰度转换
                        processedImage = new Bitmap(originalImage.Width, originalImage.Height);
                        using (Graphics g = Graphics.FromImage(processedImage))
                        {
                            float[][] colorMatrixElements = {
                                new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                                new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                                new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                                new float[] {0, 0, 0, 1, 0},
                                new float[] {0, 0, 0, 0, 1}
                            };
                            ColorMatrix colorMatrix = new ColorMatrix(colorMatrixElements);
                            ImageAttributes attributes = new ImageAttributes();
                            attributes.SetColorMatrix(colorMatrix);

                            g.DrawImage(originalImage, new System.Drawing.Rectangle(0, 0, originalImage.Width, originalImage.Height),
                                0, 0, originalImage.Width, originalImage.Height, GraphicsUnit.Pixel, attributes);
                        }
                    }
                    else
                    {
                        // 灰度图直接使用原图
                        processedImage = new Bitmap(originalImage);
                    }

                    // 调整图片尺寸以适应单元格
                    int targetWidth = (int)picWidth;
                    int targetHeight = (int)picHeight;
                    using (Bitmap resizedImage = new Bitmap(processedImage, targetWidth, targetHeight))
                    {
                        // 设置压缩参数
                        EncoderParameters encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 75L); // 压缩质量（0-100）

                        // 获取 JPEG 编码器
                        ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                            .FirstOrDefault(c => c.MimeType == "image/jpeg");

                        //保存压缩后的图片
                        resizedImage.Save(tempPath, jpegCodec, encoderParams);
                        //resizedImage.Save(tempPath);
                    }

                    processedImage.Dispose();
                }


                // 插入处理后的图片
                sheet.Shapes.AddPicture(tempPath, Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, picLeft, picTop, picWidth, picHeight);

                // 删除临时文件
                File.Delete(tempPath);

                // string sheetPath = sheet.Application.ActiveWorkbook.Path;

                //  sheet.Shapes.AddPicture(tempPath, Microsoft.Office.Core.MsoTriState.msoFalse,  Microsoft.Office.Core.MsoTriState.msoTrue, picLeft, picTop, picWidth, picHeight);
                //sheet.Shapes.AddPicture(PicturePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, PicLeft, PicTop, PicWidth, PicHeight);

            }
            catch (Exception)
            {
                //MessageBox.Show("错误：" + ex.Message);
            }
        }
    }
}
