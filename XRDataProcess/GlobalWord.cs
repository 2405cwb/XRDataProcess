using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OperateIniFile;
using MSWord = Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;

namespace XRDataProcess
{
    class GlobalWord
    {
        /// <summary>
        /// 报告复制粘贴延时ms
        /// </summary>
        public static int wd_sleep_us = 1000;

        /// <summary>
        /// 报告格式文字延时ms
        /// </summary>
        public static int wd_sleep_us2 = 500;

        public static void wordAppGoTo(MSWord.Application wordApp, MSWord.Range wordrange)
        {
            //避免频繁操作word报错
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.GoToEditableRange(wordrange);
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
        public static void wordAppAlignment(MSWord.Application wordApp, MSWord.WdParagraphAlignment wdAlign)
        {
            //避免频繁操作word报错
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Paragraphs.Last.Alignment = wdAlign;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }

        public static void wordAppSize(MSWord.Application wordApp, float size)
        {
            //避免频繁操作word报错
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Paragraphs.Last.Range.Font.Size = size;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }

        public static void wordAppFontBold(MSWord.Application wordApp)
        {
            //避免频繁操作word报错
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Paragraphs.Last.Range.Bold = 1;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
        public static void wordAppFontName(MSWord.Application wordApp, string name)
        {
            //避免频繁操作word报错
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Paragraphs.Last.Range.Font.Name = name;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
        public static void wordAppTypeText(MSWord.Application wordApp, string text)
        {
            //避免频繁操作word报错
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.TypeText(text);
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Paragraphs.Add();
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.GoToNext(MSWord.WdGoToItem.wdGoToLine);
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
        public static void wordAppSelectionPaste(MSWord.Application wordApp)
        {
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Paste();
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
        public static void wordAppHeadingFormat(MSWord.Application wordApp)
        {
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Rows.HeadingFormat = -1;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
        public static void PastExcel2Word(MSWord.Application wordApp, MSExcel.Range excelrange, MSWord.Document wordDoc, String Lable)
        {
            while (true)
            {
                try
                {
                    foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                    {
                        if (book.Name == Lable)
                        {
                            System.Windows.Forms.Clipboard.Clear();  
                            excelrange.Copy();
                            Thread.Sleep(500);
                            book.Select();
                            book.Range.Paste();
                            //Thread.Sleep(GlobalWord.wd_sleep_us);
                            break;
                        }
                    }
                    break;
                }
                catch(Exception ex)
                {
                    Thread.Sleep(200);
                }
            }
            Thread.Sleep(200);
        }
        public static void PastExcel2Word(MSWord.Application wordApp, MSExcel.Shape excelshape, MSWord.Document wordDoc, String Lable)
        {
            while (true)
            {
                try
                {
                    foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                    {
                        if (book.Name == Lable)
                        {
                            System.Windows.Forms.Clipboard.Clear();
                            excelshape.Copy();
                            book.Select();
                            book.Range.PasteAndFormat(MSWord.WdRecoveryType.wdChartPicture);
                            break;
                        }
                    }
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
            Thread.Sleep(200);
        }
        
        public static MSWord.Range GetMarkRange(MSWord.Document wordDoc, String sBookmarks)
        {
            MSWord.Range result = null;
            while (true)
            {
                try
                {
                    foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                    {
                        if (book.Name == sBookmarks)
                        {
                            book.Select();
                            result = book.Range;
                            break;
                        }
                    }
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
            return result;
        }

        // 定位到文档最后一行
        public static void gotoLastLine(MSWord.Document thisDocument)
        {
            object dummy = System.Reflection.Missing.Value;
            object what = MSWord.WdGoToItem.wdGoToLine;
            object which = MSWord.WdGoToDirection.wdGoToLast;
            object count = 99999999;
            thisDocument.Application.Selection.GoTo(ref what, ref which, ref count, ref dummy);
        }
    }
}
