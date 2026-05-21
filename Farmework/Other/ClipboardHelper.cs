using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Farmework.Other
{
    public static class ClipboardHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsClipboardFormatAvailable(uint format);

        private const uint CF_TEXT = 1;
        private const uint CF_UNICODETEXT = 13;
        private const uint CF_HDROP = 15;

        public static bool ClipboardHasText()
        {
            bool result = false;
            if (OpenClipboard(IntPtr.Zero))
            {
                result = IsClipboardFormatAvailable(CF_TEXT) || IsClipboardFormatAvailable(CF_UNICODETEXT);
                CloseClipboard();
            }
            return result;
        }

        public static void ClearClipboard()
        {
            if (OpenClipboard(IntPtr.Zero))
            {
                System.Windows.Forms.Clipboard.Clear();
                CloseClipboard();
            }
        }
    }
}
