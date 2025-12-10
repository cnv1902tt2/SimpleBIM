// SimpleBIM/Update/NativeMethods.cs
using System;
using System.Runtime.InteropServices;

namespace SimpleBIM.Update
{
    internal static class NativeMethods
    {
        // Đưa cửa sổ lên trên cùng
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        // Cấu trúc cho FlashWindowEx
        [StructLayout(LayoutKind.Sequential)]
        public struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        // Hằng số
        private const int SW_RESTORE = 9;
        private const int HWND_TOPMOST = -1;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;

        // Flash flags
        private const uint FLASHW_ALL = 0x00000003;        // Nhấp nháy title bar + taskbar
        private const uint FLASHW_TIMERNOFG = 0x0000000C;  // Nhấp nháy liên tục cho đến khi foreground

        /// <summary>
        /// Đưa cửa sổ lên trên cùng + nhấp nháy taskbar (BẮT BUỘC PHẢI THẤY)
        /// </summary>
        public static void BringToFrontAndFlash(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;

            try
            {
                // 1. Restore nếu bị minimize
                ShowWindow(hWnd, SW_RESTORE);

                // 2. Đưa lên trên cùng
                SetForegroundWindow(hWnd);
                SetWindowPos(hWnd, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0,
                             SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

                // 3. Nhấp nháy taskbar (rất bắt mắt)
                var fInfo = new FLASHWINFO
                {
                    cbSize = (uint)Marshal.SizeOf(typeof(FLASHWINFO)),
                    hwnd = hWnd,
                    dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG, // Nhấp nháy cả title + taskbar
                    uCount = 8,     // Số lần nhấp nháy
                    dwTimeout = 0   // Tốc độ mặc định
                };

                FlashWindowEx(ref fInfo);
            }
            catch
            {
                // Không crash nếu user32 lỗi (ví dụ: chạy trên Wine, RDP...)
            }
        }
    }
}