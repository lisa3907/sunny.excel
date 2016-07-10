/*
*	Singe Instance Mutex Sample (version 2)
*
*	Here is code that demonstrates one way of making a single instance application.
*
*	I've crammed all this code into a single file to make it easy for you read
*	to get an overview.
*
*	In a real application, you'd want to split these classes into multiple files.
*
*	You can use all this code directly in one of your applications, or put some
*	of the code in a class library so that you can easily reuse it in all your applications.
*	The code is designed to work either way.
*/

using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

#pragma warning disable 1589, 1591

namespace SunnyExcel
{
    /*
        *	WinApi
        *
        *	This class is just a wrapper for your various WinApi functions.
        *
        *	In this sample only the bare essentials are included.
        *	In my own WinApi class, I have all the WinApi functions that any
        *	of my applications would ever need.
        *
        */

    // using System.Runtime.InteropServices;

    public class WinApi
    {
        //-------------------------------------------------------------------------------------------------------------------------
        //
        //-------------------------------------------------------------------------------------------------------------------------
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        public struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        //Stop flashing. The system restores the window to its original state. 
        public const uint FLASHW_STOP = 0;

        //Flash the window caption. 
        public const uint FLASHW_CAPTION = 1;

        //Flash the taskbar button. 
        public const uint FLASHW_TRAY = 2;

        //Flash both the window caption and taskbar button.
        //This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags. 
        public const uint FLASHW_ALL = 3;

        //Flash continuously, until the FLASHW_STOP flag is set. 
        public const uint FLASHW_TIMER = 4;

        //Flash continuously until the window comes to the foreground. 
        public const uint FLASHW_TIMERNOFG = 12;

        /// Minor adjust to the code above
        /// <summary>
        /// Flashes a window until the window comes to the foreground
        /// Receives the form that will flash
        /// </summary>
        /// <param name="p_form">The handle to the window to flash</param>
        /// <returns>whether or not the window needed flashing</returns>
        public static bool FlashWindowEx(Form p_form)
        {
            IntPtr _hWnd = p_form.Handle;
            FLASHWINFO _fInfo = new FLASHWINFO();

            _fInfo.cbSize = Convert.ToUInt32(Marshal.SizeOf(_fInfo));
            _fInfo.hwnd = _hWnd;
            _fInfo.dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG;
            _fInfo.uCount = uint.MaxValue;
            _fInfo.dwTimeout = 0;

            return FlashWindowEx(ref _fInfo);
        }

        //-------------------------------------------------------------------------------------------------------------------------
        //
        //-------------------------------------------------------------------------------------------------------------------------
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);

        public static int RegisterWindowMessage(string format, params object[] args)
        {
            string message = String.Format(format, args);
            return RegisterWindowMessage(message);
        }

        public const int HWND_BROADCAST = 0xffff;
        public const int SW_SHOWNORMAL = 1;

        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);

        [DllImportAttribute("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImportAttribute("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public static void ShowToFront(IntPtr window)
        {
            ShowWindow(window, SW_SHOWNORMAL);
            SetForegroundWindow(window);
        }

        //-------------------------------------------------------------------------------------------------------------------------
        //
        //-------------------------------------------------------------------------------------------------------------------------
    }
}