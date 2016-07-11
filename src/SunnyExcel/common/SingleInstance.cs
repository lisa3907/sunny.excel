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
using System.Threading;

#pragma warning disable 1589, 1591

namespace SunnyExcel
{
    /* All of the code below can optionally be put in a class library and reused with all your applications. */

    /*
    *	SingeInstance
    *
    *	This is where the magic happens.
    *
    *	Start() tries to create a mutex.
    *	If it detects that another instance is already using the mutex, then it returns FALSE.
    *	Otherwise it returns TRUE.
    *	(Notice that a GUID is used for the mutex name, which is a little better than using the application name.)
    *
    *	If another instance is detected, then you can use ShowFirstInstance() to show it
    *	(which will work as long as you override WndProc as shown above).
    *
    *	ShowFirstInstance() broadcasts a message to all windows.
    *	The message is WM_SHOWFIRSTINSTANCE.
    *	(Notice that a GUID is used for WM_SHOWFIRSTINSTANCE.
    *	That allows you to reuse this code in multiple applications without getting
    *	strange results when you run them all at the same time.)
    *
    */

    // using System.Threading;

    public class SingleInstance
    {
        public static readonly int WM_QUERYENDSESSION = 0x11;

        public static readonly int WM_SHOWFIRSTINSTANCE = WinApi.RegisterWindowMessage("WM_SHOWFIRSTINSTANCE|{0}", ProgramInfo.AssemblyGuid);
        public static Mutex mutex = null;

        public static bool Start()
        {
            bool onlyInstance = false;
            string mutexName = String.Format("Local\\{0}", ProgramInfo.AssemblyGuid);

            // if you want your app to be limited to a single instance
            // across ALL SESSIONS (multiple users & terminal services), then use the following line instead:
            // string mutexName = String.Format("Global\\{0}", ProgramInfo.AssemblyGuid);

            mutex = new Mutex(true, mutexName, out onlyInstance);
            return onlyInstance;
        }

        public static void ShowFirstInstance()
        {
            WinApi.PostMessage(
                (IntPtr)WinApi.HWND_BROADCAST,
                WM_SHOWFIRSTINSTANCE,
                IntPtr.Zero,
                IntPtr.Zero);
        }

        public static void Stop()
        {
            if (mutex != null)
            {
                mutex.ReleaseMutex();
                mutex = null;
            }
        }
    }
}