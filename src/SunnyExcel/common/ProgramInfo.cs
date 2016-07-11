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
using System.Reflection;

#pragma warning disable 1589, 1591

namespace SunnyExcel
{
    /*
    *	ProgramInfo
    *
    *	This class is just for getting information about the application.
    *	Each assembly has a GUID, and that GUID is useful to us in this application,
    *	so the most important thing in this class is the AssemblyGuid property.
    *
    *	GetEntryAssembly() is used instead of GetExecutingAssembly(), so that you
    *	can put this code into a class library and still get the results you expect.
    *	(Otherwise it would return info on the DLL assembly instead of your application.)
    */

    // using System.Reflection;

    public class ProgramInfo
    {
        public static string AssemblyGuid
        {
            get
            {
                object[] attributes = Assembly.GetEntryAssembly().GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false);
                if (attributes.Length == 0)
                    return "";

                return ((System.Runtime.InteropServices.GuidAttribute)attributes[0]).Value;
            }
        }

        public static string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetEntryAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (String.IsNullOrEmpty(titleAttribute.Title) == false)
                        return titleAttribute.Title;
                }

                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().CodeBase);
            }
        }
    }
}