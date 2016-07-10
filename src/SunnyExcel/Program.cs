using System;
using System.Diagnostics;
using System.Security.Principal;
using System.ServiceModel;
using System.Threading;
using System.Windows.Forms;

namespace SunnyExcel
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (SingleInstance.Start() == false)
            {
                //MessageBox.Show("it is already running");

                SingleInstance.ShowFirstInstance();
                return;
            }

            // Add the event handler for handling UI thread exceptions to the event.
            Application.ThreadException += Application_ThreadException;

            // Set the unhandled exception mode to force all Windows Forms errors to go through our handler.
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            // Add the event handler for handling non-UI thread exceptions to the event. 
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                RunAdministrator();

                using (MainForm _main_form = new MainForm())
                    Application.Run(_main_form);
            }
            catch (EndpointNotFoundException ex)
            {
                string _message
                        = "I can not find server due to a network error or a configuration problem." + Environment.NewLine
                        + "automatically shutdown" + Environment.NewLine
                        + "Please contact your system administrator." + Environment.NewLine
                        + Environment.NewLine + ex.Message;
                MessageBox.Show(_message);
            }
            catch (Exception ex)
            {
                string _message
                        = "shutdown due to an unknown error. " + Environment.NewLine
                        + "Please contact your system administrator." + Environment.NewLine
                        + Environment.NewLine + ex.Message;
                MessageBox.Show(_message);
            }
            finally
            {
                SingleInstance.Stop();
            }
        }

        private static void RunAdministrator()
        {
            if (IsAdministrator() == false)
            {
                try
                {
                    var _proc_info = new ProcessStartInfo();
                    _proc_info.UseShellExecute = true;
                    _proc_info.FileName = Application.ExecutablePath;
                    _proc_info.WorkingDirectory = Environment.CurrentDirectory;
                    _proc_info.Verb = "runas";
                    Process.Start(_proc_info);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }

                return;
            }
        }

        public static bool IsAdministrator()
        {
            var _identity = WindowsIdentity.GetCurrent();

            if (null != _identity)
            {
                WindowsPrincipal principal = new WindowsPrincipal(_identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }

            return false;
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            var _result = DialogResult.Abort;

            try
            {
                _result = MessageBox.Show(
                        String.Format("Whoops! Please contact the developers with the following thread exception information:\n\n{0}{1}", e.Exception.Message, e.Exception.StackTrace),
                        "Application Error",
                        MessageBoxButtons.AbortRetryIgnore,
                        MessageBoxIcon.Stop
                    );
            }
            finally
            {
                if (_result == DialogResult.Abort)
                {
                    SingleInstance.Stop();
                    Application.Exit();
                }
            }
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var _ex = (Exception)e.ExceptionObject;

                MessageBox.Show(
                        String.Format("Whoops! Please contact the developers with the following unhandled exception information:\n\n{0}{1}", _ex.Message, _ex.StackTrace),
                        "Fatal Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                    );
            }
            finally
            {
                SingleInstance.Stop();
                Application.Exit();
            }
        }
    }
}