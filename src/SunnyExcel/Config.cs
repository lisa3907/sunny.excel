using System;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Security.Principal;
using System.Text.RegularExpressions;

#pragma warning disable 1589, 1591

namespace SunnyExcel
{
    public class CConfig
    {
        //-----------------------------------------------------------------------------------------------------------------------------
        // 
        //-----------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_key"></param>
        /// <returns></returns>
        public string GetAppValue(string p_key)
        {
            return CfgSettings.Get(p_key);
        }

        /// <summary>
        /// 
        /// </summary>
        public NameValueCollection CfgSettings
        {
            get
            {
                return ConfigurationManager.AppSettings;
            }
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        // 
        //-----------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string[] GetIPAddresses()
        {
            return GetIPAddresses(Dns.GetHostName());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_host_name">name of host PC</param>
        /// <returns></returns>
        public string[] GetIPAddresses(string p_host_name)
        {
            string[] _result;

            if (IsValidIP(p_host_name) == false)
            {
                IPAddress[] _ipadrs = Dns.GetHostEntry(p_host_name).AddressList;
                _result = new string[_ipadrs.Length];

                int i = 0;
                Array.ForEach(_ipadrs, _ip =>
                {
                    if (_ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        _result[i++] = _ip.ToString();
                });

                Array.Resize(ref _result, i);
            }
            else
            {
                _result = new string[] { p_host_name };
            }

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetIPAddress()
        {
            return GetIPAddress(Dns.GetHostName());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_host_name">name of host PC</param>
        /// <returns></returns>
        public string GetIPAddress(string p_host_name)
        {
            var _result = p_host_name;

            if (IsValidIP(p_host_name) == false)
            {
                IPHostEntry _ipHostInfo = Dns.GetHostEntry(p_host_name);
                foreach (IPAddress _ip in _ipHostInfo.AddressList)
                {
                    if (_ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        _result = _ip.ToString();
                        break;
                    }
                }
            }

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetIP6Address()
        {
            return GetIP6Address(Dns.GetHostName());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_host_name">name of host PC</param>
        /// <returns></returns>
        public string GetIP6Address(string p_host_name)
        {
            var _result = p_host_name;

            if (IsValidIP(p_host_name) == false)
            {
                IPHostEntry _ipHostInfo = Dns.GetHostEntry(p_host_name);
                foreach (IPAddress _ip in _ipHostInfo.AddressList)
                {
                    if (_ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
                    {
                        _result = _ip.ToString().Split('%')[0];
                        break;
                    }
                }
            }

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetMacAddress()
        {
            var _nics = NetworkInterface.GetAllNetworkInterfaces();

            var _mac_address = "";
            foreach (NetworkInterface _adapter in _nics)
            {
                // only return MAC Address from first card  
                if (String.IsNullOrEmpty(_mac_address) == true)
                {
                    //IPInterfaceProperties properties = adapter.GetIPProperties(); Line is not required
                    _mac_address = _adapter.GetPhysicalAddress().ToString();
                }
            }

            return _mac_address;
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        // 
        //-----------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_host_name">name of host PC</param>
        /// <returns></returns>
        public string GetLocalIpAddress(string p_host_name)
        {
            return IsLocalMachine(p_host_name) == true ? IPAddress : p_host_name;
        }

        /// <summary>
        /// method to validate an IP address
        /// using regular expressions. The pattern
        /// being used will validate an ip address
        /// with the range of 1.0.0.0 to 255.255.255.255
        /// </summary>
        /// <param name="p_ip_address">Address to validate</param>
        /// <returns></returns>
        public bool IsValidIP(string p_ip_address)
        {
            //create our match pattern
            const string _pattern = @"^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$";

            //create our Regular Expression object
            Regex _regCheck = new Regex(_pattern);

            //boolean variable to hold the status
            var _result = false;

            //check to make sure an ip address was provided
            if (p_ip_address == "")
            {
                //no address provided so return false
                _result = false;
            }
            else
            {
                //address provided so use the IsMatch Method
                //of the Regular Expression object
                _result = _regCheck.IsMatch(p_ip_address, 0);
            }

            //return the results
            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_host_name">name of host PC</param>
        /// <returns></returns>
        public bool IsLocalMachine(string p_host_name)
        {
            var _result = false;

            if (String.IsNullOrEmpty(p_host_name) == true ||
                p_host_name.ToLower() == "localhost" || p_host_name == "127.0.0.1" ||
                p_host_name.ToLower() == MachineName || p_host_name == "::1"
                )
            {
                _result = true;
            }
            else
            {
                string[] _serverIPs = GetIPAddresses(p_host_name);
                if (_serverIPs.Length > 0)
                {
                    string[] _ipadrs = GetIPAddresses();
                    foreach (string _ip in _ipadrs)
                    {
                        if (_ip == _serverIPs[0])
                        {
                            _result = true;
                            break;
                        }
                    }
                }
            }

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_local_host_name"></param>
        /// <param name="p_remote_host_name"></param>
        /// <returns></returns>
        public bool IsSameMachine(string p_local_host_name, string p_remote_host_name)
        {
            var _result = false;

            if (p_local_host_name == p_remote_host_name)
            {
                _result = true;
            }
            else
            {
                string[] _localIPs = GetIPAddresses(p_local_host_name);
                if (_localIPs.Length > 0)
                {
                    string[] _remoteIPs = GetIPAddresses(p_remote_host_name);
                    foreach (string _ip in _remoteIPs)
                    {
                        if (_ip == _localIPs[0])
                        {
                            _result = true;
                            break;
                        }
                    }
                }
            }

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        public string MachineName
        {
            get
            {
                return Environment.MachineName.ToLower();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string UserDomainName
        {
            get
            {
                return Environment.UserDomainName.ToLower();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string UserName
        {
            get
            {
                return Environment.UserName.ToLower();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool IsUnix
        {
            get
            {
                return Environment.OSVersion.Platform == PlatformID.Unix;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool IsWindows
        {
            get
            {
                return Environment.OSVersion.Platform == PlatformID.Win32NT;
            }
        }

        private static string[] __ip_addresses;

        /// <summary>
        /// 
        /// </summary>
        public string[] IPAddresses
        {
            get
            {
                if (__ip_addresses == null)
                    __ip_addresses = GetIPAddresses();

                return __ip_addresses;
            }
        }

        private static string __ip_address;

        /// <summary>
        /// 
        /// </summary>
        public string IPAddress
        {
            get
            {
                if (__ip_address == null)
                    __ip_address = GetIPAddress();

                return __ip_address;
            }
        }

        private static string __ip6_address;

        /// <summary>
        /// 
        /// </summary>
        public string IP6Address
        {
            get
            {
                if (__ip6_address == null)
                    __ip6_address = GetIP6Address();

                return __ip6_address;
            }
        }

        private static string __mac_address;

        /// <summary>
        /// 
        /// </summary>
        public string MacAddress
        {
            get
            {
                if (__mac_address == null)
                    __mac_address = GetMacAddress();

                return __mac_address;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_assembly_string"></param>
        /// <returns></returns>
        public string GetVersion(string p_assembly_string)
        {
            string[] _splitNames = p_assembly_string.Split(',');
            string _version = "";

            foreach (string f in _splitNames)
            {
                if (f.Trim().StartsWith("Version="))
                    _version = f.Trim().Replace("Version=", "");
            }

            return _version;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_assembly_string"></param>
        /// <returns></returns>
        public string GetPublicKeyToken(string p_assembly_string)
        {
            string[] _splitNames = p_assembly_string.Split(',');
            string _keyToken = "";

            foreach (string f in _splitNames)
            {
                if (f.Trim().StartsWith("PublicKeyToken="))
                    _keyToken = f.Trim().Replace("PublicKeyToken=", "").ToLower();
            }

            return _keyToken;
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        //
        //-----------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetCommonFolder()
        {
            var _result = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
                        + @"\OdinSoft\";

            if (Directory.Exists(_result) == false)
                Directory.CreateDirectory(_result);

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_folder"></param>
        /// <returns></returns>
        public string GetWorkingFolder(string p_folder)
        {
            var _result = GetCommonFolder();

            if (String.IsNullOrEmpty(p_folder) == false)
                _result += p_folder + @"\";

            if (Directory.Exists(_result) == false)
                Directory.CreateDirectory(_result);

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_productId">제품 ID</param>
        /// <param name="p_folder"></param>
        /// <returns></returns>
        public string GetWorkingFolder(string p_productId, string p_folder)
        {
            var _result = GetCommonFolder();

            _result += p_productId + @"\";

            if (String.IsNullOrEmpty(p_folder) == false)
                _result += p_folder + @"\";

            if (Directory.Exists(_result) == false)
                Directory.CreateDirectory(_result);

            return _result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetDownloadsFolder(string p_sub_folder = "")
        {
            var _result = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                        + @"\Downloads\";

            if (String.IsNullOrEmpty(p_sub_folder) == false)
                _result = Path.Combine(_result, p_sub_folder);

            if (Directory.Exists(_result) == false)
                Directory.CreateDirectory(_result);

            return _result;
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        // 영어(미국), 일본어(일본), 중국어(중국), 중국어(간체), 중국어(번체), 한국어(대한민국), 태국어(태국). 필리핀어(필리핀)
        //-----------------------------------------------------------------------------------------------------------------------------
        private static ArrayList __culture = new ArrayList(new string[] { "ko-kr", "en-us", "ja-jp", "zh-cn", "zh-chs", "zh-cht", "th-th", "fil-ph" });

        /// <summary>
        /// 
        /// </summary>
        public string GetDefaultCulture
        {
            get
            {
                return __culture[0].ToString();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_culture"></param>
        /// <returns></returns>
        public bool ContainsCultureName(string p_culture)
        {
            return __culture.Contains(p_culture.ToLower());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_culture"></param>
        /// <returns></returns>
        public string GetCultureName(string p_culture)
        {
            var _result = p_culture;

            for (int i = 0; i < __culture.Count; i++)
            {
                if (_result.ToLower().Equals(__culture[i].ToString()) == true)
                    return _result;
            }

            switch (_result.ToLower().Substring(0, 2))
            {
                case "ko":
                    _result = __culture[0].ToString();
                    break;

                case "en":
                    _result = __culture[1].ToString();
                    break;

                case "ja":
                    _result = __culture[2].ToString();
                    break;

                case "zh":
                    _result = __culture[3].ToString();
                    break;

                default:
                    _result = __culture[1].ToString();
                    break;
            }

            return _result;
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        //
        //-----------------------------------------------------------------------------------------------------------------------------
        private static bool? _runningFromNUnit = null;

        /// <summary>
        /// Unit Test 중인지를 확인 합니다.
        /// </summary>
        public bool IsRunningFromNunit
        {
            get
            {
                if (_runningFromNUnit == null)
                {
                    _runningFromNUnit = false;

                    const string _testAssemblyName = "Microsoft.VisualStudio.QualityTools.UnitTestFramework";
                    foreach (Assembly _assembly in AppDomain.CurrentDomain.GetAssemblies())
                    {
                        if (_assembly.FullName.StartsWith(_testAssemblyName))
                        {
                            _runningFromNUnit = true;
                            break;
                        }
                    }
                }

                return _runningFromNUnit.Value;
            }
        }

        /// <summary>
        /// SQL ConnectionString에 application name을 추가 합니다.
        /// </summary>
        /// <param name="p_connection_string"></param>
        /// <param name="p_application_name"></param>
        /// <returns></returns>
        public string AppendAppNameInConnectionString(string p_connection_string, string p_application_name)
        {
            var _result = "";

            const string _cApplication = "application name";

            string[] _nodes = p_connection_string.Split(';');
            foreach (string _node in _nodes)
            {
                if (_node.Split('=')[0].ToLower() != _cApplication)
                    _result += _node + ";";
            }

            _result += String.Format("{0}={1};", _cApplication, p_application_name);

            return _result;
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        //
        //-----------------------------------------------------------------------------------------------------------------------------

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsAdministrator()
        {
            var _result = false;

            WindowsIdentity _identity = WindowsIdentity.GetCurrent();
            if (_identity != null)
            {

                WindowsPrincipal principal = new WindowsPrincipal(_identity);
                _result = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }

            return _result;
        }

        //-----------------------------------------------------------------------------------------------------------------------------
        //
        //-----------------------------------------------------------------------------------------------------------------------------
    }
}