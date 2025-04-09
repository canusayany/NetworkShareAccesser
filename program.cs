using System.Runtime.InteropServices;
using System.Diagnostics;
using System.ComponentModel;
using System.Net;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp56;

/// <summary>
/// Windows网络共享访问类
/// </summary>
public class NetworkShareAccesser : IDisposable
{
    private string _networkName;
    private string _username;
    private string _password;
    private string _domain;
    private bool _connected;
    private bool _forceConnection;

    // 连接标志
    private const int CONNECT_INTERACTIVE = 0x00000008;
    private const int CONNECT_PROMPT = 0x00000010;
    private const int CONNECT_UPDATE_PROFILE = 0x00000001;
    private const int CONNECT_REDIRECT = 0x00000080;

    // 资源类型
    private const int RESOURCETYPE_DISK = 0x00000001;

    [DllImport("mpr.dll")]
    private static extern int WNetAddConnection2(NETRESOURCE netResource, string password, string username, int flags);

    [DllImport("mpr.dll")]
    private static extern int WNetAddConnection3(IntPtr hwndOwner, NETRESOURCE netResource, string password, string username, int flags);

    [DllImport("mpr.dll")]
    private static extern int WNetCancelConnection2(string name, int flags, bool force);

    [DllImport("mpr.dll")]
    private static extern int WNetGetLastError(out int errorCode, StringBuilder errorDesc, int errorDescSize, StringBuilder sourceName, int sourceNameSize);

    [StructLayout(LayoutKind.Sequential)]
    private class NETRESOURCE
    {
        public int dwScope = 0;
        public int dwType = 0;
        public int dwDisplayType = 0;
        public int dwUsage = 0;
        public string lpLocalName = "";
        public string lpRemoteName = "";
        public string lpComment = "";
        public string lpProvider = "";
    }

    /// <summary>
    /// 创建网络共享访问实例
    /// </summary>
    /// <param name="networkName">网络共享路径，如\\192.168.1.150\New</param>
    /// <param name="username">用户名</param>
    /// <param name="password">密码</param>
    /// <param name="domain">域名，如果不在域中可以为空</param>
    /// <param name="forceConnection">如果已存在连接是否强制重新连接</param>
    public NetworkShareAccesser(string networkName, string username, string password, string domain = "", bool forceConnection = false)
    {
        _networkName = networkName;
        _username = username;
        _password = password;
        _domain = domain;
        _forceConnection = forceConnection;
    }

    /// <summary>
    /// 获取详细的错误信息
    /// </summary>
    private string GetErrorMessage(int errorCode)
    {
        StringBuilder errorDesc = new StringBuilder(1024);
        StringBuilder sourceName = new StringBuilder(1024);
        int result = WNetGetLastError(out int _, errorDesc, errorDesc.Capacity, sourceName, sourceName.Capacity);

        if (result == 0)
        {
            return $"错误代码: {errorCode}, 描述: {errorDesc.ToString()}, 来源: {sourceName.ToString()}";
        }
        else
        {
            return $"错误代码: {errorCode}, 无法获取详细描述";
        }
    }

    /// <summary>
    /// 连接到网络共享
    /// </summary>
    /// <param name="useAlternativeMethods">是否尝试替代连接方法</param>
    /// <returns>连接是否成功</returns>
    public bool Connect(bool useAlternativeMethods = true)
    {
        if (_forceConnection)
        {
            // 强制断开现有连接
            try { WNetCancelConnection2(_networkName, 0, true); } catch { }
        }

        var netResource = new NETRESOURCE
        {
            dwType = RESOURCETYPE_DISK,
            lpRemoteName = _networkName
        };

        // 准备用户名
        string user = _username;
        if (!string.IsNullOrEmpty(_domain))
        {
            user = $"{_domain}\\{_username}";
        }

        Console.WriteLine($"尝试使用用户: {user} 连接到: {_networkName}");

        // 第一种方法: 基本连接
        int result = WNetAddConnection2(netResource, _password, user, 0);
        _connected = result == 0;

        if (!_connected && useAlternativeMethods)
        {
            // 记录错误信息
            string errorMsg = GetErrorMessage(result);
            Console.WriteLine($"基本连接方法失败: {errorMsg}");

            // 第二种方法: 使用交互式标志
            Console.WriteLine("尝试使用交互式连接...");
            result = WNetAddConnection2(netResource, _password, user, CONNECT_INTERACTIVE);
            _connected = result == 0;

            if (!_connected)
            {
                // 记录错误信息
                errorMsg = GetErrorMessage(result);
                Console.WriteLine($"交互式连接方法失败: {errorMsg}");

                // 第三种方法: 使用更新配置文件标志
                Console.WriteLine("尝试使用更新配置文件标志连接...");
                result = WNetAddConnection2(netResource, _password, user, CONNECT_UPDATE_PROFILE);
                _connected = result == 0;

                if (!_connected)
                {
                    errorMsg = GetErrorMessage(result);
                    Console.WriteLine($"更新配置文件连接方法失败: {errorMsg}");

                    // 第四种方法：不使用凭据直接连接
                    Console.WriteLine("尝试不使用凭据直接连接...");
                    result = WNetAddConnection2(netResource, null, null, 0);
                    _connected = result == 0;

                    if (!_connected)
                    {
                        errorMsg = GetErrorMessage(result);
                        throw new Win32Exception(result, $"无法连接到网络共享 {_networkName}: {errorMsg}");
                    }
                }
            }
        }
        else if (!_connected)
        {
            string errorMsg = GetErrorMessage(result);
            throw new Win32Exception(result, $"无法连接到网络共享 {_networkName}: {errorMsg}");
        }

        return _connected;
    }

    /// <summary>
    /// 尝试直接访问共享路径不创建连接
    /// </summary>
    public bool TryDirectAccess()
    {
        try
        {
            // 尝试直接列出文件
            string[] files = Directory.GetFiles(_networkName);
            Console.WriteLine($"直接访问成功，找到 {files.Length} 个文件");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"直接访问失败: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 断开网络共享连接
    /// </summary>
    public void Disconnect()
    {
        if (_connected)
        {
            WNetCancelConnection2(_networkName, 0, true);
            _connected = false;
        }
    }

    /// <summary>
    /// 获取共享文件夹中的所有文件
    /// </summary>
    /// <returns>文件列表</returns>
    public List<string> GetFiles()
    {
        try
        {
            return new List<string>(Directory.GetFiles(_networkName));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"获取文件列表失败: {ex.Message}");
            return new List<string>();
        }
    }

    /// <summary>
    /// 获取共享文件夹中的所有子文件夹
    /// </summary>
    /// <returns>文件夹列表</returns>
    public List<string> GetDirectories()
    {
        try
        {
            return new List<string>(Directory.GetDirectories(_networkName));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"获取目录列表失败: {ex.Message}");
            return new List<string>();
        }
    }

    /// <summary>
    /// 读取文件内容
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <returns>文件内容</returns>
    public string ReadFileText(string fileName)
    {
        string filePath = Path.Combine(_networkName, fileName);
        return File.ReadAllText(filePath);
    }

    /// <summary>
    /// 实现IDisposable接口
    /// </summary>
    public void Dispose()
    {
        Disconnect();
    }
}

/// <summary>
/// 程序入口类
/// </summary>
public class Program
{
    /// <summary>
    /// 主函数
    /// </summary>
    public static void Main(string[] args)
    {
        // 网络共享配置信息
        string ip = "192.168.1.150";
        string username = "*";
        string password = "*";
        string shareName = @"Users";
        string networkPath = $"\\\\{ip}\\{shareName}";

        Console.WriteLine($"尝试连接到 {networkPath}");
        Console.WriteLine("当前运行的Windows账户: " + Environment.UserName);

        try
        {
            // 先尝试直接访问
            Console.WriteLine("\n=== 尝试直接访问共享路径 ===");
            bool directAccessSuccess = false;
            try
            {
                string[] files = Directory.GetFiles(networkPath);
                Console.WriteLine($"直接访问成功，找到 {files.Length} 个文件");
                directAccessSuccess = true;

                Console.WriteLine("\n文件列表:");
                foreach (var file in files)
                {
                    Console.WriteLine($"- {Path.GetFileName(file)}");
                }

                string[] dirs = Directory.GetDirectories(networkPath);
                Console.WriteLine($"\n找到 {dirs.Length} 个目录");
                Console.WriteLine("\n目录列表:");
                foreach (var dir in dirs)
                {
                    Console.WriteLine($"- {Path.GetFileName(dir)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"直接访问失败: {ex.Message}");
            }

            if (!directAccessSuccess)
            {
                // 使用NetworkShareAccesser尝试连接
                Console.WriteLine("\n=== 尝试使用NetworkShareAccesser连接 ===");
                
                // 尝试不同的凭据组合
                var credentialOptions = new List<(string username, string password, string domain, string desc)>
                {
                    (username, password, "", "标准凭据"),
                    ("", "", "", "无凭据"),
                    (Environment.UserName, "", "", "当前Windows用户"),
                    (username, password, ip, "使用IP作为域的凭据")
                };

                bool connected = false;
                foreach (var option in credentialOptions)
                {
                    if (connected) break;
                    
                    Console.WriteLine($"\n尝试使用 {option.desc}");
                    using (var accesser = new NetworkShareAccesser(networkPath, option.username, option.password, option.domain, true))
                    {
                        try
                        {
                            connected = accesser.Connect();
                            if (connected)
                            {
                                Console.WriteLine($"使用 {option.desc} 连接成功！");
                                
                                // 获取并显示文件列表
                                Console.WriteLine("\n文件列表:");
                                var files = accesser.GetFiles();
                                if (files.Count > 0)
                                {
                                    foreach (var file in files)
                                    {
                                        Console.WriteLine($"- {Path.GetFileName(file)}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("没有找到文件");
                                }

                                // 获取并显示文件夹列表
                                Console.WriteLine("\n文件夹列表:");
                                var directories = accesser.GetDirectories();
                                if (directories.Count > 0)
                                {
                                    foreach (var dir in directories)
                                    {
                                        Console.WriteLine($"- {Path.GetFileName(dir)}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("没有找到文件夹");
                                }

                                // 尝试读取一个文件（如果有）
                                if (files.Count > 0)
                                {
                                    string firstFile = Path.GetFileName(files[0]);
                                    Console.WriteLine($"\n读取文件 {firstFile} 的内容:");
                                    try
                                    {
                                        string content = accesser.ReadFileText(firstFile);
                                        Console.WriteLine(content.Length > 100 ? content.Substring(0, 100) + "..." : content);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"读取文件失败: {ex.Message}");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"使用 {option.desc} 连接失败: {ex.Message}");
                        }
                    }
                }

                if (!connected)
                {
                    Console.WriteLine("\n所有连接方法都失败了！");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"错误: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"内部错误: {ex.InnerException.Message}");
            }
        }

        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }
}
