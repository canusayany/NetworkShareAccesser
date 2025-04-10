using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace BR.ECS.DeviceDriver.HTBA.HouzeQSEP400;
/// <summary>
/// 网络共享配置类
/// </summary>
public class ShareConfig
{
    /// <summary>
    /// 远程服务器IP地址
    /// </summary>
    public string RemoteIP { get; set; } = "192.168.1.150";

    /// <summary>
    /// 用户名
    /// </summary>
    public string Username { get; set; } = "bioyond";

    /// <summary>
    /// 密码
    /// </summary>
    public string Password { get; set; } = "bioyond";

    /// <summary>
    /// 共享文件夹名称
    /// </summary>
    public string ShareName { get; set; } = @"Result";

    /// <summary>
    /// 本地保存基础路径
    /// </summary>
    public string LocalBasePath { get; set; } = "./";

    /// <summary>
    /// 配置文件路径
    /// </summary>
    private static readonly string ConfigFilePath = Path.Combine(
        (AppDomain.CurrentDomain.BaseDirectory),
        "config.json");

    /// <summary>
    /// 获取网络路径
    /// </summary>
    public string GetNetworkPath()
    {
        return $"\\\\{RemoteIP}\\{ShareName}";
    }

    /// <summary>
    /// 获取或创建本地保存路径
    /// </summary>
    public string GetLocalBasePath()
    {
        if (string.IsNullOrEmpty(LocalBasePath))
        {
            LocalBasePath = Path.Combine(
        (AppDomain.CurrentDomain.BaseDirectory)

            );
        }

        // 确保目录存在
        if (!Directory.Exists(LocalBasePath))
        {
            Directory.CreateDirectory(LocalBasePath);
        }

        return LocalBasePath;
    }

    /// <summary>
    /// 保存配置到文件
    /// </summary>
    public void SaveConfig()
    {
        try
        {
            // 确保目录存在
            string configDir = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }

            // 序列化配置到JSON文件
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string jsonString = JsonSerializer.Serialize(this, options);
            File.WriteAllText(ConfigFilePath, jsonString);

            Console.WriteLine($"配置已保存到: {ConfigFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"保存配置时出错: {ex.Message}");
        }
    }

    /// <summary>
    /// 从文件加载配置
    /// </summary>
    /// <returns>加载的配置</returns>
    public static ShareConfig LoadConfig()
    {
        try
        {
            if (File.Exists(ConfigFilePath))
            {
                string jsonString = File.ReadAllText(ConfigFilePath);
                var config = JsonSerializer.Deserialize<ShareConfig>(jsonString);
                Console.WriteLine("已成功加载配置文件");
                return config;
            }
            else
            {
                Console.WriteLine("配置文件不存在，使用默认配置");
                var defaultConfig = new ShareConfig();
                defaultConfig.SaveConfig();
                return defaultConfig;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"加载配置时出错: {ex.Message}");
            return new ShareConfig();
        }
    }
}

/// <summary>
/// Windows网络共享访问封装类
/// </summary>
public class NetworkShareManager : IDisposable
{
    private ShareConfig _config;
    private NetworkShareAccesser _accesser;
    private bool _isConnected = false;
    private bool _disposed = false;

    /// <summary>
    /// 创建共享管理器实例
    /// </summary>
    /// <param name="useConfigFile">是否使用配置文件</param>
    public NetworkShareManager(bool useConfigFile = true)
    {
        if (useConfigFile)
        {
            _config = ShareConfig.LoadConfig();
        }
        else
        {
            _config = new ShareConfig();
        }

        InitializeAccesser();
    }

    /// <summary>
    /// 使用指定配置创建共享管理器实例
    /// </summary>
    /// <param name="config">共享配置</param>
    public NetworkShareManager(ShareConfig config)
    {
        _config = config ?? new ShareConfig();
        InitializeAccesser();
    }

    /// <summary>
    /// 初始化网络访问器
    /// </summary>
    private void InitializeAccesser()
    {
        string networkPath = _config.GetNetworkPath();
        _accesser = new NetworkShareAccesser(
            networkPath,
            _config.Username,
            _config.Password
        );
    }

    /// <summary>
    /// 尝试连接到共享
    /// </summary>
    /// <returns>是否连接成功</returns>
    public bool Connect()
    {
        try
        {
            Console.WriteLine($"尝试连接到: {_config.GetNetworkPath()}");

            try
            {
                // 首先尝试直接访问
                string[] files = Directory.GetFiles(_config.GetNetworkPath());
                Console.WriteLine($"直接访问成功，找到 {files.Length} 个文件");
                _isConnected = true;
                return true;
            }
            catch
            {
                // 如果直接访问失败，尝试使用凭据连接
                _isConnected = _accesser.Connect();
                if (_isConnected)
                {
                    Console.WriteLine("连接成功！");
                }
                return _isConnected;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"连接失败: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 断开连接
    /// </summary>
    public void Disconnect()
    {
        if (_isConnected)
        {
            _accesser.Disconnect();
            _isConnected = false;
            Console.WriteLine("已断开连接");
        }
    }

    /// <summary>
    /// 获取文件列表
    /// </summary>
    /// <returns>文件列表</returns>
    public List<string> GetFiles()
    {
        EnsureConnected();
        return _accesser.GetFiles();
    }

    /// <summary>
    /// 获取目录列表
    /// </summary>
    /// <returns>目录列表</returns>
    public List<string> GetDirectories()
    {
        EnsureConnected();
        return _accesser.GetDirectories();
    }

    /// <summary>
    /// 查找指定后缀的文件并按照时间排序
    /// </summary>
    /// <param name="extension">文件后缀，如".txt"、".csv"等，不区分大小写</param>
    /// <param name="count">要获取的文件数量，默认为10</param>
    /// <param name="ascending">是否按时间升序排序（旧的文件在前），默认为false（新的文件在前）</param>
    /// <param name="searchOption">搜索选项，默认只搜索当前文件夹</param>
    /// <returns>排序后的文件信息列表</returns>
    public List<FileInfo> FindFilesByExtension(string extension, int count = 10, bool ascending = false, SearchOption searchOption = SearchOption.TopDirectoryOnly)
    {
        EnsureConnected();
        return _accesser.FindFilesByExtension(extension, count, ascending, searchOption);
    }

    /// <summary>
    /// 确保已连接
    /// </summary>
    private void EnsureConnected()
    {
        if (!_isConnected)
        {
            if (!Connect())
            {
                throw new InvalidOperationException("未连接到网络共享");
            }
        }
    }

    /// <summary>
    /// 将网络共享中的文件复制到本地
    /// </summary>
    /// <param name="remoteFileName">远程文件名</param>
    /// <param name="localSubPath">本地子路径，相对于基础路径</param>
    /// <param name="overwrite">是否覆盖</param>
    /// <returns>是否成功</returns>
    public static bool CopyFileToLocal(string networkPath, string username, string password, string remoteFileName, string localPath, bool overwrite = true)
    {
        try
        {
            using (var accesser = new NetworkShareAccesser(networkPath, username, password))
            {
                if (accesser.Connect())
                {
                    string remoteFilePath = Path.Combine(networkPath, remoteFileName);

                    // 确保目标目录存在
                    string localDirectory = Path.GetDirectoryName(localPath);
                    if (!Directory.Exists(localDirectory))
                    {
                        Directory.CreateDirectory(localDirectory);
                    }

                    File.Copy(remoteFilePath, localPath, overwrite);

                    // 验证文件是否成功复制
                    if (File.Exists(localPath))
                    {
                        Console.WriteLine($"文件 {remoteFileName} 已成功复制到 {localPath}");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine($"文件复制失败: {remoteFileName}");
                        return false;
                    }
                }
                else
                {
                    Console.WriteLine($"无法连接到网络共享: {networkPath}");
                    return false;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"复制文件时出错: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 复制指定后缀和时间条件的文件到本地
    /// </summary>
    /// <param name="networkPath">网络共享路径</param>
    /// <param name="username">用户名</param>
    /// <param name="password">密码</param>
    /// <param name="extension">文件后缀</param>
    /// <param name="count">要复制的文件数量</param>
    /// <param name="localDirectoryPath">本地保存目录</param>
    /// <param name="ascending">是否按时间升序排序</param>
    /// <param name="overwrite">是否覆盖已存在的文件</param>
    /// <param name="searchOption">搜索选项，默认搜索所有子文件夹</param>
    /// <returns>成功复制的文件数量</returns>
    public static int CopyLatestFilesByExtension(string networkPath, string username, string password,
                                               string extension, int count, string localDirectoryPath,
                                               bool ascending = false, bool overwrite = true,
                                               SearchOption searchOption = SearchOption.AllDirectories)
    {
        int successCount = 0;

        try
        {
            using (var accesser = new NetworkShareAccesser(networkPath, username, password))
            {
                if (accesser.Connect())
                {
                    // 查找符合条件的文件
                    var files = accesser.FindFilesByExtension(extension, count, ascending, searchOption);

                    if (files.Count == 0)
                    {
                        Console.WriteLine($"没有找到符合条件的 {extension} 文件");
                        return 0;
                    }

                    // 确保本地目录存在
                    if (!Directory.Exists(localDirectoryPath))
                    {
                        Directory.CreateDirectory(localDirectoryPath);
                    }

                    // 复制每个文件
                    foreach (var file in files)
                    {
                        try
                        {
                            // 获取相对于网络共享根目录的路径
                            string relativePath = file.FullName.Substring(networkPath.Length).TrimStart('\\');
                            // 创建本地子目录结构
                            string localSubDir = Path.GetDirectoryName(relativePath);

                            if (!string.IsNullOrEmpty(localSubDir))
                            {
                                string fullLocalSubDir = Path.Combine(localDirectoryPath, localSubDir);
                                if (!Directory.Exists(fullLocalSubDir))
                                {
                                    Directory.CreateDirectory(fullLocalSubDir);
                                }
                            }

                            // 构建完整的本地文件路径
                            string localFilePath = Path.Combine(localDirectoryPath, relativePath);

                            // 复制文件
                            File.Copy(file.FullName, localFilePath, overwrite);

                            // 验证复制是否成功
                            if (File.Exists(localFilePath))
                            {
                                Console.WriteLine($"文件 {relativePath} 已成功复制到 {localFilePath}");
                                successCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"复制文件 {file.FullName} 时出错: {ex.Message}");
                        }
                    }

                    Console.WriteLine($"成功复制 {successCount}/{files.Count} 个文件到 {localDirectoryPath}");
                }
                else
                {
                    Console.WriteLine($"无法连接到网络共享: {networkPath}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"复制文件时出错: {ex.Message}");
        }

        return successCount;
    }

    /// <summary>
    /// 复制指定后缀和时间条件的文件到本地（使用配置文件）
    /// </summary>
    /// <param name="extension">文件后缀</param>
    /// <param name="count">要复制的文件数量</param>
    /// <param name="localDirectoryPath">本地保存目录</param>
    /// <param name="ascending">是否按时间升序排序</param>
    /// <param name="overwrite">是否覆盖已存在的文件</param>
    /// <param name="searchOption">搜索选项，默认搜索所有子文件夹</param>
    /// <returns>成功复制的文件数量</returns>
    public static int CopyLatestFilesByExtension(
        string extension, int count, string localDirectoryPath,
        bool ascending = false, bool overwrite = true,
        SearchOption searchOption = SearchOption.AllDirectories)
    {
        // 加载配置
        var config = ShareConfig.LoadConfig();
        string networkPath = config.GetNetworkPath();

        // 调用完整版本的方法
        return CopyLatestFilesByExtension(
            networkPath,
            config.Username,
            config.Password,
            extension,
            count,
            localDirectoryPath,
            ascending,
            overwrite,
            searchOption
        );
    }

    /// <summary>
    /// 复制指定后缀和时间条件的文件到本地（使用配置对象）
    /// </summary>
    /// <param name="config">共享配置对象</param>
    /// <param name="extension">文件后缀</param>
    /// <param name="count">要复制的文件数量</param>
    /// <param name="localDirectoryPath">本地保存目录</param>
    /// <param name="ascending">是否按时间升序排序</param>
    /// <param name="overwrite">是否覆盖已存在的文件</param>
    /// <param name="searchOption">搜索选项，默认搜索所有子文件夹</param>
    /// <returns>成功复制的文件数量</returns>
    public static int CopyLatestFilesByExtension(
        ShareConfig config,
        string extension, int count, string localDirectoryPath,
        bool ascending = false, bool overwrite = true,
        SearchOption searchOption = SearchOption.AllDirectories)
    {
        if (config == null)
        {
            config = ShareConfig.LoadConfig();
        }
        string networkPath = config.GetNetworkPath();

        // 调用完整版本的方法
        return CopyLatestFilesByExtension(
            networkPath,
            config.Username,
            config.Password,
            extension,
            count,
            localDirectoryPath,
            ascending,
            overwrite,
            searchOption
        );
    }

    /// <summary>
    /// 保存当前配置到文件
    /// </summary>
    public void SaveConfig()
    {
        _config.SaveConfig();
    }

    /// <summary>
    /// 更新配置
    /// </summary>
    /// <param name="ip">IP地址</param>
    /// <param name="username">用户名</param>
    /// <param name="password">密码</param>
    /// <param name="shareName">共享名</param>
    /// <param name="localBasePath">本地保存路径</param>
    /// <param name="saveToFile">是否保存到文件</param>
    public void UpdateConfig(string ip = null, string username = null, string password = null, string shareName = null, string localBasePath = null, bool saveToFile = true)
    {
        bool changed = false;

        if (ip != null && _config.RemoteIP != ip)
        {
            _config.RemoteIP = ip;
            changed = true;
        }

        if (username != null && _config.Username != username)
        {
            _config.Username = username;
            changed = true;
        }

        if (password != null && _config.Password != password)
        {
            _config.Password = password;
            changed = true;
        }

        if (shareName != null && _config.ShareName != shareName)
        {
            _config.ShareName = shareName;
            changed = true;
        }

        if (localBasePath != null && _config.LocalBasePath != localBasePath)
        {
            _config.LocalBasePath = localBasePath;
            changed = true;
        }

        if (changed)
        {
            // 更新访问器
            Disconnect();
            InitializeAccesser();

            if (saveToFile)
            {
                _config.SaveConfig();
            }
        }
    }

    /// <summary>
    /// 实现IDisposable接口
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否由Dispose方法调用</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // 释放托管资源
                if (_accesser != null)
                {
                    _accesser.Dispose();
                    _accesser = null;
                }
            }
            _disposed = true;
        }
    }
}

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
    /// 查找指定后缀的文件并按照时间排序
    /// </summary>
    /// <param name="extension">文件后缀，如".txt"、".csv"等，不区分大小写</param>
    /// <param name="count">要获取的文件数量，默认为10</param>
    /// <param name="ascending">是否按时间升序排序（旧的文件在前），默认为false（新的文件在前）</param>
    /// <param name="searchOption">搜索选项，默认只搜索当前文件夹</param>
    /// <returns>排序后的文件信息列表</returns>
    public List<FileInfo> FindFilesByExtension(string extension, int count = 10, bool ascending = false, SearchOption searchOption = SearchOption.TopDirectoryOnly)
    {
        try
        {
            // 确保扩展名以点开头
            if (!string.IsNullOrEmpty(extension) && !extension.StartsWith("."))
            {
                extension = "." + extension;
            }

            // 获取所有文件
            string[] files;
            if (string.IsNullOrEmpty(extension))
            {
                files = Directory.GetFiles(_networkName, "*.*", searchOption);
            }
            else
            {
                files = Directory.GetFiles(_networkName, $"*{extension}", searchOption);
            }

            // 转换为FileInfo对象列表，以便获取创建时间和修改时间
            var fileInfos = files.Select(f => new FileInfo(f)).ToList();

            // 按照修改时间排序
            if (ascending)
            {
                // 升序：旧的在前
                fileInfos = fileInfos.OrderBy(f => f.LastWriteTime).ToList();
            }
            else
            {
                // 降序：新的在前
                fileInfos = fileInfos.OrderByDescending(f => f.LastWriteTime).ToList();
            }

            // 返回指定数量的文件
            return fileInfos.Take(count).ToList();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"查找文件失败: {ex.Message}");
            return new List<FileInfo>();
        }
    }

    /// <summary>
    /// 查找指定后缀的文件并按照创建时间排序
    /// </summary>
    /// <param name="extension">文件后缀，如".txt"、".csv"等，不区分大小写</param>
    /// <param name="count">要获取的文件数量，默认为10</param>
    /// <param name="ascending">是否按时间升序排序（旧的文件在前），默认为false（新的文件在前）</param>
    /// <param name="searchOption">搜索选项，默认只搜索当前文件夹</param>
    /// <returns>排序后的文件信息列表</returns>
    public List<FileInfo> FindFilesByExtensionByCreationTime(string extension, int count = 10, bool ascending = false, SearchOption searchOption = SearchOption.TopDirectoryOnly)
    {
        try
        {
            // 确保扩展名以点开头
            if (!string.IsNullOrEmpty(extension) && !extension.StartsWith("."))
            {
                extension = "." + extension;
            }

            // 获取所有文件
            string[] files;
            if (string.IsNullOrEmpty(extension))
            {
                files = Directory.GetFiles(_networkName, "*.*", searchOption);
            }
            else
            {
                files = Directory.GetFiles(_networkName, $"*{extension}", searchOption);
            }

            // 转换为FileInfo对象列表，以便获取创建时间
            var fileInfos = files.Select(f => new FileInfo(f)).ToList();

            // 按照创建时间排序
            if (ascending)
            {
                // 升序：旧的在前
                fileInfos = fileInfos.OrderBy(f => f.CreationTime).ToList();
            }
            else
            {
                // 降序：新的在前
                fileInfos = fileInfos.OrderByDescending(f => f.CreationTime).ToList();
            }

            // 返回指定数量的文件
            return fileInfos.Take(count).ToList();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"查找文件失败: {ex.Message}");
            return new List<FileInfo>();
        }
    }

    /// <summary>
    /// 将网络共享中的文件复制到本地计算机
    /// </summary>
    /// <param name="remoteFileName">要复制的远程文件名或相对路径</param>
    /// <param name="localFilePath">本地保存路径，包括文件名</param>
    /// <param name="overwrite">如果本地文件已存在，是否覆盖</param>
    /// <returns>复制是否成功</returns>
    public bool CopyFileToLocal(string remoteFileName, string localFilePath, bool overwrite = true)
    {
        try
        {
            // 构建远程文件的完整路径
            string remoteFilePath = Path.Combine(_networkName, remoteFileName);

            // 确保目标目录存在
            string localDirectory = Path.GetDirectoryName(localFilePath);
            if (!Directory.Exists(localDirectory))
            {
                Directory.CreateDirectory(localDirectory);
            }

            // 复制文件
            File.Copy(remoteFilePath, localFilePath, overwrite);

            // 验证文件是否成功复制
            if (File.Exists(localFilePath))
            {
                Console.WriteLine($"文件 {remoteFileName} 已成功复制到 {localFilePath}");
                return true;
            }
            else
            {
                Console.WriteLine($"文件复制失败: {remoteFileName}");
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"复制文件时出错: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 将多个网络共享文件复制到本地计算机的同一目录
    /// </summary>
    /// <param name="remoteFileNames">要复制的远程文件名列表</param>
    /// <param name="localDirectoryPath">本地保存目录路径</param>
    /// <param name="overwrite">如果本地文件已存在，是否覆盖</param>
    /// <returns>成功复制的文件数量</returns>
    public int CopyFilesToLocal(List<string> remoteFileNames, string localDirectoryPath, bool overwrite = true)
    {
        int successCount = 0;

        // 确保目标目录存在
        if (!Directory.Exists(localDirectoryPath))
        {
            Directory.CreateDirectory(localDirectoryPath);
        }

        foreach (var remoteFileName in remoteFileNames)
        {
            try
            {
                // 获取文件名部分
                string fileName = Path.GetFileName(remoteFileName);
                // 构建本地完整路径
                string localFilePath = Path.Combine(localDirectoryPath, fileName);
                // 调用单文件复制方法
                if (CopyFileToLocal(remoteFileName, localFilePath, overwrite))
                {
                    successCount++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"复制文件 {remoteFileName} 时出错: {ex.Message}");
            }
        }

        Console.WriteLine($"成功复制 {successCount}/{remoteFileNames.Count} 个文件到 {localDirectoryPath}");
        return successCount;
    }

    /// <summary>
    /// 复制指定后缀和时间条件的文件到本地
    /// </summary>
    /// <param name="extension">文件后缀</param>
    /// <param name="count">要复制的文件数量</param>
    /// <param name="localDirectoryPath">本地保存目录</param>
    /// <param name="ascending">是否按时间升序排序</param>
    /// <param name="overwrite">是否覆盖已存在的文件</param>
    /// <param name="searchOption">搜索选项，默认只搜索当前文件夹</param>
    /// <returns>成功复制的文件数量</returns>
    public int CopyLatestFilesByExtensionToLocal(string extension, int count, string localDirectoryPath, bool ascending = false, bool overwrite = true, SearchOption searchOption = SearchOption.TopDirectoryOnly)
    {
        // 查找符合条件的文件
        var files = FindFilesByExtension(extension, count, ascending, searchOption);
        if (files.Count == 0)
        {
            Console.WriteLine($"没有找到符合条件的 {extension} 文件");
            return 0;
        }

        // 准备文件名列表，保持相对路径
        var fileNames = files.Select(f =>
        {
            // 获取相对于网络共享根目录的路径
            string relativePath = f.FullName.Substring(_networkName.Length).TrimStart('\\');
            return relativePath;
        }).ToList();

        // 复制文件
        return CopyFilesToLocal(fileNames, localDirectoryPath, overwrite);
    }

    /// <summary>
    /// 实现IDisposable接口
    /// </summary>
    public void Dispose()
    {
        Disconnect();
    }
}
