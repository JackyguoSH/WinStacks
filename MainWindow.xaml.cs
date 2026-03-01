using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using System.Windows.Shell;

using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using WinStacks.Services;

namespace WinStacks
{
    /// <summary>
    /// WinStacks 主窗口类，负责桌面文件叠放显示和管理
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        #region Windows API 声明

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string? className, string? windowTitle);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        #endregion

        #region Shell API 声明

        [StructLayout(LayoutKind.Sequential)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        [DllImport("shell32.dll")]
        public static extern int SHGetStockIconInfo(uint siid, uint uFlags, ref SHSTOCKICONINFO psii);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHSTOCKICONINFO
        {
            public uint cbSize;
            public IntPtr hIcon;
            public int iSysImageIndex;
            public int iIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szPath;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string? lpVerb;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string? lpFile;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string? lpParameters;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string? lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string? lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIcon;
            public IntPtr hProcess;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHOpenWithDialog(IntPtr hwndParent, ref OPENASINFO poainfo);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct OPENASINFO
        {
            [MarshalAs(UnmanagedType.LPWStr)]
            public string? pcszFile;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string? pcszClass;
            public int oaUI;
        }

        #endregion

        #region 私有字段

        private FileSystemWatcher? _watcher;
        private System.Windows.Forms.NotifyIcon? _trayIcon;
        private Window? _currentMenu;
        private bool _keepMenuAlive = false;
        private bool _iconsRestored = false;
        private int _currentLayoutMode = 0;
        private bool _disposed = false;
        private readonly LogService _logger = LogService.Instance;

        private readonly List<FileSystemEventHandler> _fileWatcherCreatedHandlers = new List<FileSystemEventHandler>();
        private readonly List<FileSystemEventHandler> _fileWatcherDeletedHandlers = new List<FileSystemEventHandler>();
        private readonly List<RenamedEventHandler> _fileWatcherRenamedHandlers = new List<RenamedEventHandler>();

        #endregion

        #region 构造函数

        /// <summary>
        /// 初始化 MainWindow 的新实例
        /// </summary>
        public MainWindow()
        {
            InitializeExceptionHandlers();
            InitializeComponent();
            SetupEventHandlers();
            SetupTrayIcon();
            
            _logger.LogInfo($"WinStacks 主窗口已初始化 - {WindowsVersionService.GetVersionInfo()}");
        }

        /// <summary>
        /// 初始化全局异常处理器
        /// </summary>
        private void InitializeExceptionHandlers()
        {
            AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
            Application.Current.DispatcherUnhandledException += OnDispatcherUnhandledException;
            Application.Current.Exit += OnApplicationExit;
            
            _logger.LogInfo("全局异常处理器已初始化");
        }

        /// <summary>
        /// 取消订阅全局异常处理器
        /// </summary>
        private void UninitializeExceptionHandlers()
        {
            AppDomain.CurrentDomain.UnhandledException -= OnAppDomainUnhandledException;
            Application.Current.DispatcherUnhandledException -= OnDispatcherUnhandledException;
            Application.Current.Exit -= OnApplicationExit;
        }

        /// <summary>
        /// 设置窗口事件处理程序
        /// </summary>
        private void SetupEventHandlers()
        {
            this.Activated += MainWindow_Activated;
            this.Loaded += MainWindow_Loaded;
            this.Closed += MainWindow_Closed;
        }

        /// <summary>
        /// 取消订阅窗口事件处理程序
        /// </summary>
        private void RemoveEventHandlers()
        {
            this.Activated -= MainWindow_Activated;
            this.Loaded -= MainWindow_Loaded;
            this.Closed -= MainWindow_Closed;
        }

        #endregion

        #region 事件处理方法

        /// <summary>
        /// 窗口激活事件处理
        /// </summary>
        private void MainWindow_Activated(object? sender, EventArgs e)
        {
            PinToDesktop();
        }

        /// <summary>
        /// 窗口关闭事件处理
        /// </summary>
        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            EnsureDesktopIconsVisible();
            CleanupResources();
            _logger.LogInfo("主窗口已关闭");
        }

        #endregion

        #region 异常处理

        /// <summary>
        /// 处理 AppDomain 未捕获的异常
        /// </summary>
        private void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception exception)
            {
                ExceptionHandler.HandleUnhandledException(exception, "AppDomain");
            }
            EnsureDesktopIconsVisible();
        }

        /// <summary>
        /// 处理 Dispatcher 未捕获的异常
        /// </summary>
        private void OnDispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            ExceptionHandler.HandleDispatcherException(e.Exception);
            e.Handled = true;
        }

        /// <summary>
        /// 处理应用程序退出事件
        /// </summary>
        private void OnApplicationExit(object sender, ExitEventArgs e)
        {
            _logger.LogInfo("应用程序正在退出");
            EnsureDesktopIconsVisible();
            CleanupResources();
        }

        #endregion

        #region 窗口核心功能

        /// <summary>
        /// 将窗口固定到桌面底层
        /// </summary>
        private void PinToDesktop()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                {
                    SetWindowPos(hwnd, AppConstants.HWND_BOTTOM, 0, 0, 0, 0, 
                        AppConstants.SWP_NOMOVE | AppConstants.SWP_NOSIZE | AppConstants.SWP_NOACTIVATE);
                }
            }, "固定窗口到桌面");
        }

        /// <summary>
        /// 窗口加载完成事件处理
        /// </summary>
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            ExceptionHandler.SafeExecute(() =>
            {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                int exStyle = GetWindowLong(hwnd, AppConstants.GWL_EXSTYLE);
                SetWindowLong(hwnd, AppConstants.GWL_EXSTYLE, exStyle | AppConstants.WS_EX_TOOLWINDOW);
                
                IntPtr progman = FindWindow("Progman", null);
                if (progman != IntPtr.Zero)
                {
                    SetWindowLongPtr(hwnd, AppConstants.GWLP_HWNDPARENT, progman);
                }
                
                PinToDesktop();
                EnsureDesktopIconsHidden();
                ScanDesktopFiles();
                StartDesktopWatcher();
                LoadLayoutSettings();
                
                _logger.LogInfo("主窗口加载完成");
            }, "主窗口加载");
        }

        #endregion

        #region 布局设置

        /// <summary>
        /// 加载布局设置
        /// </summary>
        private void LoadLayoutSettings()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                string appDataDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    AppConstants.AppDataDirectoryName);
                string configPath = Path.Combine(appDataDir, AppConstants.LayoutConfigFileName);

                if (!File.Exists(configPath))
                {
                    Directory.CreateDirectory(appDataDir);
                    File.WriteAllText(configPath, "0");
                    ApplyLayout(AppConstants.LayoutModeRight, false);

                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        OpenSettingsWindow();
                        PinToDesktop();
                    }, System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    
                    _logger.LogInfo("首次运行，已创建默认配置");
                }
                else
                {
                    string savedMode = File.ReadAllText(configPath);
                    if (int.TryParse(savedMode, out int mode))
                    {
                        ApplyLayout(mode, false);
                    }
                    else
                    {
                        ApplyLayout(AppConstants.LayoutModeRight, false);
                    }
                    _logger.LogInfo($"已加载布局设置: 模式 {_currentLayoutMode}");
                }
            }, "加载布局设置", showErrorToUser: false);
            
            ApplyLayout(AppConstants.LayoutModeRight, false);
        }

        /// <summary>
        /// 应用指定的布局模式
        /// </summary>
        /// <param name="mode">布局模式 (0: 右侧, 1: 左侧, 2: 顶部, 3: 底部)</param>
        /// <param name="saveToFile">是否保存到配置文件</param>
        private void ApplyLayout(int mode, bool saveToFile)
        {
            _currentLayoutMode = mode;
            
            switch (mode)
            {
                case AppConstants.LayoutModeRight:
                    StacksContainer.Orientation = Orientation.Vertical;
                    StacksContainer.HorizontalAlignment = HorizontalAlignment.Right;
                    StacksContainer.VerticalAlignment = VerticalAlignment.Top;
                    StacksContainer.Margin = new Thickness(0, AppConstants.TopMargin, AppConstants.LayoutMargin, 0);
                    break;
                case AppConstants.LayoutModeLeft:
                    StacksContainer.Orientation = Orientation.Vertical;
                    StacksContainer.HorizontalAlignment = HorizontalAlignment.Left;
                    StacksContainer.VerticalAlignment = VerticalAlignment.Top;
                    StacksContainer.Margin = new Thickness(AppConstants.LayoutMargin, AppConstants.TopMargin, 0, 0);
                    break;
                case AppConstants.LayoutModeTop:
                    StacksContainer.Orientation = Orientation.Horizontal;
                    StacksContainer.HorizontalAlignment = HorizontalAlignment.Center;
                    StacksContainer.VerticalAlignment = VerticalAlignment.Top;
                    StacksContainer.Margin = new Thickness(0, AppConstants.LayoutMargin, 0, 0);
                    break;
                case AppConstants.LayoutModeBottom:
                    StacksContainer.Orientation = Orientation.Horizontal;
                    StacksContainer.HorizontalAlignment = HorizontalAlignment.Center;
                    StacksContainer.VerticalAlignment = VerticalAlignment.Bottom;
                    StacksContainer.Margin = new Thickness(0, 0, 0, AppConstants.BottomMargin);
                    break;
            }

            if (saveToFile)
            {
                ExceptionHandler.SafeExecute(() =>
                {
                    string appDataDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        AppConstants.AppDataDirectoryName);
                    Directory.CreateDirectory(appDataDir);
                    File.WriteAllText(Path.Combine(appDataDir, AppConstants.LayoutConfigFileName), mode.ToString());
                    _logger.LogInfo($"布局设置已保存: 模式 {mode}");
                }, "保存布局设置");
            }
        }

        #endregion

        #region 系统托盘

        /// <summary>
        /// 设置系统托盘图标
        /// </summary>
        private void SetupTrayIcon()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                _trayIcon = new System.Windows.Forms.NotifyIcon();
                string? exePath = Environment.ProcessPath ?? AppContext.BaseDirectory;
                
                if (!string.IsNullOrEmpty(exePath))
                {
                    try
                    {
                        _trayIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"无法加载托盘图标: {ex.Message}");
                    }
                }

                _trayIcon.Text = "WinStacks - 智能叠放";
                _trayIcon.Visible = true;

                var menu = new System.Windows.Forms.ContextMenuStrip();
                menu.Items.Add("👁️ 隐藏/显示桌面图标", null, TrayMenu_ToggleDesktopIcons_Click);
                menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
                menu.Items.Add("⚙️ 排版设置", null, TrayMenu_OpenSettings_Click);
                menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
                menu.Items.Add("❌ 退出 WinStacks", null, TrayMenu_Exit_Click);

                _trayIcon.ContextMenuStrip = menu;
                _trayIcon.DoubleClick += TrayIcon_DoubleClick;
                
                _logger.LogInfo("系统托盘图标已创建");
            }, "设置托盘图标");
        }

        /// <summary>
        /// 清理托盘图标资源
        /// </summary>
        private void CleanupTrayIcon()
        {
            if (_trayIcon != null)
            {
                _trayIcon.DoubleClick -= TrayIcon_DoubleClick;
                
                if (_trayIcon.ContextMenuStrip != null)
                {
                    foreach (var item in _trayIcon.ContextMenuStrip.Items)
                    {
                        if (item is System.Windows.Forms.ToolStripMenuItem menuItem)
                        {
                            menuItem.Click -= TrayMenu_ToggleDesktopIcons_Click;
                            menuItem.Click -= TrayMenu_OpenSettings_Click;
                            menuItem.Click -= TrayMenu_Exit_Click;
                        }
                    }
                }
                
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                _trayIcon = null;
            }
        }

        /// <summary>
        /// 托盘菜单 - 切换桌面图标
        /// </summary>
        private void TrayMenu_ToggleDesktopIcons_Click(object? sender, EventArgs e)
        {
            ToggleDesktopIcons();
        }

        /// <summary>
        /// 托盘菜单 - 打开设置
        /// </summary>
        private void TrayMenu_OpenSettings_Click(object? sender, EventArgs e)
        {
            OpenSettingsWindow();
        }

        /// <summary>
        /// 托盘菜单 - 退出
        /// </summary>
        private void TrayMenu_Exit_Click(object? sender, EventArgs e)
        {
            ExitApplication();
        }

        /// <summary>
        /// 托盘图标双击事件
        /// </summary>
        private void TrayIcon_DoubleClick(object? sender, EventArgs e)
        {
            OpenSettingsWindow();
        }

        #endregion

        #region 桌面图标控制

        /// <summary>
        /// 获取桌面 DefView 窗口句柄
        /// </summary>
        /// <returns>DefView 窗口句柄</returns>
        private IntPtr GetDesktopDefView()
        {
            IntPtr hwnd = FindWindow("Progman", null);
            IntPtr defView = FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
            
            if (defView == IntPtr.Zero)
            {
                EnumWindows((tophandle, topparamhandle) =>
                {
                    IntPtr p = FindWindowEx(tophandle, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if (p != IntPtr.Zero)
                    {
                        defView = p;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
            }
            return defView;
        }

        /// <summary>
        /// 切换桌面图标显示/隐藏
        /// </summary>
        private void ToggleDesktopIcons()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                IntPtr defView = GetDesktopDefView();
                if (defView != IntPtr.Zero)
                {
                    SendMessage(defView, AppConstants.WM_COMMAND, (IntPtr)AppConstants.ToggleDesktopCommand, IntPtr.Zero);
                    _logger.LogInfo("桌面图标显示状态已切换");
                }
            }, "切换桌面图标");
        }

        /// <summary>
        /// 确保桌面图标隐藏
        /// </summary>
        private void EnsureDesktopIconsHidden()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                IntPtr defView = GetDesktopDefView();
                if (defView != IntPtr.Zero)
                {
                    IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
                    if (listView != IntPtr.Zero)
                    {
                        int style = GetWindowLong(listView, AppConstants.GWL_STYLE);
                        if ((style & AppConstants.WS_VISIBLE) != 0)
                        {
                            SendMessage(defView, AppConstants.WM_COMMAND, (IntPtr)AppConstants.ToggleDesktopCommand, IntPtr.Zero);
                            _logger.LogInfo("桌面图标已隐藏");
                        }
                    }
                }
            }, "隐藏桌面图标");
        }

        /// <summary>
        /// 确保桌面图标可见
        /// </summary>
        private void EnsureDesktopIconsVisible()
        {
            if (_iconsRestored) return;

            ExceptionHandler.SafeExecute(() =>
            {
                IntPtr defView = GetDesktopDefView();
                if (defView != IntPtr.Zero)
                {
                    IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
                    if (listView != IntPtr.Zero)
                    {
                        int style = GetWindowLong(listView, AppConstants.GWL_STYLE);
                        if ((style & AppConstants.WS_VISIBLE) == 0)
                        {
                            SendMessage(defView, AppConstants.WM_COMMAND, (IntPtr)AppConstants.ToggleDesktopCommand, IntPtr.Zero);
                            _logger.LogInfo("桌面图标已恢复显示");
                        }
                    }
                }
            }, "恢复桌面图标");
            
            _iconsRestored = true;
        }

        #endregion

        #region 文件监控

        /// <summary>
        /// 启动桌面文件监控
        /// </summary>
        private void StartDesktopWatcher()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                _watcher = new FileSystemWatcher(desktopPath)
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite,
                    EnableRaisingEvents = true
                };
                
                FileSystemEventHandler createdHandler = (s, e) => RefreshDesktop();
                FileSystemEventHandler deletedHandler = (s, e) => RefreshDesktop();
                RenamedEventHandler renamedHandler = (s, e) => RefreshDesktop();
                
                _watcher.Created += createdHandler;
                _watcher.Deleted += deletedHandler;
                _watcher.Renamed += renamedHandler;
                
                _fileWatcherCreatedHandlers.Add(createdHandler);
                _fileWatcherDeletedHandlers.Add(deletedHandler);
                _fileWatcherRenamedHandlers.Add(renamedHandler);
                
                _logger.LogInfo("桌面文件监控已启动");
            }, "启动文件监控");
        }

        /// <summary>
        /// 停止桌面文件监控
        /// </summary>
        private void StopDesktopWatcher()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                
                for (int i = 0; i < _fileWatcherCreatedHandlers.Count; i++)
                {
                    _watcher.Created -= _fileWatcherCreatedHandlers[i];
                }
                for (int i = 0; i < _fileWatcherDeletedHandlers.Count; i++)
                {
                    _watcher.Deleted -= _fileWatcherDeletedHandlers[i];
                }
                for (int i = 0; i < _fileWatcherRenamedHandlers.Count; i++)
                {
                    _watcher.Renamed -= _fileWatcherRenamedHandlers[i];
                }
                
                _fileWatcherCreatedHandlers.Clear();
                _fileWatcherDeletedHandlers.Clear();
                _fileWatcherRenamedHandlers.Clear();
                
                _watcher.Dispose();
                _watcher = null;
            }
        }

        /// <summary>
        /// 刷新桌面显示
        /// </summary>
        private void RefreshDesktop()
        {
            Application.Current.Dispatcher.Invoke(() => ScanDesktopFiles());
        }

        #endregion

        #region 文件扫描与分类

        /// <summary>
        /// 扫描桌面文件并进行分类
        /// </summary>
        private void ScanDesktopFiles()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string[] files = Directory.GetFiles(desktopPath);
                string[] directories = Directory.GetDirectories(desktopPath);

                var categories = new Dictionary<string, List<string>>()
                {
                    { "系统", new List<string>() },
                    { "照片", new List<string>() },
                    { "文档", new List<string>() },
                    { "影片", new List<string>() },
                    { "音乐", new List<string>() },
                    { "压缩包", new List<string>() },
                    { "应用程序", new List<string>() },
                    { "文件夹", new List<string>() },
                    { "其他", new List<string>() }
                };

                foreach (string dir in directories)
                {
                    categories["文件夹"].Add(dir);
                }
                
                categories["系统"].Add(AppConstants.CLSID_ThisPC);
                categories["系统"].Add(AppConstants.CLSID_RecycleBin);

                foreach (string file in files)
                {
                    CategorizeFile(file, categories);
                }

                StacksContainer.Children.Clear();
                foreach (var category in categories)
                {
                    if (category.Value.Count > 0)
                    {
                        string displayTitle = category.Key;
                        UIElement stackUI = CreateStackBlock(displayTitle, category.Value);
                        StacksContainer.Children.Add(stackUI);
                    }
                }
            }, "扫描桌面文件");
        }

        /// <summary>
        /// 将文件分类到对应类别
        /// </summary>
        /// <param name="file">文件路径</param>
        /// <param name="categories">分类字典</param>
        private void CategorizeFile(string file, Dictionary<string, List<string>> categories)
        {
            ExceptionHandler.SafeExecute(() =>
            {
                FileAttributes attributes = File.GetAttributes(file);
                if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden ||
                    (attributes & FileAttributes.System) == FileAttributes.System)
                {
                    return;
                }

                string ext = Path.GetExtension(file).ToLower();
                
                if (IsImageFile(ext)) categories["照片"].Add(file);
                else if (IsDocumentFile(ext)) categories["文档"].Add(file);
                else if (IsVideoFile(ext)) categories["影片"].Add(file);
                else if (IsAudioFile(ext)) categories["音乐"].Add(file);
                else if (IsArchiveFile(ext)) categories["压缩包"].Add(file);
                else if (IsExecutableFile(ext)) categories["应用程序"].Add(file);
                else categories["其他"].Add(file);
            }, $"分类文件: {file}");
        }

        /// <summary>
        /// 判断是否为图片文件
        /// </summary>
        private static bool IsImageFile(string ext)
        {
            return Array.IndexOf(AppConstants.ImageExtensions, ext) >= 0;
        }

        /// <summary>
        /// 判断是否为文档文件
        /// </summary>
        private static bool IsDocumentFile(string ext)
        {
            return Array.IndexOf(AppConstants.DocumentExtensions, ext) >= 0;
        }

        /// <summary>
        /// 判断是否为视频文件
        /// </summary>
        private static bool IsVideoFile(string ext)
        {
            return Array.IndexOf(AppConstants.VideoExtensions, ext) >= 0;
        }

        /// <summary>
        /// 判断是否为音频文件
        /// </summary>
        private static bool IsAudioFile(string ext)
        {
            return Array.IndexOf(AppConstants.AudioExtensions, ext) >= 0;
        }

        /// <summary>
        /// 判断是否为压缩包文件
        /// </summary>
        private static bool IsArchiveFile(string ext)
        {
            return Array.IndexOf(AppConstants.ArchiveExtensions, ext) >= 0;
        }

        /// <summary>
        /// 判断是否为可执行文件
        /// </summary>
        private static bool IsExecutableFile(string ext)
        {
            return Array.IndexOf(AppConstants.ExecutableExtensions, ext) >= 0;
        }

        #endregion

        #region 图标获取

        /// <summary>
        /// 获取文件或文件夹的原生系统图标
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>图标图像源</returns>
        private ImageSource? GetNativeIcon(string path)
        {
            return ExceptionHandler.SafeExecute(() =>
            {
                if (path == AppConstants.CLSID_ThisPC)
                {
                    return GetStockIcon(AppConstants.SIID_DRIVEFIXED);
                }
                else if (path == AppConstants.CLSID_RecycleBin)
                {
                    return GetStockIcon(AppConstants.SIID_RECYCLER);
                }

                return GetFileIcon(path);
            }, $"获取图标: {path}");
        }

        /// <summary>
        /// 获取系统内置图标
        /// </summary>
        /// <param name="stockIconId">系统图标ID</param>
        /// <returns>图标图像源</returns>
        private ImageSource? GetStockIcon(uint stockIconId)
        {
            SHSTOCKICONINFO sii = new SHSTOCKICONINFO();
            sii.cbSize = (uint)Marshal.SizeOf(typeof(SHSTOCKICONINFO));

            if (SHGetStockIconInfo(stockIconId, AppConstants.SHGSI_ICON | AppConstants.SHGSI_LARGEICON, ref sii) == 0 && sii.hIcon != IntPtr.Zero)
            {
                ImageSource img = Imaging.CreateBitmapSourceFromHIcon(sii.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                DestroyIcon(sii.hIcon);
                return img;
            }
            return null;
        }

        /// <summary>
        /// 获取文件图标
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>图标图像源</returns>
        private ImageSource? GetFileIcon(string path)
        {
            SHFILEINFO shfi = new SHFILEINFO();
            IntPtr res = SHGetFileInfo(path, 0, ref shfi, (uint)Marshal.SizeOf(shfi), AppConstants.SHGFI_ICON | AppConstants.SHGFI_LARGEICON);
            
            if (res != IntPtr.Zero && shfi.hIcon != IntPtr.Zero)
            {
                ImageSource img = Imaging.CreateBitmapSourceFromHIcon(shfi.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                DestroyIcon(shfi.hIcon);
                return img;
            }
            return null;
        }

        #endregion

        #region UI 创建

        /// <summary>
        /// 创建叠放区块 UI 元素
        /// </summary>
        /// <param name="title">区块标题</param>
        /// <param name="filePaths">文件路径列表</param>
        /// <returns>UI 元素</returns>
        private UIElement CreateStackBlock(string title, List<string> filePaths)
        {
            int count = filePaths.Count;
            Grid mainGrid = new Grid
            {
                Width = AppConstants.StackBlockWidth,
                Height = AppConstants.StackBlockHeight,
                Margin = new Thickness(AppConstants.StackBlockMargin, AppConstants.StackBlockVerticalMargin, AppConstants.StackBlockMargin, AppConstants.StackBlockVerticalMargin),
                Cursor = Cursors.Hand,
                Background = Brushes.Transparent
            };

            ScaleTransform scaleTransform = new ScaleTransform(AppConstants.NormalScaleFactor, AppConstants.NormalScaleFactor);
            mainGrid.RenderTransform = scaleTransform;
            mainGrid.RenderTransformOrigin = new Point(0.5, 0.5);

            Grid iconStackGrid = new Grid
            {
                Width = AppConstants.IconStackWidth,
                Height = AppConstants.IconStackHeight,
                VerticalAlignment = VerticalAlignment.Top
            };

            int maxIcons = Math.Min(3, filePaths.Count);
            double totalWidth = AppConstants.IconSize + (maxIcons - 1) * AppConstants.IconStackOffset;
            double centerOffset = (AppConstants.IconStackWidth - totalWidth) / 2;

            CreateIconStack(iconStackGrid, filePaths, maxIcons, centerOffset);
            CreateBadge(iconStackGrid, count);

            TextBlock titleText = new TextBlock
            {
                Text = title,
                FontSize = AppConstants.TitleFontSize,
                Foreground = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 5),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    BlurRadius = 4,
                    ShadowDepth = 1,
                    Opacity = 0.9
                }
            };

            mainGrid.Children.Add(iconStackGrid);
            mainGrid.Children.Add(titleText);

            SetupStackBlockInteractions(mainGrid, scaleTransform, title, filePaths);

            return mainGrid;
        }

        /// <summary>
        /// 创建图标堆叠效果
        /// </summary>
        private void CreateIconStack(Grid iconStackGrid, List<string> filePaths, int maxIcons, double centerOffset)
        {
            for (int i = maxIcons - 1; i >= 0; i--)
            {
                ImageSource? source = GetNativeIcon(filePaths[i]);
                if (source != null)
                {
                    Image img = new Image
                    {
                        Source = source,
                        Width = AppConstants.IconSize,
                        Height = AppConstants.IconSize,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Top,
                        Margin = new Thickness(centerOffset + (maxIcons - 1 - i) * AppConstants.IconStackOffset, centerOffset + (maxIcons - 1 - i) * AppConstants.IconStackOffset, 0, 0),
                        Effect = new DropShadowEffect
                        {
                            Color = Colors.Black,
                            BlurRadius = 6,
                            ShadowDepth = 2,
                            Opacity = 0.35
                        }
                    };
                    RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);
                    iconStackGrid.Children.Add(img);
                }
            }
        }

        /// <summary>
        /// 创建数量徽章
        /// </summary>
        private static void CreateBadge(Grid iconStackGrid, int count)
        {
            Border badge = new Border
            {
                Background = new SolidColorBrush(AppConstants.BadgeBackgroundColor),
                BorderBrush = new SolidColorBrush(AppConstants.BadgeBorderColor),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(5, 0, 5, 0),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, -6, -4, 0),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    BlurRadius = 4,
                    ShadowDepth = 2,
                    Opacity = 0.5
                }
            };
            badge.Child = new TextBlock
            {
                Text = count.ToString(),
                FontSize = AppConstants.BadgeFontSize,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            };
            iconStackGrid.Children.Add(badge);
        }

        /// <summary>
        /// 设置叠放区块的交互行为
        /// </summary>
        private void SetupStackBlockInteractions(Grid mainGrid, ScaleTransform scaleTransform, string title, List<string> filePaths)
        {
            mainGrid.AllowDrop = true;
            System.Windows.Threading.DispatcherTimer? dragOpenTimer = null;

            mainGrid.MouseEnter += (s, e) =>
            {
                DoubleAnimation anim = new DoubleAnimation(AppConstants.HoverScaleFactor, TimeSpan.FromMilliseconds(AppConstants.HoverAnimationDurationMs))
                {
                    EasingFunction = new CubicEase() { EasingMode = EasingMode.EaseOut }
                };
                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, anim);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, anim);
            };

            mainGrid.MouseLeave += (s, e) =>
            {
                DoubleAnimation anim = new DoubleAnimation(AppConstants.NormalScaleFactor, TimeSpan.FromMilliseconds(AppConstants.HoverAnimationDurationMs))
                {
                    EasingFunction = new CubicEase() { EasingMode = EasingMode.EaseOut }
                };
                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, anim);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, anim);
            };

            mainGrid.MouseLeftButtonDown += (sender, e) =>
            {
                ShowStackMenu(mainGrid, title, filePaths);
                e.Handled = true;
            };

            mainGrid.DragEnter += (s, e) =>
            {
                dragOpenTimer?.Stop();
                dragOpenTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(AppConstants.DragOpenMenuDelayMs)
                };
                dragOpenTimer.Tick += (ts, te) =>
                {
                    dragOpenTimer?.Stop();
                    ShowStackMenu(mainGrid, title, filePaths);
                };
                dragOpenTimer.Start();
            };

            mainGrid.DragLeave += (s, e) => { dragOpenTimer?.Stop(); };
            mainGrid.Drop += (s, e) => { dragOpenTimer?.Stop(); };
        }

        #endregion

        #region 设置窗口

        /// <summary>
        /// 打开设置窗口
        /// </summary>
        private void OpenSettingsWindow()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                Window settingsWin = new Window
                {
                    Title = "WinStacks 排版设置",
                    Width = AppConstants.SettingsWindowWidth,
                    Height = AppConstants.SettingsWindowHeight,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    Topmost = true,
                    Background = new SolidColorBrush(AppConstants.DefaultBackgroundColor),
                    Foreground = Brushes.White,
                    WindowStyle = WindowStyle.ToolWindow
                };

                StackPanel sp = new StackPanel { Margin = new Thickness(25) };
                sp.Children.Add(new TextBlock
                {
                    Text = "选择叠放区的位置：",
                    FontSize = AppConstants.SettingsTitleFontSize,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 15)
                });

                ComboBox cbLayout = new ComboBox { Height = 35, FontSize = AppConstants.SettingsContentFontSize };
                cbLayout.Items.Add("📍 右侧 - 垂直排版 (Mac 风格)");
                cbLayout.Items.Add("📍 左侧 - 垂直排版 (Win 风格)");
                cbLayout.Items.Add("📍 顶部 - 水平排版");
                cbLayout.Items.Add("📍 底部 - 水平排版 (Dock 风格)");
                cbLayout.SelectedIndex = _currentLayoutMode;
                sp.Children.Add(cbLayout);

                Button btnApply = new Button
                {
                    Content = "保存并应用",
                    Height = 35,
                    Margin = new Thickness(0, 20, 0, 0),
                    Background = new SolidColorBrush(AppConstants.ThemeBlueColor),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };
                btnApply.Click += (s, e) =>
                {
                    ApplyLayout(cbLayout.SelectedIndex, true);
                    settingsWin.Close();
                };
                sp.Children.Add(btnApply);

                settingsWin.Content = sp;
                settingsWin.Show();
                
                _logger.LogInfo("设置窗口已打开");
            }, "打开设置窗口");
        }

        #endregion

        #region 文件操作

        /// <summary>
        /// 安全打开文件或文件夹
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="isDir">是否为目录</param>
        private void SafeOpenItem(string path, bool isDir)
        {
            ExceptionHandler.SafeExecute(() =>
            {
                if (path.StartsWith("::"))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = path,
                        UseShellExecute = true
                    });
                    return;
                }

                if (isDir)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = $"\"{path}\"",
                        UseShellExecute = true
                    });
                }
                else
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
                
                _logger.LogInfo($"已打开: {path}");
            }, $"打开项目: {path}", showErrorToUser: true);
        }

        /// <summary>
        /// 执行删除操作（移入回收站）
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="isDir">是否为目录</param>
        private void PerformDelete(string path, bool isDir)
        {
            ExceptionHandler.SafeExecute(() =>
            {
                if (isDir)
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                        path,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                }
                else
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(
                        path,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                }
                
                _logger.LogInfo($"已移入回收站: {path}");
            }, $"删除项目: {path}");
        }

        #endregion

        #region 菜单显示

        /// <summary>
        /// 显示叠放菜单
        /// </summary>
        /// <param name="anchor">锚点元素</param>
        /// <param name="title">菜单标题</param>
        /// <param name="filePaths">文件路径列表</param>
        private void ShowStackMenu(UIElement anchor, string title, List<string> filePaths)
        {
            ExceptionHandler.SafeExecute(() =>
            {
                CloseCurrentMenu();

                _currentMenu = CreateMenuWindow();
                if (_currentMenu == null) return;

                Action safeCloseMenu = CreateSafeCloseMenuAction();
                _currentMenu.Deactivated += Menu_Deactivated;

                var menuContext = new MenuContext();
                Border menuBorder = CreateMenuContent(title, filePaths, safeCloseMenu, menuContext);
                
                _currentMenu.Content = menuBorder;
                PositionAndShowMenu(anchor, menuBorder);
            }, "显示叠放菜单");
        }

        /// <summary>
        /// 菜单失去激活事件处理
        /// </summary>
        private void Menu_Deactivated(object? sender, EventArgs e)
        {
            if (!_keepMenuAlive)
            {
                CloseCurrentMenu();
            }
        }

        /// <summary>
        /// 关闭当前菜单
        /// </summary>
        private void CloseCurrentMenu()
        {
            if (_currentMenu != null)
            {
                try
                {
                    _currentMenu.Deactivated -= Menu_Deactivated;
                    _currentMenu.Close();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"关闭菜单时发生错误: {ex.Message}");
                }
                _currentMenu = null;
            }
        }

        /// <summary>
        /// 创建菜单窗口
        /// </summary>
        private Window CreateMenuWindow()
        {
            Window menu = new Window
            {
                WindowStyle = WindowStyle.None,
                AllowsTransparency = false,
                ShowInTaskbar = false,
                Topmost = false,
                SizeToContent = SizeToContent.WidthAndHeight,
                Background = new SolidColorBrush(AppConstants.TransparentWindowColor),
                WindowStartupLocation = WindowStartupLocation.Manual
            };

            WindowChrome.SetWindowChrome(menu, new WindowChrome
            {
                GlassFrameThickness = new Thickness(-1),
                CaptionHeight = 0,
                UseAeroCaptionButtons = false
            });

            menu.SourceInitialized += Menu_SourceInitialized;

            return menu;
        }

        /// <summary>
        /// 菜单窗口源初始化事件处理
        /// </summary>
        private void Menu_SourceInitialized(object? sender, EventArgs e)
        {
            if (_currentMenu == null) return;
            
            IntPtr hwnd = new WindowInteropHelper(_currentMenu).Handle;
            
            if (WindowsVersionService.SupportsDarkModeTitleBar)
            {
                int dark = AppConstants.DWM_DARK_MODE_ENABLED;
                DwmSetWindowAttribute(hwnd, AppConstants.DWMWA_USE_IMMERSIVE_DARK_MODE, ref dark, sizeof(int));
            }
            
            if (WindowsVersionService.SupportsMicaEffect)
            {
                int backdrop = AppConstants.DWM_BACKDROP_MICA;
                DwmSetWindowAttribute(hwnd, AppConstants.DWMWA_SYSTEMBACKDROP_TYPE, ref backdrop, sizeof(int));
            }
            
            if (WindowsVersionService.SupportsWindowCorners)
            {
                int corners = AppConstants.DWM_CORNER_ROUND;
                DwmSetWindowAttribute(hwnd, AppConstants.DWMWA_WINDOW_CORNER_PREFERENCE, ref corners, sizeof(int));
            }
            
            SetWindowPos(hwnd, AppConstants.HWND_TOPMOST, 0, 0, 0, 0, 
                AppConstants.SWP_NOMOVE | AppConstants.SWP_NOSIZE);
        }

        /// <summary>
        /// 创建安全关闭菜单的操作
        /// </summary>
        private Action CreateSafeCloseMenuAction()
        {
            return () =>
            {
                if (_currentMenu != null)
                {
                    Window temp = _currentMenu;
                    _currentMenu = null;
                    try
                    {
                        temp.Deactivated -= Menu_Deactivated;
                        temp.Close();
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"关闭菜单窗口时发生错误: {ex.Message}");
                    }
                }
            };
        }

        /// <summary>
        /// 菜单上下文类，用于存储菜单状态
        /// </summary>
        private class MenuContext
        {
            public List<string> SelectedPaths { get; } = new List<string>();
            public List<bool> IsDirectory { get; } = new List<bool>();
            public List<Border> SelectedBorders { get; } = new List<Border>();
            public List<TextBlock> SelectedTexts { get; } = new List<TextBlock>();
            public List<TextBox> SelectedTextBoxes { get; } = new List<TextBox>();
        }

        /// <summary>
        /// 创建菜单内容
        /// </summary>
        private Border CreateMenuContent(string title, List<string> filePaths, Action safeCloseMenu, MenuContext context)
        {
            Border menuBorder = new Border
            {
                Background = Brushes.Transparent,
                Padding = new Thickness(15)
            };

            StackPanel outerPanel = new StackPanel { Orientation = Orientation.Vertical };
            outerPanel.Children.Add(new TextBlock
            {
                Text = title,
                Foreground = new SolidColorBrush(AppConstants.MenuTitleTextColor),
                FontWeight = FontWeights.SemiBold,
                FontSize = AppConstants.MenuTitleFontSize,
                Margin = new Thickness(4, 0, 0, 12)
            });

            int columnCount = Math.Min(AppConstants.MenuMaxColumns, filePaths.Count);
            double dynamicWidth = columnCount * (AppConstants.MenuItemWidth + AppConstants.MenuItemMargin * 2);

            WrapPanel gridPanel = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                Width = dynamicWidth
            };

            foreach (string path in filePaths)
            {
                UIElement item = CreateMenuItem(path, safeCloseMenu, context);
                gridPanel.Children.Add(item);
            }

            double maxAllowedHeight = SystemParameters.WorkArea.Height * AppConstants.MenuMaxHeightRatio;
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Hidden,
                MaxHeight = maxAllowedHeight,
                Content = gridPanel
            };

            scrollViewer.PreviewMouseWheel += ScrollViewer_PreviewMouseWheel;

            outerPanel.Children.Add(scrollViewer);
            menuBorder.Child = outerPanel;

            return menuBorder;
        }

        /// <summary>
        /// 滚动视图鼠标滚轮事件处理
        /// </summary>
        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is ScrollViewer scrollViewer)
            {
                scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta);
                e.Handled = true;
            }
        }

        /// <summary>
        /// 创建菜单项
        /// </summary>
        private UIElement CreateMenuItem(string path, Action safeCloseMenu, MenuContext context)
        {
            string fileName = GetDisplayName(path);
            bool isDir = Directory.Exists(path) || path.StartsWith("::");

            Border itemBorder = new Border
            {
                Background = Brushes.Transparent,
                CornerRadius = new CornerRadius(6),
                Width = AppConstants.MenuItemWidth,
                Height = AppConstants.MenuItemHeight,
                Margin = new Thickness(AppConstants.MenuItemMargin),
                Cursor = Cursors.Hand
            };

            StackPanel itemPanel = new StackPanel { VerticalAlignment = VerticalAlignment.Center };

            AddIconToPanel(itemPanel, path);
            (TextBlock itemText, TextBox editBox) = AddTextToPanel(itemPanel, fileName);

            itemBorder.Child = itemPanel;

            SetupMenuItemInteractions(itemBorder, path, isDir, safeCloseMenu, context, itemText, editBox);
            SetupContextMenu(itemBorder, path, isDir, safeCloseMenu, context, itemText, editBox);

            return itemBorder;
        }

        /// <summary>
        /// 获取显示名称
        /// </summary>
        private static string GetDisplayName(string path)
        {
            if (path == AppConstants.CLSID_ThisPC) return "此电脑";
            if (path == AppConstants.CLSID_RecycleBin) return "回收站";
            return Path.GetFileName(path);
        }

        /// <summary>
        /// 添加图标到面板
        /// </summary>
        private void AddIconToPanel(StackPanel itemPanel, string path)
        {
            ImageSource? iconSource = GetNativeIcon(path);
            if (iconSource != null)
            {
                Image fileIcon = new Image
                {
                    Source = iconSource,
                    Width = AppConstants.MenuItemIconSize,
                    Height = AppConstants.MenuItemIconSize,
                    Margin = new Thickness(0, 0, 0, 8),
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                RenderOptions.SetBitmapScalingMode(fileIcon, BitmapScalingMode.HighQuality);
                itemPanel.Children.Add(fileIcon);
            }
        }

        /// <summary>
        /// 添加文本到面板
        /// </summary>
        private static (TextBlock, TextBox) AddTextToPanel(StackPanel itemPanel, string fileName)
        {
            TextBlock itemText = new TextBlock
            {
                Text = fileName,
                Foreground = Brushes.White,
                FontSize = AppConstants.MenuItemFontSize,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                MaxHeight = 30,
                TextTrimming = TextTrimming.CharacterEllipsis,
                HorizontalAlignment = HorizontalAlignment.Center,
                Padding = new Thickness(2, 0, 2, 0)
            };
            itemPanel.Children.Add(itemText);

            TextBox editBox = new TextBox
            {
                Text = fileName,
                FontSize = AppConstants.MenuItemFontSize,
                TextAlignment = TextAlignment.Center,
                Visibility = Visibility.Collapsed,
                Width = 70,
                MaxHeight = 30,
                TextWrapping = TextWrapping.Wrap,
                Padding = new Thickness(0),
                Margin = new Thickness(0)
            };
            itemPanel.Children.Add(editBox);

            return (itemText, editBox);
        }

        /// <summary>
        /// 设置菜单项交互
        /// </summary>
        private void SetupMenuItemInteractions(Border itemBorder, string path, bool isDir, Action safeCloseMenu,
            MenuContext context, TextBlock itemText, TextBox editBox)
        {
            itemBorder.MouseEnter += (s, e) =>
            {
                if (!context.SelectedBorders.Contains(itemBorder))
                {
                    itemBorder.Background = new SolidColorBrush(AppConstants.HoverBackgroundColor);
                }
            };
            itemBorder.MouseLeave += (s, e) =>
            {
                if (!context.SelectedBorders.Contains(itemBorder))
                {
                    itemBorder.Background = Brushes.Transparent;
                }
            };

            Action<string> finishRename = (newName) =>
            {
                editBox.Visibility = Visibility.Collapsed;
                itemText.Visibility = Visibility.Visible;
                
                if (string.IsNullOrWhiteSpace(newName) || newName == Path.GetFileName(path)) return;

                ExceptionHandler.SafeExecute(() =>
                {
                    string? dir = Path.GetDirectoryName(path);
                    if (dir == null) return;
                    
                    string ext = isDir ? "" : Path.GetExtension(path);
                    string newFileName = isDir ? newName : (newName.EndsWith(ext, StringComparison.OrdinalIgnoreCase) ? newName : newName + ext);
                    string newPath = Path.Combine(dir, newFileName);

                    if (isDir)
                        Directory.Move(path, newPath);
                    else
                        File.Move(path, newPath);

                    itemText.Text = newFileName;

                    int idx = context.SelectedPaths.IndexOf(path);
                    if (idx >= 0) context.SelectedPaths[idx] = newPath;
                    
                    _logger.LogInfo($"重命名: {path} -> {newPath}");
                }, "重命名文件");
            };

            editBox.LostFocus += (s, e) => { finishRename(editBox.Text); };
            editBox.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    finishRename(editBox.Text);
                    e.Handled = true;
                }
                else if (e.Key == Key.Escape)
                {
                    editBox.Visibility = Visibility.Collapsed;
                    itemText.Visibility = Visibility.Visible;
                    e.Handled = true;
                }
            };

            Point startPoint = new Point();
            bool isMouseDown = false;
            System.Windows.Threading.DispatcherTimer? renameTimer = null;
            bool wasAlreadySelected = false;

            itemBorder.MouseLeftButtonDown += (s, e) =>
            {
                wasAlreadySelected = context.SelectedBorders.Contains(itemBorder);
                renameTimer?.Stop();
                renameTimer = null;

                if (e.ClickCount == 2)
                {
                    isMouseDown = false;
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => { SafeOpenItem(path, isDir); });
                    e.Handled = true;
                }
                else
                {
                    if (Keyboard.Modifiers == ModifierKeys.Control)
                    {
                        HandleCtrlClick(itemBorder, path, isDir, context, itemText, editBox, wasAlreadySelected);
                    }
                    else
                    {
                        if (wasAlreadySelected && context.SelectedPaths.Count == 1 && !path.StartsWith("::"))
                        {
                            var localRenameTimer = new System.Windows.Threading.DispatcherTimer
                            {
                                Interval = TimeSpan.FromMilliseconds(AppConstants.RenameDelayMs)
                            };
                            localRenameTimer.Tick += (ts, te) =>
                            {
                                localRenameTimer.Stop();
                                itemText.Visibility = Visibility.Collapsed;
                                editBox.Text = Path.GetFileNameWithoutExtension(path);
                                editBox.Visibility = Visibility.Visible;
                                editBox.Focus();
                                editBox.SelectAll();
                            };
                            localRenameTimer.Start();
                            renameTimer = localRenameTimer;
                        }
                        else
                        {
                            ClearSelection(context);
                            AddToSelection(itemBorder, path, isDir, context, itemText, editBox);
                        }
                    }
                    startPoint = e.GetPosition(null);
                    isMouseDown = true;
                    _currentMenu?.Focus();
                }
            };

            itemBorder.MouseMove += (s, e) =>
            {
                if (isMouseDown && e.LeftButton == MouseButtonState.Pressed)
                {
                    Vector diff = startPoint - e.GetPosition(null);
                    if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                        Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                    {
                        renameTimer?.Stop();
                        renameTimer = null;
                        isMouseDown = false;
                        StartDragDrop(itemBorder, path, context, safeCloseMenu);
                    }
                }
            };

            itemBorder.MouseLeftButtonUp += (s, e) => { isMouseDown = false; };

            SetupDragDrop(itemBorder, path, isDir, safeCloseMenu, context);
        }

        /// <summary>
        /// 处理 Ctrl+点击事件
        /// </summary>
        private void HandleCtrlClick(Border itemBorder, string path, bool isDir, MenuContext context,
            TextBlock itemText, TextBox editBox, bool wasAlreadySelected)
        {
            if (wasAlreadySelected)
            {
                RemoveFromSelection(itemBorder, path, context, itemText, editBox);
                itemBorder.Background = Brushes.Transparent;
            }
            else
            {
                AddToSelection(itemBorder, path, isDir, context, itemText, editBox);
            }
        }

        /// <summary>
        /// 清除选择
        /// </summary>
        private static void ClearSelection(MenuContext context)
        {
            foreach (var b in context.SelectedBorders)
            {
                b.Background = Brushes.Transparent;
            }
            context.SelectedBorders.Clear();
            context.SelectedTexts.Clear();
            context.SelectedTextBoxes.Clear();
            context.SelectedPaths.Clear();
            context.IsDirectory.Clear();
        }

        /// <summary>
        /// 添加到选择
        /// </summary>
        private void AddToSelection(Border itemBorder, string path, bool isDir, MenuContext context,
            TextBlock itemText, TextBox editBox)
        {
            context.SelectedBorders.Add(itemBorder);
            context.SelectedTexts.Add(itemText);
            context.SelectedTextBoxes.Add(editBox);
            context.SelectedPaths.Add(path);
            context.IsDirectory.Add(isDir);
            itemBorder.Background = new SolidColorBrush(AppConstants.SelectedBackgroundColor);
        }

        /// <summary>
        /// 从选择中移除
        /// </summary>
        private static void RemoveFromSelection(Border itemBorder, string path, MenuContext context,
            TextBlock itemText, TextBox editBox)
        {
            context.SelectedBorders.Remove(itemBorder);
            context.SelectedTexts.Remove(itemText);
            context.SelectedTextBoxes.Remove(editBox);
            int idx = context.SelectedPaths.IndexOf(path);
            if (idx >= 0)
            {
                context.SelectedPaths.RemoveAt(idx);
                context.IsDirectory.RemoveAt(idx);
            }
        }

        /// <summary>
        /// 开始拖放操作
        /// </summary>
        private void StartDragDrop(Border itemBorder, string path, MenuContext context, Action safeCloseMenu)
        {
            _keepMenuAlive = true;
            try
            {
                string[] dragFiles = context.SelectedPaths.Contains(path)
                    ? context.SelectedPaths.ToArray()
                    : new string[] { path };
                DataObject dragData = new DataObject(DataFormats.FileDrop, dragFiles);
                DragDrop.DoDragDrop(itemBorder, dragData, DragDropEffects.Copy | DragDropEffects.Move);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"拖放操作失败: {ex.Message}");
            }
            finally
            {
                _keepMenuAlive = false;
                Application.Current.Dispatcher.InvokeAsync(() => { safeCloseMenu(); });
            }
        }

        /// <summary>
        /// 设置拖放处理
        /// </summary>
        private void SetupDragDrop(Border itemBorder, string path, bool isDir, Action safeCloseMenu, MenuContext context)
        {
            bool isRecycleBin = path == AppConstants.CLSID_RecycleBin;
            
            if ((!isDir || path.StartsWith("::")) && !isRecycleBin) return;

            itemBorder.AllowDrop = true;

            itemBorder.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    itemBorder.Background = new SolidColorBrush(AppConstants.DropTargetColor);
                }
            };

            itemBorder.DragLeave += (s, e) =>
            {
                if (!context.SelectedBorders.Contains(itemBorder))
                    itemBorder.Background = Brushes.Transparent;
                else
                    itemBorder.Background = new SolidColorBrush(AppConstants.SelectedBackgroundColor);
            };

            itemBorder.Drop += (s, e) =>
            {
                if (!context.SelectedBorders.Contains(itemBorder))
                    itemBorder.Background = Brushes.Transparent;
                else
                    itemBorder.Background = new SolidColorBrush(AppConstants.SelectedBackgroundColor);

                HandleDrop(path, isRecycleBin, safeCloseMenu, e);
            };
        }

        /// <summary>
        /// 处理放置操作
        /// </summary>
        private void HandleDrop(string path, bool isRecycleBin, Action safeCloseMenu, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            bool actedAny = false;

            if (isRecycleBin)
            {
                if (MessageBox.Show($"确定要把这 {files.Length} 个项目移入回收站吗？",
                    "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    foreach (string sourceFile in files)
                    {
                        if (ExceptionHandler.SafeExecute(() =>
                        {
                            PerformDelete(sourceFile, Directory.Exists(sourceFile));
                        }, $"移入回收站: {sourceFile}"))
                        {
                            actedAny = true;
                        }
                    }
                }
            }
            else
            {
                foreach (string sourceFile in files)
                {
                    if (sourceFile.Equals(path, StringComparison.OrdinalIgnoreCase)) continue;
                    
                    if (ExceptionHandler.SafeExecute(() =>
                    {
                        string destFile = Path.Combine(path, Path.GetFileName(sourceFile));
                        if (Directory.Exists(sourceFile))
                            Directory.Move(sourceFile, destFile);
                        else
                            File.Move(sourceFile, destFile);
                        _logger.LogInfo($"移动文件: {sourceFile} -> {destFile}");
                    }, $"移动文件: {sourceFile}"))
                    {
                        actedAny = true;
                    }
                }
            }

            if (actedAny)
            {
                safeCloseMenu();
                RefreshDesktop();
            }
        }

        /// <summary>
        /// 设置右键菜单
        /// </summary>
        private void SetupContextMenu(Border itemBorder, string path, bool isDir, Action safeCloseMenu,
            MenuContext context, TextBlock itemText, TextBox editBox)
        {
            ContextMenu ctxMenu = new ContextMenu();

            MenuItem openItem = new MenuItem { Header = "打开" };
            openItem.Click += (s, e) =>
            {
                safeCloseMenu();
                Application.Current.Dispatcher.InvokeAsync(() => { SafeOpenItem(path, isDir); });
            };
            ctxMenu.Items.Add(openItem);

            if (!isDir)
            {
                MenuItem openWithItem = new MenuItem { Header = "打开方式..." };
                openWithItem.Click += (s, e) =>
                {
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        ExceptionHandler.SafeExecute(() =>
                        {
                            OPENASINFO info = new OPENASINFO
                            {
                                pcszFile = path,
                                pcszClass = null,
                                oaUI = AppConstants.OAIF_ALLOW_REGISTRATION | AppConstants.OAIF_EXEC
                            };
                            SHOpenWithDialog(IntPtr.Zero, ref info);
                        }, "打开方式对话框");
                    });
                };
                ctxMenu.Items.Add(openWithItem);
            }

            ctxMenu.Items.Add(new Separator());

            MenuItem copyItem = new MenuItem { Header = "复制" };
            copyItem.Click += (s, e) =>
            {
                safeCloseMenu();
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    ExceptionHandler.SafeExecute(() =>
                    {
                        var sc = new System.Collections.Specialized.StringCollection();
                        if (context.SelectedPaths.Contains(path))
                        {
                            foreach (var p in context.SelectedPaths) sc.Add(p);
                        }
                        else
                        {
                            sc.Add(path);
                        }
                        System.Windows.Clipboard.SetFileDropList(sc);
                        _logger.LogInfo($"已复制 {sc.Count} 个项目到剪贴板");
                    }, "复制文件");
                });
            };
            ctxMenu.Items.Add(copyItem);

            MenuItem showItem = new MenuItem { Header = "在文件夹中显示" };
            showItem.Click += (s, e) =>
            {
                safeCloseMenu();
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    ExceptionHandler.SafeExecute(() =>
                    {
                        Process.Start("explorer.exe", $"/select,\"{path}\"");
                    }, "在文件夹中显示");
                });
            };
            ctxMenu.Items.Add(showItem);

            ctxMenu.Items.Add(new Separator());

            if (!path.StartsWith("::"))
            {
                MenuItem propItem = new MenuItem { Header = "属性" };
                propItem.Click += (s, e) =>
                {
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        ExceptionHandler.SafeExecute(() =>
                        {
                            SHELLEXECUTEINFO info = new SHELLEXECUTEINFO();
                            info.cbSize = Marshal.SizeOf(typeof(SHELLEXECUTEINFO));
                            info.lpVerb = "properties";
                            info.lpFile = path;
                            info.nShow = 5;
                            info.fMask = AppConstants.SEE_MASK_INVOKEIDLIST;
                            ShellExecuteEx(ref info);
                        }, "显示属性");
                    });
                };
                ctxMenu.Items.Add(propItem);
            }

            if (!path.StartsWith("::"))
            {
                MenuItem renameItem = new MenuItem { Header = "重命名 (F2)" };
                renameItem.Click += (s, e) =>
                {
                    itemText.Visibility = Visibility.Collapsed;
                    editBox.Text = Path.GetFileNameWithoutExtension(path);
                    editBox.Visibility = Visibility.Visible;
                    editBox.Focus();
                    editBox.SelectAll();
                };
                ctxMenu.Items.Add(renameItem);
            }

            if (!path.StartsWith("::"))
            {
                MenuItem delItem = new MenuItem { Header = "删除文件", Foreground = Brushes.Red };
                delItem.Click += (s, e) =>
                {
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (context.SelectedPaths.Contains(path) && context.SelectedPaths.Count > 1)
                        {
                            if (MessageBox.Show($"确定要把这 {context.SelectedPaths.Count} 个项目移入回收站吗？",
                                "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                            {
                                for (int i = 0; i < context.SelectedPaths.Count; i++)
                                {
                                    if (!context.SelectedPaths[i].StartsWith("::"))
                                    {
                                        PerformDelete(context.SelectedPaths[i], context.IsDirectory[i]);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (MessageBox.Show($"确定要把 {GetDisplayName(path)} 移入回收站吗？",
                                "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                            {
                                PerformDelete(path, isDir);
                            }
                        }
                    });
                };
                ctxMenu.Items.Add(delItem);
            }

            itemBorder.ContextMenu = ctxMenu;

            if (_currentMenu != null)
            {
                _currentMenu.KeyDown += (s, e) =>
                {
                    HandleMenuKeyDown(e, path, isDir, safeCloseMenu, context, itemText, editBox);
                };
            }
        }

        /// <summary>
        /// 处理菜单键盘事件
        /// </summary>
        private void HandleMenuKeyDown(KeyEventArgs e, string path, bool isDir, Action safeCloseMenu,
            MenuContext context, TextBlock itemText, TextBox editBox)
        {
            if (context.SelectedPaths.Count == 0) return;

            if (e.Key == Key.F2 && context.SelectedPaths.Count == 1)
            {
                if (!context.SelectedPaths[0].StartsWith("::"))
                {
                    TextBlock tb = context.SelectedTexts[0];
                    TextBox box = context.SelectedTextBoxes[0];
                    tb.Visibility = Visibility.Collapsed;
                    box.Text = Path.GetFileNameWithoutExtension(context.SelectedPaths[0]);
                    box.Visibility = Visibility.Visible;
                    box.Focus();
                    box.SelectAll();
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
            {
                ExceptionHandler.SafeExecute(() =>
                {
                    var sc = new System.Collections.Specialized.StringCollection();
                    foreach (var p in context.SelectedPaths) sc.Add(p);
                    System.Windows.Clipboard.SetFileDropList(sc);
                    
                    foreach (var b in context.SelectedBorders)
                    {
                        if (b != null)
                        {
                            b.Background = new SolidColorBrush(AppConstants.CopyFeedbackColor);
                            var timer = new System.Windows.Threading.DispatcherTimer
                            {
                                Interval = TimeSpan.FromMilliseconds(AppConstants.CopyFeedbackDurationMs)
                            };
                            Border tb = b;
                            timer.Tick += (ts, te) =>
                            {
                                tb.Background = new SolidColorBrush(AppConstants.SelectedBackgroundColor);
                                timer.Stop();
                            };
                            timer.Start();
                        }
                    }
                    _logger.LogInfo($"已复制 {sc.Count} 个项目到剪贴板");
                }, "复制文件");
            }
            else if (e.Key == Key.Delete)
            {
                var targets = new List<string>(context.SelectedPaths);
                var dirs = new List<bool>(context.IsDirectory);
                safeCloseMenu();

                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    string msg = targets.Count == 1
                        ? $"确定要把 {Path.GetFileName(targets[0])} 移入回收站吗？"
                        : $"确定要把这 {targets.Count} 个项目移入回收站吗？";
                    
                    if (MessageBox.Show(msg, "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        for (int i = 0; i < targets.Count; i++)
                        {
                            PerformDelete(targets[i], dirs[i]);
                        }
                    }
                });
            }
        }

        /// <summary>
        /// 定位并显示菜单
        /// </summary>
        private void PositionAndShowMenu(UIElement anchor, Border menuBorder)
        {
            if (_currentMenu == null) return;

            double maxAllowedHeight = SystemParameters.WorkArea.Height * AppConstants.MenuMaxHeightRatio;
            menuBorder.Measure(new Size(Double.PositiveInfinity, maxAllowedHeight + 50));
            double menuWidth = menuBorder.DesiredSize.Width;
            double menuHeight = menuBorder.DesiredSize.Height;

            PresentationSource? source = PresentationSource.FromVisual(this);
            double dpiX = 1.0;
            double dpiY = 1.0;
            if (source != null && source.CompositionTarget != null)
            {
                dpiX = source.CompositionTarget.TransformToDevice.M11;
                dpiY = source.CompositionTarget.TransformToDevice.M22;
            }

            Point screenPos = anchor.PointToScreen(new Point(0, 0));
            double logicalX = screenPos.X / dpiX;
            double logicalY = screenPos.Y / dpiY;

            Rect workArea = SystemParameters.WorkArea;

            double left = logicalX - (menuWidth / 2) + AppConstants.MenuHorizontalOffset;
            if (StacksContainer.HorizontalAlignment == HorizontalAlignment.Right)
            {
                left = logicalX - menuWidth + AppConstants.MenuRightLayoutOffset;
            }
            else if (StacksContainer.HorizontalAlignment == HorizontalAlignment.Left)
            {
                left = logicalX + AppConstants.MenuLeftLayoutOffset;
            }

            double top = logicalY + AppConstants.MenuVerticalOffset;
            if (StacksContainer.VerticalAlignment == VerticalAlignment.Bottom)
            {
                top = logicalY - menuHeight - AppConstants.MenuBottomLayoutOffset;
            }

            if (left < workArea.Left + AppConstants.MenuEdgeMargin) left = workArea.Left + AppConstants.MenuEdgeMargin;
            if (left + menuWidth > workArea.Right - AppConstants.MenuEdgeMargin) left = workArea.Right - menuWidth - AppConstants.MenuEdgeMargin;
            if (top < workArea.Top + AppConstants.MenuEdgeMargin) top = workArea.Top + AppConstants.MenuEdgeMargin;
            if (top + menuHeight > workArea.Bottom - AppConstants.MenuEdgeMargin) top = workArea.Bottom - menuHeight - AppConstants.MenuEdgeMargin;

            _currentMenu.Left = left;
            _currentMenu.Top = top;
            _currentMenu.Show();
            _currentMenu.Activate();
        }

        #endregion

        #region 资源清理

        /// <summary>
        /// 清理资源
        /// </summary>
        private void CleanupResources()
        {
            ExceptionHandler.SafeExecute(() =>
            {
                StopDesktopWatcher();
                CleanupTrayIcon();
                RemoveEventHandlers();
                UninitializeExceptionHandlers();

                _logger.LogInfo("资源已清理");
            }, "清理资源");
        }

        /// <summary>
        /// 退出应用程序
        /// </summary>
        private void ExitApplication()
        {
            this.Close();
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                CleanupResources();
                GC.SuppressFinalize(this);
            }
        }

        #endregion
    }
}
