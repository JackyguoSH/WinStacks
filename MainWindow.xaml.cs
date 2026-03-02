
#pragma warning disable CS8625
#pragma warning disable CS8600
#pragma warning disable CS8602
#pragma warning disable CS8604

using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Linq;
using System.Security.Permissions;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using System.Windows.Shell;

using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace WinStacks
{
    public partial class MainWindow : Window
    {
        
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        const uint SWP_NOSIZE = 0x0001; const uint SWP_NOMOVE = 0x0002; const uint SWP_NOACTIVATE = 0x0010;
        const int GWL_EXSTYLE = -20; const int GWL_STYLE = -16; const int WS_EX_TOOLWINDOW = 0x00000080; const int GWLP_HWNDPARENT = -8;
        const int WS_VISIBLE = 0x10000000;

        // --- 获取 Windows 原生系统图标 ---
        [StructLayout(LayoutKind.Sequential)]
        public struct SHFILEINFO { public IntPtr hIcon; public int iIcon; public uint dwAttributes; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName; };
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);
        
        [DllImport("shell32.dll")]
        public static extern int SHGetStockIconInfo(uint siid, uint uFlags, ref SHSTOCKICONINFO psii);
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHSTOCKICONINFO { public uint cbSize; public IntPtr hIcon; public int iSysImageIndex; public int iIcon; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szPath; }
        const uint SHGSI_ICON = 0x000000100; const uint SHGSI_LARGEICON = 0x000000000;
        const uint SIID_DRIVEFIXED = 59; const uint SIID_RECYCLER = 31;

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool DestroyIcon(IntPtr hIcon);
        const uint SHGFI_ICON = 0x000000100; const uint SHGFI_LARGEICON = 0x000000000;

        // --- 调用“属性”API ---
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHELLEXECUTEINFO { public int cbSize; public uint fMask; public IntPtr hwnd; [MarshalAs(UnmanagedType.LPTStr)] public string lpVerb; [MarshalAs(UnmanagedType.LPTStr)] public string lpFile; [MarshalAs(UnmanagedType.LPTStr)] public string lpParameters; [MarshalAs(UnmanagedType.LPTStr)] public string lpDirectory; public int nShow; public IntPtr hInstApp; public IntPtr lpIDList; [MarshalAs(UnmanagedType.LPTStr)] public string lpClass; public IntPtr hkeyClass; public uint dwHotKey; public IntPtr hIcon; public IntPtr hProcess; }
        const uint SEE_MASK_INVOKEIDLIST = 0x0000000C;

        // --- 调用 Win11 “打开方式”对话框 API ---
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHOpenWithDialog(IntPtr hwndParent, ref OPENASINFO poainfo);
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct OPENASINFO { public string pcszFile; public string pcszClass; public int oaUI; }
        const int OAIF_ALLOW_REGISTRATION = 0x00000001; 
        const int OAIF_EXEC = 0x00000004;

        // --- 原生移入回收站 API ---
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEOPSTRUCT { public IntPtr hwnd; public uint wFunc; public IntPtr pFrom; public IntPtr pTo; public ushort fFlags; [MarshalAs(UnmanagedType.Bool)] public bool fAnyOperationsAborted; public IntPtr hNameMappings; public IntPtr lpszProgressTitle; }
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);
        const uint FO_DELETE = 0x0003; 
        const ushort FOF_ALLOWUNDO = 0x0040; const ushort FOF_NOCONFIRMATION = 0x0010; const ushort FOF_SILENT = 0x0004; const ushort FOF_NOERRORUI = 0x0400;

        // --- 一键隐藏/显示桌面图标 API ---
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam); public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")] public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        const uint WM_COMMAND = 0x0111; const int ToggleDesktopCommand = 0x7402;

        // --- Win11 原生亚克力材质渲染 API ---
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private FileSystemWatcher? _watcher;
        private System.Windows.Forms.NotifyIcon? _trayIcon;
        private Dictionary<string, Window> _openMenus = new Dictionary<string, Window>();
        private Dictionary<string, Rect> _menuTheoreticalPositions = new Dictionary<string, Rect>();
        private bool _keepMenuAlive = false; 
        private bool _iconsRestored = false;
        private int _currentLayoutMode = 0; 
        private Dictionary<string, string> _appConfig = new Dictionary<string, string>();

        private string Loc(string zh, string en) { return _appConfig.TryGetValue("Lang", out var lang) && lang == "en" ? en : zh; }

        // 发布版移除日志功能
        private void LogDebug(string message) { }

        public MainWindow()
        {
            AppDomain.CurrentDomain.UnhandledException += (s, e) => { EnsureDesktopIconsVisible(); };
            Application.Current.DispatcherUnhandledException += (s, e) => { e.Handled = true; };
            Application.Current.Exit += (s, e) => { EnsureDesktopIconsVisible(); };

            InitializeComponent();
            this.Activated += (s, e) => { PinToDesktop(); };
            this.Loaded += MainWindow_Loaded;
            SetupTrayIcon();
        }

        private void PinToDesktop()
        {
            try {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero) {
                    SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                }
            } catch { }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | WS_EX_TOOLWINDOW);
            IntPtr progman = FindWindow("Progman", null);
            if (progman != IntPtr.Zero) SetWindowLongPtr(hwnd, GWLP_HWNDPARENT, progman);
            
            PinToDesktop();
            EnsureDesktopIconsHidden();
            LoadLayoutSettings();
            ScanDesktopFiles();
            StartDesktopWatcher();
        }

        private void LoadLayoutSettings()
        {
            try {
                string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinStacks");
                string configPath = Path.Combine(appDataDir, "config.txt");
                _appConfig.Clear();

                if (File.Exists(configPath)) {
                    foreach (var line in File.ReadAllLines(configPath)) {
                        int idx = line.IndexOf('=');
                        if (idx > 0) _appConfig[line.Substring(0, idx).Trim()] = line.Substring(idx + 1).Trim();
                    }
                }

                if (_appConfig.TryGetValue("Layout", out string? modeStr) && int.TryParse(modeStr, out int mode)) {
                    ApplyLayout(mode, false);
                } else {
                    ApplyLayout(0, false);
                }
                
                if (!_appConfig.ContainsKey("Lang")) _appConfig["Lang"] = "zh";
                if (!_appConfig.ContainsKey("EnableRecentFiles")) _appConfig["EnableRecentFiles"] = "true";
                if (!_appConfig.ContainsKey("MaxRecentFiles")) _appConfig["MaxRecentFiles"] = "8";
                string recentName = Loc("最近使用", "Recent Files");
                if (!_appConfig.ContainsKey($"Stack_{recentName}")) _appConfig[$"Stack_{recentName}"] = "100,100";

                // Migration from old config
                string oldConfig = Path.Combine(appDataDir, "layout_config.txt");
                if (File.Exists(oldConfig) && !File.Exists(configPath)) {
                    if (int.TryParse(File.ReadAllText(oldConfig), out int oldMode)) { ApplyLayout(oldMode, true); }
                    try { File.Delete(oldConfig); } catch { }
                }
            } catch { 
                ApplyLayout(0, false); 
            }
        }

        private void SaveConfig()
        {
            try {
                string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinStacks");
                Directory.CreateDirectory(appDataDir);
                List<string> lines = new List<string>();
                foreach (var kvp in _appConfig) lines.Add($"{kvp.Key}={kvp.Value}");
                File.WriteAllLines(Path.Combine(appDataDir, "config.txt"), lines);
            } catch { }
        }

        private void SetupTrayIcon()
        {
            _trayIcon = new System.Windows.Forms.NotifyIcon();
            string exePath = Environment.ProcessPath ?? AppContext.BaseDirectory;
            try { _trayIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath); } catch { }
            
            _trayIcon.Text = Loc("WinStacks - 智能叠放", "WinStacks - Smart Stacks");
            _trayIcon.Visible = true;

            var menu = new System.Windows.Forms.ContextMenuStrip();
            menu.Items.Add(Loc("⚙️ 设置中心 / Settings", "⚙️ Settings Center"), null, (s, e) => OpenSettingsWindow());
            menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            menu.Items.Add(Loc("❌ 退出程序 / Exit", "❌ Exit WinStacks"), null, (s, e) => ExitApplication());
            
            _trayIcon.ContextMenuStrip = menu;
            _trayIcon.DoubleClick += (s, e) => OpenSettingsWindow();
        }

        private IntPtr GetDesktopDefView()
        {
            IntPtr hwnd = FindWindow("Progman", null);
            IntPtr defView = FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (defView == IntPtr.Zero) {
                EnumWindows((tophandle, topparamhandle) => {
                    IntPtr p = FindWindowEx(tophandle, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if (p != IntPtr.Zero) { defView = p; return false; }
                    return true;
                }, IntPtr.Zero);
            }
            return defView;
        }

        private void ToggleDesktopIcons()
        {
            IntPtr defView = GetDesktopDefView();
            if (defView != IntPtr.Zero) SendMessage(defView, WM_COMMAND, (IntPtr)ToggleDesktopCommand, IntPtr.Zero);
        }

        private void EnsureDesktopIconsHidden()
        {
            IntPtr defView = GetDesktopDefView();
            if (defView != IntPtr.Zero) {
                IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
                if (listView != IntPtr.Zero) {
                    int style = GetWindowLong(listView, GWL_STYLE);
                    if ((style & WS_VISIBLE) != 0) { 
                        SendMessage(defView, WM_COMMAND, (IntPtr)ToggleDesktopCommand, IntPtr.Zero);
                    }
                }
            }
        }

        private void EnsureDesktopIconsVisible()
        {
            if (_iconsRestored) return; 
            IntPtr defView = GetDesktopDefView();
            if (defView != IntPtr.Zero) {
                IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
                if (listView != IntPtr.Zero) {
                    int style = GetWindowLong(listView, GWL_STYLE);
                    if ((style & WS_VISIBLE) == 0) { 
                        SendMessage(defView, WM_COMMAND, (IntPtr)ToggleDesktopCommand, IntPtr.Zero);
                    }
                }
            }
            _iconsRestored = true;
        }

        protected override void OnClosed(EventArgs e)
        {
            EnsureDesktopIconsVisible();
            if (_trayIcon != null) { _trayIcon.Visible = false; _trayIcon.Dispose(); }
            base.OnClosed(e);
        }

        private void ExitApplication()
        {
            this.Close(); 
        }

        private void OpenSettingsWindow()
        {
            Window settingsWin = new Window {
                Title = Loc("WinStacks 设置中心", "WinStacks Settings Center"), Width = 340, Height = 560, WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize, Topmost = true, Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
                Foreground = Brushes.White, WindowStyle = WindowStyle.ToolWindow 
            };
            StackPanel sp = new StackPanel { Margin = new Thickness(25) };
            
            sp.Children.Add(new TextBlock { Text = Loc("排版布局 (Layout Position)：", "Stack Layout Position:"), FontSize = 14, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10) });
            ComboBox cbLayout = new ComboBox { Height = 35, FontSize = 14 };
            cbLayout.Items.Add(Loc("📍 右侧 - 垂直排版 (Mac 风格)", "📍 Right - Vertical layout")); cbLayout.Items.Add(Loc("📍 左侧 - 垂直排版 (Win 风格)", "📍 Left - Vertical layout"));
            cbLayout.Items.Add(Loc("📍 顶部 - 水平排版", "📍 Top - Horizontal layout")); cbLayout.Items.Add(Loc("📍 底部 - 水平排版 (Dock 风格)", "📍 Bottom - Horizontal layout"));
            cbLayout.SelectedIndex = _currentLayoutMode; 
            sp.Children.Add(cbLayout);

            sp.Children.Add(new TextBlock { Text = Loc("显示语言 (Display Language)：", "Display Language:"), FontSize = 14, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 15, 0, 10) });
            ComboBox cbLang = new ComboBox { Height = 35, FontSize = 14 };
            cbLang.Items.Add("简体中文 (Chinese)"); cbLang.Items.Add("English");
            cbLang.SelectedIndex = _appConfig.TryGetValue("Lang", out var l) && l == "en" ? 1 : 0;
            sp.Children.Add(cbLang);

            sp.Children.Add(new Separator { Margin = new Thickness(0, 15, 0, 15), Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF)) });

            sp.Children.Add(new TextBlock { Text = Loc("最近使用 (Recent Files)：", "Recent Files Feature:"), FontSize = 14, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10) });
            CheckBox chkRecents = new CheckBox { Content = Loc("启用“最近使用”智能叠放", "Enable 'Recent Files' Stack"), Foreground = Brushes.White, FontSize = 13, Margin = new Thickness(0, 0, 0, 10), VerticalContentAlignment = VerticalAlignment.Center };
            chkRecents.IsChecked = _appConfig.TryGetValue("EnableRecentFiles", out string? recEn) ? recEn == "true" : true;
            sp.Children.Add(chkRecents);
            
            StackPanel maxFilesSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
            maxFilesSp.Children.Add(new TextBlock { Text = Loc("显示数量：", "Item count: "), Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center });
            TextBox txtMaxFiles = new TextBox { Width = 40, Text = _appConfig.TryGetValue("MaxRecentFiles", out string? maxF) ? maxF : "8", TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center };
            maxFilesSp.Children.Add(txtMaxFiles);
            sp.Children.Add(maxFilesSp);

            sp.Children.Add(new Separator { Margin = new Thickness(0, 5, 0, 15), Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF)) });

            Button btnToggleIcons = new Button { Content = Loc("👁️ 隐藏/显示原生桌面图标", "👁️ Toggle Native Desktop Icons"), Height = 35, Margin = new Thickness(0, 0, 0, 15), Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF)), Foreground = Brushes.White, Cursor = Cursors.Hand };
            btnToggleIcons.Click += (s, e) => { ToggleDesktopIcons(); };
            sp.Children.Add(btnToggleIcons);

            CheckBox chkAutoStart = new CheckBox { Content = Loc("🚀 随 Windows 开机自动启动", "🚀 Start WinStacks with Windows"), Foreground = Brushes.White, FontSize = 13, Margin = new Thickness(0, 0, 0, 15), VerticalContentAlignment = VerticalAlignment.Center };
            try {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false)) {
                    chkAutoStart.IsChecked = (key != null && key.GetValue("WinStacks") != null);
                }
            } catch { }
            sp.Children.Add(chkAutoStart);

            Button btnApply = new Button { Content = Loc("✔ 保存并应用", "✔ Save & Apply"), Height = 40, Margin = new Thickness(0, 10, 0, 0), Background = new SolidColorBrush(Color.FromRgb(0, 120, 215)), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            btnApply.Click += (s, e) => { 
                _appConfig["Lang"] = cbLang.SelectedIndex == 1 ? "en" : "zh";
                
                string recentNameEn = "Recent Files"; string recentNameZh = "最近使用";
                bool isRecentsNowEnabled = chkRecents.IsChecked == true;
                _appConfig["EnableRecentFiles"] = isRecentsNowEnabled ? "true" : "false";
                
                if (!isRecentsNowEnabled) {
                    if (_openMenus.ContainsKey(recentNameZh)) { _openMenus[recentNameZh].Close(); _openMenus.Remove(recentNameZh); }
                    if (_openMenus.ContainsKey(recentNameEn)) { _openMenus[recentNameEn].Close(); _openMenus.Remove(recentNameEn); }
                }

                if (int.TryParse(txtMaxFiles.Text, out int parsedMax) && parsedMax > 0 && parsedMax <= 30) { _appConfig["MaxRecentFiles"] = parsedMax.ToString(); }
                ApplyLayout(cbLayout.SelectedIndex, true); 
                
                try {
                    string exePath = Environment.ProcessPath ?? AppContext.BaseDirectory;
                    using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true)) {
                        if (chkAutoStart.IsChecked == true) key?.SetValue("WinStacks", $"\"{exePath}\"");
                        else key?.DeleteValue("WinStacks", false);
                    }
                } catch { }

                settingsWin.Close(); 
                SetupTrayIcon();
                RefreshDesktop(); 
            };
            sp.Children.Add(btnApply); settingsWin.Content = sp; settingsWin.Show();
        }

        private void ApplyLayout(int mode, bool saveToFile)
        {
            _currentLayoutMode = mode;
            switch (mode) {
                case 0: StacksContainer.Orientation = Orientation.Vertical; StacksContainer.HorizontalAlignment = HorizontalAlignment.Right; StacksContainer.VerticalAlignment = VerticalAlignment.Top; StacksContainer.Margin = new Thickness(0, 50, 20, 0); break;
                case 1: StacksContainer.Orientation = Orientation.Vertical; StacksContainer.HorizontalAlignment = HorizontalAlignment.Left; StacksContainer.VerticalAlignment = VerticalAlignment.Top; StacksContainer.Margin = new Thickness(20, 50, 0, 0); break;
                case 2: StacksContainer.Orientation = Orientation.Horizontal; StacksContainer.HorizontalAlignment = HorizontalAlignment.Center; StacksContainer.VerticalAlignment = VerticalAlignment.Top; StacksContainer.Margin = new Thickness(0, 20, 0, 0); break;
                case 3: StacksContainer.Orientation = Orientation.Horizontal; StacksContainer.HorizontalAlignment = HorizontalAlignment.Center; StacksContainer.VerticalAlignment = VerticalAlignment.Bottom; StacksContainer.Margin = new Thickness(0, 0, 0, 60); break;
            }

            if (saveToFile) {
                _appConfig["Layout"] = mode.ToString();
                SaveConfig();
            }
        }

        private ImageSource? GetNativeIcon(string path)
        {
            try {
                if (path == "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") {
                    SHSTOCKICONINFO sii = new SHSTOCKICONINFO(); sii.cbSize = (uint)Marshal.SizeOf(typeof(SHSTOCKICONINFO));
                    if (SHGetStockIconInfo(SIID_DRIVEFIXED, SHGSI_ICON | SHGSI_LARGEICON, ref sii) == 0 && sii.hIcon != IntPtr.Zero) {
                        ImageSource img = Imaging.CreateBitmapSourceFromHIcon(sii.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()); DestroyIcon(sii.hIcon); return img;
                    }
                } else if (path == "::{645FF040-5081-101B-9F08-00AA002F954E}") {
                    SHSTOCKICONINFO sii = new SHSTOCKICONINFO(); sii.cbSize = (uint)Marshal.SizeOf(typeof(SHSTOCKICONINFO));
                    if (SHGetStockIconInfo(SIID_RECYCLER, SHGSI_ICON | SHGSI_LARGEICON, ref sii) == 0 && sii.hIcon != IntPtr.Zero) {
                        ImageSource img = Imaging.CreateBitmapSourceFromHIcon(sii.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()); DestroyIcon(sii.hIcon); return img;
                    }
                }

                SHFILEINFO shfi = new SHFILEINFO();
                IntPtr res = SHGetFileInfo(path, 0, ref shfi, (uint)Marshal.SizeOf(shfi), SHGFI_ICON | SHGFI_LARGEICON);
                if (res != IntPtr.Zero && shfi.hIcon != IntPtr.Zero) {
                    ImageSource img = Imaging.CreateBitmapSourceFromHIcon(shfi.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    DestroyIcon(shfi.hIcon); return img;
                }
            } catch { }
            return null;
        }

        private void StartDesktopWatcher()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            _watcher = new FileSystemWatcher(desktopPath) { NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite, EnableRaisingEvents = true };
            _watcher.Created += (s, e) => RefreshDesktop(); _watcher.Deleted += (s, e) => RefreshDesktop(); _watcher.Renamed += (s, e) => RefreshDesktop();
        }

        private void RefreshDesktop() { Application.Current.Dispatcher.Invoke(() => ScanDesktopFiles()); }

        private void ScanDesktopFiles()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string[] files = Directory.GetFiles(desktopPath); string[] directories = Directory.GetDirectories(desktopPath);

            var catKeys = new { Sys = Loc("系统", "System"), Pic = Loc("照片", "Pictures"), Doc = Loc("文档", "Documents"), Vid = Loc("影片", "Videos"), Mus = Loc("音乐", "Music"), Zip = Loc("压缩包", "Archives"), App = Loc("应用程序", "Applications"), Dir = Loc("文件夹", "Folders"), Other = Loc("其他", "Others"), Recent = Loc("最近使用", "Recent Files") };

            var categories = new Dictionary<string, List<string>>() {
                { catKeys.Sys, new List<string>() }, { catKeys.Pic, new List<string>() }, { catKeys.Doc, new List<string>() }, { catKeys.Vid, new List<string>() },
                { catKeys.Mus, new List<string>() }, { catKeys.Zip, new List<string>() }, { catKeys.App, new List<string>() },
                { catKeys.Dir, new List<string>() }, { catKeys.Other, new List<string>() }, { catKeys.Recent, new List<string>() }
            };

            foreach (string dir in directories) categories[catKeys.Dir].Add(dir);
            categories[catKeys.Sys].Add("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
            categories[catKeys.Sys].Add("::{645FF040-5081-101B-9F08-00AA002F954E}");

            bool enableRecent = _appConfig.TryGetValue("EnableRecentFiles", out string? er) ? er == "true" : true;
            int maxRecent = _appConfig.TryGetValue("MaxRecentFiles", out string? mr) && int.TryParse(mr, out int m) ? m : 8;
            
            if (enableRecent) {
                var validFiles = new List<string>();
                foreach (string file in files) {
                    try {
                        FileAttributes attr = File.GetAttributes(file);
                        if ((attr & FileAttributes.Hidden) == 0 && (attr & FileAttributes.System) == 0) validFiles.Add(file);
                    } catch { }
                }
                var recentFiles = validFiles.OrderByDescending(f => File.GetLastWriteTime(f)).Take(maxRecent).ToList();
                categories[catKeys.Recent].AddRange(recentFiles);
            }

            foreach (string file in files) {
                try {
                    FileAttributes attributes = File.GetAttributes(file);
                    if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden || (attributes & FileAttributes.System) == FileAttributes.System) continue;
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".jpg" || ext == ".png" || ext == ".jpeg" || ext == ".gif" || ext == ".bmp" || ext == ".webp") categories[catKeys.Pic].Add(file);
                    else if (ext == ".doc" || ext == ".docx" || ext == ".pdf" || ext == ".txt" || ext == ".xls" || ext == ".xlsx" || ext == ".ppt" || ext == ".pptx" || ext == ".md" || ext == ".csv") categories[catKeys.Doc].Add(file);
                    else if (ext == ".mp4" || ext == ".mov" || ext == ".avi" || ext == ".mkv" || ext == ".wmv") categories[catKeys.Vid].Add(file);
                    else if (ext == ".mp3" || ext == ".wav" || ext == ".flac" || ext == ".aac" || ext == ".m4a") categories[catKeys.Mus].Add(file);
                    else if (ext == ".zip" || ext == ".rar" || ext == ".7z" || ext == ".tar" || ext == ".gz") categories[catKeys.Zip].Add(file);
                    else if (ext == ".lnk" || ext == ".url" || ext == ".exe" || ext == ".bat") categories[catKeys.App].Add(file);
                    else categories[catKeys.Other].Add(file); 
                } catch { }
            }

            StacksContainer.Children.Clear();
            foreach (var category in categories) {
                if (category.Value.Count > 0) {
                    UIElement stackUI = CreateStackBlock(category.Key, category.Value);
                    StacksContainer.Children.Add(stackUI);
                }
            }

            // Auto-refresh pinned menus containing live data
            if (enableRecent && categories[catKeys.Recent].Count > 0) {
                if (_openMenus.ContainsKey(catKeys.Recent)) {
                    bool keepAliveCache = _keepMenuAlive;
                    Window oldMenu = _openMenus[catKeys.Recent];
                    double l = oldMenu.Left, t = oldMenu.Top;
                    _openMenus.Remove(catKeys.Recent); oldMenu.Close();
                    // Temporary UI refresh to bypass point-collision shifts on live reload
                    string savePosTemp = $"{l},{t}";
                    if (!_appConfig.ContainsKey($"Stack_{catKeys.Recent}")) _appConfig[$"Stack_{catKeys.Recent}"] = savePosTemp;
                    
                    ShowStackMenu(this, catKeys.Recent, categories[catKeys.Recent]);
                } else if (_appConfig.ContainsKey($"Stack_{catKeys.Recent}")) {
                    ShowStackMenu(this, catKeys.Recent, categories[catKeys.Recent]);
                }
            }
        }

        private UIElement CreateStackBlock(string title, List<string> filePaths)
        {
            int count = filePaths.Count;
            Grid mainGrid = new Grid { Width = 74, Height = 95, Margin = new Thickness(12, 15, 12, 15), Cursor = Cursors.Hand, Background = Brushes.Transparent };
            ScaleTransform scaleTransform = new ScaleTransform(1.0, 1.0); mainGrid.RenderTransform = scaleTransform; mainGrid.RenderTransformOrigin = new Point(0.5, 0.5);
            Grid iconStackGrid = new Grid { Width = 60, Height = 60, VerticalAlignment = VerticalAlignment.Top };

            int maxIcons = Math.Min(3, filePaths.Count);
            double totalWidth = 44 + (maxIcons - 1) * 6;
            double centerOffset = (60 - totalWidth) / 2;

            for (int i = maxIcons - 1; i >= 0; i--) {
                ImageSource? source = GetNativeIcon(filePaths[i]);
                if (source != null) {
                    Image img = new Image {
                        Source = source, Width = 44, Height = 44, HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Top,
                        Margin = new Thickness(centerOffset + (maxIcons - 1 - i) * 6, centerOffset + (maxIcons - 1 - i) * 6, 0, 0), 
                        Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 6, ShadowDepth = 2, Opacity = 0.35 }
                    };
                    RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality); iconStackGrid.Children.Add(img);
                }
            }

            Border badge = new Border {
                Background = new SolidColorBrush(Color.FromArgb(0xE6, 0x1E, 0x1E, 0x1E)), BorderBrush = new SolidColorBrush(Color.FromArgb(0x50, 0xFF, 0xFF, 0xFF)), BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10), Padding = new Thickness(5, 0, 5, 0), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, -6, -4, 0),
                Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 4, ShadowDepth = 2, Opacity = 0.5 }
            };
            badge.Child = new TextBlock { Text = count.ToString(), FontSize = 10, FontWeight = FontWeights.Bold, Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center };
            iconStackGrid.Children.Add(badge);

            TextBlock titleText = new TextBlock { Text = title, FontSize = 13, Foreground = Brushes.White, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Bottom, Margin = new Thickness(0, 0, 0, 5), Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 4, ShadowDepth = 1, Opacity = 0.9 } };
            mainGrid.Children.Add(iconStackGrid); mainGrid.Children.Add(titleText);
            
            mainGrid.AllowDrop = true;
            System.Windows.Threading.DispatcherTimer? dragOpenTimer = null;

            mainGrid.MouseEnter += (s, e) => {
                DoubleAnimation anim = new DoubleAnimation(1.1, TimeSpan.FromMilliseconds(150)) { EasingFunction = new CubicEase() { EasingMode = EasingMode.EaseOut } };
                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, anim); scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, anim);
            };
            mainGrid.MouseLeave += (s, e) => {
                DoubleAnimation anim = new DoubleAnimation(1.0, TimeSpan.FromMilliseconds(150)) { EasingFunction = new CubicEase() { EasingMode = EasingMode.EaseOut } };
                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, anim); scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, anim);
            };
            mainGrid.MouseLeftButtonDown += (sender, e) => { ShowStackMenu(mainGrid, title, filePaths); e.Handled = true; };
            
            mainGrid.DragEnter += (s, e) => {
                if (dragOpenTimer != null) { dragOpenTimer.Stop(); }
                dragOpenTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(600) };
                dragOpenTimer.Tick += (ts, te) => {
                    if (dragOpenTimer != null) dragOpenTimer.Stop();
                    ShowStackMenu(mainGrid, title, filePaths);
                };
                dragOpenTimer.Start();
            };
            mainGrid.DragLeave += (s, e) => { if (dragOpenTimer != null) dragOpenTimer.Stop(); };
            mainGrid.Drop += (s, e) => { if (dragOpenTimer != null) dragOpenTimer.Stop(); };

            return mainGrid;
        }

        private void SafeOpenItem(string path, bool isDir)
        {
            try {
                if (path.StartsWith("::")) {
                    Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = path, UseShellExecute = true });
                    return;
                }

                if (isDir) { Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"\"{path}\"", UseShellExecute = true }); } 
                else { Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true }); }
            } catch (Exception ex) {
                MessageBox.Show(Loc($"无法打开此项目：\n{ex.Message}", $"Cannot open this item:\n{ex.Message}"), Loc("WinStacks 错误", "WinStacks Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PerformDelete(string path, bool isDir)
        {
            try {
                if (isDir) {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                        path, 
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, 
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                } else {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(
                        path, 
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, 
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                }
            } catch { }
        }

        private void ShowStackMenu(UIElement anchor, string title, List<string> filePaths)
        {
            if (_openMenus.TryGetValue(title, out Window? existingMenu)) {
                try { existingMenu.Activate(); return; } catch { _openMenus.Remove(title); }
            }

            Window newMenu = new Window {
                WindowStyle = WindowStyle.None, AllowsTransparency = false, ShowInTaskbar = false, Topmost = false, 
                SizeToContent = SizeToContent.Height,
                Background = new SolidColorBrush(Color.FromArgb(0x01, 0x00, 0x00, 0x00)), 
                WindowStartupLocation = WindowStartupLocation.Manual
            };
            WindowChrome.SetWindowChrome(newMenu, new WindowChrome { GlassFrameThickness = new Thickness(-1), CaptionHeight = 0, UseAeroCaptionButtons = false });

            newMenu.SourceInitialized += (s, e) => {
                IntPtr hwnd = new WindowInteropHelper(newMenu).Handle;
                int dark = 1; DwmSetWindowAttribute(hwnd, 20, ref dark, sizeof(int));
                int backdrop = 3; DwmSetWindowAttribute(hwnd, 38, ref backdrop, sizeof(int)); 
                int corners = 2; DwmSetWindowAttribute(hwnd, 33, ref corners, sizeof(int));
                
                int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
                SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | WS_EX_TOOLWINDOW);
                IntPtr progman = FindWindow("Progman", null);
                if (progman != IntPtr.Zero) SetWindowLongPtr(hwnd, GWLP_HWNDPARENT, progman);
                SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            };

            bool isPinned = _appConfig.ContainsKey($"Stack_{title}");

            Action safeCloseMenu = null!;
            safeCloseMenu = () => {
                if (_openMenus.ContainsKey(title)) {
                    Window temp = _openMenus[title];
                    _openMenus.Remove(title); 
                    _menuTheoreticalPositions.Remove(title);
                    try { temp.Close(); } catch { }
                }
            };

            newMenu.Closing += (s, e) => {
                try {
                    if (isPinned) {
                        _appConfig[$"Stack_{title}"] = $"{newMenu.Left},{newMenu.Top}";
                        SaveConfig();
                    } else if (_appConfig.ContainsKey($"Stack_{title}")) {
                        _appConfig.Remove($"Stack_{title}");
                        SaveConfig();
                    }
                } catch { }
            };

            newMenu.Deactivated += (s, e) => { if (!_keepMenuAlive && !isPinned) { safeCloseMenu(); } };

            List<string> currentSelectedPaths = new List<string>();
            List<bool> isCurrentDirs = new List<bool>();
            List<Border> selectedItemBorders = new List<Border>();
            List<TextBlock> selectedItemTexts = new List<TextBlock>();
            List<TextBox> selectedItemBoxes = new List<TextBox>();

            newMenu.KeyDown += (s, e) => {
                if (currentSelectedPaths.Count == 0) return;

                if (e.Key == Key.F2 && currentSelectedPaths.Count == 1) {
                    if (!currentSelectedPaths[0].StartsWith("::")) {
                        TextBlock tb = selectedItemTexts[0];
                        TextBox box = selectedItemBoxes[0];
                        tb.Visibility = Visibility.Collapsed;
                        box.Text = Path.GetFileNameWithoutExtension(currentSelectedPaths[0]);
                        box.Visibility = Visibility.Visible;
                        box.Focus();
                        box.SelectAll();
                        e.Handled = true;
                    }
                }
                else if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control) {
                    try {
                        var sc = new System.Collections.Specialized.StringCollection();
                        foreach (var p in currentSelectedPaths) sc.Add(p);
                        System.Windows.Clipboard.SetFileDropList(sc);
                        foreach (var b in selectedItemBorders) {
                            if (b != null) {
                                b.Background = new SolidColorBrush(Color.FromArgb(0x80, 0x4A, 0x90, 0xE2));
                                var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
                                Border tb = b;
                                timer.Tick += (ts, te) => { 
                                    tb.Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF)); 
                                    timer.Stop(); 
                                };
                                timer.Start();
                            }
                        }
                    } catch { }
                }
                else if (e.Key == Key.Delete) {
                    var targets = new List<string>(currentSelectedPaths);
                    var dirs = new List<bool>(isCurrentDirs);
                    if (!isPinned) safeCloseMenu();

                    Application.Current.Dispatcher.InvokeAsync(() => {
                        string msg = targets.Count == 1 ? Loc($"确定要把 {Path.GetFileName(targets[0])} 移入回收站吗？", $"Move {Path.GetFileName(targets[0])} to Recycle Bin?") : Loc($"确定要把这 {targets.Count} 个项目移入回收站吗？", $"Move {targets.Count} items to Recycle Bin?");
                        if (MessageBox.Show(msg, Loc("删除确认", "Confirm Delete"), MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
                            for (int i = 0; i < targets.Count; i++) {
                                PerformDelete(targets[i], dirs[i]);
                            }
                        }
                    });
                }
            };

            Border menuBorder = new Border { Background = Brushes.Transparent, Padding = new Thickness(15), Cursor = Cursors.Arrow };
            menuBorder.MouseLeftButtonDown += (s, e) => {
                if (e.OriginalSource is ScrollViewer || e.OriginalSource is TextBox || e.OriginalSource is Button || e.OriginalSource is MenuItem) return;
                try { newMenu.DragMove(); } catch { }
            };

            int columnCount = Math.Min(6, filePaths.Count);
            double dynamicWidth = columnCount * 84; 
            double windowWidth = Math.Max(145, dynamicWidth + 30); 
            newMenu.Width = windowWidth;
            
            Grid outerPanel = new Grid { HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch };
            outerPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            outerPanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            Grid headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 12), Background = Brushes.Transparent, HorizontalAlignment = HorizontalAlignment.Stretch };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            
            TextBlock titleTb = new TextBlock { Text = title, Foreground = new SolidColorBrush(Color.FromArgb(0xC0, 0xFF, 0xFF, 0xFF)), FontWeight = FontWeights.SemiBold, FontSize = 12, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(titleTb, 0); headerGrid.Children.Add(titleTb);
            
            TextBlock pinBtn = new TextBlock { Text = "📌", Foreground = isPinned ? new SolidColorBrush(Color.FromRgb(0, 120, 215)) : new SolidColorBrush(Color.FromArgb(0x60, 0xFF, 0xFF, 0xFF)), FontSize = 12, Cursor = Cursors.Hand, ToolTip = Loc("固定位置 (支持自由拖拽保存坐标)", "Pin Position (Draggable)"), VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Right, Padding = new Thickness(5, 0, 0, 0) };
            Grid.SetColumn(pinBtn, 1); headerGrid.Children.Add(pinBtn);
            pinBtn.MouseLeftButtonDown += (s, e) => {
                isPinned = !isPinned;
                pinBtn.Foreground = isPinned ? new SolidColorBrush(Color.FromRgb(0, 120, 215)) : new SolidColorBrush(Color.FromArgb(0x60, 0xFF, 0xFF, 0xFF));
                if (!isPinned && _appConfig.ContainsKey($"Stack_{title}")) {
                    _appConfig.Remove($"Stack_{title}"); SaveConfig();
                }
                e.Handled = true;
            };

            Grid.SetRow(headerGrid, 0);
            outerPanel.Children.Add(headerGrid);

            WrapPanel gridPanel = new WrapPanel { 
                Orientation = Orientation.Horizontal, 
                Width = dynamicWidth,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            foreach (string path in filePaths) {
                string fileName = Path.GetFileName(path);
                if (path == "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") fileName = Loc("此电脑", "This PC");
                else if (path == "::{645FF040-5081-101B-9F08-00AA002F954E}") fileName = Loc("回收站", "Recycle Bin");

                bool isDir = Directory.Exists(path) || path.StartsWith("::");

                Border itemBorder = new Border { Background = Brushes.Transparent, CornerRadius = new CornerRadius(6), Width = 76, Height = 88, Margin = new Thickness(4), Cursor = Cursors.Hand };
                StackPanel itemPanel = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
                
                ImageSource? iconSource = GetNativeIcon(path);
                if (iconSource != null) {
                    Image fileIcon = new Image { Source = iconSource, Width = 40, Height = 40, Margin = new Thickness(0, 0, 0, 8), HorizontalAlignment = HorizontalAlignment.Center };
                    RenderOptions.SetBitmapScalingMode(fileIcon, BitmapScalingMode.HighQuality);
                    itemPanel.Children.Add(fileIcon);
                }
                
                TextBlock itemText = new TextBlock { Text = fileName, Foreground = Brushes.White, FontSize = 11, TextAlignment = TextAlignment.Center, TextWrapping = TextWrapping.Wrap, MaxHeight = 30, TextTrimming = TextTrimming.CharacterEllipsis, HorizontalAlignment = HorizontalAlignment.Center, Padding = new Thickness(2,0,2,0) };
                itemPanel.Children.Add(itemText); 

                TextBox editBox = new TextBox { Text = fileName, FontSize = 11, TextAlignment = TextAlignment.Center, Visibility = Visibility.Collapsed, Width = 70, MaxHeight = 30, TextWrapping = TextWrapping.Wrap, Padding = new Thickness(0), Margin = new Thickness(0) };
                
                Action<string> finishRename = (newName) => {
                    editBox.Visibility = Visibility.Collapsed;
                    itemText.Visibility = Visibility.Visible;
                    if (string.IsNullOrWhiteSpace(newName) || newName == Path.GetFileName(path)) return;
                    
                    try {
                        string dir = Path.GetDirectoryName(path);
                        string ext = isDir ? "" : Path.GetExtension(path);
                        string newFileName = isDir ? newName : (newName.EndsWith(ext, StringComparison.OrdinalIgnoreCase) ? newName : newName + ext);
                        string newPath = Path.Combine(dir, newFileName);
                        
                        if (isDir) Directory.Move(path, newPath);
                        else File.Move(path, newPath);
                        
                        itemText.Text = newFileName;
                        
                        int idx = currentSelectedPaths.IndexOf(path);
                        if (idx >= 0) currentSelectedPaths[idx] = newPath;
                    } catch { }
                };

                editBox.LostFocus += (s, e) => { finishRename(editBox.Text); };
                editBox.KeyDown += (s, e) => {
                    if (e.Key == Key.Enter) { finishRename(editBox.Text); e.Handled = true; }
                    else if (e.Key == Key.Escape) { editBox.Visibility = Visibility.Collapsed; itemText.Visibility = Visibility.Visible; e.Handled = true; }
                };
                itemPanel.Children.Add(editBox);

                itemBorder.Child = itemPanel;

                itemBorder.MouseEnter += (s, e) => { if (!selectedItemBorders.Contains(itemBorder)) itemBorder.Background = new SolidColorBrush(Color.FromArgb(0x20, 0xFF, 0xFF, 0xFF)); };
                itemBorder.MouseLeave += (s, e) => { if (!selectedItemBorders.Contains(itemBorder)) itemBorder.Background = Brushes.Transparent; };
                
                ContextMenu ctxMenu = new ContextMenu();
                
                MenuItem openItem = new MenuItem { Header = Loc("打开", "Open") };
                openItem.Click += (s, e) => { 
                    if (!isPinned) safeCloseMenu(); 
                    Application.Current.Dispatcher.InvokeAsync(() => { SafeOpenItem(path, isDir); });
                };
                ctxMenu.Items.Add(openItem);

                if (!isDir) {
                    MenuItem openWithItem = new MenuItem { Header = Loc("打开方式...", "Open with...") };
                    openWithItem.Click += (s, e) => { 
                        if (!isPinned) safeCloseMenu();
                        Application.Current.Dispatcher.InvokeAsync(() => {
                            try {
                                OPENASINFO info = new OPENASINFO { pcszFile = path, pcszClass = null, oaUI = OAIF_ALLOW_REGISTRATION | OAIF_EXEC };
                                SHOpenWithDialog(IntPtr.Zero, ref info);
                            } catch { }
                        });
                    };
                    ctxMenu.Items.Add(openWithItem);
                }
                
                ctxMenu.Items.Add(new Separator());
                MenuItem copyItem = new MenuItem { Header = Loc("复制", "Copy") };
                copyItem.Click += (s, e) => { 
                    if (!isPinned) safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        try { 
                            var sc = new System.Collections.Specialized.StringCollection(); 
                            if (currentSelectedPaths.Contains(path)) { foreach (var p in currentSelectedPaths) sc.Add(p); } else { sc.Add(path); }
                            System.Windows.Clipboard.SetFileDropList(sc); 
                        } catch { } 
                    });
                };
                ctxMenu.Items.Add(copyItem);

                MenuItem showItem = new MenuItem { Header = Loc("在文件夹中显示", "Show in folder") };
                showItem.Click += (s, e) => { 
                    if (!isPinned) safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        try { Process.Start("explorer.exe", $"/select,\"{path}\""); } catch { } 
                    });
                };
                ctxMenu.Items.Add(showItem);
                ctxMenu.Items.Add(new Separator());

                MenuItem propItem = new MenuItem { Header = Loc("属性", "Properties") };
                propItem.Click += (s, e) => { 
                    if (!isPinned) safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        try {
                            SHELLEXECUTEINFO info = new SHELLEXECUTEINFO(); info.cbSize = Marshal.SizeOf(info); info.lpVerb = "properties"; info.lpFile = path; info.nShow = 5; info.fMask = SEE_MASK_INVOKEIDLIST;
                            ShellExecuteEx(ref info); 
                        } catch { } 
                    });
                };
                if (!path.StartsWith("::")) ctxMenu.Items.Add(propItem);

                MenuItem renameItem = new MenuItem { Header = Loc("重命名 (F2)", "Rename (F2)") };
                renameItem.Click += (s, e) => {
                    if (!path.StartsWith("::")) {
                        itemText.Visibility = Visibility.Collapsed;
                        editBox.Text = Path.GetFileNameWithoutExtension(path);
                        editBox.Visibility = Visibility.Visible;
                        editBox.Focus();
                        editBox.SelectAll();
                    }
                };
                if (!path.StartsWith("::")) ctxMenu.Items.Add(renameItem);

                MenuItem delItem = new MenuItem { Header = Loc("删除文件", "Delete"), Foreground = Brushes.Red };
                delItem.Click += (s, e) => {
                    if (!isPinned) safeCloseMenu(); 
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        if (currentSelectedPaths.Contains(path) && currentSelectedPaths.Count > 1) {
                            if (MessageBox.Show(Loc($"确定要把这 {currentSelectedPaths.Count} 个项目移入回收站吗？", $"Move {currentSelectedPaths.Count} items to Recycle Bin?"), Loc("删除确认", "Confirm Delete"), MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
                                for(int i=0; i<currentSelectedPaths.Count; i++) {
                                    if (!currentSelectedPaths[i].StartsWith("::")) PerformDelete(currentSelectedPaths[i], isCurrentDirs[i]);
                                }
                            }
                        } else {
                            if (MessageBox.Show(Loc($"确定要把 {fileName} 移入回收站吗？", $"Move {fileName} to Recycle Bin?"), Loc("删除确认", "Confirm Delete"), MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
                                PerformDelete(path, isDir);
                            }
                        }
                    });
                };
                if (!path.StartsWith("::")) ctxMenu.Items.Add(delItem);
                itemBorder.ContextMenu = ctxMenu;

                Point startPoint = new Point();
                bool isMouseDown = false;
                System.Windows.Threading.DispatcherTimer? renameTimer = null;
                bool wasAlreadySelected = false;

                itemBorder.MouseLeftButtonDown += (s, e) => {
                    wasAlreadySelected = selectedItemBorders.Contains(itemBorder);
                    if (renameTimer != null) { renameTimer.Stop(); renameTimer = null; }

                    if (e.ClickCount == 2) {
                        isMouseDown = false;
                        if (!isPinned) safeCloseMenu();
                        Application.Current.Dispatcher.InvokeAsync(() => { SafeOpenItem(path, isDir); });
                        e.Handled = true;
                    } else {
                        if (Keyboard.Modifiers == ModifierKeys.Control) {
                            if (wasAlreadySelected) {
                                selectedItemBorders.Remove(itemBorder);
                                selectedItemTexts.Remove(itemText);
                                selectedItemBoxes.Remove(editBox);
                                int idx = currentSelectedPaths.IndexOf(path);
                                if(idx >= 0) { currentSelectedPaths.RemoveAt(idx); isCurrentDirs.RemoveAt(idx); }
                                itemBorder.Background = Brushes.Transparent;
                            } else {
                                selectedItemBorders.Add(itemBorder);
                                selectedItemTexts.Add(itemText);
                                selectedItemBoxes.Add(editBox);
                                currentSelectedPaths.Add(path);
                                isCurrentDirs.Add(isDir);
                                itemBorder.Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF)); 
                            }
                        } else {
                            if (wasAlreadySelected && currentSelectedPaths.Count == 1 && !path.StartsWith("::")) {
                                renameTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                                renameTimer.Tick += (ts, te) => {
                                    if (renameTimer != null) renameTimer.Stop();
                                    itemText.Visibility = Visibility.Collapsed;
                                    editBox.Text = Path.GetFileNameWithoutExtension(path);
                                    editBox.Visibility = Visibility.Visible;
                                    editBox.Focus();
                                    editBox.SelectAll();
                                };
                                renameTimer.Start();
                            } else {
                                foreach (var b in selectedItemBorders) b.Background = Brushes.Transparent;
                                selectedItemBorders.Clear();
                                selectedItemTexts.Clear();
                                selectedItemBoxes.Clear();
                                currentSelectedPaths.Clear();
                                isCurrentDirs.Clear();

                                selectedItemBorders.Add(itemBorder);
                                selectedItemTexts.Add(itemText);
                                selectedItemBoxes.Add(editBox);
                                currentSelectedPaths.Add(path);
                                isCurrentDirs.Add(isDir);
                                itemBorder.Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF)); 
                            }
                        }
                        
                        startPoint = e.GetPosition(null); 
                        isMouseDown = true;
                        
                        newMenu.Focus(); 
                    }
                };

                itemBorder.MouseMove += (s, e) => {
                    if (isMouseDown && e.LeftButton == MouseButtonState.Pressed) {
                        Vector diff = startPoint - e.GetPosition(null);
                        if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance) {
                            if (renameTimer != null) { renameTimer.Stop(); renameTimer = null; }
                            isMouseDown = false;
                            _keepMenuAlive = true; 
                            try {
                                string[] dragFiles = currentSelectedPaths.Contains(path) ? currentSelectedPaths.ToArray() : new string[] { path };
                                DataObject dragData = new DataObject(DataFormats.FileDrop, dragFiles);
                                DragDrop.DoDragDrop(itemBorder, dragData, DragDropEffects.Copy | DragDropEffects.Move);
                            } catch { } 
                            finally {
                                _keepMenuAlive = false; 
                                Application.Current.Dispatcher.InvokeAsync(() => { safeCloseMenu(); });
                            }
                        }
                    }
                };

                itemBorder.MouseLeftButtonUp += (s, e) => { 
                    isMouseDown = false; 
                };

                bool isRecycleBin = path == "::{645FF040-5081-101B-9F08-00AA002F954E}";
                if ((isDir && !path.StartsWith("::")) || isRecycleBin) {
                    itemBorder.AllowDrop = true;
                    itemBorder.DragEnter += (s, e) => {
                        if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                            itemBorder.Background = new SolidColorBrush(Color.FromArgb(0x60, 0x4A, 0x90, 0xE2)); 
                        }
                    };
                    itemBorder.DragLeave += (s, e) => {
                        if (!selectedItemBorders.Contains(itemBorder)) itemBorder.Background = Brushes.Transparent;
                        else itemBorder.Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF));
                    };
                    itemBorder.Drop += (s, e) => {
                        if (!selectedItemBorders.Contains(itemBorder)) itemBorder.Background = Brushes.Transparent;
                        else itemBorder.Background = new SolidColorBrush(Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF));
                        
                        if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                            bool actedAny = false;
                            
                            if (isRecycleBin) {
                                if (MessageBox.Show($"确定要把这 {files.Length} 个项目移入回收站吗？", "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
                                    foreach (string sourceFile in files) {
                                        try { PerformDelete(sourceFile, Directory.Exists(sourceFile)); actedAny = true; } catch { }
                                    }
                                }
                            } else {
                                foreach (string sourceFile in files) {
                                    if (sourceFile.Equals(path, StringComparison.OrdinalIgnoreCase)) continue;
                                    try {
                                        string destFile = Path.Combine(path, Path.GetFileName(sourceFile));
                                        if (Directory.Exists(sourceFile)) {
                                            Directory.Move(sourceFile, destFile);
                                        } else {
                                            File.Move(sourceFile, destFile);
                                        }
                                        actedAny = true;
                                    } catch { }
                                }
                            }
                            
                            if (actedAny) {
                                safeCloseMenu();
                                RefreshDesktop();
                            }
                        }
                    };
                }

                gridPanel.Children.Add(itemBorder);
            }

            double maxAllowedHeight = SystemParameters.WorkArea.Height * 0.75; 
            ScrollViewer scrollViewer = new ScrollViewer {
                VerticalScrollBarVisibility = ScrollBarVisibility.Hidden, 
                MaxHeight = maxAllowedHeight,
                Content = gridPanel
            };
            
            scrollViewer.PreviewMouseWheel += (s, e) => {
                scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta);
                e.Handled = true;
            };

            Grid.SetRow(scrollViewer, 1);
            outerPanel.Children.Add(scrollViewer);
            menuBorder.Child = outerPanel; 
            newMenu.Content = menuBorder;

            menuBorder.Measure(new Size(Double.PositiveInfinity, maxAllowedHeight + 50)); 
            double menuWidth = windowWidth; 
            double menuHeight = menuBorder.DesiredSize.Height;
            
            PresentationSource source = PresentationSource.FromVisual(this);
            double dpiX = 1.0; double dpiY = 1.0;
            if (source != null && source.CompositionTarget != null) { dpiX = source.CompositionTarget.TransformToDevice.M11; dpiY = source.CompositionTarget.TransformToDevice.M22; }
            Point screenPos = anchor.PointToScreen(new Point(0, 0));
            double logicalX = screenPos.X / dpiX; double logicalY = screenPos.Y / dpiY;

            Rect workArea = SystemParameters.WorkArea;
            
            double left = logicalX - (menuWidth / 2) + 37; 
            if (StacksContainer.HorizontalAlignment == HorizontalAlignment.Right) {
                left = logicalX - menuWidth + 20; 
            } else if (StacksContainer.HorizontalAlignment == HorizontalAlignment.Left) {
                left = logicalX + 80; 
            }

            double top = logicalY + 110;
            if (StacksContainer.VerticalAlignment == VerticalAlignment.Bottom) { top = logicalY - menuHeight - 15; }

            double margin = 25;
            if (left < workArea.Left + margin) left = workArea.Left + margin; 
            if (left + menuWidth > workArea.Right - margin) left = workArea.Right - menuWidth - margin; 
            if (top < workArea.Top + margin) top = workArea.Top + margin; 
            if (top + menuHeight > workArea.Bottom - margin) top = workArea.Bottom - menuHeight - margin; 

            if (_appConfig.TryGetValue($"Stack_{title}", out string? savedPos)) {
                var parts = savedPos.Split(',');
                if (parts.Length == 2 && double.TryParse(parts[0], out double sx) && double.TryParse(parts[1], out double sy)) {
                    left = sx; top = sy;
                }
            } else {
                bool overlapping = true;
                int attempts = 0;
                while (overlapping && attempts < 15) {
                    overlapping = false;
                    Rect newRect = new Rect(left - 5, top - 5, menuWidth + 10, menuHeight + 10);
                    foreach (var kvp in _openMenus) {
                        Window win = kvp.Value;
                        double wl = win.Left;
                        double wt = win.Top;
                        double ww = win.ActualWidth > 0 ? win.ActualWidth : (_menuTheoreticalPositions.ContainsKey(kvp.Key) ? _menuTheoreticalPositions[kvp.Key].Width : 200);
                        double wh = win.ActualHeight > 0 ? win.ActualHeight : (_menuTheoreticalPositions.ContainsKey(kvp.Key) ? _menuTheoreticalPositions[kvp.Key].Height : 200);
                        
                        if (double.IsNaN(wl) || double.IsNaN(wt)) {
                            if (_menuTheoreticalPositions.ContainsKey(kvp.Key)) {
                                wl = _menuTheoreticalPositions[kvp.Key].Left;
                                wt = _menuTheoreticalPositions[kvp.Key].Top;
                            } else continue;
                        }

                        Rect existingRect = new Rect(wl, wt, ww, wh);
                        if (newRect.IntersectsWith(existingRect)) {
                            if (StacksContainer.HorizontalAlignment == HorizontalAlignment.Right) {
                                left = existingRect.Left - menuWidth - 15;
                            } else if (StacksContainer.HorizontalAlignment == HorizontalAlignment.Left) {
                                left = existingRect.Right + 15;
                            } else {
                                left -= 40; top -= 40;
                            }
                            overlapping = true;
                            break;
                        }
                    }
                    attempts++;
                }

                if (left < workArea.Left + margin) left = workArea.Left + margin; 
                if (left + menuWidth > workArea.Right - margin) left = workArea.Right - menuWidth - margin; 
                if (top < workArea.Top + margin) top = workArea.Top + margin; 
                if (top + menuHeight > workArea.Bottom - margin) top = workArea.Bottom - menuHeight - margin; 
            }

            _menuTheoreticalPositions[title] = new Rect(left, top, menuWidth, menuHeight);
            newMenu.Left = left; newMenu.Top = top; 
            _openMenus[title] = newMenu;
            newMenu.Show(); 
            newMenu.Activate(); 
        }
    }
}