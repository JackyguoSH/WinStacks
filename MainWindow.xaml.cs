
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
        private Window? _currentMenu;
        private bool _keepMenuAlive = false; 
        private bool _iconsRestored = false;
        private int _currentLayoutMode = 0; 

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
            ScanDesktopFiles();
            StartDesktopWatcher();
            LoadLayoutSettings();
        }

        private void LoadLayoutSettings()
        {
            try {
                string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinStacks");
                string configPath = Path.Combine(appDataDir, "layout_config.txt");

                if (!File.Exists(configPath)) {
                    Directory.CreateDirectory(appDataDir);
                    File.WriteAllText(configPath, "0"); 
                    ApplyLayout(0, false); 
                    
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        OpenSettingsWindow();
                        PinToDesktop();
                    }, System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                } else {
                    string savedMode = File.ReadAllText(configPath);
                    if (int.TryParse(savedMode, out int mode)) {
                        ApplyLayout(mode, false);
                    } else {
                        ApplyLayout(0, false);
                    }
                }
            } catch { 
                ApplyLayout(0, false); 
            }
        }

        private void SetupTrayIcon()
        {
            _trayIcon = new System.Windows.Forms.NotifyIcon();
            string exePath = Environment.ProcessPath ?? AppContext.BaseDirectory;
            try { _trayIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath); } catch { }
            
            _trayIcon.Text = "WinStacks - 智能叠放";
            _trayIcon.Visible = true;

            var menu = new System.Windows.Forms.ContextMenuStrip();
            menu.Items.Add("👁️ 隐藏/显示桌面图标", null, (s, e) => ToggleDesktopIcons());
            menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            menu.Items.Add("⚙️ 排版设置", null, (s, e) => OpenSettingsWindow());
            menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            menu.Items.Add("❌ 退出 WinStacks", null, (s, e) => ExitApplication());
            
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
                Title = "WinStacks 排版设置", Width = 320, Height = 200, WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize, Topmost = true, Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
                Foreground = Brushes.White, WindowStyle = WindowStyle.ToolWindow 
            };
            StackPanel sp = new StackPanel { Margin = new Thickness(25) };
            sp.Children.Add(new TextBlock { Text = "选择叠放区的位置：", FontSize = 14, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 15) });
            ComboBox cbLayout = new ComboBox { Height = 35, FontSize = 14 };
            cbLayout.Items.Add("📍 右侧 - 垂直排版 (Mac 风格)"); cbLayout.Items.Add("📍 左侧 - 垂直排版 (Win 风格)");
            cbLayout.Items.Add("📍 顶部 - 水平排版"); cbLayout.Items.Add("📍 底部 - 水平排版 (Dock 风格)");
            cbLayout.SelectedIndex = _currentLayoutMode; 
            sp.Children.Add(cbLayout);

            Button btnApply = new Button { Content = "保存并应用", Height = 35, Margin = new Thickness(0, 20, 0, 0), Background = new SolidColorBrush(Color.FromRgb(0, 120, 215)), Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            btnApply.Click += (s, e) => { ApplyLayout(cbLayout.SelectedIndex, true); settingsWin.Close(); };
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
                try {
                    string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinStacks");
                    Directory.CreateDirectory(appDataDir);
                    File.WriteAllText(Path.Combine(appDataDir, "layout_config.txt"), mode.ToString());
                } catch { }
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

            var categories = new Dictionary<string, List<string>>() {
                { "系统", new List<string>() }, { "照片", new List<string>() }, { "文档", new List<string>() }, { "影片", new List<string>() },
                { "音乐", new List<string>() }, { "压缩包", new List<string>() }, { "应用程序", new List<string>() },
                { "文件夹", new List<string>() }, { "其他", new List<string>() }
            };

            foreach (string dir in directories) categories["文件夹"].Add(dir);
            categories["系统"].Add("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
            categories["系统"].Add("::{645FF040-5081-101B-9F08-00AA002F954E}");

            foreach (string file in files) {
                try {
                    FileAttributes attributes = File.GetAttributes(file);
                    if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden || (attributes & FileAttributes.System) == FileAttributes.System) continue;
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".jpg" || ext == ".png" || ext == ".jpeg" || ext == ".gif" || ext == ".bmp" || ext == ".webp") categories["照片"].Add(file);
                    else if (ext == ".doc" || ext == ".docx" || ext == ".pdf" || ext == ".txt" || ext == ".xls" || ext == ".xlsx" || ext == ".ppt" || ext == ".pptx" || ext == ".md" || ext == ".csv") categories["文档"].Add(file);
                    else if (ext == ".mp4" || ext == ".mov" || ext == ".avi" || ext == ".mkv" || ext == ".wmv") categories["影片"].Add(file);
                    else if (ext == ".mp3" || ext == ".wav" || ext == ".flac" || ext == ".aac" || ext == ".m4a") categories["音乐"].Add(file);
                    else if (ext == ".zip" || ext == ".rar" || ext == ".7z" || ext == ".tar" || ext == ".gz") categories["压缩包"].Add(file);
                    else if (ext == ".lnk" || ext == ".url" || ext == ".exe" || ext == ".bat") categories["应用程序"].Add(file);
                    else categories["其他"].Add(file); 
                } catch { }
            }

            StacksContainer.Children.Clear();
            foreach (var category in categories) {
                if (category.Value.Count > 0) {
                    string displayTitle = category.Key == "其他" ? $"其他" : category.Key;
                    UIElement stackUI = CreateStackBlock(displayTitle, category.Value);
                    StacksContainer.Children.Add(stackUI);
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
                MessageBox.Show($"无法打开此项目：\n{ex.Message}", "WinStacks 错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
            if (_currentMenu != null) { 
                try { _currentMenu.Close(); } catch { } 
                _currentMenu = null; 
            }

            _currentMenu = new Window {
                WindowStyle = WindowStyle.None, AllowsTransparency = false, ShowInTaskbar = false, Topmost = false, 
                SizeToContent = SizeToContent.WidthAndHeight,
                Background = new SolidColorBrush(Color.FromArgb(0x01, 0x00, 0x00, 0x00)), 
                WindowStartupLocation = WindowStartupLocation.Manual
            };
            WindowChrome.SetWindowChrome(_currentMenu, new WindowChrome { GlassFrameThickness = new Thickness(-1), CaptionHeight = 0, UseAeroCaptionButtons = false });

            _currentMenu.SourceInitialized += (s, e) => {
                IntPtr hwnd = new WindowInteropHelper(_currentMenu).Handle;
                int dark = 1; DwmSetWindowAttribute(hwnd, 20, ref dark, sizeof(int));
                int backdrop = 3; DwmSetWindowAttribute(hwnd, 38, ref backdrop, sizeof(int)); 
                int corners = 2; DwmSetWindowAttribute(hwnd, 33, ref corners, sizeof(int));
                SetWindowPos(hwnd, new IntPtr(-1), 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            };

            Action safeCloseMenu = null!;
            safeCloseMenu = () => {
                if (_currentMenu != null) {
                    Window temp = _currentMenu;
                    _currentMenu = null; 
                    try { temp.Close(); } catch { }
                }
            };

            _currentMenu.Deactivated += (s, e) => { if (!_keepMenuAlive) { safeCloseMenu(); } };

            List<string> currentSelectedPaths = new List<string>();
            List<bool> isCurrentDirs = new List<bool>();
            List<Border> selectedItemBorders = new List<Border>();
            List<TextBlock> selectedItemTexts = new List<TextBlock>();
            List<TextBox> selectedItemBoxes = new List<TextBox>();

            _currentMenu.KeyDown += (s, e) => {
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
                    safeCloseMenu();

                    Application.Current.Dispatcher.InvokeAsync(() => {
                        string msg = targets.Count == 1 ? $"确定要把 {Path.GetFileName(targets[0])} 移入回收站吗？" : $"确定要把这 {targets.Count} 个项目移入回收站吗？";
                        if (MessageBox.Show(msg, "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
                            for (int i = 0; i < targets.Count; i++) {
                                PerformDelete(targets[i], dirs[i]);
                            }
                        }
                    });
                }
            };

            Border menuBorder = new Border { Background = Brushes.Transparent, Padding = new Thickness(15) };
            StackPanel outerPanel = new StackPanel { Orientation = Orientation.Vertical };
            outerPanel.Children.Add(new TextBlock { Text = title, Foreground = new SolidColorBrush(Color.FromArgb(0xC0, 0xFF, 0xFF, 0xFF)), FontWeight = FontWeights.SemiBold, FontSize = 12, Margin = new Thickness(4, 0, 0, 12) });
            

            int columnCount = Math.Min(6, filePaths.Count);
            double dynamicWidth = columnCount * 84; 
            
            WrapPanel gridPanel = new WrapPanel { 
                Orientation = Orientation.Horizontal, 
                Width = dynamicWidth 
            };

            foreach (string path in filePaths) {
                string fileName = Path.GetFileName(path);
                if (path == "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") fileName = "此电脑";
                else if (path == "::{645FF040-5081-101B-9F08-00AA002F954E}") fileName = "回收站";

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
                
                MenuItem openItem = new MenuItem { Header = "打开" };
                openItem.Click += (s, e) => { 
                    safeCloseMenu(); 
                    Application.Current.Dispatcher.InvokeAsync(() => { SafeOpenItem(path, isDir); });
                };
                ctxMenu.Items.Add(openItem);

                if (!isDir) {
                    MenuItem openWithItem = new MenuItem { Header = "打开方式..." };
                    openWithItem.Click += (s, e) => { 
                        safeCloseMenu();
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
                MenuItem copyItem = new MenuItem { Header = "复制" };
                copyItem.Click += (s, e) => { 
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        try { 
                            var sc = new System.Collections.Specialized.StringCollection(); 
                            if (currentSelectedPaths.Contains(path)) { foreach (var p in currentSelectedPaths) sc.Add(p); } else { sc.Add(path); }
                            System.Windows.Clipboard.SetFileDropList(sc); 
                        } catch { } 
                    });
                };
                ctxMenu.Items.Add(copyItem);

                MenuItem showItem = new MenuItem { Header = "在文件夹中显示" };
                showItem.Click += (s, e) => { 
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        try { Process.Start("explorer.exe", $"/select,\"{path}\""); } catch { } 
                    });
                };
                ctxMenu.Items.Add(showItem);
                ctxMenu.Items.Add(new Separator());

                MenuItem propItem = new MenuItem { Header = "属性" };
                propItem.Click += (s, e) => { 
                    safeCloseMenu();
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        try {
                            SHELLEXECUTEINFO info = new SHELLEXECUTEINFO(); info.cbSize = Marshal.SizeOf(info); info.lpVerb = "properties"; info.lpFile = path; info.nShow = 5; info.fMask = SEE_MASK_INVOKEIDLIST;
                            ShellExecuteEx(ref info); 
                        } catch { } 
                    });
                };
                if (!path.StartsWith("::")) ctxMenu.Items.Add(propItem);

                MenuItem renameItem = new MenuItem { Header = "重命名 (F2)" };
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

                MenuItem delItem = new MenuItem { Header = "删除文件", Foreground = Brushes.Red };
                delItem.Click += (s, e) => {
                    safeCloseMenu(); 
                    Application.Current.Dispatcher.InvokeAsync(() => {
                        if (currentSelectedPaths.Contains(path) && currentSelectedPaths.Count > 1) {
                            if (MessageBox.Show($"确定要把这 {currentSelectedPaths.Count} 个项目移入回收站吗？", "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
                                for(int i=0; i<currentSelectedPaths.Count; i++) {
                                    if (!currentSelectedPaths[i].StartsWith("::")) PerformDelete(currentSelectedPaths[i], isCurrentDirs[i]);
                                }
                            }
                        } else {
                            if (MessageBox.Show($"确定要把 {fileName} 移入回收站吗？", "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) {
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
                        safeCloseMenu();
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
                        
                        if (_currentMenu != null) { _currentMenu.Focus(); } 
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

            outerPanel.Children.Add(scrollViewer);
            menuBorder.Child = outerPanel; 
            if (_currentMenu != null) _currentMenu.Content = menuBorder;

            if (_currentMenu != null)
            {
                menuBorder.Measure(new Size(Double.PositiveInfinity, maxAllowedHeight + 50)); 
                double menuWidth = menuBorder.DesiredSize.Width; 
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

                _currentMenu.Left = left; _currentMenu.Top = top; 
                _currentMenu.Show(); 
                _currentMenu.Activate(); 
            }
        }
    }
}