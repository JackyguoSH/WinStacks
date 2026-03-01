using System;
using System.Windows;
using System.Windows.Media;

namespace WinStacks.Services
{
    /// <summary>
    /// 应用程序常量配置类，集中管理所有硬编码值
    /// </summary>
    public static class AppConstants
    {
        #region 窗口 API 常量

        /// <summary>
        /// 窗口放置在底层的句柄值
        /// </summary>
        public static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

        /// <summary>
        /// 窗口位置标志：保持当前大小
        /// </summary>
        public const uint SWP_NOSIZE = 0x0001;

        /// <summary>
        /// 窗口位置标志：保持当前位置
        /// </summary>
        public const uint SWP_NOMOVE = 0x0002;

        /// <summary>
        /// 窗口位置标志：不激活窗口
        /// </summary>
        public const uint SWP_NOACTIVATE = 0x0010;

        /// <summary>
        /// 扩展窗口样式索引
        /// </summary>
        public const int GWL_EXSTYLE = -20;

        /// <summary>
        /// 窗口样式索引
        /// </summary>
        public const int GWL_STYLE = -16;

        /// <summary>
        /// 工具窗口样式
        /// </summary>
        public const int WS_EX_TOOLWINDOW = 0x00000080;

        /// <summary>
        /// 窗口父句柄索引
        /// </summary>
        public const int GWLP_HWNDPARENT = -8;

        /// <summary>
        /// 可见窗口样式
        /// </summary>
        public const int WS_VISIBLE = 0x10000000;

        /// <summary>
        /// 顶层窗口句柄
        /// </summary>
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        #endregion

        #region Shell 图标常量

        /// <summary>
        /// 获取图标标志
        /// </summary>
        public const uint SHGFI_ICON = 0x000000100;

        /// <summary>
        /// 获取大图标标志
        /// </summary>
        public const uint SHGFI_LARGEICON = 0x000000000;

        /// <summary>
        /// 获取系统图标标志
        /// </summary>
        public const uint SHGSI_ICON = 0x000000100;

        /// <summary>
        /// 获取系统大图标标志
        /// </summary>
        public const uint SHGSI_LARGEICON = 0x000000000;

        /// <summary>
        /// 固定驱动器图标ID
        /// </summary>
        public const uint SIID_DRIVEFIXED = 59;

        /// <summary>
        /// 回收站图标ID
        /// </summary>
        public const uint SIID_RECYCLER = 31;

        #endregion

        #region Shell 操作常量

        /// <summary>
        /// Shell 执行标志：调用 ID 列表
        /// </summary>
        public const uint SEE_MASK_INVOKEIDLIST = 0x0000000C;

        /// <summary>
        /// 打开方式对话框标志：允许注册
        /// </summary>
        public const int OAIF_ALLOW_REGISTRATION = 0x00000001;

        /// <summary>
        /// 打开方式对话框标志：执行
        /// </summary>
        public const int OAIF_EXEC = 0x00000004;

        #endregion

        #region 文件操作常量

        /// <summary>
        /// 删除操作
        /// </summary>
        public const uint FO_DELETE = 0x0003;

        /// <summary>
        /// 允许撤销（回收站）
        /// </summary>
        public const ushort FOF_ALLOWUNDO = 0x0040;

        /// <summary>
        /// 不显示确认对话框
        /// </summary>
        public const ushort FOF_NOCONFIRMATION = 0x0010;

        /// <summary>
        /// 静默模式
        /// </summary>
        public const ushort FOF_SILENT = 0x0004;

        /// <summary>
        /// 不显示错误界面
        /// </summary>
        public const ushort FOF_NOERRORUI = 0x0400;

        #endregion

        #region 桌面图标控制常量

        /// <summary>
        /// WM_COMMAND 消息
        /// </summary>
        public const uint WM_COMMAND = 0x0111;

        /// <summary>
        /// 切换桌面图标命令
        /// </summary>
        public const int ToggleDesktopCommand = 0x7402;

        #endregion

        #region DWM 属性常量

        /// <summary>
        /// DWM 窗口属性：暗色模式
        /// </summary>
        public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        /// <summary>
        /// DWM 窗口属性：系统圆角
        /// </summary>
        public const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;

        /// <summary>
        /// DWM 窗口属性：Mica 效果
        /// </summary>
        public const int DWMWA_SYSTEMBACKDROP_TYPE = 38;

        /// <summary>
        /// DWM 暗色模式值
        /// </summary>
        public const int DWM_DARK_MODE_ENABLED = 1;

        /// <summary>
        /// DWM Mica 效果值
        /// </summary>
        public const int DWM_BACKDROP_MICA = 3;

        /// <summary>
        /// DWM 圆角值
        /// </summary>
        public const int DWM_CORNER_ROUND = 2;

        #endregion

        #region UI 尺寸常量

        /// <summary>
        /// 叠放区块宽度
        /// </summary>
        public const double StackBlockWidth = 74;

        /// <summary>
        /// 叠放区块高度
        /// </summary>
        public const double StackBlockHeight = 95;

        /// <summary>
        /// 叠放区块边距
        /// </summary>
        public const double StackBlockMargin = 12;

        /// <summary>
        /// 叠放区块垂直间距
        /// </summary>
        public const double StackBlockVerticalMargin = 15;

        /// <summary>
        /// 图标堆叠区域宽度
        /// </summary>
        public const double IconStackWidth = 60;

        /// <summary>
        /// 图标堆叠区域高度
        /// </summary>
        public const double IconStackHeight = 60;

        /// <summary>
        /// 单个图标大小
        /// </summary>
        public const double IconSize = 44;

        /// <summary>
        /// 图标堆叠偏移量
        /// </summary>
        public const double IconStackOffset = 6;

        /// <summary>
        /// 菜单项宽度
        /// </summary>
        public const double MenuItemWidth = 76;

        /// <summary>
        /// 菜单项高度
        /// </summary>
        public const double MenuItemHeight = 88;

        /// <summary>
        /// 菜单项边距
        /// </summary>
        public const double MenuItemMargin = 4;

        /// <summary>
        /// 菜单列数
        /// </summary>
        public const int MenuMaxColumns = 6;

        /// <summary>
        /// 菜单项图标大小
        /// </summary>
        public const double MenuItemIconSize = 40;

        /// <summary>
        /// 菜单最大高度比例（相对于工作区）
        /// </summary>
        public const double MenuMaxHeightRatio = 0.75;

        /// <summary>
        /// 设置窗口宽度
        /// </summary>
        public const double SettingsWindowWidth = 320;

        /// <summary>
        /// 设置窗口高度
        /// </summary>
        public const double SettingsWindowHeight = 200;

        /// <summary>
        /// 布局边距
        /// </summary>
        public const double LayoutMargin = 20;

        /// <summary>
        /// 顶部边距
        /// </summary>
        public const double TopMargin = 50;

        /// <summary>
        /// 底部边距
        /// </summary>
        public const double BottomMargin = 60;

        #endregion

        #region UI 颜色常量

        /// <summary>
        /// 默认背景颜色
        /// </summary>
        public static Color DefaultBackgroundColor => Color.FromRgb(30, 30, 30);

        /// <summary>
        /// 徽章背景颜色
        /// </summary>
        public static Color BadgeBackgroundColor => Color.FromArgb(0xE6, 0x1E, 0x1E, 0x1E);

        /// <summary>
        /// 徽章边框颜色
        /// </summary>
        public static Color BadgeBorderColor => Color.FromArgb(0x50, 0xFF, 0xFF, 0xFF);

        /// <summary>
        /// 主题蓝色
        /// </summary>
        public static Color ThemeBlueColor => Color.FromRgb(0, 120, 215);

        /// <summary>
        /// 悬停背景颜色
        /// </summary>
        public static Color HoverBackgroundColor => Color.FromArgb(0x20, 0xFF, 0xFF, 0xFF);

        /// <summary>
        /// 选中背景颜色
        /// </summary>
        public static Color SelectedBackgroundColor => Color.FromArgb(0x40, 0xFF, 0xFF, 0xFF);

        /// <summary>
        /// 拖放目标颜色
        /// </summary>
        public static Color DropTargetColor => Color.FromArgb(0x60, 0x4A, 0x90, 0xE2);

        /// <summary>
        /// 复制反馈颜色
        /// </summary>
        public static Color CopyFeedbackColor => Color.FromArgb(0x80, 0x4A, 0x90, 0xE2);

        /// <summary>
        /// 菜单标题文字颜色
        /// </summary>
        public static Color MenuTitleTextColor => Color.FromArgb(0xC0, 0xFF, 0xFF, 0xFF);

        /// <summary>
        /// 透明背景颜色（用于窗口）
        /// </summary>
        public static Color TransparentWindowColor => Color.FromArgb(0x01, 0x00, 0x00, 0x00);

        #endregion

        #region 动画常量

        /// <summary>
        /// 悬停动画持续时间（毫秒）
        /// </summary>
        public const int HoverAnimationDurationMs = 150;

        /// <summary>
        /// 悬停缩放比例
        /// </summary>
        public const double HoverScaleFactor = 1.1;

        /// <summary>
        /// 正常缩放比例
        /// </summary>
        public const double NormalScaleFactor = 1.0;

        /// <summary>
        /// 拖放打开菜单延迟（毫秒）
        /// </summary>
        public const int DragOpenMenuDelayMs = 600;

        /// <summary>
        /// 重命名延迟（毫秒）
        /// </summary>
        public const int RenameDelayMs = 500;

        /// <summary>
        /// 复制反馈持续时间（毫秒）
        /// </summary>
        public const int CopyFeedbackDurationMs = 200;

        #endregion

        #region 文件类型扩展名

        /// <summary>
        /// 图片文件扩展名
        /// </summary>
        public static readonly string[] ImageExtensions = { ".jpg", ".png", ".jpeg", ".gif", ".bmp", ".webp" };

        /// <summary>
        /// 文档文件扩展名
        /// </summary>
        public static readonly string[] DocumentExtensions = { ".doc", ".docx", ".pdf", ".txt", ".xls", ".xlsx", ".ppt", ".pptx", ".md", ".csv" };

        /// <summary>
        /// 视频文件扩展名
        /// </summary>
        public static readonly string[] VideoExtensions = { ".mp4", ".mov", ".avi", ".mkv", ".wmv" };

        /// <summary>
        /// 音频文件扩展名
        /// </summary>
        public static readonly string[] AudioExtensions = { ".mp3", ".wav", ".flac", ".aac", ".m4a" };

        /// <summary>
        /// 压缩包文件扩展名
        /// </summary>
        public static readonly string[] ArchiveExtensions = { ".zip", ".rar", ".7z", ".tar", ".gz" };

        /// <summary>
        /// 可执行文件扩展名
        /// </summary>
        public static readonly string[] ExecutableExtensions = { ".lnk", ".url", ".exe", ".bat" };

        #endregion

        #region 特殊文件夹 CLSID

        /// <summary>
        /// 此电脑 CLSID
        /// </summary>
        public const string CLSID_ThisPC = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

        /// <summary>
        /// 回收站 CLSID
        /// </summary>
        public const string CLSID_RecycleBin = "::{645FF040-5081-101B-9F08-00AA002F954E}";

        #endregion

        #region 配置文件

        /// <summary>
        /// 布局配置文件名
        /// </summary>
        public const string LayoutConfigFileName = "layout_config.txt";

        /// <summary>
        /// 应用程序数据目录名
        /// </summary>
        public const string AppDataDirectoryName = "WinStacks";

        /// <summary>
        /// 日志保留天数
        /// </summary>
        public const int LogRetentionDays = 30;

        #endregion

        #region UI 字体大小

        /// <summary>
        /// 标题字体大小
        /// </summary>
        public const double TitleFontSize = 13;

        /// <summary>
        /// 菜单标题字体大小
        /// </summary>
        public const double MenuTitleFontSize = 12;

        /// <summary>
        /// 菜单项字体大小
        /// </summary>
        public const double MenuItemFontSize = 11;

        /// <summary>
        /// 徽章字体大小
        /// </summary>
        public const double BadgeFontSize = 10;

        /// <summary>
        /// 设置标题字体大小
        /// </summary>
        public const double SettingsTitleFontSize = 14;

        /// <summary>
        /// 设置内容字体大小
        /// </summary>
        public const double SettingsContentFontSize = 14;

        #endregion

        #region 布局模式

        /// <summary>
        /// 布局模式：右侧垂直
        /// </summary>
        public const int LayoutModeRight = 0;

        /// <summary>
        /// 布局模式：左侧垂直
        /// </summary>
        public const int LayoutModeLeft = 1;

        /// <summary>
        /// 布局模式：顶部水平
        /// </summary>
        public const int LayoutModeTop = 2;

        /// <summary>
        /// 布局模式：底部水平
        /// </summary>
        public const int LayoutModeBottom = 3;

        #endregion

        #region 菜单位置计算

        /// <summary>
        /// 菜单水平偏移量
        /// </summary>
        public const double MenuHorizontalOffset = 37;

        /// <summary>
        /// 菜单右侧布局偏移量
        /// </summary>
        public const double MenuRightLayoutOffset = 20;

        /// <summary>
        /// 菜单左侧布局偏移量
        /// </summary>
        public const double MenuLeftLayoutOffset = 80;

        /// <summary>
        /// 菜单垂直偏移量
        /// </summary>
        public const double MenuVerticalOffset = 110;

        /// <summary>
        /// 菜单底部布局偏移量
        /// </summary>
        public const double MenuBottomLayoutOffset = 15;

        /// <summary>
        /// 菜单边距
        /// </summary>
        public const double MenuEdgeMargin = 25;

        #endregion
    }
}
