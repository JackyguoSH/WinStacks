using System;
using System.Runtime.InteropServices;

namespace WinStacks.Services
{
    /// <summary>
    /// Windows 版本检测服务，提供操作系统版本相关功能
    /// </summary>
    public static class WindowsVersionService
    {
        /// <summary>
        /// Windows 11 的最低构建版本号
        /// </summary>
        private const int Windows11BuildNumber = 22000;

        /// <summary>
        /// Windows 10 1903 的最低构建版本号（支持暗色模式标题栏）
        /// </summary>
        private const int Windows10_1903_BuildNumber = 18362;

        /// <summary>
        /// Windows 10 22H2 的最低构建版本号（支持 Mica 效果）
        /// </summary>
        private const int Windows10_22H2_BuildNumber = 19045;

        /// <summary>
        /// 缓存的构建版本号
        /// </summary>
        private static readonly int _cachedBuildNumber = GetBuildNumber();

        /// <summary>
        /// 获取当前系统的构建版本号
        /// </summary>
        public static int BuildNumber => _cachedBuildNumber;

        /// <summary>
        /// 判断是否为 Windows 11 或更高版本
        /// </summary>
        public static bool IsWindows11OrLater => _cachedBuildNumber >= Windows11BuildNumber;

        /// <summary>
        /// 判断是否支持暗色模式标题栏
        /// </summary>
        public static bool SupportsDarkModeTitleBar => _cachedBuildNumber >= Windows10_1903_BuildNumber;

        /// <summary>
        /// 判断是否支持 Mica 效果
        /// </summary>
        public static bool SupportsMicaEffect => _cachedBuildNumber >= Windows11BuildNumber;

        /// <summary>
        /// 判断是否支持系统圆角
        /// </summary>
        public static bool SupportsWindowCorners => _cachedBuildNumber >= Windows11BuildNumber;

        /// <summary>
        /// 判断是否支持系统背景效果（Mica/Acrylic）
        /// </summary>
        public static bool SupportsSystemBackdrop => _cachedBuildNumber >= Windows11BuildNumber;

        /// <summary>
        /// 获取操作系统版本信息字符串
        /// </summary>
        /// <returns>版本信息字符串</returns>
        public static string GetVersionInfo()
        {
            string versionName = IsWindows11OrLater ? "Windows 11" : "Windows 10";
            return $"{versionName} (Build {BuildNumber})";
        }

        /// <summary>
        /// 获取系统构建版本号
        /// </summary>
        /// <returns>构建版本号，如果无法获取则返回 0</returns>
        private static int GetBuildNumber()
        {
            try
            {
                return RtlGetVersionBuildNumber();
            }
            catch
            {
                return Environment.OSVersion.Version.Build;
            }
        }

        /// <summary>
        /// 通过 RtlGetVersion 获取真实的构建版本号
        /// </summary>
        /// <returns>构建版本号</returns>
        private static int RtlGetVersionBuildNumber()
        {
            var osvi = new OSVERSIONINFOEX();
            osvi.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFOEX));

            if (RtlGetVersion(ref osvi) == 0)
            {
                return osvi.dwBuildNumber;
            }

            return 0;
        }

        #region Native Methods

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int RtlGetVersion(ref OSVERSIONINFOEX lpVersionInformation);

        [StructLayout(LayoutKind.Sequential)]
        private struct OSVERSIONINFOEX
        {
            public int dwOSVersionInfoSize;
            public int dwMajorVersion;
            public int dwMinorVersion;
            public int dwBuildNumber;
            public int dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;
            public short wServicePackMajor;
            public short wServicePackMinor;
            public short wSuiteMask;
            public byte wProductType;
            public byte wReserved;
        }

        #endregion
    }
}
