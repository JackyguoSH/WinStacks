using System;
using System.IO;
using System.Text;

namespace WinStacks.Services
{
    /// <summary>
    /// 日志服务类，提供统一的日志记录功能
    /// </summary>
    public sealed class LogService : IDisposable
    {
        private static LogService? _instance;
        private static readonly object _lock = new object();
        private readonly string _logFilePath;
        private readonly object _fileLock = new object();
        private bool _disposed;

        /// <summary>
        /// 获取日志服务单例实例
        /// </summary>
        public static LogService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new LogService();
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// 私有构造函数，初始化日志文件路径
        /// </summary>
        private LogService()
        {
            string appDataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WinStacks",
                "Logs");
            
            Directory.CreateDirectory(appDataDir);
            
            _logFilePath = Path.Combine(
                appDataDir,
                $"WinStacks_{DateTime.Now:yyyyMMdd}.log");
        }

        /// <summary>
        /// 记录信息级别日志
        /// </summary>
        /// <param name="message">日志消息</param>
        public void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        /// <summary>
        /// 记录警告级别日志
        /// </summary>
        /// <param name="message">日志消息</param>
        public void LogWarning(string message)
        {
            WriteLog("WARN", message);
        }

        /// <summary>
        /// 记录错误级别日志
        /// </summary>
        /// <param name="message">日志消息</param>
        public void LogError(string message)
        {
            WriteLog("ERROR", message);
        }

        /// <summary>
        /// 记录异常信息
        /// </summary>
        /// <param name="context">异常发生的上下文描述</param>
        /// <param name="exception">异常对象</param>
        public void LogException(string context, Exception exception)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"上下文: {context}");
            sb.AppendLine($"异常类型: {exception.GetType().FullName}");
            sb.AppendLine($"异常消息: {exception.Message}");
            sb.AppendLine($"堆栈跟踪: {exception.StackTrace}");
            
            if (exception.InnerException != null)
            {
                sb.AppendLine($"内部异常: {exception.InnerException.Message}");
            }
            
            WriteLog("EXCEPTION", sb.ToString());
        }

        /// <summary>
        /// 写入日志到文件
        /// </summary>
        /// <param name="level">日志级别</param>
        /// <param name="message">日志消息</param>
        private void WriteLog(string level, string message)
        {
            try
            {
                lock (_fileLock)
                {
                    string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
                    File.AppendAllText(_logFilePath, logEntry, Encoding.UTF8);
                }
            }
            catch
            {
                // 日志写入失败时静默处理，避免影响主程序
            }
        }

        /// <summary>
        /// 获取日志文件路径
        /// </summary>
        /// <returns>日志文件的完整路径</returns>
        public string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// 清理超过指定天数的旧日志文件
        /// </summary>
        /// <param name="daysToKeep">保留天数</param>
        public void CleanOldLogs(int daysToKeep = 30)
        {
            try
            {
                string? directory = Path.GetDirectoryName(_logFilePath);
                if (directory == null) return;

                string[] logFiles = Directory.GetFiles(directory, "WinStacks_*.log");
                DateTime cutoffDate = DateTime.Now.AddDays(-daysToKeep);

                foreach (string file in logFiles)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    if (fileInfo.CreationTime < cutoffDate)
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch
                        {
                            // 删除失败时忽略
                        }
                    }
                }
            }
            catch
            {
                // 清理失败时忽略
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                LogInfo("日志服务已关闭");
            }
        }
    }
}
