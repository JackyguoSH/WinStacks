using System;
using System.Windows;

namespace WinStacks.Services
{
    /// <summary>
    /// 异常处理服务类，提供统一的异常处理功能
    /// </summary>
    public static class ExceptionHandler
    {
        /// <summary>
        /// 处理未捕获的异常
        /// </summary>
        /// <param name="exception">异常对象</param>
        /// <param name="source">异常来源</param>
        public static void HandleUnhandledException(Exception exception, string source)
        {
            LogService.Instance.LogException($"未处理的异常 - {source}", exception);
            
            ShowErrorMessage("程序发生意外错误", 
                $"抱歉，程序遇到了一个意外错误。\n\n错误信息: {exception.Message}\n\n日志已保存到: {LogService.Instance.GetLogFilePath()}");
        }

        /// <summary>
        /// 处理 Dispatcher 未捕获的异常
        /// </summary>
        /// <param name="exception">异常对象</param>
        public static void HandleDispatcherException(Exception exception)
        {
            LogService.Instance.LogException("Dispatcher 未处理异常", exception);
            
            // 对于非致命异常，仅记录日志，不中断程序
            if (IsNonFatalException(exception))
            {
                LogService.Instance.LogWarning($"非致命异常已处理: {exception.Message}");
            }
            else
            {
                ShowErrorMessage("程序发生错误", 
                    $"程序遇到错误: {exception.Message}\n\n日志已保存到: {LogService.Instance.GetLogFilePath()}");
            }
        }

        /// <summary>
        /// 安全执行操作，捕获并记录异常
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <param name="context">操作上下文描述</param>
        /// <param name="showErrorToUser">是否向用户显示错误</param>
        /// <returns>操作是否成功执行</returns>
        public static bool SafeExecute(Action action, string context, bool showErrorToUser = false)
        {
            try
            {
                action();
                return true;
            }
            catch (Exception ex)
            {
                LogService.Instance.LogException(context, ex);
                
                if (showErrorToUser)
                {
                    ShowErrorMessage("操作失败", $"{context} 失败: {ex.Message}");
                }
                
                return false;
            }
        }

        /// <summary>
        /// 安全执行操作并返回结果
        /// </summary>
        /// <typeparam name="T">返回值类型</typeparam>
        /// <param name="func">要执行的函数</param>
        /// <param name="context">操作上下文描述</param>
        /// <param name="defaultValue">发生异常时的默认返回值</param>
        /// <returns>函数执行结果或默认值</returns>
        public static T? SafeExecute<T>(Func<T> func, string context, T? defaultValue = default)
        {
            try
            {
                return func();
            }
            catch (Exception ex)
            {
                LogService.Instance.LogException(context, ex);
                return defaultValue;
            }
        }

        /// <summary>
        /// 判断是否为非致命异常
        /// </summary>
        /// <param name="exception">异常对象</param>
        /// <returns>是否为非致命异常</returns>
        private static bool IsNonFatalException(Exception exception)
        {
            return exception is OutOfMemoryException ||
                   exception is System.ComponentModel.Win32Exception ||
                   exception is InvalidOperationException;
        }

        /// <summary>
        /// 显示错误消息对话框
        /// </summary>
        /// <param name="title">对话框标题</param>
        /// <param name="message">错误消息</param>
        private static void ShowErrorMessage(string title, string message)
        {
            try
            {
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
            catch
            {
                // 如果无法显示消息框，记录日志
                LogService.Instance.LogError($"无法显示错误消息: {title} - {message}");
            }
        }
    }
}
