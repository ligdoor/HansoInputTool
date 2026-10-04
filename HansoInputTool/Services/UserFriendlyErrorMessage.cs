using System;
using System.IO;

namespace HansoInputTool.Services
{
    /// <summary>ユーザー向けには原因と次の操作を示し、技術的な詳細はログで確認できるようにする。</summary>
    public static class UserFriendlyErrorMessage
    {
        public static string For(Exception exception, string operation)
        {
            var cause = FindCause(exception);
            string detail = cause switch
            {
                UnauthorizedAccessException => "保存先のフォルダを使う権限がない可能性があります。別の保存先を選ぶか、管理者に確認してください。",
                FileNotFoundException => "必要なファイルが見つかりません。ファイルが移動・削除されていないか確認してください。",
                DirectoryNotFoundException => "指定したフォルダが見つかりません。フォルダを選び直してください。",
                IOException => "対象のファイルがExcelなどで開かれているか、保存先を利用できない可能性があります。ファイルを閉じて、もう一度お試しください。",
                InvalidDataException => "ファイルの内容を読み取れません。正しいファイルを選び直してください。",
                _ => $"一時的な問題の可能性があります。もう一度お試しください。解決しない場合は、logsフォルダ内の最新の日付のログファイルを管理者にお渡しください。\n保存場所: {Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs")}"
            };

            return $"{operation}ができませんでした。\n\n{detail}";
        }

        private static Exception FindCause(Exception exception)
        {
            while (exception?.InnerException != null)
                exception = exception.InnerException;
            return exception;
        }
    }
}
