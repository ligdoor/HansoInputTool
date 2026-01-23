using System.Windows;
using System.Windows.Threading;
using NLog;

namespace HansoInputTool
{
    public partial class App : Application
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public App()
        {
            // UIスレッドで発生した、キャッチされなかったすべての例外を処理
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            // 例外をログに記録
            Logger.Error(e.Exception, "予期せぬUIエラーが発生しました。");

            // ユーザーに通知
            MessageBox.Show("予期せぬエラーが発生しました。アプリケーションを終了します。\n詳細はログファイルを確認してください。", "重大なエラー", MessageBoxButton.OK, MessageBoxImage.Error);

            // 例外を処理済みにし、アプリケーションのクラッシュを防ぐ (場合による)
            e.Handled = true;

            // アプリケーションを終了
            Application.Current.Shutdown();
        }
    }
}