using System.Windows;
using HansoInputTool.Services;
using System.Windows.Threading;
using NLog;

namespace HansoInputTool
{
    public partial class App : Application
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        // データパスをApp全体で共有する
        public static string DataPath { get; private set; }

        public App()
        {
            // UIスレッドで発生した、キャッチされなかったすべての例外を処理
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // データフォルダのセットアップ（初回起動時はフォルダ選択ダイアログを表示）
            DataPath = DataSetupService.EnsureDataPath();
            if (DataPath == null)
            {
                // キャンセルされた場合はアプリ終了
                Logger.Warn("データフォルダが設定されなかったためアプリを終了します。");
                Shutdown();
            }
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