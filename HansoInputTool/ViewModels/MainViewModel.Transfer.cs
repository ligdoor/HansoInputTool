using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;

namespace HansoInputTool.ViewModels
{
    public partial class MainViewModel
    {
        #region 転記

        private async Task StartTransfer()
        {
            if (!int.TryParse(Period, out var period) || !int.TryParse(Month, out var month) || !int.TryParse(RNumber, out var rNum))
            {
                MessageBox.Show("期、月、R番号を正しく入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool shouldContinue = false;
            var confirmVM = new TransferConfirmationViewModel(_excelHandler, Rates, _columnMap, Period, Month, RNumber, r => shouldContinue = r, _flagService);
            new Views.TransferConfirmationWindow(confirmVM) { Owner = Application.Current.MainWindow }.ShowDialog();
            if (!shouldContinue) { Log("転記処理がキャンセルされました。"); return; }

            string outputDir = ShowFolderBrowserDialog("出力先のベースフォルダを選択してください");
            if (outputDir == null) { Log("フォルダ選択がキャンセルされました。"); return; }

            IsBusy = true;
            var progressVM = new ProgressWindowViewModel();
            var progressWindow = new Views.ProgressWindow(progressVM) { Owner = Application.Current.MainWindow };
            var progress = new Progress<TransferProgressReport>(report =>
            {
                if (report.Total > 0)
                    progressVM.UpdateProgress(report.Current, report.Total, report.Message ?? "");
                else if (!string.IsNullOrEmpty(report.Message))
                    progressVM.AppendLog(report.Message);
            });
            progressWindow.Show();

            try
            {
                _excelHandler.Save();
                await new TransferService().ExecuteAsync(
                    InputFilePath, TemplateFilePath, outputDir,
                    period, month, rNum, _allSheetNames, Rates, _columnMap, progress, _flagService, _dbService, EraName,
                    _vehicleSettingsService);

                Log("========\n転記完了\n========");
                Period = Month = RNumber = string.Empty;
                progressVM.Complete("2つのファイルの作成が完了しました。");
                if (_dbService != null)
                {
                    try
                    {
                        _dbService.ClearAllData();
                        _excelHandler.InvalidateCacheAll();
                        EastSheet.ClearRegisteredSheets();
                        UpdatePreview();
                        Log("[DB] 転記完了につきDBデータをクリアしました。");

                        // Input.xlsxにも残存データがあるためクリアして保存
                        foreach (var msg in _excelHandler.ClearData()) Log(msg);
                        _excelHandler.Save();
                        Log("[Excel] Input.xlsxのデータをクリアしました。");
                    }
                    catch (InvalidOperationException)
                    {
                        // 転記自体は完了済み。セッションが確定済みのため自動クリアのみスキップする
                        // （誤操作防止のため意図的な仕様。手動で確定解除すればクリアできる）。
                        Log("[DB] このセッションは確定済みのため、DBデータの自動クリアはスキップされました。");
                    }
                }
                else
                {
                    ClearInputData(false);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "転記処理中にエラーが発生しました。");
                Log($"エラー: {ex.Message}");
                progressVM.ErrorComplete("エラーが発生しました: 詳細はログファイルを確認してください。");
            }
            finally
            {
                IsBusy = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }

        #endregion
    }
}
