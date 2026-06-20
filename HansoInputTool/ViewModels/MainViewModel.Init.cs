using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using HansoInputTool.Models;
using HansoInputTool.Services;
using Newtonsoft.Json;

namespace HansoInputTool.ViewModels
{
    public partial class MainViewModel
    {
        #region 初期化

        private async Task OnWindowLoaded()
        {
            try
            {
                Logger.Info("アプリケーションの初期化を開始します。");

                if (!Directory.Exists(BaseDataPath))
                {
                    MessageBox.Show("データフォルダが見つかりません。\n実行ファイルと同じ場所に 'data' フォルダを配置してください。",
                        "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    Application.Current.Shutdown();
                    return;
                }

                _backupService.CreateAutoBackup(InputFilePath);
                _backupService.CreateAutoBackup(TemplateFilePath);
                Log("起動時の自動バックアップを作成しました。");

                var ratesJson = await File.ReadAllTextAsync(RatesFilePath);
                Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(ratesJson);

                var columnMapJson = await File.ReadAllTextAsync(ColumnMapFilePath);
                _columnMap = JsonConvert.DeserializeObject<ColumnMapping>(columnMapJson);

                _flagService = new FlagDefinitionService(CustomFlagsFilePath);
                _excelHandler = new ExcelHandler(InputFilePath, TemplateFilePath, _columnMap);
                _shortcutService = new ShortcutService(ShortcutSettingsFilePath);

                // SQLiteサービスを初期化してExcelHandlerに注入
                _dbService = new DatabaseService(DatabaseFilePath);
                _excelHandler.DbService = _dbService;

                // 車両設定サービスを初期化
                _vehicleSettingsService = new VehicleSettingsService(VehicleSettingsFilePath);
                _excelHandler.VehicleSettingsService = _vehicleSettingsService;

                // 起動時フラグ自動同期
                _excelHandler.SyncFlagsOnStartup(_flagService);
                Log("ショートカット設定を読み込みました。");

                // EraName（元号）をappsettings.jsonから読み込む
                EraName = Services.DataSetupService.ReadEraNameFromSettings();

                // 期・R（前回値）をappsettings.jsonから読み込む
                var (lastPeriod, lastRNumber) = Services.DataSetupService.ReadLastPeriodRNumber();
                if (!string.IsNullOrEmpty(lastPeriod))  _period  = lastPeriod;
                if (!string.IsNullOrEmpty(lastRNumber)) _rNumber = lastRNumber;
                OnPropertyChanged(nameof(Period));
                OnPropertyChanged(nameof(RNumber));

                // 月末日チェック用に年・月を渡す（Month は "1"〜"12" の文字列）
                NormalSheet.Initialize(_excelHandler, Log, UpdatePreview, _flagService,
                    _vehicleSettingsService,
                    getYearMonth: () =>
                    {
                        int.TryParse(Month, out var m);
                        return (DateTime.Now.Year, m);
                    });
                _excelHandler.FlagService = _flagService;
                EastSheet.Initialize(_excelHandler, Log);

                ReloadAllData();
                await CheckForUpdate();

                if (_excelHandler.CheckRemainingData())
                {
                    var result = MessageBox.Show("前回のデータが残っています。\n全ての入力データをクリアして新規に開始しますか？",
                        "データクリア確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                        ClearInputData(true);
                }

                Logger.Info("アプリケーションの初期化が完了しました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "アプリケーションの初期化中に致命的なエラーが発生しました。");
                MessageBox.Show($"初期化エラー:\n{ex.GetType().Name}\n{ex.Message}\n\n内部エラー:{ex.InnerException?.Message}",
                    "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        #endregion
    }
}
