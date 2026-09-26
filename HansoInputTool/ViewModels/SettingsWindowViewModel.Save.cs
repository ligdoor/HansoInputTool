using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using Newtonsoft.Json;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// SettingsWindowViewModel の分割定義：設定保存(SaveSettings)。
    /// 複数の保存処理（車両シート/料金/ショートカット/バックアップ/フラグ/元号/列マッピング/深夜方式）を
    /// まとめて実行する一連の処理のため、内部は分割せずそのまま移動。
    /// </summary>
    public partial class SettingsWindowViewModel
    {
        private void SaveSettings(object parameter)
        {
            // 車両シートのバリデーション
            if (VehicleSheetList.Any(v => string.IsNullOrWhiteSpace(v.VehicleTypeName)))
            {
                MessageBox.Show("車両名が空の項目があります。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var duplicate = VehicleSheetList.GroupBy(v => v.VehicleTypeName).FirstOrDefault(g => g.Count() > 1);
            if (duplicate != null)
            {
                MessageBox.Show($"車両名 '{duplicate.Key}' が重複しています。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ショートカットの重複チェック
            if (ShortcutSettingsVM.HasDuplicates(out string duplicateInfo))
            {
                var result = MessageBox.Show(
                    $"{duplicateInfo}\n\nそのまま保存しますか？",
                    "ショートカットの重複",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            try
            {
                // 車両シート設定の保存
                var originalSheetNames = _excelHandler.GetVehicleSheetNames();
                var finalSheetVMs = VehicleSheetList.ToList();

                var finalOriginalNames = finalSheetVMs.Where(vm => vm.OriginalSheetName != null).Select(vm => vm.OriginalSheetName).ToList();
                var sheetsToDelete = originalSheetNames.Except(finalOriginalNames).ToList();

                var renamedVMs = finalSheetVMs.Where(vm => vm.OriginalSheetName != null && vm.OriginalSheetName != vm.VehicleTypeName).ToList();
                var renameMap = renamedVMs.ToDictionary(vm => vm.OriginalSheetName, vm => vm.VehicleTypeName);

                var addedVMs = finalSheetVMs.Where(vm => vm.OriginalSheetName == null).ToList();
                var sheetsToAdd = new List<(string newName, string templateName)>();

                foreach (var vehicleVM in addedVMs)
                {
                    string templateSheetName = vehicleVM.Selected事業所カテゴリ == "東日本セレモニー"
                        ? "Template2"
                        : "Template1";

                    if (!_excelHandler.InputSheetExists(templateSheetName))
                    {
                        MessageBox.Show($"コピー元となるテンプレートシート '{templateSheetName}' が見つかりません。\nInput.xlsxに'{templateSheetName}'という名前のシートを作成してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        _excelHandler.Load();
                        return;
                    }

                    sheetsToAdd.Add((vehicleVM.VehicleTypeName, templateSheetName));
                }

                _excelHandler.SyncAllVehicleSheets(sheetsToDelete, renameMap, sheetsToAdd);
                _excelHandler.Save();

                // 料金設定の保存
                string json = JsonConvert.SerializeObject(Rates, Formatting.Indented);
                File.WriteAllText(_ratesFilePath, json);

                // ショートカット設定の保存
                if (_shortcutService != null)
                {
                    var newShortcutSettings = ShortcutSettingsVM.ToShortcutSettings();
                    _shortcutService.UpdateSettings(newShortcutSettings);
                    _shortcutService.Save();
                }

                // バックアップ保持数を反映
                if (_backupService != null)
                {
                    _backupService.MaxBackupFiles       = MaxAutoBackupFiles;
                    _backupService.MaxManualBackupFiles = MaxManualBackupFiles;
                }

                // フラグ設定の保存
                if (FlagSettingsVM != null)
                {
                    if (!FlagSettingsVM.Validate(out string flagError))
                    {
                        MessageBox.Show(flagError, "フラグ設定エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // 変更前のフラグ一覧をディープコピーで保存（差分検出用）
                    // ※ FlagDefinitionは参照型のため ToList() だけでは不十分。
                    //   ApplyChanges→RebuildColumns で同一オブジェクトのプロパティが
                    //   書き換わり oldFlags と newFlags が同じ内容になるのを防ぐ。
                    var oldFlags = _flagService.Flags
                        .Select(f => new HansoInputTool.Models.FlagDefinition
                        {
                            Id          = f.Id,
                            DisplayName = f.DisplayName,
                            Type        = f.Type,
                            AmountType  = f.AmountType,
                            AmountValue = f.AmountValue,
                            Order       = f.Order,
                            ExcelColumn = f.ExcelColumn
                        })
                        .ToList();

                    FlagSettingsVM.ApplyChanges();

                    // 変更後のフラグ一覧
                    var newFlags = _flagService.Flags.ToList();

                    // Excel列を同期（追加・削除）
                    _excelHandler.SyncFlagColumns(oldFlags, newFlags);
                    _excelHandler.Save();

                    // NormalSheetのチェックボックスを再構築
                    _mainViewModel.NormalSheet.RebuildFlagItems();

                    // フラグショートカットをShortcutServiceに同期
                    _mainViewModel.SyncFlagShortcuts();
                }

                // 元号設定の保存
                var saveEra = string.IsNullOrWhiteSpace(EraName) ? "R" : EraName.Trim();
                Services.DataSetupService.SaveEraNameToSettings(saveEra);
                Services.DataSetupService.SaveEraStartYearToSettings(EraStartYear);
                _mainViewModel.EraName = saveEra;

                // 列マッピング保存
                var cm = new Models.ColumnMapping
                {
                    NormalSheet = new Models.SheetColumnMap
                    {
                        Day           = CmDay,
                        HansoCount    = CmHansoCount,
                        YuryoKm       = CmYuryoKm,
                        MuryoKm       = CmMuryoKm,
                        KihonFee      = CmKihonFee,
                        SokoFee       = CmSokoFee,
                        ShinyaFee     = CmShinyaFee,
                        TotalFee      = CmTotalFee,
                        ShinyaMinutes = CmShinyaMinutes
                    },
                    EastSheet = new Models.CellAddressMap
                    {
                        Jitsudo     = CmEastJitsudo,
                        Hanso       = CmEastHanso,
                        YuryoKm     = CmEastYuryoKm,
                        MuryoKm     = CmEastMuryoKm,
                        UnsoJisseki = CmEastUnsoJisseki
                    },
                    ShukeiSheet = new Models.CellAddressMap
                    {
                        Days      = CmShukeiDays,
                        Hanso     = CmShukeiHanso,
                        YuryoKm   = CmShukeiYuryoKm,
                        MuryoKm   = CmShukeiMuryoKm,
                        Total     = CmShukeiTotal
                    }
                };
                Services.DataSetupService.SaveColumnMap(cm);
                _mainViewModel.ReloadColumnMap(cm);

                // 車両ごとの深夜入力方式を保存
                if (_vehicleSettingsService != null)
                {
                    var vs = new Models.VehicleSettings();
                    foreach (var v in VehicleSheetList)
                        vs[v.VehicleTypeName] = new Models.VehicleConfig { LateInputMode = v.LateInputMode, IsFuelTracked = v.IsFuelTracked };
                    _vehicleSettingsService.Save(vs);
                    _mainViewModel.ReloadVehicleSettings(vs);
                }

                _mainViewModel.UpdateRatesAndReload(Rates);

                MessageBox.Show("設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                if (parameter is Window window)
                {
                    window.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定の保存中にエラーが発生しました。\n{ex.Message}", "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                _excelHandler.Load();
            }
        }
    }
}
