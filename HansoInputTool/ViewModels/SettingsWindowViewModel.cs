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
    public class SettingsWindowViewModel : ObservableObject
    {
        private readonly MainViewModel _mainViewModel;
        private readonly ExcelHandler _excelHandler;
        private readonly VehicleSettingsService _vehicleSettingsService;
        private readonly string _ratesFilePath;
        private readonly ShortcutService _shortcutService;
        private readonly BackupService _backupService;
        private readonly FlagDefinitionService _flagService;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ObservableCollection<VehicleSheetViewModel> VehicleSheetList { get; set; }

        private VehicleSheetViewModel _selectedVehicle;
        public VehicleSheetViewModel SelectedVehicle
        {
            get => _selectedVehicle;
            set
            {
                if (SetProperty(ref _selectedVehicle, value))
                {
                    OnPropertyChanged(nameof(CanMoveUp));
                    OnPropertyChanged(nameof(CanMoveDown));
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        // ショートカット設定
        public ShortcutSettingsViewModel ShortcutSettingsVM { get; }

        // バックアップ設定
        private int _maxAutoBackupFiles;

        // フラグ管理
        public FlagSettingsViewModel FlagSettingsVM { get; }
        public int MaxAutoBackupFiles
        {
            get => _maxAutoBackupFiles;
            set => SetProperty(ref _maxAutoBackupFiles, Math.Max(1, Math.Min(50, value)));
        }

        private int _maxManualBackupFiles;
        public int MaxManualBackupFiles
        {
            get => _maxManualBackupFiles;
            set => SetProperty(ref _maxManualBackupFiles, Math.Max(1, Math.Min(100, value)));
        }

        // 元号設定
        private string _eraName;
        public string EraName
        {
            get => _eraName;
            set => SetProperty(ref _eraName, value);
        }

        private int _eraStartYear;
        public int EraStartYear
        {
            get => _eraStartYear;
            set => SetProperty(ref _eraStartYear, value);
        }

        // 列マッピング（通常シート）
        private int _cmDay;           public int CmDay           { get => _cmDay;           set { SetProperty(ref _cmDay, value);           UpdateColLabel(ref _cmDayLabel,           value); } }
        private int _cmHansoCount;    public int CmHansoCount    { get => _cmHansoCount;    set { SetProperty(ref _cmHansoCount, value);    UpdateColLabel(ref _cmHansoCountLabel,    value); } }
        private int _cmYuryoKm;       public int CmYuryoKm       { get => _cmYuryoKm;       set { SetProperty(ref _cmYuryoKm, value);       UpdateColLabel(ref _cmYuryoKmLabel,       value); } }
        private int _cmMuryoKm;       public int CmMuryoKm       { get => _cmMuryoKm;       set { SetProperty(ref _cmMuryoKm, value);       UpdateColLabel(ref _cmMuryoKmLabel,       value); } }
        private int _cmKihonFee;      public int CmKihonFee      { get => _cmKihonFee;      set { SetProperty(ref _cmKihonFee, value);      UpdateColLabel(ref _cmKihonFeeLabel,      value); } }
        private int _cmSokoFee;       public int CmSokoFee       { get => _cmSokoFee;       set { SetProperty(ref _cmSokoFee, value);       UpdateColLabel(ref _cmSokoFeeLabel,       value); } }
        private int _cmShinyaFee;     public int CmShinyaFee     { get => _cmShinyaFee;     set { SetProperty(ref _cmShinyaFee, value);     UpdateColLabel(ref _cmShinyaFeeLabel,     value); } }
        private int _cmTotalFee;      public int CmTotalFee      { get => _cmTotalFee;      set { SetProperty(ref _cmTotalFee, value);      UpdateColLabel(ref _cmTotalFeeLabel,      value); } }
        private int _cmShinyaMinutes; public int CmShinyaMinutes { get => _cmShinyaMinutes; set { SetProperty(ref _cmShinyaMinutes, value); UpdateColLabel(ref _cmShinyaMinutesLabel, value); } }

        // → 列名ラベル（A,B,C…）
        private string _cmDayLabel;           public string CmDayLabel           { get => _cmDayLabel;           set => SetProperty(ref _cmDayLabel, value); }
        private string _cmHansoCountLabel;    public string CmHansoCountLabel    { get => _cmHansoCountLabel;    set => SetProperty(ref _cmHansoCountLabel, value); }
        private string _cmYuryoKmLabel;       public string CmYuryoKmLabel       { get => _cmYuryoKmLabel;       set => SetProperty(ref _cmYuryoKmLabel, value); }
        private string _cmMuryoKmLabel;       public string CmMuryoKmLabel       { get => _cmMuryoKmLabel;       set => SetProperty(ref _cmMuryoKmLabel, value); }
        private string _cmKihonFeeLabel;      public string CmKihonFeeLabel      { get => _cmKihonFeeLabel;      set => SetProperty(ref _cmKihonFeeLabel, value); }
        private string _cmSokoFeeLabel;       public string CmSokoFeeLabel       { get => _cmSokoFeeLabel;       set => SetProperty(ref _cmSokoFeeLabel, value); }
        private string _cmShinyaFeeLabel;     public string CmShinyaFeeLabel     { get => _cmShinyaFeeLabel;     set => SetProperty(ref _cmShinyaFeeLabel, value); }
        private string _cmTotalFeeLabel;      public string CmTotalFeeLabel      { get => _cmTotalFeeLabel;      set => SetProperty(ref _cmTotalFeeLabel, value); }
        private string _cmShinyaMinutesLabel; public string CmShinyaMinutesLabel { get => _cmShinyaMinutesLabel; set => SetProperty(ref _cmShinyaMinutesLabel, value); }

        // 列マッピング（東日本シート・集計シート）セルアドレス
        private string _cmEastJitsudo;     public string CmEastJitsudo     { get => _cmEastJitsudo;     set => SetProperty(ref _cmEastJitsudo, value); }
        private string _cmEastHanso;       public string CmEastHanso       { get => _cmEastHanso;       set => SetProperty(ref _cmEastHanso, value); }
        private string _cmEastYuryoKm;     public string CmEastYuryoKm     { get => _cmEastYuryoKm;     set => SetProperty(ref _cmEastYuryoKm, value); }
        private string _cmEastMuryoKm;     public string CmEastMuryoKm     { get => _cmEastMuryoKm;     set => SetProperty(ref _cmEastMuryoKm, value); }
        private string _cmEastUnsoJisseki; public string CmEastUnsoJisseki { get => _cmEastUnsoJisseki; set => SetProperty(ref _cmEastUnsoJisseki, value); }
        private string _cmShukeiDays;      public string CmShukeiDays      { get => _cmShukeiDays;      set => SetProperty(ref _cmShukeiDays, value); }
        private string _cmShukeiHanso;     public string CmShukeiHanso     { get => _cmShukeiHanso;     set => SetProperty(ref _cmShukeiHanso, value); }
        private string _cmShukeiYuryoKm;   public string CmShukeiYuryoKm   { get => _cmShukeiYuryoKm;   set => SetProperty(ref _cmShukeiYuryoKm, value); }
        private string _cmShukeiMuryoKm;   public string CmShukeiMuryoKm   { get => _cmShukeiMuryoKm;   set => SetProperty(ref _cmShukeiMuryoKm, value); }
        private string _cmShukeiTotal;     public string CmShukeiTotal     { get => _cmShukeiTotal;     set => SetProperty(ref _cmShukeiTotal, value); }

        private static string ColNumToLetter(int col)
        {
            if (col < 1) return "?";
            string result = "";
            while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }
            return result;
        }
        private void UpdateColLabel(ref string field, int col)
        {
            field = $"→ {ColNumToLetter(col)}列";
            OnPropertyChanged(nameof(CmDayLabel));
            OnPropertyChanged(nameof(CmHansoCountLabel));
            OnPropertyChanged(nameof(CmYuryoKmLabel));
            OnPropertyChanged(nameof(CmMuryoKmLabel));
            OnPropertyChanged(nameof(CmKihonFeeLabel));
            OnPropertyChanged(nameof(CmSokoFeeLabel));
            OnPropertyChanged(nameof(CmShinyaFeeLabel));
            OnPropertyChanged(nameof(CmTotalFeeLabel));
            OnPropertyChanged(nameof(CmShinyaMinutesLabel));
        }

        // 選択中のタブインデックス
        private int _selectedTabIndex;
        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public ICommand AddVehicleCommand    { get; }
        public ICommand DeleteVehicleCommand { get; }
        public ICommand MoveUpCommand        { get; }
        public ICommand MoveDownCommand      { get; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand ResetShortcutsCommand { get; }

        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel,
            ShortcutService shortcutService,
            BackupService backupService = null,
            FlagDefinitionService flagService = null,
            VehicleSettingsService vehicleSettingsService = null)
        {
            _excelHandler           = excelHandler;
            _ratesFilePath          = ratesFilePath;
            _mainViewModel          = mainViewModel;
            _shortcutService        = shortcutService;
            _backupService          = backupService;
            _flagService            = flagService;
            _vehicleSettingsService = vehicleSettingsService;
            Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(JsonConvert.SerializeObject(currentRates));
            var currentSheets = _excelHandler.GetVehicleSheetNames();
            VehicleSheetList = new ObservableCollection<VehicleSheetViewModel>(
                currentSheets.Select(s =>
                {
                    var vm = new VehicleSheetViewModel(s);
                    if (_vehicleSettingsService != null)
                    {
                        vm.LateInputMode = _vehicleSettingsService.IsFeeMode(s) ? "fee" : "time";
                        vm.IsFuelTracked = _vehicleSettingsService.IsFuelTracked(s);
                    }
                    return vm;
                }));

            // ショートカット設定VMを初期化
            ShortcutSettingsVM = new ShortcutSettingsViewModel(_shortcutService.CurrentSettings);

            // フラグ管理VMを初期化
            FlagSettingsVM = flagService != null ? new FlagSettingsViewModel(flagService) : null;

            // バックアップ設定の初期値を読み込み
            MaxAutoBackupFiles   = _backupService?.MaxBackupFiles       ?? 10;
            MaxManualBackupFiles = _backupService?.MaxManualBackupFiles ?? 20;
            EraName = Services.DataSetupService.ReadEraNameFromSettings();
            EraStartYear = Services.DataSetupService.ReadEraStartYearFromSettings();

            // 列マッピング読み込み
            var cm = Services.DataSetupService.ReadColumnMap();
            CmDay           = cm.NormalSheet.Day;
            CmHansoCount    = cm.NormalSheet.HansoCount;
            CmYuryoKm       = cm.NormalSheet.YuryoKm;
            CmMuryoKm       = cm.NormalSheet.MuryoKm;
            CmKihonFee      = cm.NormalSheet.KihonFee;
            CmSokoFee       = cm.NormalSheet.SokoFee;
            CmShinyaFee     = cm.NormalSheet.ShinyaFee;
            CmTotalFee      = cm.NormalSheet.TotalFee;
            CmShinyaMinutes = cm.NormalSheet.ShinyaMinutes;
            CmEastJitsudo     = cm.EastSheet.Jitsudo;
            CmEastHanso       = cm.EastSheet.Hanso;
            CmEastYuryoKm     = cm.EastSheet.YuryoKm;
            CmEastMuryoKm     = cm.EastSheet.MuryoKm;
            CmEastUnsoJisseki = cm.EastSheet.UnsoJisseki;
            CmShukeiDays      = cm.ShukeiSheet.Days;
            CmShukeiHanso     = cm.ShukeiSheet.Hanso;
            CmShukeiYuryoKm   = cm.ShukeiSheet.YuryoKm;
            CmShukeiMuryoKm   = cm.ShukeiSheet.MuryoKm;
            CmShukeiTotal     = cm.ShukeiSheet.Total;

            AddVehicleCommand    = new RelayCommand(p => AddVehicle());
            DeleteVehicleCommand = new RelayCommand(p => DeleteVehicle(), p => SelectedVehicle != null);
            MoveUpCommand        = new RelayCommand(_ => MoveVehicle(-1), _ => CanMoveUp);
            MoveDownCommand      = new RelayCommand(_ => MoveVehicle(1),  _ => CanMoveDown);
            SaveCommand          = new RelayCommand(p => SaveSettings(p));
            CancelCommand        = new RelayCommand(p => ((Window)p).Close());
            ResetShortcutsCommand = new RelayCommand(p => ResetShortcuts());
        }

        // 旧コンストラクタ（後方互換性のため）
        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel)
            : this(currentRates, excelHandler, ratesFilePath, mainViewModel, null, null, null)
        {
        }

        public bool CanMoveUp   => SelectedVehicle != null && VehicleSheetList.IndexOf(SelectedVehicle) > 0;
        public bool CanMoveDown => SelectedVehicle != null && VehicleSheetList.IndexOf(SelectedVehicle) < VehicleSheetList.Count - 1;

        private void MoveVehicle(int direction)
        {
            if (SelectedVehicle == null) return;
            int idx    = VehicleSheetList.IndexOf(SelectedVehicle);
            int newIdx = idx + direction;
            if (newIdx < 0 || newIdx >= VehicleSheetList.Count) return;

            var moving = SelectedVehicle;
            VehicleSheetList.Move(idx, newIdx);

            // 既存シートの場合はExcelのシート順も即時同期
            var neighbor = VehicleSheetList[direction > 0 ? newIdx - 1 : newIdx + 1];
            if (moving.OriginalSheetName != null && neighbor.OriginalSheetName != null)
                _excelHandler.MoveVehicleSheet(moving.OriginalSheetName, neighbor.OriginalSheetName, direction < 0);

            OnPropertyChanged(nameof(CanMoveUp));
            OnPropertyChanged(nameof(CanMoveDown));
            CommandManager.InvalidateRequerySuggested();
        }

        private void AddVehicle()
        {
            var newVehicle = new VehicleSheetViewModel();
            VehicleSheetList.Add(newVehicle);
            SelectedVehicle = newVehicle;
        }

        private void DeleteVehicle()
        {
            if (SelectedVehicle == null) return;
            var sheetName = SelectedVehicle.OriginalSheetName ?? "新しい車両";
            var result = MessageBox.Show($"車両 '{sheetName}' をリストから削除しますか？\n（実際のファイルからの削除は「保存」ボタンを押した時に実行されます）", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                VehicleSheetList.Remove(SelectedVehicle);
                SelectedVehicle = null;
            }
        }

        private void ResetShortcuts()
        {
            var result = MessageBox.Show(
                "ショートカット設定をデフォルトに戻しますか？",
                "リセット確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                ShortcutSettingsVM.ResetToDefaultsCommand.Execute(null);
            }
        }

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
