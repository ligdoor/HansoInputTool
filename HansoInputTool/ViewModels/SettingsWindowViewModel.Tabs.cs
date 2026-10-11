using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// SettingsWindowViewModel の分割定義：入力タブの管理（追加・名前変更・入力方式・並び替え・削除）と、
    /// 車両の所属タブの扱い。
    /// </summary>
    public partial class SettingsWindowViewModel
    {
        public ObservableCollection<TabEditItem> TabList { get; private set; }

        private TabEditItem _selectedTabItem;
        public TabEditItem SelectedTabItem
        {
            get => _selectedTabItem;
            set
            {
                if (!SetProperty(ref _selectedTabItem, value)) return;
                OnPropertyChanged(nameof(CanMoveTabUp));
                OnPropertyChanged(nameof(CanMoveTabDown));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool CanMoveTabUp   => SelectedTabItem != null && TabList.IndexOf(SelectedTabItem) > 0;
        public bool CanMoveTabDown => SelectedTabItem != null && TabList.IndexOf(SelectedTabItem) < TabList.Count - 1;

        public ICommand AddTabCommand      { get; private set; }
        public ICommand DeleteTabCommand   { get; private set; }
        public ICommand MoveTabUpCommand   { get; private set; }
        public ICommand MoveTabDownCommand { get; private set; }

        private bool _revertingTab;

        /// <summary>コンストラクタの最後（VehicleSheetList の作成後）に呼ぶ</summary>
        private void InitializeTabs()
        {
            var tabs = _mainViewModel?.TabSettingsService?.Tabs ?? TabSettingsService.CreateDefaults();
            TabList = new ObservableCollection<TabEditItem>(tabs.Select(t => new TabEditItem(t.Id, t.Name, t.InputMode)));

            foreach (var v in VehicleSheetList) v.PropertyChanged += OnVehiclePropertyChanged;
            VehicleSheetList.CollectionChanged += OnVehicleListChanged;
            RefreshTabVehicleCounts();

            AddTabCommand      = new RelayCommand(_ => AddTab());
            DeleteTabCommand   = new RelayCommand(_ => DeleteTab(), _ => SelectedTabItem != null);
            MoveTabUpCommand   = new RelayCommand(_ => MoveTab(-1), _ => CanMoveTabUp);
            MoveTabDownCommand = new RelayCommand(_ => MoveTab(1),  _ => CanMoveTabDown);

            SelectedTabItem = TabList.FirstOrDefault();
        }

        private void OnVehicleListChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
                foreach (VehicleSheetViewModel v in e.OldItems) v.PropertyChanged -= OnVehiclePropertyChanged;
            if (e.NewItems != null)
                foreach (VehicleSheetViewModel v in e.NewItems) v.PropertyChanged += OnVehiclePropertyChanged;
            RefreshTabVehicleCounts();
        }

        private void OnVehiclePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(VehicleSheetViewModel.TabId) || _revertingTab) return;
            var vehicle = (VehicleSheetViewModel)sender;
            var tab = TabList.FirstOrDefault(t => t.Id == vehicle.TabId);
            if (tab == null) return;

            // 既存の車両は、入力方式の違うタブへは移せない（Excelシートのレイアウトが違うため）
            if (vehicle.OriginalSheetName != null)
            {
                var originalTab  = _mainViewModel?.SheetTabResolver?.GetTab(vehicle.OriginalSheetName);
                var originalMode = _mainViewModel?.SheetTabResolver?.GetInputMode(vehicle.OriginalSheetName) ?? TabInputMode.Daily;
                if (originalMode != tab.InputMode)
                {
                    MessageBox.Show(
                        $"登録済みの車両は、入力方式の違うタブへは移せません。\n（'{vehicle.VehicleTypeName}' は {TabInputMode.DisplayName(originalMode)} のタブにのみ移せます）\n\n別の入力方式にしたい場合は、新しい車両として追加し直してください。",
                        "移動できません", MessageBoxButton.OK, MessageBoxImage.Warning);

                    var revertTo = originalTab?.Id;
                    Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _revertingTab = true;
                        try { vehicle.TabId = revertTo; }
                        finally { _revertingTab = false; }
                        RefreshTabVehicleCounts();
                    }));
                    return;
                }
            }

            vehicle.IsTotalModeTab = tab.IsTotalMode;
            RefreshTabVehicleCounts();
        }

        private void RefreshTabVehicleCounts()
        {
            if (TabList == null) return;
            foreach (var tab in TabList)
                tab.VehicleCount = VehicleSheetList.Count(v => v.TabId == tab.Id);
        }

        private void AddTab()
        {
            const string baseName = "新しいタブ";
            string name = baseName;
            int n = 2;
            while (TabList.Any(t => t.Name == name)) name = $"{baseName} {n++}";

            var item = new TabEditItem("tab-" + Guid.NewGuid().ToString("N").Substring(0, 8), name, TabInputMode.Daily);
            TabList.Add(item);
            RefreshTabVehicleCounts();
            SelectedTabItem = item;
        }

        private void DeleteTab()
        {
            var tab = SelectedTabItem;
            if (tab == null) return;

            if (tab.VehicleCount > 0)
            {
                MessageBox.Show($"タブ '{tab.Name}' には {tab.VehicleCount} 台の車両が所属しています。\n先に車両を別のタブへ移すか、車両を削除してください。",
                    "削除できません", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (TabList.Count <= 1)
            {
                MessageBox.Show("入力タブは最低1つ必要です。", "削除できません", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show($"タブ '{tab.Name}' を削除しますか？\n（［保存］を押した時に反映されます）", "削除確認",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes) return;

            int index = TabList.IndexOf(tab);
            TabList.Remove(tab);
            SelectedTabItem = TabList[Math.Min(index, TabList.Count - 1)];
        }

        private void MoveTab(int direction)
        {
            if (SelectedTabItem == null) return;
            int idx = TabList.IndexOf(SelectedTabItem);
            int newIdx = idx + direction;
            if (newIdx < 0 || newIdx >= TabList.Count) return;

            TabList.Move(idx, newIdx);
            OnPropertyChanged(nameof(CanMoveTabUp));
            OnPropertyChanged(nameof(CanMoveTabDown));
            CommandManager.InvalidateRequerySuggested();
        }

        /// <summary>保存前のチェック。問題があれば error にメッセージを入れて false を返す</summary>
        private bool ValidateTabs(out string error)
        {
            error = null;
            if (TabList == null || TabList.Count == 0) { error = "入力タブが1つもありません。"; return false; }
            if (TabList.Any(t => string.IsNullOrWhiteSpace(t.Name))) { error = "名前が空の入力タブがあります。"; return false; }

            var dup = TabList.GroupBy(t => t.Name.Trim()).FirstOrDefault(g => g.Count() > 1);
            if (dup != null) { error = $"入力タブ名 '{dup.Key}' が重複しています。"; return false; }

            foreach (var v in VehicleSheetList)
            {
                var tab = TabList.FirstOrDefault(t => t.Id == v.TabId);
                if (tab == null) { error = $"車両 '{v.VehicleTypeName}' の所属タブが選択されていません。"; return false; }
                if (v.Selected事業所カテゴリ == "東日本セレモニー" && !tab.IsTotalMode)
                {
                    error = $"車両 '{v.VehicleTypeName}' は営業所が「東日本セレモニー」のため、月合計入力のタブに所属させてください。";
                    return false;
                }
            }
            return true;
        }

        /// <summary>月合計入力のタブに所属する車両（保存後の名前）の一覧。Excelシート作成・名前変更時の判定に使う</summary>
        private System.Collections.Generic.HashSet<string> BuildTotalModeSheetNames()
        {
            var totalTabIds = TabList.Where(t => t.IsTotalMode).Select(t => t.Id).ToHashSet();
            return VehicleSheetList.Where(v => v.TabId != null && totalTabIds.Contains(v.TabId))
                                   .Select(v => v.VehicleTypeName)
                                   .ToHashSet();
        }

        /// <summary>入力タブの定義と、車両ごとの設定（深夜入力方式・給油管理・所属タブ）を保存し、メイン画面へ反映する</summary>
        private void SaveTabsAndVehicleSettings()
        {
            var tabService = _mainViewModel?.TabSettingsService;
            if (tabService != null)
            {
                tabService.Save(TabList.Select(t => new TabDefinition
                {
                    Id = t.Id,
                    Name = t.Name.Trim(),
                    InputMode = t.InputMode
                }));
            }

            if (_vehicleSettingsService != null)
            {
                var vs = new VehicleSettings();
                foreach (var v in VehicleSheetList)
                    vs[v.VehicleTypeName] = new VehicleConfig
                    {
                        LateInputMode = v.LateInputMode,
                        IsFuelTracked = v.IsFuelTracked,
                        TabId = v.TabId
                    };
                _vehicleSettingsService.Save(vs);
                _mainViewModel.ReloadVehicleSettings(vs);
            }

            _mainViewModel?.ReloadInputTabs();
        }
    }
}
