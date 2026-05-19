using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 設定画面の「フラグ管理」タブ用ViewModel
    /// </summary>
    public class FlagSettingsViewModel : ObservableObject
    {
        private readonly FlagDefinitionService _flagService;

        public ObservableCollection<FlagEditItem> FlagItems { get; } = new();

        private FlagEditItem _selectedItem;
        public FlagEditItem SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (SetProperty(ref _selectedItem, value))
                    CommandManager.InvalidateRequerySuggested();
            }
        }

        public ICommand AddCommand    { get; }
        public ICommand DeleteCommand { get; }
        public ICommand MoveUpCommand { get; }
        public ICommand MoveDownCommand { get; }

        public FlagSettingsViewModel(FlagDefinitionService flagService)
        {
            _flagService = flagService;

            // 現在のフラグをリストに展開
            foreach (var flag in flagService.Flags.OrderBy(f => f.Order))
                FlagItems.Add(new FlagEditItem(flag));

            AddCommand     = new RelayCommand(_ => AddFlag());
            DeleteCommand  = new RelayCommand(_ => DeleteFlag(),  _ => SelectedItem != null);
            MoveUpCommand  = new RelayCommand(_ => MoveUp(),      _ => CanMoveUp());
            MoveDownCommand = new RelayCommand(_ => MoveDown(),   _ => CanMoveDown());
        }

        private void AddFlag()
        {
            var newItem = new FlagEditItem
            {
                DisplayName = "新しいフラグ",
                Type        = FlagType.CountOnly,
                AmountType  = Models.AmountType.Rate,
                AmountValue = null
            };
            FlagItems.Add(newItem);
            SelectedItem = newItem;
        }

        private void DeleteFlag()
        {
            if (SelectedItem == null) return;
            var result = MessageBox.Show(
                $"「{SelectedItem.DisplayName}」を削除しますか？\n※ Excelに既に記録されたデータは消えません",
                "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;
            FlagItems.Remove(SelectedItem);
            SelectedItem = FlagItems.FirstOrDefault();
        }

        private bool CanMoveUp()   => SelectedItem != null && FlagItems.IndexOf(SelectedItem) > 0;
        private bool CanMoveDown() => SelectedItem != null && FlagItems.IndexOf(SelectedItem) < FlagItems.Count - 1;

        private void MoveUp()
        {
            int i = FlagItems.IndexOf(SelectedItem);
            if (i <= 0) return;
            FlagItems.Move(i, i - 1);
            CommandManager.InvalidateRequerySuggested();
        }

        private void MoveDown()
        {
            int i = FlagItems.IndexOf(SelectedItem);
            if (i < 0 || i >= FlagItems.Count - 1) return;
            FlagItems.Move(i, i + 1);
            CommandManager.InvalidateRequerySuggested();
        }

        /// <summary>
        /// 保存前バリデーション。エラーがあれば内容を返す。
        /// </summary>
        public bool Validate(out string errorMessage)
        {
            foreach (var item in FlagItems)
            {
                if (string.IsNullOrWhiteSpace(item.DisplayName))
                {
                    errorMessage = "フラグの表示名が空の項目があります。";
                    return false;
                }
                if (item.Type == FlagType.WithAmount)
                {
                    if (item.AmountType == null)
                    {
                        errorMessage = $"「{item.DisplayName}」の金額タイプが選択されていません。";
                        return false;
                    }
                    if (item.AmountValue == null || item.AmountValueText == null)
                    {
                        errorMessage = $"「{item.DisplayName}」の金額/倍率が入力されていません。";
                        return false;
                    }
                }
            }
            errorMessage = null;
            return true;
        }

        /// <summary>
        /// 編集内容を FlagDefinitionService に保存する
        /// </summary>
        public void ApplyChanges()
        {
            // 既存フラグを全削除して再構築
            var existingIds = _flagService.Flags.Select(f => f.Id).ToList();
            foreach (var id in existingIds)
                _flagService.Remove(id);

            // 順番通りに追加（既存IDは保持、新規はAdd内でGUID生成）
            for (int i = 0; i < FlagItems.Count; i++)
            {
                var item = FlagItems[i];
                var def  = new FlagDefinition
                {
                    Id          = item.OriginalId ?? string.Empty, // Add内で上書きされる（新規の場合）
                    DisplayName = item.DisplayName,
                    Type        = item.Type,
                    AmountType  = item.Type == FlagType.WithAmount ? item.AmountType : null,
                    AmountValue = item.Type == FlagType.WithAmount ? item.AmountValue : null,
                    TargetFee   = item.Type == FlagType.WithAmount ? item.TargetFee : Models.TargetFee.BaseFee,
                    Order       = i + 1
                };

                if (string.IsNullOrEmpty(item.OriginalId))
                {
                    // 新規追加
                    _flagService.Add(def);
                }
                else
                {
                    // 既存を復元（IDを維持するため手動でリストに追加する処理）
                    _flagService.RestoreWithId(def);
                }
            }
            _flagService.Save();
        }
    }

    /// <summary>
    /// フラグ設定画面の1行分のVM（編集用）
    /// </summary>
    public class FlagEditItem : ObservableObject
    {
        /// <summary>既存フラグのID（新規は null）</summary>
        public string OriginalId { get; }

        private string _displayName;
        public string DisplayName
        {
            get => _displayName;
            set => SetProperty(ref _displayName, value);
        }

        private FlagType _type;
        public FlagType Type
        {
            get => _type;
            set
            {
                if (SetProperty(ref _type, value))
                    OnPropertyChanged(nameof(IsWithAmount));
            }
        }

        public bool IsWithAmount => Type == FlagType.WithAmount;

        private AmountType? _amountType;
        public AmountType? AmountType
        {
            get => _amountType;
            set
            {
                if (SetProperty(ref _amountType, value))
                    OnPropertyChanged(nameof(IsRate));
            }
        }

        public bool IsRate => AmountType == Models.AmountType.Rate;

        private double? _amountValue;
        public double? AmountValue
        {
            get => _amountValue;
            set => SetProperty(ref _amountValue, value);
        }

        // TextBoxバインディング用（文字列←→double?）
        private string _amountValueText;
        public string AmountValueText
        {
            get => _amountValueText;
            set
            {
                if (SetProperty(ref _amountValueText, value))
                    AmountValue = double.TryParse(value, out double d) ? d : null;
            }
        }

        // RadioButton バインディング用
        public bool IsTypeCountOnly
        {
            get => Type == FlagType.CountOnly;
            set { if (value) Type = FlagType.CountOnly; }
        }
        public bool IsTypeWithAmount
        {
            get => Type == FlagType.WithAmount;
            set { if (value) Type = FlagType.WithAmount; }
        }
        public bool IsAmountTypeRate
        {
            get => AmountType == Models.AmountType.Rate;
            set { if (value) AmountType = Models.AmountType.Rate; }
        }
        public bool IsAmountTypeFixed
        {
            get => AmountType == Models.AmountType.Fixed;
            set { if (value) AmountType = Models.AmountType.Fixed; }
        }

        // 適用対象料金 RadioButton バインディング用
        private Models.TargetFee _targetFee = Models.TargetFee.BaseFee;
        public Models.TargetFee TargetFee
        {
            get => _targetFee;
            set => SetProperty(ref _targetFee, value);
        }
        public bool IsTargetBaseFee
        {
            get => TargetFee == Models.TargetFee.BaseFee;
            set { if (value) TargetFee = Models.TargetFee.BaseFee; }
        }
        public bool IsTargetMileageFee
        {
            get => TargetFee == Models.TargetFee.MileageFee;
            set { if (value) TargetFee = Models.TargetFee.MileageFee; }
        }
        public bool IsTargetBoth
        {
            get => TargetFee == Models.TargetFee.Both;
            set { if (value) TargetFee = Models.TargetFee.Both; }
        }

        // 新規作成用
        public FlagEditItem()
        {
            OriginalId  = null;
            _amountType = Models.AmountType.Rate;
        }

        // 既存フラグから生成
        public FlagEditItem(FlagDefinition def)
        {
            OriginalId       = def.Id;
            _displayName     = def.DisplayName;
            _type            = def.Type;
            _amountType      = def.AmountType;
            _amountValue     = def.AmountValue;
            _amountValueText = def.AmountValue?.ToString() ?? string.Empty;
            _targetFee       = def.TargetFee;
        }
    }
}
