// ViewModels/TransferConfirmationViewModel.cs
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Views;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace HansoInputTool.ViewModels
{
    public partial class TransferConfirmationViewModel : ObservableObject
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly ExcelHandler _excelHandler;
        private readonly Dictionary<string, RateInfo> _rates;
        private readonly ColumnMapping _columnMap;
        private readonly Action<bool> _callback;
        private readonly FlagDefinitionService _flagService;

        // 基本情報
        public string Period { get; set; }
        public string Month { get; set; }
        public string RNumber { get; set; }

        // 車両リスト
        public ObservableCollection<VehicleTab> VehicleTabs { get; set; }

        private VehicleTab _selectedVehicle;
        public VehicleTab SelectedVehicle
        {
            get => _selectedVehicle;
            set
            {
                if (SetProperty(ref _selectedVehicle, value))
                {
                    LoadVehicleData();
                }
            }
        }

        // 現在の車両データ
        public ObservableCollection<TransferRowData> CurrentVehicleRows { get; set; }

        private VehicleSummary _currentVehicleSummary;
        public VehicleSummary CurrentVehicleSummary
        {
            get => _currentVehicleSummary;
            set => SetProperty(ref _currentVehicleSummary, value);
        }

        // エラー・警告リスト
        public ObservableCollection<ValidationIssue> ValidationIssues { get; set; }
        private List<ValidationIssue> _allValidationIssues = new List<ValidationIssue>();

        private bool _showErrorsOnly;
        public bool ShowErrorsOnly
        {
            get => _showErrorsOnly;
            set
            {
                if (SetProperty(ref _showErrorsOnly, value))
                {
                    FilterValidationIssues();
                }
            }
        }

        // 統計情報
        public string TotalVehicles { get; set; }
        public string VehiclesWithErrors { get; set; }
        public string VehiclesWithWarnings { get; set; }
        public string TotalEstimatedRevenue { get; set; }

        // コマンド
        public ICommand PreviousVehicleCommand { get; }
        public ICommand NextVehicleCommand { get; }
        public ICommand EditDataCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand ConfirmTransferCommand { get; }
        public ICommand JumpToIssueCommand { get; }

        public TransferConfirmationViewModel(
            ExcelHandler excelHandler,
            Dictionary<string, RateInfo> rates,
            ColumnMapping columnMap,
            string period,
            string month,
            string rNumber,
            Action<bool> callback,
            FlagDefinitionService flagService = null)
        {
            _excelHandler = excelHandler;
            _rates        = rates;
            _columnMap    = columnMap;
            _callback     = callback;
            _flagService  = flagService;

            Period = period;
            Month = month;
            RNumber = rNumber;

            VehicleTabs = new ObservableCollection<VehicleTab>();
            CurrentVehicleRows = new ObservableCollection<TransferRowData>();
            ValidationIssues = new ObservableCollection<ValidationIssue>();

            // コマンド初期化
            PreviousVehicleCommand = new RelayCommand(_ => MoveToPreviousVehicle(), _ => CanMoveToPrevious());
            NextVehicleCommand = new RelayCommand(_ => MoveToNextVehicle(), _ => CanMoveToNext());
            EditDataCommand = new RelayCommand(_ => EditData());
            CancelCommand = new RelayCommand(_ => Cancel());
            ConfirmTransferCommand = new RelayCommand(_ => ConfirmTransfer());
            JumpToIssueCommand = new RelayCommand(param => JumpToIssue(param as ValidationIssue));

            // データ読み込み
            LoadAllVehicles();
        }

    }
}
