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
    /// <summary>
    /// TransferConfirmationViewModel の分割定義：車両切り替え・ステータス表示・転記確定などのコマンド処理。
    /// </summary>
    public partial class TransferConfirmationViewModel
    {
        private SolidColorBrush GetStatusColor(int errors, int warnings)
        {
            if (errors > 0)
                return new SolidColorBrush(Color.FromRgb(239, 68, 68));
            if (warnings > 0)
                return new SolidColorBrush(Color.FromRgb(251, 191, 36));
            return new SolidColorBrush(Color.FromRgb(16, 185, 129));
        }

        private string GetStatusIcon(int errors, int warnings)
        {
            if (errors > 0) return "❌";
            if (warnings > 0) return "⚠️";
            return "✓";
        }

        private void MoveToPreviousVehicle()
        {
            var index = VehicleTabs.IndexOf(SelectedVehicle);
            if (index > 0)
            {
                SelectedVehicle = VehicleTabs[index - 1];
            }
        }

        private void MoveToNextVehicle()
        {
            var index = VehicleTabs.IndexOf(SelectedVehicle);
            if (index < VehicleTabs.Count - 1)
            {
                SelectedVehicle = VehicleTabs[index + 1];
            }
        }

        private bool CanMoveToPrevious()
        {
            return SelectedVehicle != null && VehicleTabs.IndexOf(SelectedVehicle) > 0;
        }

        private bool CanMoveToNext()
        {
            return SelectedVehicle != null && VehicleTabs.IndexOf(SelectedVehicle) < VehicleTabs.Count - 1;
        }

        private void JumpToIssue(ValidationIssue issue)
        {
            if (issue == null) return;

            var vehicle = VehicleTabs.FirstOrDefault(v => v.SheetName == issue.SheetName);
            if (vehicle != null)
            {
                SelectedVehicle = vehicle;
            }
        }

        private void EditData()
        {
            _callback?.Invoke(false);
            Application.Current.Windows.OfType<TransferConfirmationWindow>().FirstOrDefault()?.Close();
        }

        private void Cancel()
        {
            _callback?.Invoke(false);
            Application.Current.Windows.OfType<TransferConfirmationWindow>().FirstOrDefault()?.Close();
        }

        private void ConfirmTransfer()
        {
            var totalErrors = VehicleTabs.Sum(v => v.ErrorCount);
            if (totalErrors > 0)
            {
                var result = MessageBox.Show(
                    $"{totalErrors}件のエラーが検出されています。\nこのまま転記を実行しますか？",
                    "エラー確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            _callback?.Invoke(true);
            Application.Current.Windows.OfType<TransferConfirmationWindow>().FirstOrDefault()?.Close();
        }
    }
}
