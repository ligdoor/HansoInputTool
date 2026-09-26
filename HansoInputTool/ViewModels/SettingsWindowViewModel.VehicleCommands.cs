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
    /// SettingsWindowViewModel の分割定義：車両リストの並び替え・追加・削除・ショートカットリセット。
    /// </summary>
    public partial class SettingsWindowViewModel
    {
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

    }
}
