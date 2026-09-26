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
    /// SettingsWindowViewModel の分割定義：Excel列マッピング設定（通常/東日本/集計シート）まわり。
    /// </summary>
    public partial class SettingsWindowViewModel
    {
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
    }
}
