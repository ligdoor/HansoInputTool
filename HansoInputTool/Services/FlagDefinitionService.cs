using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HansoInputTool.Models;
using Newtonsoft.Json;
using NLog;

namespace HansoInputTool.Services
{
    /// <summary>
    /// custom_flags.json の読み書きと、フラグ定義リストの管理を担当するサービス
    /// </summary>
    public class FlagDefinitionService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        // ExcelのフラグはIsKoryo(列12)から始まる
        private const int FirstFlagColumn = 12;

        private readonly string _filePath;
        private List<FlagDefinition> _flags;

        public IReadOnlyList<FlagDefinition> Flags => _flags.AsReadOnly();

        public FlagDefinitionService(string filePath)
        {
            _filePath = filePath;
            Load();
        }

        /// <summary>
        /// JSONを読み込む。ファイルがなければデフォルト（行旅・エンバーミング）を生成して保存する。
        /// </summary>
        public void Load()
        {
            if (File.Exists(_filePath))
            {
                try
                {
                    var json = File.ReadAllText(_filePath);
                    _flags = JsonConvert.DeserializeObject<List<FlagDefinition>>(json)
                             ?? CreateDefaults();
                    RebuildColumns();
                    Logger.Info($"custom_flags.json を読み込みました（{_flags.Count}件）");
                    return;
                }
                catch (Exception ex)
                {
                    Logger.Warn(ex, "custom_flags.json の読み込みに失敗しました。デフォルトを使用します。");
                }
            }

            _flags = CreateDefaults();
            Save();
        }

        /// <summary>
        /// 現在のフラグリストをJSONに保存する
        /// </summary>
        public void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(_filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var json = JsonConvert.SerializeObject(_flags, Formatting.Indented);
                File.WriteAllText(_filePath, json);
                Logger.Info("custom_flags.json を保存しました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "custom_flags.json の保存に失敗しました。");
                throw;
            }
        }

        /// <summary>
        /// IDを保持したままフラグを追加する（設定保存時の再構築用）
        /// </summary>
        public void RestoreWithId(FlagDefinition flag)
        {
            _flags.Add(flag);
            RebuildColumns();
        }

        /// <summary>
        /// フラグを追加する（列番号は自動割り当て）
        /// </summary>
        public void Add(FlagDefinition flag)
        {
            flag.Id    = Guid.NewGuid().ToString("N")[..8];
            flag.Order = _flags.Count + 1;
            _flags.Add(flag);
            RebuildColumns();
        }

        /// <summary>
        /// フラグを更新する（Id以外を上書き）
        /// </summary>
        public void Update(FlagDefinition updated)
        {
            var existing = _flags.FirstOrDefault(f => f.Id == updated.Id);
            if (existing == null) return;
            existing.DisplayName = updated.DisplayName;
            existing.Type        = updated.Type;
            existing.AmountType  = updated.AmountType;
            existing.AmountValue = updated.AmountValue;
            // Order・ExcelColumnはRebuildColumnsで再計算
        }

        /// <summary>
        /// フラグを削除する
        /// </summary>
        public void Remove(string id)
        {
            var flag = _flags.FirstOrDefault(f => f.Id == id);
            if (flag != null) _flags.Remove(flag);
            RebuildColumns();
        }

        /// <summary>
        /// 順番を変更する（indexFrom → indexTo に移動）
        /// </summary>
        public void Reorder(int fromIndex, int toIndex)
        {
            if (fromIndex < 0 || fromIndex >= _flags.Count) return;
            if (toIndex   < 0 || toIndex   >= _flags.Count) return;
            var item = _flags[fromIndex];
            _flags.RemoveAt(fromIndex);
            _flags.Insert(toIndex, item);
            RebuildColumns();
        }

        /// <summary>
        /// Orderと ExcelColumn をリスト順に振り直す
        /// </summary>
        private void RebuildColumns()
        {
            for (int i = 0; i < _flags.Count; i++)
            {
                _flags[i].Order       = i + 1;
                _flags[i].ExcelColumn = FirstFlagColumn + i;
            }
        }

        /// <summary>
        /// デフォルトフラグ（行旅死亡人・エンバーミング）を生成する
        /// </summary>
        private static List<FlagDefinition> CreateDefaults()
        {
            return new List<FlagDefinition>
            {
                new FlagDefinition
                {
                    Id          = "koryo",
                    DisplayName = "行旅死亡人",
                    Type        = FlagType.WithAmount,
                    AmountType  = Models.AmountType.Rate,
                    AmountValue = 0.5,
                    Order       = 1,
                    ExcelColumn = 12
                },
                new FlagDefinition
                {
                    Id          = "embalming",
                    DisplayName = "エンバーミング",
                    Type        = FlagType.CountOnly,
                    AmountType  = null,
                    AmountValue = null,
                    Order       = 2,
                    ExcelColumn = 13
                }
            };
        }
    }
}
