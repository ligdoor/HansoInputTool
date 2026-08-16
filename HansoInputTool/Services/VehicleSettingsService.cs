using System;
using System.IO;
using HansoInputTool.Models;
using Newtonsoft.Json;
using NLog;

namespace HansoInputTool.Services
{
    public class VehicleSettingsService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly string _filePath;
        private VehicleSettings _settings;

        public VehicleSettingsService(string filePath)
        {
            _filePath = filePath;
            _settings = Load();
        }

        public VehicleSettings Settings => _settings;

        public bool IsFeeMode(string sheetName) => _settings.IsFeeMode(sheetName);

        /// <summary>指定シートが給油管理の対象かどうかを返す</summary>
        public bool IsFuelTracked(string sheetName) => _settings.IsFuelTracked(sheetName);

        private VehicleSettings Load()
        {
            try
            {
                if (File.Exists(_filePath))
                {
                    var json = File.ReadAllText(_filePath);
                    var result = JsonConvert.DeserializeObject<VehicleSettings>(json);
                    if (result != null)
                    {
                        Logger.Info($"vehicle_settings.json を読み込みました ({result.Count}件)");
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "vehicle_settings.json の読み込みに失敗しました");
            }
            return new VehicleSettings();
        }

        public void Save(VehicleSettings settings)
        {
            try
            {
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(_filePath, json);
                _settings = settings;
                Logger.Info($"vehicle_settings.json を保存しました ({settings.Count}件)");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "vehicle_settings.json の保存に失敗しました");
            }
        }

        public void Reload()
        {
            _settings = Load();
        }
    }
}
