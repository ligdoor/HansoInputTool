namespace HansoInputTool.Models
{
    /// <summary>fuel_records テーブルの1行を表すモデル（給油記録）</summary>
    public class FuelRecord
    {
        public long   Id               { get; set; }
        public long   SessionId        { get; set; }
        /// <summary>対象車両のシート名（例: "CH富士吉田 寝台車 29"）</summary>
        public string VehicleSheetName { get; set; }
        /// <summary>給油日（1〜31）</summary>
        public int    Day              { get; set; }
        /// <summary>給油時のメーター(Km)</summary>
        public double OdometerKm       { get; set; }
        /// <summary>給油量(㍑)</summary>
        public double Liters           { get; set; }
        public string CreatedAt        { get; set; }
    }
}
