namespace HansoInputTool.Models
{
    public class TransportRecord
    {
        public string Date { get; set; }
        public string RNumber { get; set; }
        public string VehicleType { get; set; }
        public string VehicleNumber { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string StartMeter { get; set; }
        public string EndMeter { get; set; }
        public string Distance { get; set; }
        public string TollFee { get; set; }
        public string ParkingFee { get; set; }
        public string OtherFee { get; set; }
        public string NightFee { get; set; }
        public string NightHours { get; set; }
        public string Driver { get; set; }
        public string DeceasedName { get; set; }
        public string Location { get; set; }
        public string Remarks { get; set; }
    }
}