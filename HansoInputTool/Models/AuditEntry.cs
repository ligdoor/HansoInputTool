namespace HansoInputTool.Models
{
    /// <summary>One immutable history event describing a change to monthly input data.</summary>
    public class AuditEntry
    {
        public long Id { get; set; }
        public string EventAt { get; set; }
        public string UserName { get; set; }
        public string Action { get; set; }
        public string EntityType { get; set; }
        public long? SessionId { get; set; }
        public string SheetName { get; set; }
        public long? RecordId { get; set; }
        public int? Day { get; set; }
        public string BeforeJson { get; set; }
        public string AfterJson { get; set; }
        public string ChangeSummary { get; set; }
    }
}
