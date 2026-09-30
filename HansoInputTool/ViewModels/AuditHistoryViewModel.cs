using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class AuditHistoryViewModel : ObservableObject
    {
        private readonly List<AuditEntry> _allEntries;
        private string _searchText = "";

        public ObservableCollection<AuditEntry> Entries { get; } = new();
        public string ResultCount => $"{Entries.Count:N0}件（最新2,000件まで表示）";

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value)) ApplyFilter();
            }
        }

        public AuditHistoryViewModel(DatabaseService databaseService)
        {
            _allEntries = databaseService?.GetAuditHistory() ?? new List<AuditEntry>();
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            var query = (SearchText ?? "").Trim();
            var filtered = string.IsNullOrEmpty(query)
                ? _allEntries
                : _allEntries.Where(e =>
                    (e.EventAt ?? "").Contains(query, System.StringComparison.OrdinalIgnoreCase) ||
                    (e.UserName ?? "").Contains(query, System.StringComparison.OrdinalIgnoreCase) ||
                    (e.Action ?? "").Contains(query, System.StringComparison.OrdinalIgnoreCase) ||
                    (e.EntityType ?? "").Contains(query, System.StringComparison.OrdinalIgnoreCase) ||
                    (e.SheetName ?? "").Contains(query, System.StringComparison.OrdinalIgnoreCase) ||
                    (e.ChangeSummary ?? "").Contains(query, System.StringComparison.OrdinalIgnoreCase));

            Entries.Clear();
            foreach (var entry in filtered) Entries.Add(entry);
            OnPropertyChanged(nameof(ResultCount));
        }
    }
}
