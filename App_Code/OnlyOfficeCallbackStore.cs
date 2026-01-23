using System;
using System.Collections.Concurrent;

namespace WebEditor
{
    public sealed class OnlyOfficeCallbackStore
    {
        public sealed class SavedUrlInfo
        {
            public string Url { get; set; }
            public DateTime UtcSavedAt { get; set; }
        }

        private readonly ConcurrentDictionary<string, SavedUrlInfo> _savedByKey;

        public OnlyOfficeCallbackStore()
        {
            _savedByKey = new ConcurrentDictionary<string, SavedUrlInfo>(StringComparer.OrdinalIgnoreCase);
        }

        public bool TryGet(string lookupKey, out SavedUrlInfo saved) => _savedByKey.TryGetValue(lookupKey ?? string.Empty, out saved);

        public void Upsert(string key, string url)
        {
            if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(url)) return;

            _savedByKey[key] = new SavedUrlInfo
            {
                Url = url,
                UtcSavedAt = DateTime.UtcNow,
            };
        }

        public bool TryFindByContainedFileId(string fileId, out SavedUrlInfo match)
        {
            match = null;
            if (string.IsNullOrWhiteSpace(fileId)) return false;

            foreach (var kvp in _savedByKey)
            {
                var u = kvp.Value?.Url;
                if (!string.IsNullOrWhiteSpace(u) && u.IndexOf(fileId, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    match = kvp.Value;
                    return true;
                }
            }

            return false;
        }
    }
}
