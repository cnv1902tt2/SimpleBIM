using System;

namespace SimpleBIM.Update
{
    /// <summary>
    /// Thông tin về update từ server
    /// </summary>
    public class UpdateInfo
    {
        public bool UpdateAvailable { get; set; }
        public string LatestVersion { get; set; }
        public string MinimumRequiredVersion { get; set; }
        public DateTime ReleaseDate { get; set; }
        public string ReleaseNotes { get; set; }
        public string DownloadUrl { get; set; }
        public long FileSize { get; set; }
        public string ChecksumSHA256 { get; set; }
        public UpdateType UpdateType { get; set; }
        public bool ForceUpdate { get; set; }
        public string NotificationMessage { get; set; }
    }

    /// <summary>
    /// Request gửi lên server để check version
    /// </summary>
    public class UpdateCheckRequest
    {
        public string Product { get; set; } = "SimpleBIM";
        public string CurrentVersion { get; set; }
        public string RevitVersion { get; set; }
        public string MachineHash { get; set; }
        public string OS { get; set; }
    }

    /// <summary>
    /// Response từ server
    /// </summary>
    public class UpdateCheckResponse
    {
        public bool UpdateAvailable { get; set; }
        public string LatestVersion { get; set; }
        public string MinimumRequiredVersion { get; set; }
        public string ReleaseDate { get; set; }
        public string ReleaseNotes { get; set; }
        public string DownloadUrl { get; set; }
        public long FileSize { get; set; }
        public string ChecksumSHA256 { get; set; }
        public string UpdateType { get; set; }
        public bool ForceUpdate { get; set; }
        public string NotificationMessage { get; set; }
    }

    /// <summary>
    /// Loại update
    /// </summary>
    public enum UpdateType
    {
        Optional,    // User có thể bỏ qua
        Recommended, // Khuyến khích nhưng không bắt buộc
        Mandatory    // Bắt buộc phải update
    }

    /// <summary>
    /// Trạng thái của quá trình update
    /// </summary>
    public enum UpdateStatus
    {
        Idle,
        Checking,
        Downloading,
        Verifying,
        Installing,
        Completed,
        Failed,
        Cancelled
    }

    /// <summary>
    /// Kết quả của update check
    /// </summary>
    public class UpdateCheckResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public UpdateInfo UpdateInfo { get; set; }
    }

    /// <summary>
    /// Progress event args
    /// </summary>
    public class UpdateProgressEventArgs : EventArgs
    {
        public UpdateStatus Status { get; set; }
        public int ProgressPercentage { get; set; }
        public string Message { get; set; }
        public long BytesDownloaded { get; set; }
        public long TotalBytes { get; set; }
    }
}
