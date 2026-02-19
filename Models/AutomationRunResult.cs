namespace VisorQuotationWebApp.Models;

/// <summary>
/// Result of running the Cortizo automation
/// </summary>
public class AutomationRunResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int TotalItems { get; set; }
    public int SuccessfulItems { get; set; }
    public int FailedItems { get; set; }
    public List<AutomationLogEntry> Logs { get; set; } = new();
    public string? ScreenshotPath { get; set; }
    public string? TracePath { get; set; }
    
    /// <summary>
    /// List of profile rows that could not be filled or had missing amounts
    /// </summary>
    public List<UnfilledItem> UnfilledProfiles { get; set; } = new();
    
    /// <summary>
    /// List of accessory rows that could not be filled or had missing amounts
    /// </summary>
    public List<UnfilledItem> UnfilledAccessories { get; set; } = new();
    
    /// <summary>
    /// Total from Cortizo (ESTIMATE TOTAL)
    /// </summary>
    public decimal CortizoTotal { get; set; }
    
    /// <summary>
    /// Path to the downloaded Cortizo report ZIP file
    /// </summary>
    public string? ReportZipPath { get; set; }
    
    /// <summary>
    /// URL to download the Cortizo report ZIP (relative)
    /// </summary>
    public string? ReportDownloadUrl { get; set; }
    
    /// <summary>
    /// Original filename of the Cortizo report ZIP
    /// </summary>
    public string? ReportFileName { get; set; }
}

/// <summary>
/// Represents an item that could not be fully filled during automation
/// </summary>
public class UnfilledItem
{
    public string RowNumber { get; set; } = string.Empty;
    public string RefNumber { get; set; } = string.Empty;
    public int Amount { get; set; }
    public string Description { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
}

/// <summary>
/// A single log entry from the automation run
/// </summary>
public class AutomationLogEntry
{
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public AutomationLogLevel Level { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? Details { get; set; }
}

public enum AutomationLogLevel
{
    Info,
    Success,
    Warning,
    Error
}
