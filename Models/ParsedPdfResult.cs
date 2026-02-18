namespace VisorQuotationWebApp.Models;

/// <summary>
/// Complete result from parsing a PDF stock list
/// </summary>
public class ParsedPdfResult
{
    public PdfHeaderInfo Header { get; set; } = new();
    public List<ProfileItem> Profiles { get; set; } = new();
    public List<AccessoryItem> Accessories { get; set; } = new();
    public List<AccessoryItem> HardwareItems { get; set; } = new();
    public List<string> ParseWarnings { get; set; } = new();
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
}
