namespace VisorQuotationWebApp.Models;

/// <summary>
/// Configuration for Cortizo automation
/// </summary>
public class AutomationConfig
{
    /// <summary>
    /// Cortizo Center base URL
    /// </summary>
    public string BaseUrl { get; set; } = "https://cortizocenter.com";
    
    /// <summary>
    /// Default microns value
    /// </summary>
    public int DefaultMicrons { get; set; } = 15;
    
    /// <summary>
    /// Default CIF value
    /// </summary>
    public string DefaultCif { get; set; } = "CORTIZO";
    
    /// <summary>
    /// Default client code
    /// </summary>
    public string DefaultClientCode { get; set; } = "991238";
    
    /// <summary>
    /// Default language
    /// </summary>
    public string DefaultLanguage { get; set; } = "ENGLISH";
    
    /// <summary>
    /// Finish/shade mappings based on raw colour text
    /// </summary>
    public Dictionary<string, FinishMapping> FinishMappings { get; set; } = new();
    
    /// <summary>
    /// Whether to run browser in headless mode
    /// </summary>
    public bool Headless { get; set; } = true;
    
    /// <summary>
    /// Timeout for browser operations in milliseconds
    /// </summary>
    public int TimeoutMs { get; set; } = 30000;
}

/// <summary>
/// Mapping from raw colour to finish/shade dropdown values
/// </summary>
public class FinishMapping
{
    public string Finish1 { get; set; } = string.Empty;
    public string Shade1 { get; set; } = string.Empty;
    public string Finish2 { get; set; } = string.Empty;
    public string Shade2 { get; set; } = string.Empty;
}
