namespace VisorQuotationWebApp.Models;

/// <summary>
/// Represents a profile line item extracted from the PDF stock list
/// </summary>
public class ProfileItem
{
    public int Id { get; set; }
    
    /// <summary>
    /// The profile reference number (e.g., 2015, 2019, 4503)
    /// </summary>
    public string RefNumber { get; set; } = string.Empty;
    
    /// <summary>
    /// Quantity of profiles needed
    /// </summary>
    public int Amount { get; set; }
    
    /// <summary>
    /// Raw colour text from PDF (e.g., "Special 1 Powder Coating P1019M")
    /// </summary>
    public string RawColour { get; set; } = string.Empty;
    
    /// <summary>
    /// Mapped Finish 1 dropdown value
    /// </summary>
    public string Finish1 { get; set; } = string.Empty;
    
    /// <summary>
    /// Mapped Shade 1 dropdown value
    /// </summary>
    public string Shade1 { get; set; } = string.Empty;
    
    /// <summary>
    /// Mapped Finish 2 dropdown value
    /// </summary>
    public string Finish2 { get; set; } = string.Empty;
    
    /// <summary>
    /// Mapped Shade 2 dropdown value
    /// </summary>
    public string Shade2 { get; set; } = string.Empty;
    
    /// <summary>
    /// Profile description from PDF
    /// </summary>
    public string Description { get; set; } = string.Empty;
    
    /// <summary>
    /// Total length in meters (e.g., 210.2 from "(210.2)")
    /// </summary>
    public decimal TotalLength { get; set; }
    
    /// <summary>
    /// Whether this item should be included in the automation
    /// </summary>
    public bool IsSelected { get; set; } = true;
}
