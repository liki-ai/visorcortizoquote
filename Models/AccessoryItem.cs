namespace VisorQuotationWebApp.Models;

/// <summary>
/// Represents an accessory line item extracted from the PDF stock list
/// </summary>
public class AccessoryItem
{
    public int Id { get; set; }
    
    /// <summary>
    /// The accessory reference number
    /// </summary>
    public string RefNumber { get; set; } = string.Empty;
    
    /// <summary>
    /// Quantity of accessories needed
    /// </summary>
    public int Amount { get; set; }
    
    /// <summary>
    /// Accessory description
    /// </summary>
    public string Description { get; set; } = string.Empty;
    
    /// <summary>
    /// Finish/coating type
    /// </summary>
    public string Finish { get; set; } = string.Empty;
    
    /// <summary>
    /// Shade/colour code
    /// </summary>
    public string Shade { get; set; } = string.Empty;
    
    /// <summary>
    /// Whether this item should be included in the automation
    /// </summary>
    public bool IsSelected { get; set; } = true;
    
    /// <summary>
    /// Source section in the PDF: "Accessory" or "Hardware"
    /// </summary>
    public string Source { get; set; } = "Accessory";
}
