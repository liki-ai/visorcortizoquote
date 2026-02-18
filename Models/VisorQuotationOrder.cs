namespace VisorQuotationWebApp.Models;

/// <summary>
/// Visor company information for quotation documents
/// </summary>
public class VisorCompanyInfo
{
    public string CompanyName { get; set; } = "VISOR Sh.p.k";
    public string Address { get; set; } = "Rruga e Kavajes, Km 7";
    public string City { get; set; } = "Tirane, Albania";
    public string Phone { get; set; } = "+355 69 XXX XXXX";
    public string Email { get; set; } = "info@visor.al";
    public string Website { get; set; } = "www.visor.al";
    public string Nipt { get; set; } = "L XXXXXXXXX"; // Albanian tax ID
    public string BankAccount { get; set; } = ""; 
    public string BankName { get; set; } = "";
}

/// <summary>
/// Represents a complete quotation order for Visor
/// </summary>
public class VisorQuotationOrder
{
    // Company info
    public VisorCompanyInfo CompanyInfo { get; set; } = new();
    
    // Quotation metadata
    public string QuotationNumber { get; set; } = string.Empty;
    public DateTime QuotationDate { get; set; } = DateTime.Now;
    public DateTime ValidUntil { get; set; } = DateTime.Now.AddDays(30);
    
    // Project info (from PDF header)
    public string ProjectName { get; set; } = string.Empty;
    public string ClientName { get; set; } = string.Empty;
    public string ClientAddress { get; set; } = string.Empty;
    public string ClientPhone { get; set; } = string.Empty;
    public string ClientEmail { get; set; } = string.Empty;
    
    // Profile items with prices (from Cortizo)
    public List<VisorQuotationItem> Items { get; set; } = new();
    
    // Totals
    public decimal Subtotal { get; set; }
    public decimal VatRate { get; set; } = 20m; // Albania VAT rate
    public decimal VatAmount { get; set; }
    public decimal Total { get; set; }
    public string Currency { get; set; } = "EUR";
    
    // Additional info
    public string DeliveryTerms { get; set; } = "Ex Works";
    public string PaymentTerms { get; set; } = "50% advance, 50% on delivery";
    public int DeliveryDays { get; set; } = 21;
    public string Notes { get; set; } = string.Empty;
    
    // Reference to Cortizo quotation
    public string CortizoReference { get; set; } = string.Empty;
    public decimal CortizoTotal { get; set; }
    
    // Unfilled items that need manual review
    public List<UnfilledItem> UnfilledProfiles { get; set; } = new();
    public List<UnfilledItem> UnfilledAccessories { get; set; } = new();
    
    // Items missing from price calculation (not found in Excel or with no price)
    public List<MissingCalculationItem> MissingCalculationItems { get; set; } = new();
}

/// <summary>
/// A single item in the Visor quotation
/// </summary>
public class VisorQuotationItem
{
    public int LineNumber { get; set; }
    public string RefNumber { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Finish { get; set; } = string.Empty;
    public string Shade { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Length { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal TotalPrice { get; set; }
}

/// <summary>
/// Represents an item missing from price calculation that needs manual entry
/// </summary>
public class MissingCalculationItem
{
    public string RefNumber { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public string Description { get; set; } = string.Empty;
    public string Finish { get; set; } = string.Empty;
    /// <summary>
    /// "Profile", "Accessory", or "Hardware"
    /// </summary>
    public string Category { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
    /// <summary>
    /// Placeholder for manual price entry (shown as blank in the report)
    /// </summary>
    public decimal ManualPrice { get; set; }
}
