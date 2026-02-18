namespace VisorQuotationWebApp.Models;

/// <summary>
/// View model for the main quotation page
/// </summary>
public class QuotationViewModel
{
    // Parsed PDF data
    public ParsedPdfResult? ParsedPdf { get; set; }
    
    // Cortizo header fields (editable)
    public int Microns { get; set; } = 15;
    public string Cif { get; set; } = "CORTIZO";
    public string ClientCode { get; set; } = "991238";
    public string Language { get; set; } = "ENGLISH";
    public string ClientPurchaseOrder { get; set; } = string.Empty;
    
    // Delivery settings
    public string ProfDeliv { get; set; } = "CIP-15901 CARRETERA NOYA-PADRÓN, P.";
    public string AccDeliv { get; set; } = "CIP-15901 CARRETERA NOYA-PADRÓN, P.";
    public string PvcDeliv { get; set; } = "CIP-15901 CARRETERA NOYA-PADRÓN, P.";
    
    // General colour settings (values match Cortizo dropdown option values)
    // 90 = "SPECIAL 1 POWDER COATING", 8 = "MILL FINISH", 9 = "STANDARD POWDER COATING"
    public string GeneralFinish1 { get; set; } = "90";
    public string GeneralShade1 { get; set; } = "P1019M";
    public string GeneralFinish2 { get; set; } = "90";
    public string GeneralShade2 { get; set; } = "P1019M";
    
    // Automation options
    public bool GenerateReport { get; set; } = false;
    public bool CreateProforma { get; set; } = false;
    
    // Uploaded file info
    public string? UploadedFileName { get; set; }
    public string? UploadedFilePath { get; set; }
}
