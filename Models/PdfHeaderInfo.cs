namespace VisorQuotationWebApp.Models;

/// <summary>
/// Header information extracted from the PDF stock list
/// </summary>
public class PdfHeaderInfo
{
    /// <summary>
    /// Project name (e.g., "Hyrja Bdhe C_Dumnica_Vetem Alumin")
    /// </summary>
    public string ProjectName { get; set; } = string.Empty;
    
    /// <summary>
    /// Date and time from PDF (e.g., "2/16/2026 / 15:46")
    /// </summary>
    public string PdfDateTime { get; set; } = string.Empty;
    
    /// <summary>
    /// Directory path from PDF
    /// </summary>
    public string Directory { get; set; } = string.Empty;
    
    /// <summary>
    /// Person in charge
    /// </summary>
    public string PersonInCharge { get; set; } = string.Empty;
    
    /// <summary>
    /// Assembly place (e.g., "Shop Floor 1")
    /// </summary>
    public string AssemblyPlace { get; set; } = string.Empty;
    
    /// <summary>
    /// Company name (e.g., "Visor LLC")
    /// </summary>
    public string CompanyName { get; set; } = string.Empty;
    
    /// <summary>
    /// Total pages in the PDF
    /// </summary>
    public int TotalPages { get; set; }
}
