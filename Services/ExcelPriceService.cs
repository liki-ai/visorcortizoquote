using ClosedXML.Excel;
using VisorQuotationWebApp.Models;

namespace VisorQuotationWebApp.Services;

/// <summary>
/// Service for calculating prices from Excel price lists without Cortizo automation
/// </summary>
public class ExcelPriceService
{
    private readonly ILogger<ExcelPriceService> _logger;
    
    // Cache for loaded price data
    private Dictionary<string, ProfilePriceData>? _profilePrices;
    private Dictionary<string, AccessoryPriceData>? _accessoryPrices;
    private Dictionary<string, ColorData>? _colorData;
    
    public ExcelPriceService(ILogger<ExcelPriceService> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Load profile prices from Excel file (Precios_XXXXX_XXXXXX.xlsx)
    /// </summary>
    public Task<bool> LoadProfilePricesAsync(string filePath)
    {
        try
        {
            _logger.LogInformation("Loading profile prices from: {FilePath}", filePath);
            
            if (!File.Exists(filePath))
            {
                _logger.LogError("Profile prices file not found: {FilePath}", filePath);
                return Task.FromResult(false);
            }

            _profilePrices = new Dictionary<string, ProfilePriceData>(StringComparer.OrdinalIgnoreCase);

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null)
            {
                _logger.LogError("No worksheet found in profile prices file");
                return Task.FromResult(false);
            }

            // Find the data range - skip header rows
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            
            _logger.LogInformation("Found {Rows} rows and {Cols} columns in profile prices", lastRow, lastCol);

            // Try to detect header row - look for common headers like "REFERENCIA", "REF", "PRECIO"
            int headerRow = 1;
            for (int row = 1; row <= Math.Min(10, lastRow); row++)
            {
                var cellValue = worksheet.Cell(row, 1).GetString().ToUpperInvariant();
                if (cellValue.Contains("REF") || cellValue.Contains("CODIGO") || cellValue.Contains("REFERENCE"))
                {
                    headerRow = row;
                    break;
                }
            }

            // Detect column indices
            int refCol = -1, priceCol = -1, descCol = -1, weightCol = -1;
            
            for (int col = 1; col <= lastCol; col++)
            {
                var headerValue = worksheet.Cell(headerRow, col).GetString().ToUpperInvariant();
                
                if (headerValue.Contains("REF") || headerValue.Contains("CODIGO"))
                    refCol = col;
                else if (headerValue.Contains("PRECIO") || headerValue.Contains("PRICE") || headerValue.Contains("PVP"))
                    priceCol = col;
                else if (headerValue.Contains("DESC") || headerValue.Contains("NOMBRE") || headerValue.Contains("NAME"))
                    descCol = col;
                else if (headerValue.Contains("PESO") || headerValue.Contains("WEIGHT") || headerValue.Contains("KG"))
                    weightCol = col;
            }

            _logger.LogInformation("Detected columns - Ref: {RefCol}, Price: {PriceCol}, Desc: {DescCol}, Weight: {WeightCol}",
                refCol, priceCol, descCol, weightCol);

            // If we couldn't detect columns, use defaults based on typical Cortizo format
            if (refCol < 0) refCol = 1;
            if (priceCol < 0) priceCol = 2;

            // Read data rows
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                var refValue = worksheet.Cell(row, refCol).GetString().Trim();
                
                if (string.IsNullOrWhiteSpace(refValue))
                    continue;

                // Clean reference number (remove spaces, leading zeros in some cases)
                var cleanRef = CleanReferenceNumber(refValue);

                var priceData = new ProfilePriceData
                {
                    Reference = cleanRef,
                    OriginalReference = refValue
                };

                // Try to get price
                if (priceCol > 0)
                {
                    var priceCell = worksheet.Cell(row, priceCol);
                    if (priceCell.TryGetValue<decimal>(out var price))
                    {
                        priceData.PricePerKg = price;
                    }
                    else
                    {
                        // Try parsing as string
                        var priceStr = priceCell.GetString().Replace(",", ".").Replace("€", "").Trim();
                        if (decimal.TryParse(priceStr, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out price))
                        {
                            priceData.PricePerKg = price;
                        }
                    }
                }

                // Try to get description
                if (descCol > 0)
                {
                    priceData.Description = worksheet.Cell(row, descCol).GetString().Trim();
                }

                // Try to get weight
                if (weightCol > 0)
                {
                    var weightCell = worksheet.Cell(row, weightCol);
                    if (weightCell.TryGetValue<decimal>(out var weight))
                    {
                        priceData.WeightPerMeter = weight;
                    }
                }

                if (!_profilePrices.ContainsKey(cleanRef))
                {
                    _profilePrices[cleanRef] = priceData;
                }
            }

            _logger.LogInformation("Loaded {Count} profile prices", _profilePrices.Count);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading profile prices from {FilePath}", filePath);
            return Task.FromResult(false);
        }
    }

    /// <summary>
    /// Load accessory prices from Excel file (TarifaAcc_XXXXX_XXXXXX.xlsx)
    /// </summary>
    public Task<bool> LoadAccessoryPricesAsync(string filePath)
    {
        try
        {
            _logger.LogInformation("Loading accessory prices from: {FilePath}", filePath);
            
            if (!File.Exists(filePath))
            {
                _logger.LogError("Accessory prices file not found: {FilePath}", filePath);
                return Task.FromResult(false);
            }

            _accessoryPrices = new Dictionary<string, AccessoryPriceData>(StringComparer.OrdinalIgnoreCase);

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null)
            {
                _logger.LogError("No worksheet found in accessory prices file");
                return Task.FromResult(false);
            }

            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            
            _logger.LogInformation("Found {Rows} rows and {Cols} columns in accessory prices", lastRow, lastCol);

            // Detect header row
            int headerRow = 1;
            for (int row = 1; row <= Math.Min(10, lastRow); row++)
            {
                var cellValue = worksheet.Cell(row, 1).GetString().ToUpperInvariant();
                if (cellValue.Contains("REF") || cellValue.Contains("CODIGO") || cellValue.Contains("REFERENCE"))
                {
                    headerRow = row;
                    break;
                }
            }

            // Detect column indices
            int refCol = -1, priceCol = -1, descCol = -1;
            
            for (int col = 1; col <= lastCol; col++)
            {
                var headerValue = worksheet.Cell(headerRow, col).GetString().ToUpperInvariant();
                
                if (headerValue.Contains("REF") || headerValue.Contains("CODIGO"))
                    refCol = col;
                else if (headerValue.Contains("PRECIO") || headerValue.Contains("PRICE") || headerValue.Contains("PVP"))
                    priceCol = col;
                else if (headerValue.Contains("DESC") || headerValue.Contains("NOMBRE") || headerValue.Contains("NAME"))
                    descCol = col;
            }

            if (refCol < 0) refCol = 1;
            if (priceCol < 0) priceCol = 2;

            // Read data rows
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                var refValue = worksheet.Cell(row, refCol).GetString().Trim();
                
                if (string.IsNullOrWhiteSpace(refValue))
                    continue;

                var cleanRef = CleanReferenceNumber(refValue);

                var priceData = new AccessoryPriceData
                {
                    Reference = cleanRef,
                    OriginalReference = refValue
                };

                if (priceCol > 0)
                {
                    var priceCell = worksheet.Cell(row, priceCol);
                    if (priceCell.TryGetValue<decimal>(out var price))
                    {
                        priceData.PricePerUnit = price;
                    }
                    else
                    {
                        var priceStr = priceCell.GetString().Replace(",", ".").Replace("€", "").Trim();
                        if (decimal.TryParse(priceStr, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out price))
                        {
                            priceData.PricePerUnit = price;
                        }
                    }
                }

                if (descCol > 0)
                {
                    priceData.Description = worksheet.Cell(row, descCol).GetString().Trim();
                }

                if (!_accessoryPrices.ContainsKey(cleanRef))
                {
                    _accessoryPrices[cleanRef] = priceData;
                }
            }

            _logger.LogInformation("Loaded {Count} accessory prices", _accessoryPrices.Count);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading accessory prices from {FilePath}", filePath);
            return Task.FromResult(false);
        }
    }

    /// <summary>
    /// Load color/finish data from Excel file (COLORES CORTIZO.xlsx)
    /// </summary>
    public Task<bool> LoadColorDataAsync(string filePath)
    {
        try
        {
            _logger.LogInformation("Loading color data from: {FilePath}", filePath);
            
            if (!File.Exists(filePath))
            {
                _logger.LogWarning("Color data file not found: {FilePath}", filePath);
                return Task.FromResult(false);
            }

            _colorData = new Dictionary<string, ColorData>(StringComparer.OrdinalIgnoreCase);

            using var workbook = new XLWorkbook(filePath);
            
            foreach (var worksheet in workbook.Worksheets)
            {
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

                // Read all color codes from the worksheet
                for (int row = 1; row <= lastRow; row++)
                {
                    for (int col = 1; col <= lastCol; col++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 3)
                        {
                            // Check if it looks like a color code (alphanumeric)
                            if (IsValidColorCode(cellValue) && !_colorData.ContainsKey(cellValue))
                            {
                                _colorData[cellValue] = new ColorData
                                {
                                    Code = cellValue,
                                    SheetName = worksheet.Name
                                };
                            }
                        }
                    }
                }
            }

            _logger.LogInformation("Loaded {Count} color codes", _colorData.Count);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading color data from {FilePath}", filePath);
            return Task.FromResult(false);
        }
    }

    /// <summary>
    /// Calculate totals for profiles and accessories
    /// </summary>
    public PriceCalculationResult CalculateTotals(
        List<ProfileItem> profiles, 
        List<AccessoryItem>? accessories,
        decimal? customWithBreakPrice = null,
        decimal? customWithoutBreakPrice = null,
        decimal? accessoryDiscount = null)
    {
        var result = new PriceCalculationResult
        {
            CalculationDate = DateTime.Now,
            ProfilesCalculated = new List<ProfileCalculationItem>(),
            AccessoriesCalculated = new List<AccessoryCalculationItem>(),
            UnmatchedProfiles = new List<string>(),
            UnmatchedAccessories = new List<string>()
        };

        if (_profilePrices == null)
        {
            result.Warnings.Add("Profile prices not loaded. Please load Excel price list first.");
            return result;
        }

        // Calculate profile totals
        decimal profileTotal = 0;
        foreach (var profile in profiles.Where(p => p.IsSelected))
        {
            var cleanRef = CleanReferenceNumber(profile.RefNumber);
            
            if (_profilePrices.TryGetValue(cleanRef, out var priceData))
            {
                var calcItem = new ProfileCalculationItem
                {
                    RefNumber = profile.RefNumber,
                    Amount = profile.Amount,
                    Description = profile.Description ?? priceData.Description,
                    PricePerKg = priceData.PricePerKg,
                    WeightPerMeter = priceData.WeightPerMeter
                };

                // If we have weight, calculate total
                if (priceData.WeightPerMeter > 0 && priceData.PricePerKg > 0)
                {
                    // Assuming Amount is in meters or pieces
                    // Standard bar length is typically 6m
                    decimal totalWeight = profile.Amount * priceData.WeightPerMeter * 6; // 6m bars
                    calcItem.TotalWeight = totalWeight;
                    calcItem.TotalPrice = totalWeight * priceData.PricePerKg;
                }
                else if (priceData.PricePerKg > 0)
                {
                    // Use price directly if no weight info
                    calcItem.TotalPrice = profile.Amount * priceData.PricePerKg;
                }

                // Apply custom prices if provided (for special powder coating)
                if (customWithBreakPrice.HasValue && IsWithBreakProfile(profile))
                {
                    calcItem.CustomPriceApplied = true;
                    calcItem.TotalPrice = calcItem.TotalWeight > 0 
                        ? calcItem.TotalWeight * customWithBreakPrice.Value 
                        : profile.Amount * customWithBreakPrice.Value;
                }
                else if (customWithoutBreakPrice.HasValue && !IsWithBreakProfile(profile))
                {
                    calcItem.CustomPriceApplied = true;
                    calcItem.TotalPrice = calcItem.TotalWeight > 0 
                        ? calcItem.TotalWeight * customWithoutBreakPrice.Value 
                        : profile.Amount * customWithoutBreakPrice.Value;
                }

                profileTotal += calcItem.TotalPrice;
                result.ProfilesCalculated.Add(calcItem);
            }
            else
            {
                result.UnmatchedProfiles.Add(profile.RefNumber);
                _logger.LogWarning("Profile not found in price list: {Ref}", profile.RefNumber);
            }
        }
        result.ProfilesTotal = profileTotal;

        // Calculate accessory totals
        if (accessories != null && _accessoryPrices != null)
        {
            decimal accessoryTotal = 0;
            foreach (var accessory in accessories.Where(a => a.IsSelected))
            {
                var cleanRef = CleanReferenceNumber(accessory.RefNumber);
                
                if (_accessoryPrices.TryGetValue(cleanRef, out var priceData))
                {
                    var calcItem = new AccessoryCalculationItem
                    {
                        RefNumber = accessory.RefNumber,
                        Amount = accessory.Amount,
                        Description = accessory.Description ?? priceData.Description,
                        PricePerUnit = priceData.PricePerUnit,
                        TotalPrice = accessory.Amount * priceData.PricePerUnit
                    };

                    // Apply discount if provided
                    if (accessoryDiscount.HasValue && accessoryDiscount.Value > 0)
                    {
                        calcItem.DiscountApplied = accessoryDiscount.Value;
                        calcItem.TotalPrice *= (1 - accessoryDiscount.Value / 100);
                    }

                    accessoryTotal += calcItem.TotalPrice;
                    result.AccessoriesCalculated.Add(calcItem);
                }
                else
                {
                    result.UnmatchedAccessories.Add(accessory.RefNumber);
                    _logger.LogWarning("Accessory not found in price list: {Ref}", accessory.RefNumber);
                }
            }
            result.AccessoriesTotal = accessoryTotal;
        }

        result.GrandTotal = result.ProfilesTotal + result.AccessoriesTotal;

        _logger.LogInformation("Calculation complete - Profiles: {ProfileTotal:C}, Accessories: {AccTotal:C}, Total: {Total:C}",
            result.ProfilesTotal, result.AccessoriesTotal, result.GrandTotal);

        return result;
    }

    /// <summary>
    /// Get diagnostic info about loaded data
    /// </summary>
    public ExcelDataStatus GetStatus()
    {
        return new ExcelDataStatus
        {
            ProfilePricesLoaded = _profilePrices != null,
            ProfilePricesCount = _profilePrices?.Count ?? 0,
            AccessoryPricesLoaded = _accessoryPrices != null,
            AccessoryPricesCount = _accessoryPrices?.Count ?? 0,
            ColorDataLoaded = _colorData != null,
            ColorDataCount = _colorData?.Count ?? 0
        };
    }

    /// <summary>
    /// Look up a single profile price
    /// </summary>
    public ProfilePriceData? GetProfilePrice(string reference)
    {
        if (_profilePrices == null) return null;
        var cleanRef = CleanReferenceNumber(reference);
        return _profilePrices.TryGetValue(cleanRef, out var data) ? data : null;
    }

    /// <summary>
    /// Look up a single accessory price
    /// </summary>
    public AccessoryPriceData? GetAccessoryPrice(string reference)
    {
        if (_accessoryPrices == null) return null;
        var cleanRef = CleanReferenceNumber(reference);
        return _accessoryPrices.TryGetValue(cleanRef, out var data) ? data : null;
    }

    private string CleanReferenceNumber(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return string.Empty;

        // Remove spaces, leading/trailing characters
        var cleaned = reference.Trim().Replace(" ", "");
        
        // Remove any leading zeros for numeric references
        if (cleaned.All(char.IsDigit) && cleaned.Length > 1)
        {
            cleaned = cleaned.TrimStart('0');
        }

        return cleaned;
    }

    private bool IsValidColorCode(string value)
    {
        // Color codes are typically alphanumeric, 4-10 characters
        if (value.Length < 3 || value.Length > 15)
            return false;

        // Should start with letter or number
        if (!char.IsLetterOrDigit(value[0]))
            return false;

        // Should be mostly alphanumeric
        int alphanumericCount = value.Count(char.IsLetterOrDigit);
        return alphanumericCount >= value.Length * 0.8;
    }

    private bool IsWithBreakProfile(ProfileItem profile)
    {
        // Determine if profile is "with thermal break" based on reference pattern
        // This is a simplified check - actual logic may depend on Cortizo's naming convention
        var desc = (profile.Description ?? "").ToUpperInvariant();
        return desc.Contains("RPT") || desc.Contains("ROTURA") || desc.Contains("BREAK");
    }
}

#region Models

public class ProfilePriceData
{
    public string Reference { get; set; } = string.Empty;
    public string OriginalReference { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public decimal PricePerKg { get; set; }
    public decimal WeightPerMeter { get; set; }
}

public class AccessoryPriceData
{
    public string Reference { get; set; } = string.Empty;
    public string OriginalReference { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public decimal PricePerUnit { get; set; }
}

public class ColorData
{
    public string Code { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}

public class PriceCalculationResult
{
    public DateTime CalculationDate { get; set; }
    public List<ProfileCalculationItem> ProfilesCalculated { get; set; } = new();
    public List<AccessoryCalculationItem> AccessoriesCalculated { get; set; } = new();
    public List<string> UnmatchedProfiles { get; set; } = new();
    public List<string> UnmatchedAccessories { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    
    public decimal ProfilesTotal { get; set; }
    public decimal AccessoriesTotal { get; set; }
    public decimal GrandTotal { get; set; }
    
    public int TotalProfilesMatched => ProfilesCalculated.Count;
    public int TotalAccessoriesMatched => AccessoriesCalculated.Count;
}

public class ProfileCalculationItem
{
    public string RefNumber { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public string Description { get; set; } = string.Empty;
    public decimal PricePerKg { get; set; }
    public decimal WeightPerMeter { get; set; }
    public decimal TotalWeight { get; set; }
    public decimal TotalPrice { get; set; }
    public bool CustomPriceApplied { get; set; }
}

public class AccessoryCalculationItem
{
    public string RefNumber { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public string Description { get; set; } = string.Empty;
    public decimal PricePerUnit { get; set; }
    public decimal TotalPrice { get; set; }
    public decimal DiscountApplied { get; set; }
}

public class ExcelDataStatus
{
    public bool ProfilePricesLoaded { get; set; }
    public int ProfilePricesCount { get; set; }
    public bool AccessoryPricesLoaded { get; set; }
    public int AccessoryPricesCount { get; set; }
    public bool ColorDataLoaded { get; set; }
    public int ColorDataCount { get; set; }
}

#endregion
