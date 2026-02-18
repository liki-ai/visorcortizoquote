using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using VisorQuotationWebApp.Models;

namespace VisorQuotationWebApp.Services;

/// <summary>
/// Service for parsing PDF stock lists into structured data
/// </summary>
public class PdfParseService
{
    private readonly ILogger<PdfParseService> _logger;
    private readonly AutomationConfig _config;

    public PdfParseService(ILogger<PdfParseService> logger, AutomationConfig config)
    {
        _logger = logger;
        _config = config;
    }

    /// <summary>
    /// Parse a PDF file and extract profiles
    /// </summary>
    public ParsedPdfResult ParsePdf(Stream pdfStream)
    {
        var result = new ParsedPdfResult();

        try
        {
            using var document = PdfDocument.Open(pdfStream);
            result.Header.TotalPages = document.NumberOfPages;

            // Extract all text from all pages
            var allText = new List<string>();
            foreach (var page in document.GetPages())
            {
                var pageText = page.Text;
                allText.Add(pageText);
            }

            var fullText = string.Join("\n", allText);

            // Parse header info
            result.Header = ExtractHeader(fullText);
            result.Header.TotalPages = document.NumberOfPages;

            // Parse profiles section
            result.Profiles = ExtractProfiles(fullText, result.ParseWarnings);

            // Apply finish mappings
            foreach (var profile in result.Profiles)
            {
                ApplyFinishMapping(profile);
            }

            // Parse hardware section (between Profiles and Accessories)
            result.HardwareItems = ExtractHardware(fullText, result.ParseWarnings);

            // Parse accessories section
            result.Accessories = ExtractAccessories(fullText, result.ParseWarnings);

            // Merge hardware items into accessories with proper IDs
            if (result.HardwareItems.Count > 0)
            {
                int nextId = result.Accessories.Count > 0 
                    ? result.Accessories.Max(a => a.Id) + 1 
                    : 1;
                foreach (var hw in result.HardwareItems)
                {
                    hw.Id = nextId++;
                    result.Accessories.Add(hw);
                }
                _logger.LogInformation($"Merged {result.HardwareItems.Count} hardware items into accessories list");
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to parse PDF");
            result.Success = false;
            result.ErrorMessage = $"Failed to parse PDF: {ex.Message}";
        }

        return result;
    }

    /// <summary>
    /// Parse a PDF from a file path
    /// </summary>
    public ParsedPdfResult ParsePdf(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return ParsePdf(stream);
    }

    private PdfHeaderInfo ExtractHeader(string text)
    {
        var header = new PdfHeaderInfo();

        // Extract date/time (pattern: "2/16/2026 / 15:46")
        var dateMatch = Regex.Match(text, @"(\d{1,2}/\d{1,2}/\d{4}\s*/\s*\d{1,2}:\d{2})");
        if (dateMatch.Success)
        {
            header.PdfDateTime = dateMatch.Groups[1].Value.Trim();
        }

        // Extract project name (appears after date, before "Stock List")
        // Pattern: date line followed by project name
        var projectMatch = Regex.Match(text, @"\d{1,2}/\d{1,2}/\d{4}\s*/\s*\d{1,2}:\d{2}\s*\n?([^\n]+)", RegexOptions.Multiline);
        if (projectMatch.Success)
        {
            var projectName = projectMatch.Groups[1].Value.Trim();
            // Clean up - remove "Date:" if it appears
            if (!projectName.StartsWith("Date:") && !string.IsNullOrWhiteSpace(projectName))
            {
                header.ProjectName = projectName;
            }
        }

        // Try alternative pattern for project name
        if (string.IsNullOrEmpty(header.ProjectName))
        {
            var altProjectMatch = Regex.Match(text, @"Project:\s*\n?([^\n]+)", RegexOptions.Multiline);
            if (altProjectMatch.Success)
            {
                header.ProjectName = altProjectMatch.Groups[1].Value.Trim();
            }
        }

        // Extract assembly place
        var assemblyMatch = Regex.Match(text, @"Assembly Place:\s*([^\n]+)", RegexOptions.IgnoreCase);
        if (assemblyMatch.Success)
        {
            header.AssemblyPlace = assemblyMatch.Groups[1].Value.Trim();
        }

        // Extract directory
        var directoryMatch = Regex.Match(text, @"Directory:\s*\n?([^\n]+)", RegexOptions.Multiline);
        if (directoryMatch.Success)
        {
            var dir = directoryMatch.Groups[1].Value.Trim();
            if (!dir.StartsWith("Administrator"))
            {
                header.Directory = dir;
            }
        }

        // Extract person in charge
        var personMatch = Regex.Match(text, @"Person in Charge:\s*\n?([^\n]+)", RegexOptions.Multiline);
        if (personMatch.Success)
        {
            header.PersonInCharge = personMatch.Groups[1].Value.Trim();
        }

        return header;
    }

    private List<ProfileItem> ExtractProfiles(string text, List<string> warnings)
    {
        var profiles = new List<ProfileItem>();
        int id = 1;

        // Find the Profiles section - it starts after "Profiles" header and ends at "Hardware" or end of text
        var profilesSectionMatch = Regex.Match(text, @"Profiles\s*Quantity.*?(?=Hardware|Accessories|Gaskets|$)", 
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!profilesSectionMatch.Success)
        {
            warnings.Add("Could not find Profiles section in PDF");
            return profiles;
        }

        var profilesSection = profilesSectionMatch.Value;

        // Pattern to match profile lines
        // Format: "34 x 6.5 m" or "34 x 6.5 m (210.2)" followed by number and colour
        // The quantity pattern: NUMBER x 6.5 m (TOTAL) or NUMBER x 6.5 m
        var linePattern = @"(\d+)\s*x\s*6\.5\s*m\s*(?:\(([0-9,\.]+)\))?\s*(\d{4})\s+(.+?)(?=\d+\s*x\s*6\.5\s*m|\d+\s*pc|$)";

        var matches = Regex.Matches(profilesSection, linePattern, RegexOptions.Singleline);

        foreach (Match match in matches)
        {
            try
            {
                var profile = new ProfileItem
                {
                    Id = id++,
                    Amount = int.Parse(match.Groups[1].Value),
                    RefNumber = match.Groups[3].Value,
                };

                // Parse total length if present
                if (!string.IsNullOrEmpty(match.Groups[2].Value))
                {
                    var lengthStr = match.Groups[2].Value.Replace(",", "");
                    if (decimal.TryParse(lengthStr, out var length))
                    {
                        profile.TotalLength = length;
                    }
                }

                // Parse colour and description from the remaining text
                var colourDescText = match.Groups[4].Value.Trim();
                ParseColourAndDescription(colourDescText, profile);

                profiles.Add(profile);
            }
            catch (Exception ex)
            {
                warnings.Add($"Failed to parse profile line: {match.Value}. Error: {ex.Message}");
            }
        }

        // If regex didn't work well, try line-by-line parsing
        if (profiles.Count == 0)
        {
            profiles = ExtractProfilesLineByLine(profilesSection, warnings, ref id);
        }

        return profiles;
    }

    private List<ProfileItem> ExtractProfilesLineByLine(string text, List<string> warnings, ref int id)
    {
        var profiles = new List<ProfileItem>();
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        ProfileItem? currentProfile = null;
        var colourBuffer = new List<string>();

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();

            // Skip header lines
            if (line.StartsWith("Profiles") || line.StartsWith("Quantity") || 
                line.StartsWith("(PU)") || line.StartsWith("Drawing Number") ||
                line.StartsWith("Description") || line.StartsWith("Stock Order") ||
                line.StartsWith("Inside/Outside"))
            {
                continue;
            }

            // Check for quantity pattern: "34 x 6.5 m" or "34 x 6.5 m (210.2)"
            var qtyMatch = Regex.Match(line, @"^(\d+)\s*x\s*6\.5\s*m\s*(?:\(([0-9,\.]+)\))?");
            if (qtyMatch.Success)
            {
                // Save previous profile if exists
                if (currentProfile != null)
                {
                    FinalizeProfile(currentProfile, colourBuffer);
                    profiles.Add(currentProfile);
                }

                // Start new profile
                currentProfile = new ProfileItem
                {
                    Id = id++,
                    Amount = int.Parse(qtyMatch.Groups[1].Value)
                };

                if (!string.IsNullOrEmpty(qtyMatch.Groups[2].Value))
                {
                    var lengthStr = qtyMatch.Groups[2].Value.Replace(",", "");
                    if (decimal.TryParse(lengthStr, out var length))
                    {
                        currentProfile.TotalLength = length;
                    }
                }

                colourBuffer.Clear();

                // Check if there's more on this line (ref number + colour start)
                var remainder = line.Substring(qtyMatch.Length).Trim();
                if (!string.IsNullOrEmpty(remainder))
                {
                    // Try to extract ref number
                    var refMatch = Regex.Match(remainder, @"^(\d{4})\s*(.*)");
                    if (refMatch.Success)
                    {
                        currentProfile.RefNumber = refMatch.Groups[1].Value;
                        var colourStart = refMatch.Groups[2].Value.Trim();
                        if (!string.IsNullOrEmpty(colourStart))
                        {
                            colourBuffer.Add(colourStart);
                        }
                    }
                }
                continue;
            }

            // Check for just a ref number at start of line
            if (currentProfile != null && string.IsNullOrEmpty(currentProfile.RefNumber))
            {
                var refOnlyMatch = Regex.Match(line, @"^(\d{4})\s*(.*)");
                if (refOnlyMatch.Success)
                {
                    currentProfile.RefNumber = refOnlyMatch.Groups[1].Value;
                    var colourStart = refOnlyMatch.Groups[2].Value.Trim();
                    if (!string.IsNullOrEmpty(colourStart))
                    {
                        colourBuffer.Add(colourStart);
                    }
                    continue;
                }
            }

            // Otherwise, this is continuation of colour/description
            if (currentProfile != null && !string.IsNullOrEmpty(line))
            {
                colourBuffer.Add(line);
            }
        }

        // Don't forget the last profile
        if (currentProfile != null)
        {
            FinalizeProfile(currentProfile, colourBuffer);
            profiles.Add(currentProfile);
        }

        return profiles;
    }

    private void FinalizeProfile(ProfileItem profile, List<string> colourBuffer)
    {
        if (colourBuffer.Count == 0) return;

        var fullText = string.Join(" ", colourBuffer);
        ParseColourAndDescription(fullText, profile);
    }

    private void ParseColourAndDescription(string text, ProfileItem profile)
    {
        // Known colour patterns
        var colourPatterns = new[]
        {
            @"Special\s*1\s*Powder\s*Coating\s*P\d+[A-Z]*",
            @"COR_MILL\s*FINISH",
            @"Black\s*M",
            @"Silver\s*M",
            @"White\s*M?",
            @"Mill\s*Finish",
            @"Anodized.*?(?=\s{2}|$)",
        };

        string? matchedColour = null;
        int colourEndIndex = 0;

        foreach (var pattern in colourPatterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                matchedColour = match.Value.Trim();
                colourEndIndex = match.Index + match.Length;
                break;
            }
        }

        if (matchedColour != null)
        {
            // Normalize the colour text
            profile.RawColour = NormalizeColour(matchedColour);
            profile.Description = text.Substring(colourEndIndex).Trim();
        }
        else
        {
            // If no known pattern matched, try to split intelligently
            // Assume format: COLOUR_WORDS DESCRIPTION_WORDS
            var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            
            // Look for common colour keywords
            var colourKeywords = new[] { "Special", "Powder", "Coating", "COR_MILL", "FINISH", "Black", "Silver", "White", "Anodized", "Mill" };
            var colourWords = new List<string>();
            var descWords = new List<string>();
            bool inDescription = false;

            foreach (var word in words)
            {
                if (!inDescription && colourKeywords.Any(k => word.Contains(k, StringComparison.OrdinalIgnoreCase)))
                {
                    colourWords.Add(word);
                }
                else if (!inDescription && Regex.IsMatch(word, @"^P\d+[A-Z]*$"))
                {
                    colourWords.Add(word);
                    inDescription = true; // Shade code usually ends the colour
                }
                else
                {
                    inDescription = true;
                    descWords.Add(word);
                }
            }

            profile.RawColour = NormalizeColour(string.Join(" ", colourWords));
            profile.Description = string.Join(" ", descWords);
        }
    }

    private string NormalizeColour(string colour)
    {
        // Normalize whitespace and common variations
        colour = Regex.Replace(colour, @"\s+", " ").Trim();
        
        // Normalize "Special 1 Powder Coating P1019M" variations
        colour = Regex.Replace(colour, @"Special\s*1\s*Powder\s*Coating", "Special 1 Powder Coating", RegexOptions.IgnoreCase);
        
        // Normalize COR_MILL FINISH
        colour = Regex.Replace(colour, @"COR_MILL\s*FINISH", "COR_MILL FINISH", RegexOptions.IgnoreCase);

        return colour;
    }

    private void ApplyFinishMapping(ProfileItem profile)
    {
        var rawColour = profile.RawColour.ToUpperInvariant();

        // Check for exact mapping in config
        foreach (var mapping in _config.FinishMappings)
        {
            if (rawColour.Contains(mapping.Key.ToUpperInvariant()))
            {
                profile.Finish1 = mapping.Value.Finish1;
                profile.Shade1 = mapping.Value.Shade1;
                profile.Finish2 = mapping.Value.Finish2;
                profile.Shade2 = mapping.Value.Shade2;
                return;
            }
        }

        // Default mappings based on colour text patterns
        if (rawColour.Contains("SPECIAL 1") && rawColour.Contains("POWDER"))
        {
            profile.Finish1 = "SPECIAL 1 POWDER COATING";
            profile.Finish2 = "SPECIAL 1 POWDER COATING";
            
            // Extract shade code (e.g., P1019M)
            var shadeMatch = Regex.Match(profile.RawColour, @"P\d+[A-Z]*", RegexOptions.IgnoreCase);
            if (shadeMatch.Success)
            {
                profile.Shade1 = shadeMatch.Value.ToUpperInvariant();
                profile.Shade2 = shadeMatch.Value.ToUpperInvariant();
            }
        }
        else if (rawColour.Contains("COR_MILL") || rawColour.Contains("MILL FINISH"))
        {
            profile.Finish1 = "COR_MILL FINISH";
            profile.Finish2 = "COR_MILL FINISH";
            profile.Shade1 = "";
            profile.Shade2 = "";
        }
        else if (rawColour.Contains("BLACK"))
        {
            profile.Finish1 = "STANDARD";
            profile.Shade1 = "BLACK M";
            profile.Finish2 = "STANDARD";
            profile.Shade2 = "BLACK M";
        }
        else if (rawColour.Contains("SILVER"))
        {
            profile.Finish1 = "STANDARD";
            profile.Shade1 = "SILVER M";
            profile.Finish2 = "STANDARD";
            profile.Shade2 = "SILVER M";
        }
        else if (rawColour.Contains("WHITE"))
        {
            profile.Finish1 = "STANDARD";
            profile.Shade1 = "WHITE";
            profile.Finish2 = "STANDARD";
            profile.Shade2 = "WHITE";
        }
        else
        {
            // Default fallback
            profile.Finish1 = "SPECIAL 1 POWDER COATING";
            profile.Shade1 = "P1019M";
            profile.Finish2 = "SPECIAL 1 POWDER COATING";
            profile.Shade2 = "P1019M";
        }
    }

    /// <summary>
    /// Extract hardware items from the PDF text (between "Hardware" and "Accessories" sections)
    /// </summary>
    private List<AccessoryItem> ExtractHardware(string text, List<string> warnings)
    {
        var hardware = new List<AccessoryItem>();
        int id = 1;

        // Hardware section: starts after "Hardware" header and ends at "Accessories" or "Gaskets"
        var hardwareSectionMatch = Regex.Match(text, @"Hardware\s*Quantity.*?(?=Accessories|Gaskets|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!hardwareSectionMatch.Success)
        {
            _logger.LogInformation("No Hardware section found in PDF");
            return hardware;
        }

        var hardwareSection = hardwareSectionMatch.Value;
        _logger.LogInformation($"Found Hardware section: {hardwareSection.Length} chars");

        // Pattern: "14 pc 290540 [Colour] Description" or "14 pc (PU info) 290540 [Colour] Description"
        var linePattern = @"(\d+)\s*(?:pc|pcs)\s*(?:\([^)]*\))?\s*(\d{6})\s+(?:(Black|White|Silver|Grey|Brown)\s+)?(.+?)(?=\d+\s*(?:pc|pcs)|-- \d+ of|$)";

        var matches = Regex.Matches(hardwareSection, linePattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        foreach (Match match in matches)
        {
            try
            {
                var item = new AccessoryItem
                {
                    Id = id++,
                    Amount = int.Parse(match.Groups[1].Value),
                    RefNumber = match.Groups[2].Value,
                    Finish = match.Groups[3].Value.Trim(),
                    Description = match.Groups[4].Value.Trim(),
                    Source = "Hardware"
                };

                item.Description = Regex.Replace(item.Description, @"\s*Page:.*$", "", RegexOptions.IgnoreCase | RegexOptions.Singleline).Trim();
                item.Description = Regex.Replace(item.Description, @"\s*Logikal.*$", "", RegexOptions.IgnoreCase | RegexOptions.Singleline).Trim();
                item.Description = Regex.Replace(item.Description, @"\s*Stock List.*$", "", RegexOptions.IgnoreCase | RegexOptions.Singleline).Trim();
                item.Description = Regex.Replace(item.Description, @"\s{2,}", " ").Trim();

                if (item.Description.Contains("Drawing Number") ||
                    item.Description.Contains("Colour Description") ||
                    string.IsNullOrWhiteSpace(item.Description))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.RefNumber))
                {
                    hardware.Add(item);
                    _logger.LogInformation($"Parsed hardware: {item.Amount} x {item.RefNumber} - {item.Description} (Colour: {item.Finish})");
                }
            }
            catch (Exception ex)
            {
                warnings.Add($"Failed to parse hardware line: {match.Value}. Error: {ex.Message}");
            }
        }

        // Fallback: line-by-line parsing
        if (hardware.Count == 0)
        {
            hardware = ExtractHardwareLineByLine(hardwareSection, warnings, ref id);
        }

        _logger.LogInformation($"Extracted {hardware.Count} hardware items total");
        return hardware;
    }

    private List<AccessoryItem> ExtractHardwareLineByLine(string text, List<string> warnings, ref int id)
    {
        var hardware = new List<AccessoryItem>();
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();

            if (line.StartsWith("Hardware", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Quantity") ||
                line.StartsWith("Drawing") ||
                line.StartsWith("Description") ||
                line.StartsWith("(PU)") ||
                line.StartsWith("Page:") ||
                line.StartsWith("Logikal") ||
                line.StartsWith("Stock List") ||
                line.StartsWith("Stock Order") ||
                line.StartsWith("Inside") ||
                line.Contains("-- ") ||
                string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            // Pattern: "14 pc 290540 [Colour] Description"
            var match = Regex.Match(line, @"^(\d+)\s*(?:pc|pcs)\s*(?:\([^)]*\))?\s*(\d{6})\s*(?:(Black|White|Silver|Grey|Brown)\s+)?(.*)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                try
                {
                    var item = new AccessoryItem
                    {
                        Id = id++,
                        Amount = int.Parse(match.Groups[1].Value),
                        RefNumber = match.Groups[2].Value,
                        Finish = match.Groups[3].Value.Trim(),
                        Description = match.Groups[4].Value.Trim(),
                        Source = "Hardware"
                    };

                    if (!string.IsNullOrWhiteSpace(item.RefNumber) && item.RefNumber.Length >= 6)
                    {
                        hardware.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add($"Failed to parse hardware: {line}. Error: {ex.Message}");
                }
            }
        }

        return hardware;
    }

    /// <summary>
    /// Extract accessories from the PDF text
    /// </summary>
    private List<AccessoryItem> ExtractAccessories(string text, List<string> warnings)
    {
        var accessories = new List<AccessoryItem>();
        int id = 1;

        // Find the Accessories section - it starts after "Accessories" header and ends at "Gaskets" or end
        var accessoriesSectionMatch = Regex.Match(text, @"Accessories\s*Quantity.*?(?=Gaskets|$)", 
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!accessoriesSectionMatch.Success)
        {
            _logger.LogInformation("No Accessories section found in PDF");
            return accessories;
        }

        var accessoriesSection = accessoriesSectionMatch.Value;
        _logger.LogInformation($"Found Accessories section: {accessoriesSection.Length} chars");

        // Pattern to match accessory lines from the PDF format:
        // "104 pc (5 PU @ 25) 222085 Die cut cleat" or "104 pc 222085 Die cut cleat"
        // Also handles multi-line where quantity might be separate
        var linePattern = @"(\d+)\s*(?:pc|pcs)\s*(?:\([^)]*\))?\s*(\d{6})\s+(?:(Black|White|Silver)\s+)?(.+?)(?=\d+\s*(?:pc|pcs)|-- \d+ of|$)";

        var matches = Regex.Matches(accessoriesSection, linePattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        foreach (Match match in matches)
        {
            try
            {
                var accessory = new AccessoryItem
                {
                    Id = id++,
                    Amount = int.Parse(match.Groups[1].Value),
                    RefNumber = match.Groups[2].Value,
                    Finish = match.Groups[3].Value.Trim(),
                    Description = match.Groups[4].Value.Trim()
                };

                // Clean up description - remove page numbers, headers, etc.
                accessory.Description = Regex.Replace(accessory.Description, @"\s*Page:.*$", "", RegexOptions.IgnoreCase | RegexOptions.Singleline).Trim();
                accessory.Description = Regex.Replace(accessory.Description, @"\s*Logikal.*$", "", RegexOptions.IgnoreCase | RegexOptions.Singleline).Trim();
                accessory.Description = Regex.Replace(accessory.Description, @"\s*Stock List.*$", "", RegexOptions.IgnoreCase | RegexOptions.Singleline).Trim();
                accessory.Description = Regex.Replace(accessory.Description, @"\s{2,}", " ").Trim();

                // Skip if description looks like header content
                if (accessory.Description.Contains("Drawing Number") || 
                    accessory.Description.Contains("Colour Description") ||
                    string.IsNullOrWhiteSpace(accessory.Description))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(accessory.RefNumber))
                {
                    accessories.Add(accessory);
                    _logger.LogInformation($"Parsed accessory: {accessory.Amount} x {accessory.RefNumber} - {accessory.Description}");
                }
            }
            catch (Exception ex)
            {
                warnings.Add($"Failed to parse accessory line: {match.Value}. Error: {ex.Message}");
            }
        }

        // If regex didn't work well, try line-by-line parsing
        if (accessories.Count == 0)
        {
            accessories = ExtractAccessoriesLineByLine(accessoriesSection, warnings, ref id);
        }

        _logger.LogInformation($"Extracted {accessories.Count} accessories total");
        return accessories;
    }

    private List<AccessoryItem> ExtractAccessoriesLineByLine(string text, List<string> warnings, ref int id)
    {
        var accessories = new List<AccessoryItem>();
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();

            // Skip header lines and page markers
            if (line.StartsWith("Accessories", StringComparison.OrdinalIgnoreCase) || 
                line.StartsWith("Quantity") || 
                line.StartsWith("Drawing") ||
                line.StartsWith("Description") ||
                line.StartsWith("(PU)") ||
                line.StartsWith("Page:") ||
                line.StartsWith("Logikal") ||
                line.StartsWith("Stock List") ||
                line.Contains("-- ") ||
                string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            // Pattern: "104 pc (5 PU @ 25) 222085 [Color] Description" 
            // or "104 pc 222085 Description"
            var match = Regex.Match(line, @"^(\d+)\s*(?:pc|pcs)\s*(?:\([^)]*\))?\s*(\d{6})\s*(?:(Black|White|Silver)\s+)?(.*)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                try
                {
                    var accessory = new AccessoryItem
                    {
                        Id = id++,
                        Amount = int.Parse(match.Groups[1].Value),
                        RefNumber = match.Groups[2].Value,
                        Finish = match.Groups[3].Value.Trim(),
                        Description = match.Groups[4].Value.Trim()
                    };

                    if (!string.IsNullOrWhiteSpace(accessory.RefNumber) && accessory.RefNumber.Length >= 6)
                    {
                        accessories.Add(accessory);
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add($"Failed to parse accessory: {line}. Error: {ex.Message}");
                }
            }
            // Also try simpler pattern for meters: "1.26 m (4) 364512 Description"
            else
            {
                var meterMatch = Regex.Match(line, @"^([\d\.]+)\s*m\s*\(\d+\)\s*(\d{6})\s*(?:(Black|White|Silver|Without)\s+)?(.*)$", RegexOptions.IgnoreCase);
                if (meterMatch.Success)
                {
                    try
                    {
                        // Convert meters to quantity (1 = 1m)
                        var meters = decimal.Parse(meterMatch.Groups[1].Value.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture);
                        var accessory = new AccessoryItem
                        {
                            Id = id++,
                            Amount = (int)Math.Ceiling(meters),
                            RefNumber = meterMatch.Groups[2].Value,
                            Finish = meterMatch.Groups[3].Value.Trim(),
                            Description = $"{meterMatch.Groups[4].Value.Trim()} ({meters}m)"
                        };

                        if (!string.IsNullOrWhiteSpace(accessory.RefNumber))
                        {
                            accessories.Add(accessory);
                        }
                    }
                    catch { }
                }
            }
        }

        return accessories;
    }
}
