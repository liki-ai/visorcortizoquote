using System.Diagnostics;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using VisorQuotationWebApp.Hubs;
using VisorQuotationWebApp.Models;
using VisorQuotationWebApp.Services;

namespace VisorQuotationWebApp.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly ILoggerFactory _loggerFactory;
    private readonly PdfParseService _pdfParseService;
    private readonly AutomationConfig _automationConfig;
    private readonly IHubContext<AutomationHub> _hubContext;
    private readonly IConfiguration _configuration;
    private readonly VisorQuotationService _quotationService;

    // Static storage for parsed PDFs (in production, use distributed cache)
    private static readonly Dictionary<string, ParsedPdfResult> ParsedPdfs = new();
    private static readonly Dictionary<string, string> UploadedFiles = new();
    private static readonly Dictionary<string, QuotationViewModel> SavedViewModels = new();
    private static readonly Dictionary<string, decimal> CortizoTotals = new();
    private static CancellationTokenSource? _automationCts;

    public HomeController(
        ILogger<HomeController> logger,
        ILoggerFactory loggerFactory,
        PdfParseService pdfParseService,
        AutomationConfig automationConfig,
        IHubContext<AutomationHub> hubContext,
        IConfiguration configuration,
        VisorQuotationService quotationService)
    {
        _logger = logger;
        _loggerFactory = loggerFactory;
        _pdfParseService = pdfParseService;
        _automationConfig = automationConfig;
        _hubContext = hubContext;
        _configuration = configuration;
        _quotationService = quotationService;
    }

    public IActionResult Index()
    {
        var viewModel = new QuotationViewModel
        {
            Microns = _automationConfig.DefaultMicrons,
            Cif = _automationConfig.DefaultCif,
            ClientCode = _automationConfig.DefaultClientCode,
            Language = _automationConfig.DefaultLanguage
        };

        // Check if there's parsed data in session
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (!string.IsNullOrEmpty(sessionId) && ParsedPdfs.TryGetValue(sessionId, out var parsedPdf))
        {
            viewModel.ParsedPdf = parsedPdf;
            if (UploadedFiles.TryGetValue(sessionId, out var fileName))
            {
                viewModel.UploadedFileName = fileName;
            }
        }

        return View(viewModel);
    }

    [HttpPost]
    public async Task<IActionResult> UploadPdf(IFormFile pdfFile)
    {
        if (pdfFile == null || pdfFile.Length == 0)
        {
            return Json(new { success = false, message = "No file uploaded" });
        }

        if (!pdfFile.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return Json(new { success = false, message = "Please upload a PDF file" });
        }

        try
        {
            // Save the file temporarily
            var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            Directory.CreateDirectory(uploadsDir);

            var sessionId = Guid.NewGuid().ToString();
            var filePath = Path.Combine(uploadsDir, $"{sessionId}.pdf");

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await pdfFile.CopyToAsync(stream);
            }

            // Parse the PDF
            var result = _pdfParseService.ParsePdf(filePath);

            // Store in session
            HttpContext.Session.SetString("ParsedPdfId", sessionId);
            ParsedPdfs[sessionId] = result;
            UploadedFiles[sessionId] = pdfFile.FileName;

            return Json(new
            {
                success = result.Success,
                message = result.Success ? "PDF parsed successfully" : result.ErrorMessage,
                sessionId = sessionId,
                header = result.Header,
                profiles = result.Profiles,
                accessories = result.Accessories,
                hardwareItems = result.HardwareItems,
                hardwareCount = result.HardwareItems.Count,
                warnings = result.ParseWarnings,
                pdfUrl = $"/uploads/{sessionId}.pdf"
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to upload and parse PDF");
            return Json(new { success = false, message = $"Error: {ex.Message}" });
        }
    }

    [HttpPost]
    public IActionResult RunAutomation([FromForm] string username, [FromForm] string password,
        [FromForm] int microns, [FromForm] string cif, [FromForm] string clientCode, [FromForm] string language,
        [FromForm] string? clientPurchaseOrder, [FromForm] string generalFinish1, [FromForm] string generalShade1,
        [FromForm] string generalFinish2, [FromForm] string generalShade2, [FromForm] bool generateReport,
        [FromForm] bool createProforma, [FromForm] string? selectedProfileIds)
    {
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (string.IsNullOrEmpty(sessionId) || !ParsedPdfs.TryGetValue(sessionId, out var parsedPdf))
        {
            return Json(new { success = false, message = "No PDF parsed. Please upload a PDF first." });
        }

        var viewModel = new QuotationViewModel
        {
            Microns = microns,
            Cif = cif,
            ClientCode = clientCode,
            Language = language,
            ClientPurchaseOrder = clientPurchaseOrder ?? "",
            GeneralFinish1 = generalFinish1,
            GeneralShade1 = generalShade1,
            GeneralFinish2 = generalFinish2,
            GeneralShade2 = generalShade2,
            GenerateReport = generateReport,
            CreateProforma = createProforma
        };

        var credentials = new CortizoCredentials
        {
            Username = username,
            Password = password
        };

        // Parse selected profile IDs
        var selectedIds = new HashSet<int>();
        if (!string.IsNullOrEmpty(selectedProfileIds))
        {
            try
            {
                var ids = JsonSerializer.Deserialize<List<int>>(selectedProfileIds);
                if (ids != null)
                {
                    selectedIds = new HashSet<int>(ids);
                }
            }
            catch { /* Use all profiles */ }
        }

        // Mark selected profiles
        var profiles = parsedPdf.Profiles.ToList();
        foreach (var profile in profiles)
        {
            profile.IsSelected = selectedIds.Count == 0 || selectedIds.Contains(profile.Id);
        }

        // Get accessories (all selected by default)
        var accessories = parsedPdf.Accessories.ToList();

        // Cancel any previous automation and create a new token
        _automationCts?.Cancel();
        _automationCts = new CancellationTokenSource();
        var cts = _automationCts;

        // Run automation in background - capture services for closure
        var loggerFactory = _loggerFactory;
        var automationConfig = _automationConfig;
        var hubContext = _hubContext;
        
        _ = Task.Run(async () =>
        {
            await using var automationService = new CortizoAutomationService(
                loggerFactory.CreateLogger<CortizoAutomationService>(), automationConfig);

            // Wire up log events to SignalR
            automationService.OnLog += async entry =>
            {
                await hubContext.Clients.All.SendAsync("ReceiveLog", entry);
            };

            // Send log file path info
            var logFileName = Path.GetFileName(automationService.LogFilePath);
            await hubContext.Clients.All.SendAsync("ReceiveLog", new AutomationLogEntry
            {
                Timestamp = DateTime.Now,
                Level = AutomationLogLevel.Info,
                Message = $"Automation log file: {logFileName}",
                Details = $"View at /Home/Logs"
            });

            try
            {
                var result = await automationService.RunAutomationAsync(credentials, viewModel, profiles, accessories, cts.Token);

                // Add log file path to result
                result.TracePath = automationService.LogFilePath;

                await hubContext.Clients.All.SendAsync("ReceiveComplete", result);
            }
            catch (OperationCanceledException)
            {
                await hubContext.Clients.All.SendAsync("ReceiveLog", new AutomationLogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = AutomationLogLevel.Warning,
                    Message = "Automation was stopped by user"
                });

                var cancelResult = new AutomationRunResult
                {
                    Success = false,
                    ErrorMessage = "Automation stopped by user",
                    TracePath = automationService.LogFilePath
                };
                await hubContext.Clients.All.SendAsync("ReceiveComplete", cancelResult);
            }
        });

        return Json(new { success = true, message = "Automation started" });
    }

    [HttpGet]
    public IActionResult DownloadFile(string path)
    {
        if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var fileName = Path.GetFileName(path);
        var contentType = "application/octet-stream";

        if (fileName.EndsWith(".png"))
            contentType = "image/png";
        else if (fileName.EndsWith(".zip"))
            contentType = "application/zip";

        return PhysicalFile(path, contentType, fileName);
    }

    [HttpPost]
    public IActionResult StopAutomation()
    {
        if (_automationCts != null && !_automationCts.IsCancellationRequested)
        {
            _automationCts.Cancel();
            _logger.LogInformation("Automation stop requested by user");
            return Json(new { success = true, message = "Stop signal sent" });
        }
        return Json(new { success = false, message = "No automation running" });
    }

    [HttpGet]
    public IActionResult GetParsedProfiles()
    {
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (string.IsNullOrEmpty(sessionId) || !ParsedPdfs.TryGetValue(sessionId, out var parsedPdf))
        {
            return Json(new { success = false, profiles = new List<ProfileItem>() });
        }

        return Json(new { success = true, profiles = parsedPdf.Profiles });
    }

    [HttpPost]
    public IActionResult UpdateProfile([FromBody] ProfileItem profile)
    {
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (string.IsNullOrEmpty(sessionId) || !ParsedPdfs.TryGetValue(sessionId, out var parsedPdf))
        {
            return Json(new { success = false });
        }

        var existing = parsedPdf.Profiles.FirstOrDefault(p => p.Id == profile.Id);
        if (existing != null)
        {
            existing.RefNumber = profile.RefNumber;
            existing.Amount = profile.Amount;
            existing.Finish1 = profile.Finish1;
            existing.Shade1 = profile.Shade1;
            existing.Finish2 = profile.Finish2;
            existing.Shade2 = profile.Shade2;
            existing.IsSelected = profile.IsSelected;
        }

        return Json(new { success = true });
    }

    /// <summary>
    /// List all available automation log files
    /// </summary>
    [HttpGet]
    public IActionResult ListLogs()
    {
        var logsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        if (!Directory.Exists(logsFolder))
        {
            return Json(new { logs = new List<object>() });
        }

        var logs = Directory.GetFiles(logsFolder, "automation-*.log")
            .Select(f => new
            {
                name = Path.GetFileName(f),
                path = f,
                size = new FileInfo(f).Length,
                created = new FileInfo(f).CreationTime
            })
            .OrderByDescending(f => f.created)
            .Take(20)
            .ToList();

        return Json(new { logs });
    }

    /// <summary>
    /// View or download a specific log file
    /// </summary>
    [HttpGet]
    public IActionResult ViewLog(string name, bool download = false)
    {
        var logsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        var logPath = Path.Combine(logsFolder, name);

        if (!System.IO.File.Exists(logPath) || !name.StartsWith("automation-") || !name.EndsWith(".log"))
        {
            return NotFound("Log file not found");
        }

        if (download)
        {
            return PhysicalFile(logPath, "text/plain", name);
        }

        // Return content for viewing
        var content = System.IO.File.ReadAllText(logPath);
        return Content(content, "text/plain");
    }

    /// <summary>
    /// Get the latest log file content
    /// </summary>
    [HttpGet]
    public IActionResult GetLatestLog()
    {
        var logsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        if (!Directory.Exists(logsFolder))
        {
            return Json(new { success = false, message = "No logs directory" });
        }

        var latestLog = Directory.GetFiles(logsFolder, "automation-*.log")
            .OrderByDescending(f => new FileInfo(f).CreationTime)
            .FirstOrDefault();

        if (latestLog == null)
        {
            return Json(new { success = false, message = "No log files found" });
        }

        var content = System.IO.File.ReadAllText(latestLog);
        var fileName = Path.GetFileName(latestLog);

        return Json(new { success = true, name = fileName, content });
    }

    /// <summary>
    /// Logs viewer page
    /// </summary>
    public IActionResult Logs()
    {
        return View();
    }

    public IActionResult Privacy()
    {
        return View();
    }

    /// <summary>
    /// Store the Cortizo total after automation completes (called from SignalR)
    /// </summary>
    [HttpPost]
    public IActionResult SaveCortizoTotal([FromBody] SaveCortizoTotalRequest request)
    {
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (!string.IsNullOrEmpty(sessionId))
        {
            CortizoTotals[sessionId] = request.Total;
            _logger.LogInformation($"Saved Cortizo total: {request.Total} EUR for session {sessionId}");
        }
        return Json(new { success = true });
    }

    /// <summary>
    /// Generate Visor quotation PDF based on parsed data and Cortizo results
    /// </summary>
    [HttpPost]
    public IActionResult GenerateVisorQuotation([FromBody] GenerateQuotationRequest request)
    {
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (string.IsNullOrEmpty(sessionId) || !ParsedPdfs.TryGetValue(sessionId, out var parsedPdf))
        {
            return Json(new { success = false, message = "No PDF data available. Please upload and parse a PDF first." });
        }

        try
        {
            // Get Cortizo total from stored value or request
            decimal cortizoTotal = request.CortizoTotal;
            if (cortizoTotal == 0 && CortizoTotals.TryGetValue(sessionId, out var savedTotal))
            {
                cortizoTotal = savedTotal;
            }

            // Create view model from request
            var viewModel = new QuotationViewModel
            {
                GeneralFinish1 = request.GeneralFinish1 ?? "90",
                GeneralShade1 = request.GeneralShade1 ?? "P1019M",
                GeneralFinish2 = request.GeneralFinish2 ?? "90",
                GeneralShade2 = request.GeneralShade2 ?? "P1019M",
                ClientCode = request.ClientCode ?? "",
                ClientPurchaseOrder = request.ClientPurchaseOrder ?? ""
            };

            // Create the quotation order
            var order = _quotationService.CreateQuotationFromProfiles(parsedPdf, viewModel, cortizoTotal);

            // Update with client info from request
            if (!string.IsNullOrEmpty(request.ClientName))
                order.ClientName = request.ClientName;
            if (!string.IsNullOrEmpty(request.ClientAddress))
                order.ClientAddress = request.ClientAddress;
            if (!string.IsNullOrEmpty(request.ClientPhone))
                order.ClientPhone = request.ClientPhone;
            if (!string.IsNullOrEmpty(request.ClientEmail))
                order.ClientEmail = request.ClientEmail;
            if (!string.IsNullOrEmpty(request.Notes))
                order.Notes = request.Notes;
            if (request.VatRate > 0)
            {
                order.VatRate = request.VatRate;
                order.VatAmount = Math.Round(order.Subtotal * (order.VatRate / 100), 2);
                order.Total = order.Subtotal + order.VatAmount;
            }
            
            // Add unfilled items to report
            if (request.UnfilledProfiles != null)
                order.UnfilledProfiles = request.UnfilledProfiles;
            if (request.UnfilledAccessories != null)
                order.UnfilledAccessories = request.UnfilledAccessories;
            
            // Add missing calculation items for summary report
            if (request.MissingCalculationItems != null)
                order.MissingCalculationItems = request.MissingCalculationItems;

            // Save PDF to quotations folder
            var quotationsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "quotations");
            var pdfPath = _quotationService.SavePdf(order, quotationsFolder);
            var fileName = Path.GetFileName(pdfPath);

            _logger.LogInformation($"Generated Visor quotation: {order.QuotationNumber}");

            return Json(new
            {
                success = true,
                quotationNumber = order.QuotationNumber,
                downloadUrl = $"/quotations/{fileName}",
                total = order.Total,
                currency = order.Currency
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to generate Visor quotation");
            return Json(new { success = false, message = $"Error generating quotation: {ex.Message}" });
        }
    }

    /// <summary>
    /// Download a generated quotation PDF
    /// </summary>
    [HttpGet]
    public IActionResult DownloadQuotation(string fileName)
    {
        var quotationsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "quotations");
        var filePath = Path.Combine(quotationsFolder, fileName);

        if (!System.IO.File.Exists(filePath))
        {
            return NotFound("Quotation not found");
        }

        return PhysicalFile(filePath, "application/pdf", fileName);
    }

    /// <summary>
    /// List all generated quotations
    /// </summary>
    [HttpGet]
    public IActionResult ListQuotations()
    {
        var quotationsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "quotations");
        if (!Directory.Exists(quotationsFolder))
        {
            return Json(new { quotations = new List<object>() });
        }

        var quotations = Directory.GetFiles(quotationsFolder, "Visor_Quotation_*.pdf")
            .Select(f => new
            {
                name = Path.GetFileName(f),
                created = new FileInfo(f).CreationTime,
                size = new FileInfo(f).Length
            })
            .OrderByDescending(f => f.created)
            .Take(50)
            .ToList();

        return Json(new { quotations });
    }

    #region Excel Price Calculation Endpoints

    /// <summary>
    /// Upload and load profile prices Excel file
    /// </summary>
    [HttpPost]
    public async Task<IActionResult> UploadProfilePrices(IFormFile excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            return Json(new { success = false, message = "No file uploaded" });
        }

        if (!excelFile.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !excelFile.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
        {
            return Json(new { success = false, message = "Please upload an Excel file (.xlsx or .xls)" });
        }

        try
        {
            var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", "prices");
            Directory.CreateDirectory(uploadsDir);
            
            var filePath = Path.Combine(uploadsDir, $"profiles_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
            var result = await excelService.LoadProfilePricesAsync(filePath);
            var status = excelService.GetStatus();

            return Json(new
            {
                success = result,
                message = result ? $"Loaded {status.ProfilePricesCount} profile prices" : "Failed to load profile prices",
                profilesCount = status.ProfilePricesCount,
                fileName = excelFile.FileName
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to upload profile prices Excel");
            return Json(new { success = false, message = $"Error: {ex.Message}" });
        }
    }

    /// <summary>
    /// Upload and load accessory prices Excel file
    /// </summary>
    [HttpPost]
    public async Task<IActionResult> UploadAccessoryPrices(IFormFile excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
        {
            return Json(new { success = false, message = "No file uploaded" });
        }

        try
        {
            var uploadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", "prices");
            Directory.CreateDirectory(uploadsDir);
            
            var filePath = Path.Combine(uploadsDir, $"accessories_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
            var result = await excelService.LoadAccessoryPricesAsync(filePath);
            var status = excelService.GetStatus();

            return Json(new
            {
                success = result,
                message = result ? $"Loaded {status.AccessoryPricesCount} accessory prices" : "Failed to load accessory prices",
                accessoriesCount = status.AccessoryPricesCount,
                fileName = excelFile.FileName
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to upload accessory prices Excel");
            return Json(new { success = false, message = $"Error: {ex.Message}" });
        }
    }

    /// <summary>
    /// Load Excel files from specified paths (for known/default locations)
    /// </summary>
    [HttpPost]
    public async Task<IActionResult> LoadExcelPrices([FromBody] LoadExcelPricesRequest request)
    {
        var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
        var results = new List<string>();

        if (!string.IsNullOrEmpty(request.ProfilePricesPath))
        {
            var profileResult = await excelService.LoadProfilePricesAsync(request.ProfilePricesPath);
            results.Add(profileResult ? $"Loaded profile prices from {Path.GetFileName(request.ProfilePricesPath)}" 
                : $"Failed to load profile prices");
        }

        if (!string.IsNullOrEmpty(request.AccessoryPricesPath))
        {
            var accResult = await excelService.LoadAccessoryPricesAsync(request.AccessoryPricesPath);
            results.Add(accResult ? $"Loaded accessory prices from {Path.GetFileName(request.AccessoryPricesPath)}" 
                : $"Failed to load accessory prices");
        }

        if (!string.IsNullOrEmpty(request.ColorDataPath))
        {
            var colorResult = await excelService.LoadColorDataAsync(request.ColorDataPath);
            results.Add(colorResult ? $"Loaded color data from {Path.GetFileName(request.ColorDataPath)}" 
                : $"Failed to load color data");
        }

        var status = excelService.GetStatus();
        return Json(new
        {
            success = status.ProfilePricesLoaded || status.AccessoryPricesLoaded,
            messages = results,
            status
        });
    }

    /// <summary>
    /// Calculate totals from PDF data using Excel prices (no Cortizo automation needed)
    /// </summary>
    [HttpPost]
    public IActionResult CalculateTotalsFromExcel([FromBody] CalculateTotalsRequest request)
    {
        var sessionId = HttpContext.Session.GetString("ParsedPdfId");
        if (string.IsNullOrEmpty(sessionId) || !ParsedPdfs.TryGetValue(sessionId, out var parsedPdf))
        {
            return Json(new { success = false, message = "No PDF data available. Please upload and parse a PDF first." });
        }

        var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
        var status = excelService.GetStatus();

        if (!status.ProfilePricesLoaded)
        {
            return Json(new { success = false, message = "Profile prices not loaded. Please upload the price Excel file first." });
        }

        try
        {
            // Mark selected profiles based on request
            var profiles = parsedPdf.Profiles.ToList();
            if (request.SelectedProfileIds != null && request.SelectedProfileIds.Count > 0)
            {
                var selectedIds = new HashSet<int>(request.SelectedProfileIds);
                foreach (var profile in profiles)
                {
                    profile.IsSelected = selectedIds.Contains(profile.Id);
                }
            }

            // Mark selected accessories
            var accessories = parsedPdf.Accessories.ToList();
            if (request.SelectedAccessoryIds != null && request.SelectedAccessoryIds.Count > 0)
            {
                var selectedIds = new HashSet<int>(request.SelectedAccessoryIds);
                foreach (var acc in accessories)
                {
                    acc.IsSelected = selectedIds.Contains(acc.Id);
                }
            }

            var result = excelService.CalculateTotals(
                profiles, 
                accessories,
                request.CustomWithBreakPrice,
                request.CustomWithoutBreakPrice,
                request.AccessoryDiscount
            );

            // Build missing calculation items with full details
            var missingItems = new List<MissingCalculationItem>();
            
            foreach (var refNum in result.UnmatchedProfiles)
            {
                var profile = profiles.FirstOrDefault(p => p.RefNumber == refNum);
                missingItems.Add(new MissingCalculationItem
                {
                    RefNumber = refNum,
                    Quantity = profile?.Amount ?? 0,
                    Description = profile?.Description ?? "",
                    Finish = profile?.RawColour ?? "",
                    Category = "Profile",
                    Reason = "Not found in profile price list"
                });
            }
            
            foreach (var refNum in result.UnmatchedAccessories)
            {
                var acc = accessories.FirstOrDefault(a => a.RefNumber == refNum);
                missingItems.Add(new MissingCalculationItem
                {
                    RefNumber = refNum,
                    Quantity = acc?.Amount ?? 0,
                    Description = acc?.Description ?? "",
                    Finish = acc?.Finish ?? "",
                    Category = acc?.Source ?? "Accessory",
                    Reason = "Not found in accessory price list"
                });
            }

            return Json(new
            {
                success = true,
                calculationDate = result.CalculationDate,
                profilesTotal = result.ProfilesTotal,
                accessoriesTotal = result.AccessoriesTotal,
                grandTotal = result.GrandTotal,
                profilesCalculated = result.ProfilesCalculated,
                accessoriesCalculated = result.AccessoriesCalculated,
                unmatchedProfiles = result.UnmatchedProfiles,
                unmatchedAccessories = result.UnmatchedAccessories,
                missingCalculationItems = missingItems,
                warnings = result.Warnings,
                summary = new
                {
                    profilesMatched = result.TotalProfilesMatched,
                    profilesUnmatched = result.UnmatchedProfiles.Count,
                    accessoriesMatched = result.TotalAccessoriesMatched,
                    accessoriesUnmatched = result.UnmatchedAccessories.Count,
                    totalMissing = missingItems.Count
                }
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to calculate totals from Excel");
            return Json(new { success = false, message = $"Calculation error: {ex.Message}" });
        }
    }

    /// <summary>
    /// Get current Excel price data status
    /// </summary>
    [HttpGet]
    public IActionResult GetExcelPriceStatus()
    {
        var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
        return Json(excelService.GetStatus());
    }

    /// <summary>
    /// Look up a single profile price
    /// </summary>
    [HttpGet]
    public IActionResult LookupProfilePrice(string reference)
    {
        var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
        var price = excelService.GetProfilePrice(reference);
        
        if (price == null)
        {
            return Json(new { found = false, reference });
        }

        return Json(new { found = true, data = price });
    }

    /// <summary>
    /// Look up a single accessory price
    /// </summary>
    [HttpGet]
    public IActionResult LookupAccessoryPrice(string reference)
    {
        var excelService = HttpContext.RequestServices.GetRequiredService<ExcelPriceService>();
        var price = excelService.GetAccessoryPrice(reference);
        
        if (price == null)
        {
            return Json(new { found = false, reference });
        }

        return Json(new { found = true, data = price });
    }

    #endregion

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}

/// <summary>
/// Request to load Excel price files from paths
/// </summary>
public class LoadExcelPricesRequest
{
    public string? ProfilePricesPath { get; set; }
    public string? AccessoryPricesPath { get; set; }
    public string? ColorDataPath { get; set; }
}

/// <summary>
/// Request to calculate totals from Excel
/// </summary>
public class CalculateTotalsRequest
{
    public List<int>? SelectedProfileIds { get; set; }
    public List<int>? SelectedAccessoryIds { get; set; }
    public decimal? CustomWithBreakPrice { get; set; }
    public decimal? CustomWithoutBreakPrice { get; set; }
    public decimal? AccessoryDiscount { get; set; }
}

/// <summary>
/// Request to save Cortizo total
/// </summary>
public class SaveCortizoTotalRequest
{
    public decimal Total { get; set; }
}

/// <summary>
/// Request to generate Visor quotation
/// </summary>
public class GenerateQuotationRequest
{
    public decimal CortizoTotal { get; set; }
    public string? ClientName { get; set; }
    public string? ClientAddress { get; set; }
    public string? ClientPhone { get; set; }
    public string? ClientEmail { get; set; }
    public string? ClientCode { get; set; }
    public string? ClientPurchaseOrder { get; set; }
    public string? GeneralFinish1 { get; set; }
    public string? GeneralShade1 { get; set; }
    public string? GeneralFinish2 { get; set; }
    public string? GeneralShade2 { get; set; }
    public decimal VatRate { get; set; } = 20;
    public string? Notes { get; set; }
    
    // Unfilled items from automation
    public List<UnfilledItem>? UnfilledProfiles { get; set; }
    public List<UnfilledItem>? UnfilledAccessories { get; set; }
    
    // Missing items from price calculation (for manual fill in summary report)
    public List<MissingCalculationItem>? MissingCalculationItems { get; set; }
}
