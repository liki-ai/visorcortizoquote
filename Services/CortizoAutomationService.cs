using Microsoft.Playwright;
using System.Text;
using VisorQuotationWebApp.Models;

namespace VisorQuotationWebApp.Services;

/// <summary>
/// Service for automating Cortizo Center quotation entry using Playwright
/// </summary>
public class CortizoAutomationService : IAsyncDisposable
{
    private readonly ILogger<CortizoAutomationService> _logger;
    private readonly AutomationConfig _config;
    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private IBrowserContext? _context;
    private IPage? _page;
    
    // File logging
    private readonly string _logFilePath;
    private readonly StringBuilder _logBuffer = new();
    
    // Event for real-time logging
    public event Action<AutomationLogEntry>? OnLog;

    public CortizoAutomationService(ILogger<CortizoAutomationService> logger, AutomationConfig config)
    {
        _logger = logger;
        _config = config;
        
        // Create log file path
        var logsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        Directory.CreateDirectory(logsFolder);
        _logFilePath = Path.Combine(logsFolder, $"automation-{DateTime.Now:yyyyMMdd-HHmmss}.log");
    }
    
    /// <summary>
    /// Get the path to the current log file
    /// </summary>
    public string LogFilePath => _logFilePath;

    /// <summary>
    /// Run the automation to fill in the Cortizo quotation
    /// </summary>
    public async Task<AutomationRunResult> RunAutomationAsync(
        CortizoCredentials credentials,
        QuotationViewModel viewModel,
        List<ProfileItem> profiles,
        List<AccessoryItem>? accessories = null,
        CancellationToken cancellationToken = default)
    {
        var selectedAccessories = accessories?.Where(a => a.IsSelected).ToList() ?? new List<AccessoryItem>();
        
        var result = new AutomationRunResult
        {
            TotalItems = profiles.Count(p => p.IsSelected) + selectedAccessories.Count
        };

        var tracePath = Path.Combine(Path.GetTempPath(), $"cortizo-trace-{DateTime.Now:yyyyMMdd-HHmmss}.zip");
        var screenshotPath = Path.Combine(Path.GetTempPath(), $"cortizo-screenshot-{DateTime.Now:yyyyMMdd-HHmmss}.png");

        try
        {
            // Initialize Playwright
            Log(result, AutomationLogLevel.Info, "Initializing browser automation...");
            await InitializeBrowserAsync(tracePath);

            // Step 1: Login
            Log(result, AutomationLogLevel.Info, "Navigating to Cortizo Center login page...");
            Log(result, AutomationLogLevel.Info, $"Target URL: {_config.BaseUrl}/Login.aspx");
            await _page!.GotoAsync($"{_config.BaseUrl}/Login.aspx", new PageGotoOptions
            {
                WaitUntil = WaitUntilState.NetworkIdle,
                Timeout = _config.TimeoutMs
            });
            
            // Log page state after navigation
            await LogPageStateAsync("After loading login page");

            Log(result, AutomationLogLevel.Info, "Logging in...");
            await LoginAsync(credentials);

            // Wait for login to complete
            try
            {
                await _page.WaitForURLAsync(url => !url.Contains("Login.aspx"), new PageWaitForURLOptions
                {
                    Timeout = _config.TimeoutMs
                });
                Log(result, AutomationLogLevel.Success, "Login successful!");
            }
            catch (TimeoutException)
            {
                Log(result, AutomationLogLevel.Warning, "Login redirect timeout - checking if already logged in...");
                await LogPageStateAsync("After login timeout");
            }
            
            // Log state after login
            await LogPageStateAsync("After login");

            cancellationToken.ThrowIfCancellationRequested();

            // Step 2: Navigate to Quotations
            Log(result, AutomationLogLevel.Info, "Navigating to Quotations / Online Orders...");
            await NavigateToQuotationsAsync();
            await LogPageStateAsync("After navigating to quotations");
            await LogAvailableSelectorsAsync();

            cancellationToken.ThrowIfCancellationRequested();

            // Step 3: Create new valuation
            Log(result, AutomationLogLevel.Info, "Creating new valuation...");
            await CreateNewValuationAsync();
            await LogPageStateAsync("After creating new valuation");

            // Step 4: Set header fields
            Log(result, AutomationLogLevel.Info, "Setting header fields...");
            await SetHeaderFieldsAsync(viewModel);
            await LogPageStateAsync("After setting header fields");

            // Step 4.5: Set customized prices (static values)
            Log(result, AutomationLogLevel.Info, "Setting customized prices...");
            await SetCustomizedPricesAsync();
            await LogPageStateAsync("After setting customized prices");

            cancellationToken.ThrowIfCancellationRequested();

            // Step 5: Ensure enough rows exist
            var selectedProfiles = profiles.Where(p => p.IsSelected).ToList();
            Log(result, AutomationLogLevel.Info, $"Ensuring {selectedProfiles.Count} rows are available in the grid...");
            await EnsureEnoughRowsAsync(selectedProfiles.Count);

            // Step 6: Fill profiles using fast batch method
            Log(result, AutomationLogLevel.Info, $"Filling {selectedProfiles.Count} profile items (fast batch mode)...");
            
            try
            {
                await FillAllProfileRowsFastAsync(selectedProfiles, result, cancellationToken);
                result.SuccessfulItems = selectedProfiles.Count;
                Log(result, AutomationLogLevel.Success, $"All {selectedProfiles.Count} profile rows filled successfully");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                result.FailedItems = selectedProfiles.Count;
                Log(result, AutomationLogLevel.Error, $"Profile batch fill failed: {ex.Message}", ex.ToString());
            }

            // Step 6.5: Fill accessories if any
            if (selectedAccessories.Count > 0)
            {
                Log(result, AutomationLogLevel.Info, $"Filling {selectedAccessories.Count} accessories...");
                try
                {
                    await FillAccessoriesAsync(selectedAccessories, result, cancellationToken);
                    result.SuccessfulItems += selectedAccessories.Count;
                    Log(result, AutomationLogLevel.Success, $"All {selectedAccessories.Count} accessories filled successfully");
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    result.FailedItems += selectedAccessories.Count;
                    Log(result, AutomationLogLevel.Error, $"Accessories fill failed: {ex.Message}", ex.ToString());
                }
            }

            // ====================================================================
            // FINAL VERIFICATION: Wait for all AJAX to settle, then check ALL rows
            // ====================================================================
            cancellationToken.ThrowIfCancellationRequested();
            
            Log(result, AutomationLogLevel.Info, "All items filled. Waiting 10 seconds for all AJAX calculations to settle...");
            await Task.Delay(10000, cancellationToken);
            
            Log(result, AutomationLogLevel.Info, "[FINAL CHECK] Starting comprehensive verification of all rows...");
            
            // Final check: Profiles
            var selectedProfiles2 = profiles.Where(p => p.IsSelected).ToList();
            for (int i = 0; i < selectedProfiles2.Count; i++)
            {
                var profile = selectedProfiles2[i];
                var rowNum = (i + 1).ToString("D4");
                var checkScript = $@"
                    (function() {{
                        var importeEl = document.getElementById('txtImporte_{rowNum}');
                        var descEl = document.getElementById('txtDescripcion_{rowNum}');
                        return {{
                            amount: importeEl ? importeEl.value : '',
                            desc: descEl ? descEl.value : ''
                        }};
                    }})();
                ";
                var checkResult = await _page.EvaluateAsync<Dictionary<string, string>>(checkScript);
                var amount = checkResult.GetValueOrDefault("amount", "");
                
                if (string.IsNullOrWhiteSpace(amount))
                {
                    result.UnfilledProfiles.Add(new UnfilledItem
                    {
                        RowNumber = rowNum,
                        RefNumber = profile.RefNumber,
                        Amount = profile.Amount,
                        Description = profile.Description,
                        Reason = "Amount not calculated - needs manual review"
                    });
                }
            }
            
            Log(result, AutomationLogLevel.Info, 
                $"[FINAL CHECK] Profiles: {selectedProfiles2.Count - result.UnfilledProfiles.Count}/{selectedProfiles2.Count} have amounts");
            
            // Final check: Accessories
            if (selectedAccessories.Count > 0)
            {
                for (int i = 0; i < selectedAccessories.Count; i++)
                {
                    var acc = selectedAccessories[i];
                    var rowNum = (i + 1).ToString("D4");
                    var checkScript = $@"
                        (function() {{
                            var importeEl = document.getElementById('txtImporteAcc_{rowNum}');
                            var descEl = document.getElementById('txtDescripcionAcc_{rowNum}');
                            var priceEl = document.getElementById('txtPrecioAcc_{rowNum}');
                            return {{
                                amount: importeEl ? importeEl.value : '',
                                desc: descEl ? descEl.value : '',
                                price: priceEl ? priceEl.value : ''
                            }};
                        }})();
                    ";
                    var checkResult = await _page.EvaluateAsync<Dictionary<string, string>>(checkScript);
                    var amount = checkResult.GetValueOrDefault("amount", "");
                    var pageDesc = checkResult.GetValueOrDefault("desc", "");
                    var pagePrice = checkResult.GetValueOrDefault("price", "");
                    
                    if (string.IsNullOrWhiteSpace(amount))
                    {
                        result.UnfilledAccessories.Add(new UnfilledItem
                        {
                            RowNumber = rowNum,
                            RefNumber = acc.RefNumber,
                            Amount = acc.Amount,
                            Description = acc.Description,
                            Reason = $"Amount not calculated (price={pagePrice}, desc={pageDesc})"
                        });
                    }
                }
                
                Log(result, AutomationLogLevel.Info, 
                    $"[FINAL CHECK] Accessories: {selectedAccessories.Count - result.UnfilledAccessories.Count}/{selectedAccessories.Count} have amounts");
            }

            // Step 7: Optional actions
            if (viewModel.GenerateReport)
            {
                Log(result, AutomationLogLevel.Info, "Generating report...");
                await ClickButtonAsync("GENERATE REPORT");
                await Task.Delay(2000, cancellationToken);
            }

            if (viewModel.CreateProforma)
            {
                Log(result, AutomationLogLevel.Info, "Creating proforma...");
                await ClickButtonAsync("CREATE A PROFORMA");
                await Task.Delay(2000, cancellationToken);
            }

            // Take final screenshot
            await _page.ScreenshotAsync(new PageScreenshotOptions { Path = screenshotPath, FullPage = true });
            result.ScreenshotPath = screenshotPath;

            // Capture the Cortizo total (ESTIMATE TOTAL)
            result.CortizoTotal = await GetCortizoTotalAsync();
            Log(result, AutomationLogLevel.Info, $"Cortizo ESTIMATE TOTAL: {result.CortizoTotal} EUR");

            // Log unfilled items summary
            if (result.UnfilledProfiles.Count > 0)
            {
                Log(result, AutomationLogLevel.Warning, 
                    $"UNFILLED PROFILES: {result.UnfilledProfiles.Count} items need manual review");
                foreach (var item in result.UnfilledProfiles)
                {
                    Log(result, AutomationLogLevel.Warning, 
                        $"  - Row {item.RowNumber}: REF {item.RefNumber} x {item.Amount} - {item.Reason}");
                }
            }
            else
            {
                Log(result, AutomationLogLevel.Success, "All profiles have calculated amounts!");
            }
            
            if (result.UnfilledAccessories.Count > 0)
            {
                Log(result, AutomationLogLevel.Warning, 
                    $"UNFILLED ACCESSORIES: {result.UnfilledAccessories.Count} items need manual review");
                foreach (var item in result.UnfilledAccessories)
                {
                    Log(result, AutomationLogLevel.Warning, 
                        $"  - Row {item.RowNumber}: REF {item.RefNumber} x {item.Amount} - {item.Reason}");
                }
            }
            else if (selectedAccessories.Count > 0)
            {
                Log(result, AutomationLogLevel.Success, "All accessories have calculated amounts!");
            }

            result.Success = result.UnfilledProfiles.Count == 0 && result.UnfilledAccessories.Count == 0;
            Log(result, result.Success ? AutomationLogLevel.Success : AutomationLogLevel.Warning,
                $"Automation completed. {result.SuccessfulItems}/{result.TotalItems} items processed. " +
                $"Browser remains open for manual review - close application to close browser.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Automation failed");
            result.Success = false;
            result.ErrorMessage = ex.Message;
            Log(result, AutomationLogLevel.Error, $"Automation failed: {ex.Message}", ex.ToString());

            // Try to take error screenshot
            try
            {
                if (_page != null)
                {
                    await _page.ScreenshotAsync(new PageScreenshotOptions { Path = screenshotPath, FullPage = true });
                    result.ScreenshotPath = screenshotPath;
                }
            }
            catch { /* Ignore screenshot errors */ }
        }
        finally
        {
            // Stop tracing but DON'T dispose - keep browser open for manual review
            if (_context != null)
            {
                try
                {
                    await _context.Tracing.StopAsync(new TracingStopOptions { Path = tracePath });
                    result.TracePath = tracePath;
                }
                catch { /* Ignore tracing errors */ }
            }

            // NOTE: Browser is intentionally kept open for manual data entry
            // It will be closed when the application stops or DisposeAsync is called manually
            Log(result, AutomationLogLevel.Info, "Browser kept open for manual review. Close application to close browser.");
        }

        return result;
    }

    private async Task InitializeBrowserAsync(string tracePath)
    {
        _playwright = await Playwright.CreateAsync();
        _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = _config.Headless,
            SlowMo = _config.Headless ? 0 : 100 // Slow down in headed mode for visibility
        });

        _context = await _browser.NewContextAsync(new BrowserNewContextOptions
        {
            ViewportSize = new ViewportSize { Width = 1920, Height = 1080 }
        });

        // Start tracing for debugging
        await _context.Tracing.StartAsync(new TracingStartOptions
        {
            Screenshots = true,
            Snapshots = true,
            Sources = true
        });

        _page = await _context.NewPageAsync();
        _page.SetDefaultTimeout(_config.TimeoutMs);
    }

    private async Task LoginAsync(CortizoCredentials credentials)
    {
        WriteToLogFile(AutomationLogLevel.Info, "[LOGIN] Starting login process...");
        
        // Cortizo login page uses Spanish labels: USUARIO / CONTRASEÑA
        // Wait for the login form to be visible
        WriteToLogFile(AutomationLogLevel.Info, "[LOGIN] Waiting for login form...");
        await _page!.WaitForSelectorAsync("input[type='text'], input[type='password']");
        
        // Fill username - the first text input on the page
        var usernameInput = await _page.QuerySelectorAsync("input[type='text']");
        if (usernameInput != null)
        {
            WriteToLogFile(AutomationLogLevel.Info, $"[LOGIN] Filling username: {credentials.Username}");
            await usernameInput.FillAsync(credentials.Username);
        }
        else
        {
            WriteToLogFile(AutomationLogLevel.Warning, "[LOGIN] Username input not found!");
        }
        
        // Fill password
        var passwordInput = await _page.QuerySelectorAsync("input[type='password']");
        if (passwordInput != null)
        {
            WriteToLogFile(AutomationLogLevel.Info, "[LOGIN] Filling password: ********");
            await passwordInput.FillAsync(credentials.Password);
        }
        else
        {
            WriteToLogFile(AutomationLogLevel.Warning, "[LOGIN] Password input not found!");
        }

        // Click login button - look for ACCEDER or submit button
        var loginSelectors = new[]
        {
            "input[type='submit']",
            "button[type='submit']",
            "text=ACCEDER",
            "text=ACCESS",
            ".loginbotones",
            "[onclick*='Login']"
        };

        WriteToLogFile(AutomationLogLevel.Info, "[LOGIN] Looking for login button...");
        bool buttonClicked = false;
        foreach (var selector in loginSelectors)
        {
            try
            {
                var btn = await _page.QuerySelectorAsync(selector);
                if (btn != null)
                {
                    WriteToLogFile(AutomationLogLevel.Info, $"[LOGIN] Found button with selector: {selector}");
                    await btn.ClickAsync();
                    buttonClicked = true;
                    WriteToLogFile(AutomationLogLevel.Info, "[LOGIN] Clicked login button");
                    break;
                }
            }
            catch (Exception ex)
            {
                WriteToLogFile(AutomationLogLevel.Info, $"[LOGIN] Selector '{selector}' failed: {ex.Message}");
            }
        }
        
        if (!buttonClicked)
        {
            WriteToLogFile(AutomationLogLevel.Warning, "[LOGIN] No login button found!");
        }
        
        // Wait a moment for login to process
        await Task.Delay(2000);
        WriteToLogFile(AutomationLogLevel.Info, $"[LOGIN] Current URL after login attempt: {_page.Url}");
        FlushLogBuffer();
    }

    private async Task NavigateToQuotationsAsync()
    {
        WriteToLogFile(AutomationLogLevel.Info, "[NAVIGATE] Starting navigation to Quotations...");
        WriteToLogFile(AutomationLogLevel.Info, $"[NAVIGATE] Current URL: {_page!.Url}");
        
        // The quotations link has id="ctl00_ico7" and class="ico7"
        // It uses a redirect through ControlRedirect.ashx
        var selectors = new[]
        {
            "#ctl00_ico7",       // Direct ID from HTML
            "a.ico7",            // Class selector
            "a[href*='ControlRedirect'][href*='CgAAAB']",  // Redirect URL pattern
            "text=QUOTATIONS",
            "text=ONLINE ORDERS",
            "[href*='Valoraciones']",
            "//a[contains(@class, 'ico7')]"
        };

        foreach (var selector in selectors)
        {
            try
            {
                WriteToLogFile(AutomationLogLevel.Info, $"[NAVIGATE] Trying selector: {selector}");
                var element = await _page.QuerySelectorAsync(selector);
                if (element != null)
                {
                    var isVisible = await element.IsVisibleAsync();
                    WriteToLogFile(AutomationLogLevel.Info, $"[NAVIGATE] Found element! Visible: {isVisible}");
                    
                    if (isVisible)
                    {
                        await element.ClickAsync();
                        WriteToLogFile(AutomationLogLevel.Info, "[NAVIGATE] Clicked element, waiting for navigation...");
                        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                        await Task.Delay(2000); // Wait for page to fully load
                        WriteToLogFile(AutomationLogLevel.Info, $"[NAVIGATE] Navigation complete. New URL: {_page.Url}");
                        FlushLogBuffer();
                        return;
                    }
                }
                else
                {
                    WriteToLogFile(AutomationLogLevel.Info, $"[NAVIGATE] Selector not found: {selector}");
                }
            }
            catch (Exception ex)
            {
                WriteToLogFile(AutomationLogLevel.Info, $"[NAVIGATE] Selector '{selector}' error: {ex.Message}");
            }
        }

        // If menu click didn't work, try direct navigation to Valoraciones.aspx
        // Note: The actual URL may require session parameters
        WriteToLogFile(AutomationLogLevel.Warning, "[NAVIGATE] Could not find quotations menu, page may already be on quotations");
        _logger.LogWarning("Could not find quotations menu, page may already be on quotations");
        FlushLogBuffer();
    }

    private async Task CreateNewValuationAsync()
    {
        // Look for "NEW VALUATION" or "NUEVA VALORACIÓN" button
        var selectors = new[]
        {
            "text=NEW VALUATION",
            "text=NUEVA VALORACIÓN",
            "text=NUEVA VALORACION",
            "text=New Valuation",
            "[value='NEW VALUATION']",
            "[value*='NUEVA']",
            "input[value*='NEW']",
            "button:has-text('NEW')",
            "button:has-text('NUEVA')",
            ".btn-new-valuation",
            "#btnNewValuation",
            "a:has-text('NEW VALUATION')",
            ".botonesvaloraciones:has-text('NEW')"
        };

        foreach (var selector in selectors)
        {
            try
            {
                var element = await _page!.QuerySelectorAsync(selector);
                if (element != null)
                {
                    await element.ClickAsync();
                    await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    await Task.Delay(1000); // Wait for new form to load
                    return;
                }
            }
            catch
            {
                // Try next selector
            }
        }

        // Check if we're already on a valuation form with the profiles grid
        var profilesGrid = await _page!.QuerySelectorAsync("#gvPerfiles");
        if (profilesGrid != null)
        {
            _logger.LogInformation("Profiles grid already visible, continuing without creating new valuation");
        }
        else
        {
            _logger.LogWarning("Could not find NEW VALUATION button, continuing anyway...");
        }
    }

    /// <summary>
    /// Ensure enough rows exist in the profiles grid
    /// </summary>
    private async Task EnsureEnoughRowsAsync(int requiredRows)
    {
        // Check how many rows currently exist
        var rows = await _page!.QuerySelectorAllAsync("#gvPerfiles tr");
        var currentRows = rows.Count;

        _logger.LogInformation($"Current profile rows: {currentRows}, required: {requiredRows}");

        // The page adds rows dynamically via dgPerfilesAddRow() JavaScript function
        while (currentRows < requiredRows)
        {
            // Call the JavaScript function to add a row
            try
            {
                await _page.EvaluateAsync("dgPerfilesAddRow()");
                await Task.Delay(100); // Reduced delay
                currentRows++;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to add profile row via JavaScript");
                break;
            }
        }
    }

    /// <summary>
    /// Fill all profile rows - uses header values as fallback for finish/shade
    /// </summary>
    private async Task FillAllProfileRowsFastAsync(List<ProfileItem> profiles, AutomationRunResult result, CancellationToken cancellationToken = default)
    {
        WriteToLogFile(AutomationLogLevel.Info, $"[FILL] Starting fill for {profiles.Count} profiles...");
        
        // First, get the header (General Color) values to use as fallback
        var headerValuesScript = @"
            (function() {
                return {
                    finish1: document.getElementById('ddlAcabado_1_ColorGeneral')?.value || '',
                    shade1: document.getElementById('ddlMatiz_1_ColorGeneral')?.value || '',
                    finish2: document.getElementById('ddlAcabado_2_ColorGeneral')?.value || '',
                    shade2: document.getElementById('ddlMatiz_2_ColorGeneral')?.value || ''
                };
            })();
        ";
        var headerValues = await _page!.EvaluateAsync<Dictionary<string, string>>(headerValuesScript);
        
        var headerFinish1 = headerValues.GetValueOrDefault("finish1", "90"); // Default to SPECIAL 1 POWDER COATING
        var headerShade1 = headerValues.GetValueOrDefault("shade1", "");
        var headerFinish2 = headerValues.GetValueOrDefault("finish2", "90");
        var headerShade2 = headerValues.GetValueOrDefault("shade2", "");
        
        WriteToLogFile(AutomationLogLevel.Info, $"[FILL] Header values - Finish1: {headerFinish1}, Shade1: {headerShade1}, Finish2: {headerFinish2}, Shade2: {headerShade2}");
        
        // Process each row sequentially to ensure proper validation and price calculation
        for (int i = 0; i < profiles.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var profile = profiles[i];
            var rowNum = (i + 1).ToString("D4");
            
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Processing: REF={profile.RefNumber}, AMT={profile.Amount}, DESC={profile.Description}");
            
            // Step 1: Set reference and trigger validation (this fetches profile data from server)
            var setRefScript = $@"
                (function() {{
                    const refInput = document.getElementById('txtReferencia_{rowNum}');
                    if (refInput) {{
                        refInput.value = '{profile.RefNumber}';
                        if (typeof ValidarFormatoDatosPerfil === 'function') {{
                            ValidarFormatoDatosPerfil(refInput, false);
                        }}
                        return true;
                    }}
                    return false;
                }})();
            ";
            await _page.EvaluateAsync(setRefScript);
            
            // Wait for the reference validation AJAX to complete
            await Task.Delay(350);
            
            // Step 2: Set amount
            var setAmtScript = $@"
                (function() {{
                    const amtInput = document.getElementById('txtCantidad_{rowNum}');
                    if (amtInput) {{
                        amtInput.value = '{profile.Amount}';
                        return true;
                    }}
                    return false;
                }})();
            ";
            await _page.EvaluateAsync(setAmtScript);
            
            // Step 3: Set Finish 1 - use profile value or fall back to header
            var finish1Value = MapFinishToValue(profile.Finish1 ?? "");
            if (string.IsNullOrEmpty(finish1Value))
            {
                finish1Value = headerFinish1;
                WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Using header Finish1: {finish1Value}");
            }
            
            var setFinish1Script = $@"
                (function() {{
                    const finish1Select = document.getElementById('ddlAcabado1_{rowNum}');
                    if (finish1Select) {{
                        finish1Select.value = '{finish1Value}';
                        if (typeof RellenarComboMatices === 'function') {{
                            RellenarComboMatices(finish1Select);
                        }}
                        return true;
                    }}
                    return false;
                }})();
            ";
            await _page.EvaluateAsync(setFinish1Script);
            await Task.Delay(250); // Wait for shade options to populate
            
            // Step 4: Set Shade 1 - use profile value or fall back to header, or first available option
            var shade1Value = profile.Shade1 ?? "";
            if (string.IsNullOrEmpty(shade1Value))
            {
                shade1Value = headerShade1;
                WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Using header Shade1: {shade1Value}");
            }
            
            var setShade1Script = $@"
                (function() {{
                    const shade1Select = document.getElementById('ddlMatiz1_{rowNum}');
                    if (shade1Select && shade1Select.options.length > 0) {{
                        // Try to find by value first
                        let found = false;
                        for (let opt of shade1Select.options) {{
                            if (opt.value === '{shade1Value}' || opt.text.includes('{shade1Value}')) {{
                                shade1Select.value = opt.value;
                                found = true;
                                break;
                            }}
                        }}
                        // If not found, select first non-empty option
                        if (!found && shade1Select.options.length > 1) {{
                            shade1Select.selectedIndex = 1;
                        }}
                        return shade1Select.value;
                    }}
                    return '';
                }})();
            ";
            await _page.EvaluateAsync(setShade1Script);
            
            // Step 5: Set Finish 2 - use profile value or fall back to header
            var finish2Value = MapFinishToValue(profile.Finish2 ?? "");
            if (string.IsNullOrEmpty(finish2Value))
            {
                finish2Value = headerFinish2;
                WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Using header Finish2: {finish2Value}");
            }
            
            var setFinish2Script = $@"
                (function() {{
                    const finish2Select = document.getElementById('ddlAcabado2_{rowNum}');
                    if (finish2Select) {{
                        finish2Select.value = '{finish2Value}';
                        if (typeof RellenarComboMatices === 'function') {{
                            RellenarComboMatices(finish2Select);
                        }}
                        return true;
                    }}
                    return false;
                }})();
            ";
            await _page.EvaluateAsync(setFinish2Script);
            await Task.Delay(250); // Wait for shade options to populate
            
            // Step 6: Set Shade 2 - use profile value or fall back to header, or first available option
            var shade2Value = profile.Shade2 ?? "";
            if (string.IsNullOrEmpty(shade2Value))
            {
                shade2Value = headerShade2;
                WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Using header Shade2: {shade2Value}");
            }
            
            var setShade2Script = $@"
                (function() {{
                    const shade2Select = document.getElementById('ddlMatiz2_{rowNum}');
                    if (shade2Select && shade2Select.options.length > 0) {{
                        let found = false;
                        for (let opt of shade2Select.options) {{
                            if (opt.value === '{shade2Value}' || opt.text.includes('{shade2Value}')) {{
                                shade2Select.value = opt.value;
                                found = true;
                                break;
                            }}
                        }}
                        if (!found && shade2Select.options.length > 1) {{
                            shade2Select.selectedIndex = 1;
                        }}
                        return shade2Select.value;
                    }}
                    return '';
                }})();
            ";
            await _page.EvaluateAsync(setShade2Script);
            
            // Step 7: Trigger ValidarCamposLinea to calculate price
            var triggerPriceScript = $@"
                (function() {{
                    const shade1Select = document.getElementById('ddlMatiz1_{rowNum}');
                    if (shade1Select && typeof ValidarCamposLinea === 'function') {{
                        ValidarCamposLinea(shade1Select, false);
                        return true;
                    }}
                    return false;
                }})();
            ";
            await _page.EvaluateAsync(triggerPriceScript);
            
            // Wait for price calculation AJAX to complete
            await Task.Delay(500);
            
            // Verify the amount was calculated
            var checkAmountScript = $@"
                (function() {{
                    const importeInput = document.getElementById('txtImporte_{rowNum}');
                    return importeInput ? importeInput.value : '';
                }})();
            ";
            var amount = await _page.EvaluateAsync<string>(checkAmountScript);
            
            if (string.IsNullOrWhiteSpace(amount))
            {
                WriteToLogFile(AutomationLogLevel.Warning, $"[ROW {rowNum}] Amount not calculated, retrying with longer wait...");
                
                // Retry: trigger validation again with longer wait
                await _page.EvaluateAsync(triggerPriceScript);
                await Task.Delay(800);
                
                amount = await _page.EvaluateAsync<string>(checkAmountScript);
                if (string.IsNullOrWhiteSpace(amount))
                {
                    // Try triggering from shade2 as well
                    var triggerPrice2Script = $@"
                        (function() {{
                            const shade2Select = document.getElementById('ddlMatiz2_{rowNum}');
                            if (shade2Select && typeof ValidarCamposLinea === 'function') {{
                                ValidarCamposLinea(shade2Select, false);
                                return true;
                            }}
                            return false;
                        }})();
                    ";
                    await _page.EvaluateAsync(triggerPrice2Script);
                    await Task.Delay(800);
                    
                    amount = await _page.EvaluateAsync<string>(checkAmountScript);
                    if (string.IsNullOrWhiteSpace(amount))
                    {
                        WriteToLogFile(AutomationLogLevel.Error, $"[ROW {rowNum}] Amount still not calculated after retries");
                    }
                    else
                    {
                        WriteToLogFile(AutomationLogLevel.Success, $"[ROW {rowNum}] Amount calculated on second retry: {amount}");
                    }
                }
                else
                {
                    WriteToLogFile(AutomationLogLevel.Success, $"[ROW {rowNum}] Amount calculated on retry: {amount}");
                }
            }
            else
            {
                WriteToLogFile(AutomationLogLevel.Success, $"[ROW {rowNum}] Amount calculated: {amount}");
            }
        }
        
        WriteToLogFile(AutomationLogLevel.Info, "[FILL] All profile rows processed! Final verification will happen after all items are filled.");
        FlushLogBuffer();
    }

    /// <summary>
    /// Fill accessories section on the Cortizo page
    /// </summary>
    private async Task FillAccessoriesAsync(List<AccessoryItem> accessories, AutomationRunResult result, CancellationToken cancellationToken = default)
    {
        WriteToLogFile(AutomationLogLevel.Info, $"[ACCESSORIES] Starting fill for {accessories.Count} accessories...");
        
        // Check how many accessory rows currently exist
        var countRowsScript = @"
            (function() {
                var count = 0;
                for (var i = 1; i <= 100; i++) {
                    var rowNum = i.toString().padStart(4, '0');
                    if (document.getElementById('txtReferenciaAcc_' + rowNum)) {
                        count++;
                    } else {
                        break;
                    }
                }
                return count;
            })();
        ";
        var currentRows = await _page!.EvaluateAsync<int>(countRowsScript);
        WriteToLogFile(AutomationLogLevel.Info, $"[ACCESSORIES] Current rows: {currentRows}, needed: {accessories.Count}");
        
        // Add more rows if needed using dgAccesoriosAddRow()
        if (currentRows < accessories.Count)
        {
            WriteToLogFile(AutomationLogLevel.Info, $"[ACCESSORIES] Adding {accessories.Count - currentRows} more rows...");
            while (currentRows < accessories.Count)
            {
                var addRowScript = @"
                    (function() {
                        if (typeof dgAccesoriosAddRow === 'function') {
                            dgAccesoriosAddRow();
                            return true;
                        }
                        return false;
                    })();
                ";
                var added = await _page.EvaluateAsync<bool>(addRowScript);
                if (!added) break;
                currentRows++;
                await Task.Delay(200);
            }
            // Allow page to stabilize after adding all rows
            await Task.Delay(1000);
            WriteToLogFile(AutomationLogLevel.Info, $"[ACCESSORIES] Rows ready: {currentRows}");
        }
        
        // Process each accessory using DIRECT JS calls (same pattern that works for profiles)
        for (int i = 0; i < accessories.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var accessory = accessories[i];
            var rowNum = (i + 1).ToString("D4");
            
            WriteToLogFile(AutomationLogLevel.Info, $"[ACC {rowNum}] Processing: REF={accessory.RefNumber}, AMT={accessory.Amount}, DESC={accessory.Description}");
            
            // Step 1: Set reference value and call ValidarFormatoDatosAcc directly (like profiles do)
            var escapedRef = accessory.RefNumber.Replace("'", "\\'");
            var setRefScript = $@"
                (function() {{
                    var refEl = document.getElementById('txtReferenciaAcc_{rowNum}');
                    if (!refEl) return 'not-found';
                    refEl.value = '{escapedRef}';
                    if (typeof ValidarFormatoDatosAcc === 'function') {{
                        ValidarFormatoDatosAcc(refEl);
                    }}
                    return 'ok';
                }})();
            ";
            var refResult = await _page.EvaluateAsync<string>(setRefScript);
            if (refResult == "not-found")
            {
                WriteToLogFile(AutomationLogLevel.Warning, $"[ACC {rowNum}] Reference input not found");
                continue;
            }
            
            // Wait for reference format validation AJAX to complete (fetches description)
            await Task.Delay(1500);
            
            // Step 2: Check if description was populated (confirms reference was recognized)
            var checkDescScript = $@"
                (function() {{
                    var descEl = document.getElementById('txtDescripcionAcc_{rowNum}');
                    return descEl ? descEl.value : '';
                }})();
            ";
            var descAfterRef = await _page.EvaluateAsync<string>(checkDescScript);
            if (string.IsNullOrWhiteSpace(descAfterRef))
            {
                // Description not populated yet, wait longer and retry format validation
                WriteToLogFile(AutomationLogLevel.Info, $"[ACC {rowNum}] Description not yet populated, waiting longer...");
                await Task.Delay(1500);
                descAfterRef = await _page.EvaluateAsync<string>(checkDescScript);
                if (string.IsNullOrWhiteSpace(descAfterRef))
                {
                    // Re-trigger format validation
                    await _page.EvaluateAsync(setRefScript);
                    await Task.Delay(2000);
                    descAfterRef = await _page.EvaluateAsync<string>(checkDescScript);
                }
            }
            
            // Step 3: Set quantity value and call ValidarCamposLineaAcc directly
            var setAmtScript = $@"
                (function() {{
                    var amtEl = document.getElementById('txtCantidadAcc_{rowNum}');
                    if (!amtEl) return false;
                    amtEl.value = '{accessory.Amount}';
                    if (typeof ValidarCamposLineaAcc === 'function') {{
                        ValidarCamposLineaAcc(amtEl, false);
                    }}
                    return true;
                }})();
            ";
            await _page.EvaluateAsync<bool>(setAmtScript);
            
            // Wait for the AJAX price calculation to complete
            await Task.Delay(2000);
            
            // Step 4: Check results
            var checkAmountScript = $@"
                (function() {{
                    var importeEl = document.getElementById('txtImporteAcc_{rowNum}');
                    var descEl = document.getElementById('txtDescripcionAcc_{rowNum}');
                    var priceEl = document.getElementById('txtPrecioAcc_{rowNum}');
                    return {{
                        amount: importeEl ? importeEl.value : '',
                        desc: descEl ? descEl.value : '',
                        price: priceEl ? priceEl.value : ''
                    }};
                }})();
            ";
            var accResult = await _page.EvaluateAsync<Dictionary<string, string>>(checkAmountScript);
            var amount = accResult.GetValueOrDefault("amount", "");
            var pageDesc = accResult.GetValueOrDefault("desc", "");
            var pagePrice = accResult.GetValueOrDefault("price", "");
            
            if (!string.IsNullOrWhiteSpace(amount))
            {
                WriteToLogFile(AutomationLogLevel.Success, $"[ACC {rowNum}] Amount={amount}, Price={pagePrice}, DESC={pageDesc}");
            }
            else
            {
                // Retry 1: Re-trigger ValidarCamposLineaAcc with longer wait
                WriteToLogFile(AutomationLogLevel.Warning, $"[ACC {rowNum}] Amount empty (price={pagePrice}, desc={pageDesc}), retry 1...");
                
                var retryCalcScript = $@"
                    (function() {{
                        var amtEl = document.getElementById('txtCantidadAcc_{rowNum}');
                        if (amtEl && typeof ValidarCamposLineaAcc === 'function') {{
                            ValidarCamposLineaAcc(amtEl, false);
                            return true;
                        }}
                        return false;
                    }})();
                ";
                await _page.EvaluateAsync<bool>(retryCalcScript);
                await Task.Delay(2500);
                
                accResult = await _page.EvaluateAsync<Dictionary<string, string>>(checkAmountScript);
                amount = accResult.GetValueOrDefault("amount", "");
                pageDesc = accResult.GetValueOrDefault("desc", "");
                pagePrice = accResult.GetValueOrDefault("price", "");
                
                if (!string.IsNullOrWhiteSpace(amount))
                {
                    WriteToLogFile(AutomationLogLevel.Success, $"[ACC {rowNum}] Amount on retry 1: {amount}, Price={pagePrice}, DESC={pageDesc}");
                }
                else
                {
                    // Retry 2: Re-validate reference first, then re-trigger amount calculation
                    WriteToLogFile(AutomationLogLevel.Warning, $"[ACC {rowNum}] Still empty, retry 2 (re-validate ref + amt)...");
                    
                    await _page.EvaluateAsync(setRefScript);
                    await Task.Delay(2000);
                    await _page.EvaluateAsync<bool>(setAmtScript);
                    await Task.Delay(3000);
                    
                    accResult = await _page.EvaluateAsync<Dictionary<string, string>>(checkAmountScript);
                    amount = accResult.GetValueOrDefault("amount", "");
                    pageDesc = accResult.GetValueOrDefault("desc", "");
                    pagePrice = accResult.GetValueOrDefault("price", "");
                    
                    if (!string.IsNullOrWhiteSpace(amount))
                    {
                        WriteToLogFile(AutomationLogLevel.Success, $"[ACC {rowNum}] Amount on retry 2: {amount}, Price={pagePrice}, DESC={pageDesc}");
                    }
                    else
                    {
                        WriteToLogFile(AutomationLogLevel.Warning, $"[ACC {rowNum}] Not yet calculated after retries (AJAX may still be pending). Will verify at the end.");
                    }
                }
            }
        }
        
        WriteToLogFile(AutomationLogLevel.Info, "[ACCESSORIES] All accessory rows processed! Final verification will happen after all items are filled.");
        FlushLogBuffer();
    }

    /// <summary>
    /// Get the Cortizo ESTIMATE TOTAL from the page
    /// </summary>
    private async Task<decimal> GetCortizoTotalAsync()
    {
        try
        {
            // The total is usually in an element with id containing "Total" or "Importe"
            var totalScript = @"
                (function() {
                    // Try various selectors for the total
                    var selectors = [
                        '#ctl00_ContentPlaceHolderCortizoCenter_lblTotalValoracion',
                        '#lblTotalValoracion',
                        '.total-valoracion',
                        'input[id*=""Total""]',
                        'span[id*=""Total""]'
                    ];
                    
                    for (var sel of selectors) {
                        var el = document.querySelector(sel);
                        if (el) {
                            var text = el.innerText || el.value || '';
                            // Extract number from text like '60084.31 €' or '60,084.31'
                            var match = text.match(/[\d,\.]+/);
                            if (match) {
                                return match[0].replace(',', '');
                            }
                        }
                    }
                    
                    // Try to find by text content
                    var elements = document.querySelectorAll('*');
                    for (var el of elements) {
                        if (el.innerText && el.innerText.includes('ESTIMATE TOTAL')) {
                            var parent = el.parentElement;
                            if (parent) {
                                var inputs = parent.querySelectorAll('input');
                                for (var inp of inputs) {
                                    if (inp.value) {
                                        return inp.value.replace(',', '');
                                    }
                                }
                            }
                        }
                    }
                    
                    return '';
                })();
            ";
            
            var totalStr = await _page!.EvaluateAsync<string>(totalScript);
            
            if (!string.IsNullOrWhiteSpace(totalStr))
            {
                // Parse the total
                if (decimal.TryParse(totalStr.Replace(",", "").Replace("€", "").Trim(), 
                    System.Globalization.NumberStyles.Any, 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    out var total))
                {
                    return total;
                }
            }
        }
        catch (Exception ex)
        {
            WriteToLogFile(AutomationLogLevel.Warning, $"Could not capture Cortizo total: {ex.Message}");
        }
        
        return 0;
    }

    private async Task SetHeaderFieldsAsync(QuotationViewModel viewModel)
    {
        WriteToLogFile(AutomationLogLevel.Info, "[HEADER] Setting header fields...");
        
        // Set Microns dropdown - id: ctl00_ContentPlaceHolderCortizoCenter_LstMicraje
        WriteToLogFile(AutomationLogLevel.Info, $"[HEADER] Setting Microns to: {viewModel.Microns}");
        await TrySetSelectValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_LstMicraje, select[name*='Micraje']", viewModel.Microns.ToString());

        // Set Language dropdown - id: ctl00_ContentPlaceHolderCortizoCenter_ddlIdioma
        WriteToLogFile(AutomationLogLevel.Info, $"[HEADER] Setting Language to: {viewModel.Language}");
        await TrySetSelectValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_ddlIdioma, select[name*='Idioma']", viewModel.Language);

        // Set General Color settings using the correct IDs from the HTML
        // Finish 1: #ddlAcabado_1_ColorGeneral
        // Value mapping: 90 = "SPECIAL 1 POWDER COATING", 8 = "MILL FINISH", 9 = "STANDARD POWDER COATING"
        var finish1Value = MapFinishToValue(viewModel.GeneralFinish1);
        WriteToLogFile(AutomationLogLevel.Info, $"[HEADER] Setting Finish 1: {viewModel.GeneralFinish1} -> value={finish1Value}");
        await TrySetSelectValueAsync("#ddlAcabado_1_ColorGeneral", finish1Value);
        
        // Wait for the shade dropdown to be populated after selecting finish
        await Task.Delay(500);
        
        // Shade 1: #ddlMatiz_1_ColorGeneral
        WriteToLogFile(AutomationLogLevel.Info, $"[HEADER] Setting Shade 1: {viewModel.GeneralShade1}");
        await TrySetSelectValueAsync("#ddlMatiz_1_ColorGeneral", viewModel.GeneralShade1);
        
        // Finish 2: #ddlAcabado_2_ColorGeneral
        var finish2Value = MapFinishToValue(viewModel.GeneralFinish2);
        WriteToLogFile(AutomationLogLevel.Info, $"[HEADER] Setting Finish 2: {viewModel.GeneralFinish2} -> value={finish2Value}");
        await TrySetSelectValueAsync("#ddlAcabado_2_ColorGeneral", finish2Value);
        
        // Wait for the shade dropdown to be populated
        await Task.Delay(500);
        
        // Shade 2: #ddlMatiz_2_ColorGeneral
        WriteToLogFile(AutomationLogLevel.Info, $"[HEADER] Setting Shade 2: {viewModel.GeneralShade2}");
        await TrySetSelectValueAsync("#ddlMatiz_2_ColorGeneral", viewModel.GeneralShade2);
        
        WriteToLogFile(AutomationLogLevel.Info, "[HEADER] Header fields set complete");
        FlushLogBuffer();
    }

    /// <summary>
    /// Set the customized prices section with static values
    /// </summary>
    private async Task SetCustomizedPricesAsync()
    {
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Setting customized prices...");
        
        // SYSTEMS ALUMINIUM
        // WITH BREAK: 7.56 (*/Kg)
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Setting WITH BREAK: 7.56");
        await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPrecioConRotura", "7.56");
        
        // WITHOUT BREAK: 5.16 (*/Kg)
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Setting WITHOUT BREAK: 5.16");
        await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPrecioSinRotura", "5.16");
        
        // LACQUERED: (empty in the image, but can set if needed)
        // await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPrecioLacado", "");
        
        // ANODIZED: (empty in the image, but can set if needed)
        // await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPrecioAnodizado", "");
        
        // SYSTEMS PVC
        // PVC: 15 (%)
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Setting PVC: 15");
        await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPVCDescuentoPVC", "15");
        
        // ALUMINIUM: 10 (%)
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Setting ALUMINIUM: 10");
        await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPVCDescuentoAluminio", "10");
        
        // STEEL: (empty in the image)
        // await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtPVCDescuentoAcero", "");
        
        // ACCESSORIES
        // DISCOUNT IN ALUMINUM AND PVC ACCESSORIES: 10 (%)
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Setting ACCESSORIES DISCOUNT: 10");
        await TrySetInputValueAsync("#ctl00_ContentPlaceHolderCortizoCenter_txtDescuentoAccesorios", "10");
        
        // Trigger the onchange event to update prices
        await _page!.EvaluateAsync("if(typeof precioPersonalizadoChanged === 'function') precioPersonalizadoChanged();");
        
        WriteToLogFile(AutomationLogLevel.Info, "[PRICES] Customized prices set complete");
        FlushLogBuffer();
    }

    /// <summary>
    /// Map finish display text to Cortizo dropdown value
    /// </summary>
    private string MapFinishToValue(string finishText)
    {
        var upperText = finishText.ToUpperInvariant();
        
        if (upperText.Contains("SPECIAL 1")) return "90";
        if (upperText.Contains("SPECIAL 2")) return "91";
        if (upperText.Contains("SPECIAL 3")) return "92";
        if (upperText.Contains("SPECIAL 4")) return "93";
        if (upperText.Contains("SPECIAL 5")) return "94";
        if (upperText.Contains("SPECIAL 6")) return "95";
        if (upperText.Contains("SPECIAL 7")) return "96";
        if (upperText.Contains("STANDARD")) return "9";
        if (upperText.Contains("MILL FINISH") || upperText.Contains("COR_MILL")) return "8";
        if (upperText.Contains("WHITE POWDER")) return "4";
        if (upperText.Contains("SILVER ANODISED")) return "1";
        if (upperText.Contains("BLACK ANODISED")) return "10";
        if (upperText.Contains("PVC")) return "0";
        
        // Default to SPECIAL 1 POWDER COATING
        return "90";
    }

    private async Task FillProfileRowAsync(int rowIndex, ProfileItem profile)
    {
        // Cortizo uses 4-digit padded row numbers starting from 0001
        // Row IDs follow pattern: txtReferencia_XXXX, txtCantidad_XXXX, ddlAcabado1_XXXX, etc.
        var rowNum = (rowIndex + 1).ToString("D4"); // 0001, 0002, etc.
        
        WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Starting to fill profile: REF={profile.RefNumber}, AMT={profile.Amount}");
        
        // Wait for row to be available
        await Task.Delay(300);

        // Set REF (Reference Number) - triggers ValidarFormatoDatosPerfil on change
        var refSelector = $"#txtReferencia_{rowNum}";
        WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Looking for reference input: {refSelector}");
        var refInput = await _page!.QuerySelectorAsync(refSelector);
        if (refInput != null)
        {
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Found reference input, filling with: {profile.RefNumber}");
            await refInput.FillAsync(profile.RefNumber);
            // Trigger the onchange event by dispatching
            await refInput.DispatchEventAsync("change");
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Dispatched change event on reference input");
        }
        else
        {
            WriteToLogFile(AutomationLogLevel.Error, $"[ROW {rowNum}] Could not find reference input: {refSelector}");
            _logger.LogWarning($"Could not find reference input: {refSelector}");
        }

        // Wait for the system to validate and load profile data
        WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Waiting for profile validation...");
        await Task.Delay(1000);

        // Set Amount (Quantity)
        var amtSelector = $"#txtCantidad_{rowNum}";
        var amtInput = await _page.QuerySelectorAsync(amtSelector);
        if (amtInput != null)
        {
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Setting amount to: {profile.Amount}");
            await amtInput.FillAsync(profile.Amount.ToString());
        }
        else
        {
            WriteToLogFile(AutomationLogLevel.Warning, $"[ROW {rowNum}] Amount input not found: {amtSelector}");
        }

        // The finish/shade may be auto-set by the General Color settings
        // Only override if different from general settings
        
        // Set Finish 1 (Acabado 1) - value is numeric code
        if (!string.IsNullOrEmpty(profile.Finish1))
        {
            var finish1Selector = $"#ddlAcabado1_{rowNum}";
            var finish1Value = MapFinishToValue(profile.Finish1);
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Setting Finish 1: {profile.Finish1} -> {finish1Value}");
            await TrySetSelectValueAsync(finish1Selector, finish1Value);
            
            // Wait for shade dropdown to populate
            await Task.Delay(500);
        }

        // Set Shade 1 (Matiz 1)
        if (!string.IsNullOrEmpty(profile.Shade1))
        {
            var shade1Selector = $"#ddlMatiz1_{rowNum}";
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Setting Shade 1: {profile.Shade1}");
            await TrySetSelectValueAsync(shade1Selector, profile.Shade1);
        }

        // Set Finish 2 (Acabado 2)
        if (!string.IsNullOrEmpty(profile.Finish2))
        {
            var finish2Selector = $"#ddlAcabado2_{rowNum}";
            var finish2Value = MapFinishToValue(profile.Finish2);
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Setting Finish 2: {profile.Finish2} -> {finish2Value}");
            await TrySetSelectValueAsync(finish2Selector, finish2Value);
            
            // Wait for shade dropdown to populate
            await Task.Delay(500);
        }

        // Set Shade 2 (Matiz 2)
        if (!string.IsNullOrEmpty(profile.Shade2))
        {
            var shade2Selector = $"#ddlMatiz2_{rowNum}";
            WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Setting Shade 2: {profile.Shade2}");
            await TrySetSelectValueAsync(shade2Selector, profile.Shade2);
        }

        // Trigger the line validation which calls GetPrecio to calculate amounts
        // This is done by triggering change on one of the controls
        WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Triggering line validation...");
        var shade1Control = await _page.QuerySelectorAsync($"#ddlMatiz1_{rowNum}");
        if (shade1Control != null)
        {
            await shade1Control.DispatchEventAsync("change");
        }

        // Wait for the price calculation to complete (GetPrecio AJAX call)
        await WaitForPriceCalculationAsync(rowNum);
        WriteToLogFile(AutomationLogLevel.Info, $"[ROW {rowNum}] Profile fill complete");
        FlushLogBuffer();
    }

    /// <summary>
    /// Wait for the price calculation to complete by checking if txtImporte has a value
    /// </summary>
    private async Task WaitForPriceCalculationAsync(string rowNum, int maxWaitMs = 10000)
    {
        var importeSelector = $"#txtImporte_{rowNum}";
        var startTime = DateTime.Now;

        while ((DateTime.Now - startTime).TotalMilliseconds < maxWaitMs)
        {
            var importeInput = await _page!.QuerySelectorAsync(importeSelector);
            if (importeInput != null)
            {
                var value = await importeInput.GetAttributeAsync("value");
                if (!string.IsNullOrWhiteSpace(value))
                {
                    _logger.LogInformation($"Price calculated for row {rowNum}: {value}");
                    return;
                }
            }
            await Task.Delay(500);
        }

        _logger.LogWarning($"Price calculation timeout for row {rowNum}");
    }

    private async Task TrySetInputValueAsync(string selector, string value)
    {
        try
        {
            var element = await _page!.QuerySelectorAsync(selector);
            if (element != null)
            {
                await element.FillAsync(value);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, $"Failed to set input value with selector: {selector}");
        }
    }

    private async Task TrySetInputValueWithSelectorsAsync(string[] selectors, string value)
    {
        foreach (var selector in selectors)
        {
            try
            {
                var element = await _page!.QuerySelectorAsync(selector);
                if (element != null)
                {
                    await element.FillAsync(value);
                    return;
                }
            }
            catch
            {
                // Try next selector
            }
        }

        _logger.LogWarning($"Could not find input element with any selector for value: {value}");
    }

    private async Task TrySetSelectValueAsync(string selector, string value)
    {
        try
        {
            WriteToLogFileOnly(AutomationLogLevel.Info, $"[SELECT] Trying selector: {selector} with value: {value}");
            var element = await _page!.QuerySelectorAsync(selector);
            if (element != null)
            {
                var isVisible = await element.IsVisibleAsync();
                WriteToLogFileOnly(AutomationLogLevel.Info, $"[SELECT] Found element, visible: {isVisible}");
                
                var options = await element.QuerySelectorAllAsync("option");
                var optionValues = new List<string>();
                foreach (var opt in options.Take(10))
                {
                    var optVal = await opt.GetAttributeAsync("value") ?? "";
                    var optText = (await opt.TextContentAsync() ?? "").Trim();
                    optionValues.Add($"{optVal}='{optText}'");
                }
                WriteToLogFileOnly(AutomationLogLevel.Info, $"[SELECT] Options ({options.Count} total): {string.Join(", ", optionValues)}");
                
                try
                {
                    await element.SelectOptionAsync(new SelectOptionValue { Value = value });
                    WriteToLogFileOnly(AutomationLogLevel.Info, $"[SELECT] Successfully set by value: {value}");
                }
                catch (Exception ex1)
                {
                    WriteToLogFileOnly(AutomationLogLevel.Info, $"[SELECT] Set by value failed: {ex1.Message}, trying by label...");
                    try
                    {
                        await element.SelectOptionAsync(new SelectOptionValue { Label = value });
                        WriteToLogFileOnly(AutomationLogLevel.Info, $"[SELECT] Successfully set by label: {value}");
                    }
                    catch (Exception ex2)
                    {
                        WriteToLogFileOnly(AutomationLogLevel.Warning, $"[SELECT] Set by label also failed: {ex2.Message}");
                    }
                }
            }
            else
            {
                WriteToLogFileOnly(AutomationLogLevel.Warning, $"[SELECT] Element not found: {selector}");
            }
        }
        catch (Exception ex)
        {
            WriteToLogFileOnly(AutomationLogLevel.Error, $"[SELECT] Error: {ex.Message}");
            _logger.LogWarning(ex, $"Failed to set select value with selector: {selector}");
        }
    }

    private async Task TrySetSelectValueWithSelectorsAsync(string[] selectors, string value)
    {
        foreach (var selector in selectors)
        {
            try
            {
                var element = await _page!.QuerySelectorAsync(selector);
                if (element != null)
                {
                    // Try by value first, then by label
                    try
                    {
                        await element.SelectOptionAsync(new SelectOptionValue { Value = value });
                        return;
                    }
                    catch
                    {
                        try
                        {
                            await element.SelectOptionAsync(new SelectOptionValue { Label = value });
                            return;
                        }
                        catch
                        {
                            // Try partial match
                            var options = await element.QuerySelectorAllAsync("option");
                            foreach (var option in options)
                            {
                                var text = await option.TextContentAsync();
                                if (text?.Contains(value, StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    var optionValue = await option.GetAttributeAsync("value");
                                    if (optionValue != null)
                                    {
                                        await element.SelectOptionAsync(new SelectOptionValue { Value = optionValue });
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Try next selector
            }
        }

        _logger.LogWarning($"Could not find select element with any selector for value: {value}");
    }

    private async Task ClickButtonAsync(string buttonText)
    {
        var selectors = new[]
        {
            $"text={buttonText}",
            $"button:has-text('{buttonText}')",
            $"input[value*='{buttonText}']",
            $"a:has-text('{buttonText}')"
        };

        foreach (var selector in selectors)
        {
            try
            {
                var element = await _page!.QuerySelectorAsync(selector);
                if (element != null)
                {
                    await element.ClickAsync();
                    await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    return;
                }
            }
            catch
            {
                // Try next selector
            }
        }

        _logger.LogWarning($"Could not find button: {buttonText}");
    }

    private void Log(AutomationRunResult result, AutomationLogLevel level, string message, string? details = null)
    {
        var entry = new AutomationLogEntry
        {
            Timestamp = DateTime.Now,
            Level = level,
            Message = message,
            Details = details
        };

        result.Logs.Add(entry);
        OnLog?.Invoke(entry);

        // Write to log file (skip SignalR since we already fired it above)
        WriteToLogFileOnly(level, message, details);

        // Also log to standard logger
        switch (level)
        {
            case AutomationLogLevel.Error:
                _logger.LogError(message);
                break;
            case AutomationLogLevel.Warning:
                _logger.LogWarning(message);
                break;
            default:
                _logger.LogInformation(message);
                break;
        }
    }
    
    /// <summary>
    /// Write a log entry to the log file
    /// </summary>
    private void WriteToLogFile(AutomationLogLevel level, string message, string? details = null)
    {
        // Fire SignalR event so the UI receives detailed messages
        OnLog?.Invoke(new AutomationLogEntry
        {
            Timestamp = DateTime.Now,
            Level = level,
            Message = message,
            Details = details
        });

        WriteToLogFileOnly(level, message, details);
    }

    private void WriteToLogFileOnly(AutomationLogLevel level, string message, string? details = null)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var levelStr = level.ToString().ToUpper().PadRight(7);
            var logLine = $"[{timestamp}] [{levelStr}] {message}";
            
            if (!string.IsNullOrEmpty(details))
            {
                logLine += $"\n    Details: {details}";
            }
            
            _logBuffer.AppendLine(logLine);
            
            // Flush to file periodically or on important events
            if (level == AutomationLogLevel.Error || level == AutomationLogLevel.Success || _logBuffer.Length > 4096)
            {
                FlushLogBuffer();
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to write to log file");
        }
    }
    
    /// <summary>
    /// Log detailed page state for debugging
    /// </summary>
    private async Task LogPageStateAsync(string context)
    {
        try
        {
            if (_page == null) return;
            
            var url = _page.Url;
            var title = await _page.TitleAsync();
            
            WriteToLogFileOnly(AutomationLogLevel.Info, $"[PAGE STATE - {context}]");
            WriteToLogFileOnly(AutomationLogLevel.Info, $"  URL: {url}");
            WriteToLogFileOnly(AutomationLogLevel.Info, $"  Title: {title}");
            
            var elementsToCheck = new Dictionary<string, string>
            {
                { "Login Form", "input[type='password']" },
                { "Quotations Link", "#ctl00_ico7, a.ico7" },
                { "Profiles Grid", "#gvPerfiles" },
                { "General Color Finish 1", "#ddlAcabado_1_ColorGeneral" },
                { "First Reference Input", "#txtReferencia_0001" }
            };
            
            foreach (var (name, selector) in elementsToCheck)
            {
                try
                {
                    var element = await _page.QuerySelectorAsync(selector);
                    var exists = element != null;
                    var visible = exists && await element!.IsVisibleAsync();
                    WriteToLogFileOnly(AutomationLogLevel.Info, $"  {name}: {(exists ? "Found" : "Not Found")}{(visible ? " (Visible)" : exists ? " (Hidden)" : "")}");
                }
                catch
                {
                    WriteToLogFileOnly(AutomationLogLevel.Info, $"  {name}: Check failed");
                }
            }
            
            FlushLogBuffer();
        }
        catch (Exception ex)
        {
            WriteToLogFileOnly(AutomationLogLevel.Warning, $"Failed to log page state: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Log all visible selectors on the page for debugging
    /// </summary>
    private async Task LogAvailableSelectorsAsync(string containerSelector = "body")
    {
        try
        {
            if (_page == null) return;
            
            WriteToLogFileOnly(AutomationLogLevel.Info, $"[AVAILABLE SELECTORS in {containerSelector}]");
            
            var selects = await _page.QuerySelectorAllAsync($"{containerSelector} select");
            WriteToLogFileOnly(AutomationLogLevel.Info, $"  Found {selects.Count} select elements:");
            foreach (var select in selects.Take(20))
            {
                var id = await select.GetAttributeAsync("id") ?? "";
                var name = await select.GetAttributeAsync("name") ?? "";
                var visible = await select.IsVisibleAsync();
                WriteToLogFileOnly(AutomationLogLevel.Info, $"    - id='{id}' name='{name}' visible={visible}");
            }
            
            var inputs = await _page.QuerySelectorAllAsync($"{containerSelector} input[type='text']");
            WriteToLogFileOnly(AutomationLogLevel.Info, $"  Found {inputs.Count} text input elements:");
            foreach (var input in inputs.Take(20))
            {
                var id = await input.GetAttributeAsync("id") ?? "";
                var name = await input.GetAttributeAsync("name") ?? "";
                var visible = await input.IsVisibleAsync();
                WriteToLogFileOnly(AutomationLogLevel.Info, $"    - id='{id}' name='{name}' visible={visible}");
            }
            
            var buttons = await _page.QuerySelectorAllAsync($"{containerSelector} button, {containerSelector} input[type='submit'], {containerSelector} a[onclick]");
            WriteToLogFileOnly(AutomationLogLevel.Info, $"  Found {buttons.Count} button/link elements:");
            foreach (var btn in buttons.Take(20))
            {
                var id = await btn.GetAttributeAsync("id") ?? "";
                var text = (await btn.TextContentAsync() ?? "").Trim().Replace("\n", " ");
                if (text.Length > 50) text = text.Substring(0, 50) + "...";
                var visible = await btn.IsVisibleAsync();
                WriteToLogFileOnly(AutomationLogLevel.Info, $"    - id='{id}' text='{text}' visible={visible}");
            }
            
            FlushLogBuffer();
        }
        catch (Exception ex)
        {
            WriteToLogFileOnly(AutomationLogLevel.Warning, $"Failed to log available selectors: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Flush log buffer to file
    /// </summary>
    private void FlushLogBuffer()
    {
        if (_logBuffer.Length == 0) return;
        
        try
        {
            File.AppendAllText(_logFilePath, _logBuffer.ToString());
            _logBuffer.Clear();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to flush log buffer to file");
        }
    }

    public async ValueTask DisposeAsync()
    {
        // Flush any remaining logs
        WriteToLogFile(AutomationLogLevel.Info, "Disposing automation service...");
        FlushLogBuffer();
        
        if (_page != null)
        {
            await _page.CloseAsync();
            _page = null;
        }

        if (_context != null)
        {
            await _context.CloseAsync();
            _context = null;
        }

        if (_browser != null)
        {
            await _browser.CloseAsync();
            _browser = null;
        }

        _playwright?.Dispose();
        _playwright = null;
        
        WriteToLogFile(AutomationLogLevel.Info, "Automation service disposed.");
        FlushLogBuffer();
    }
}
