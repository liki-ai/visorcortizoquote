# Visor Quotation Web App

A .NET 8 web application that automates the creation of quotations in Cortizo Center by parsing PDF stock lists and using browser automation.

## Features

- **PDF Upload & Parsing**: Upload PDF stock lists (from Logikal) and automatically extract profile data
- **PDF Preview**: Built-in PDF viewer using PDF.js
- **Editable Grid**: Review and edit extracted profile items before submission
- **Browser Automation**: Uses Playwright to log into Cortizo Center and fill out quotation forms
- **Real-time Logging**: SignalR-based live progress updates during automation
- **Configurable Mappings**: Customizable finish/shade mappings in `appsettings.json`

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- [Playwright browsers](https://playwright.dev/dotnet/docs/intro)

## Installation

### 1. Clone or Download the Project

```bash
cd VisorQuotationWebApp
```

### 2. Restore NuGet Packages

```bash
dotnet restore
```

### 3. Install Playwright Browsers

This step downloads the Chromium browser that Playwright will use for automation:

```bash
# Install the Playwright CLI tool
dotnet tool install --global Microsoft.Playwright.CLI

# Install browsers
playwright install chromium
```

Alternatively, you can use the PowerShell script provided:

```powershell
# From the project directory
pwsh bin/Debug/net8.0/playwright.ps1 install chromium
```

### 4. Run the Application

```bash
dotnet run
```

The application will start at `https://localhost:5001` (or `http://localhost:5000`).

## Usage

### 1. Upload a PDF Stock List

- Click the upload area or drag-and-drop a PDF file
- The application will parse the PDF and extract profile items
- The PDF will be displayed in the preview panel on the right

### 2. Review Extracted Data

- Check the extracted profiles in the grid on the left
- Select/deselect items to include in the automation
- The header information panel shows project details from the PDF

### 3. Configure Settings

- Adjust the Cortizo header settings (Microns, CIF, Client Code, Language)
- Modify general colour settings (Finish 1/2, Shade 1/2)
- Enter your Cortizo Center credentials

### 4. Run Automation

- Click "Run Automation" to start the browser automation
- Watch the progress in the execution log
- The automation will:
  1. Log into Cortizo Center
  2. Navigate to Quotations / Online Orders
  3. Create a new valuation
  4. Fill in the header fields
  5. Enter each profile row
  6. Optionally generate report or create proforma

## Configuration

### appsettings.json

```json
{
  "Cortizo": {
    "BaseUrl": "https://cortizocenter.com",
    "DefaultMicrons": 15,
    "DefaultCif": "CORTIZO",
    "DefaultClientCode": "991238",
    "DefaultLanguage": "ENGLISH",
    "Headless": false,
    "TimeoutMs": 30000,
    "FinishMappings": {
      "Special 1 Powder Coating P1019M": {
        "Finish1": "SPECIAL 1 POWDER COATING",
        "Shade1": "P1019M",
        "Finish2": "SPECIAL 1 POWDER COATING",
        "Shade2": "P1019M"
      },
      "COR_MILL FINISH": {
        "Finish1": "COR_MILL FINISH",
        "Shade1": "",
        "Finish2": "COR_MILL FINISH",
        "Shade2": ""
      }
    }
  }
}
```

### Configuration Options

| Setting | Description |
|---------|-------------|
| `BaseUrl` | Cortizo Center URL |
| `DefaultMicrons` | Default microns dropdown value |
| `DefaultCif` | Default CIF value |
| `DefaultClientCode` | Default client code |
| `DefaultLanguage` | Default language (ENGLISH/SPANISH) |
| `Headless` | Run browser in headless mode (true/false) |
| `TimeoutMs` | Browser operation timeout in milliseconds |
| `FinishMappings` | Dictionary mapping raw colour text to Finish/Shade values |

### Adding Custom Finish Mappings

To add a new colour mapping, add an entry to the `FinishMappings` section:

```json
"FinishMappings": {
  "Custom Colour Name": {
    "Finish1": "DROPDOWN_VALUE_1",
    "Shade1": "SHADE_VALUE_1",
    "Finish2": "DROPDOWN_VALUE_2",
    "Shade2": "SHADE_VALUE_2"
  }
}
```

## Debugging Selectors

If the automation fails to find elements on the Cortizo website (due to UI changes), you can debug by:

1. Set `Headless: false` in appsettings.json to watch the browser
2. Use browser DevTools to inspect elements
3. Update selectors in `CortizoAutomationService.cs`

### Key Selector Locations

- `LoginAsync()`: Username/password inputs, login button
- `NavigateToQuotationsAsync()`: Menu navigation
- `CreateNewValuationAsync()`: New valuation button
- `SetHeaderFieldsAsync()`: Header field inputs/dropdowns
- `FillProfileRowAsync()`: Grid row inputs

### Playwright Trace Files

On automation failure, trace files are saved to the temp directory:
- `cortizo-trace-{timestamp}.zip`: Playwright trace (can be viewed in Playwright Trace Viewer)
- `cortizo-screenshot-{timestamp}.png`: Final screenshot

To view a trace file:
```bash
playwright show-trace <path-to-trace.zip>
```

## Project Structure

```
VisorQuotationWebApp/
├── Controllers/
│   └── HomeController.cs       # Main controller with upload/automation endpoints
├── Hubs/
│   └── AutomationHub.cs        # SignalR hub for real-time updates
├── Models/
│   ├── AutomationConfig.cs     # Configuration models
│   ├── AutomationRunResult.cs  # Automation result models
│   ├── CortizoCredentials.cs   # Credentials model
│   ├── ParsedPdfResult.cs      # PDF parsing result
│   ├── PdfHeaderInfo.cs        # PDF header data
│   ├── ProfileItem.cs          # Profile line item
│   └── QuotationViewModel.cs   # Main view model
├── Services/
│   ├── CortizoAutomationService.cs  # Playwright automation
│   └── PdfParseService.cs           # PDF parsing with PdfPig
├── Views/
│   ├── Home/
│   │   └── Index.cshtml        # Main UI view
│   └── Shared/
│       └── _Layout.cshtml      # Layout template
├── wwwroot/
│   ├── css/site.css            # Application styles
│   └── js/
│       ├── automation.js       # Client-side automation/SignalR
│       └── pdf-viewer.js       # PDF.js wrapper
├── appsettings.json            # Configuration
├── Program.cs                  # Application entry point
└── README.md                   # This file
```

## Troubleshooting

### "Playwright browsers not installed"

Run:
```bash
playwright install chromium
```

### "Login failed"

- Check credentials in the UI
- Verify Cortizo Center is accessible
- Check for CAPTCHA or 2FA requirements

### "Could not find element" errors

- The Cortizo website UI may have changed
- Run in headed mode (`Headless: false`) to debug
- Update selectors in `CortizoAutomationService.cs`

### PDF parsing issues

- Check the execution log for parse warnings
- The parser expects PDFs from Logikal software
- Different PDF formats may require parser updates

## Security Notes

- Credentials are not stored persistently (only in session)
- For production, use environment variables or Azure Key Vault
- Consider using user secrets for development:
  ```bash
  dotnet user-secrets set "CortizoCredentials:Username" "your-username"
  dotnet user-secrets set "CortizoCredentials:Password" "your-password"
  ```

## License

Proprietary - Visor LLC

## Support

For issues or feature requests, contact the development team.
