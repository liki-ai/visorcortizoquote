using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using VisorQuotationWebApp.Models;

namespace VisorQuotationWebApp.Services;

/// <summary>
/// Service for generating Visor quotation PDF documents
/// </summary>
public class VisorQuotationService
{
    private readonly ILogger<VisorQuotationService> _logger;

    public VisorQuotationService(ILogger<VisorQuotationService> logger)
    {
        _logger = logger;
        
        // Set QuestPDF license (Community license for open source)
        QuestPDF.Settings.License = LicenseType.Community;
    }

    /// <summary>
    /// Create a quotation order from parsed PDF data and Cortizo results
    /// </summary>
    public VisorQuotationOrder CreateQuotationOrder(
        ParsedPdfResult parsedPdf,
        QuotationViewModel viewModel,
        List<CortizoLineResult> cortizoResults)
    {
        var order = new VisorQuotationOrder
        {
            QuotationNumber = GenerateQuotationNumber(),
            QuotationDate = DateTime.Now,
            ValidUntil = DateTime.Now.AddDays(30),
            ProjectName = parsedPdf.Header?.ProjectName ?? "N/A",
            ClientName = parsedPdf.Header?.CompanyName ?? viewModel.ClientPurchaseOrder,
            CortizoReference = viewModel.ClientCode
        };

        // Add items from profiles
        int lineNum = 1;
        decimal subtotal = 0;

        foreach (var result in cortizoResults.Where(r => r.Amount > 0))
        {
            var profile = parsedPdf.Profiles.FirstOrDefault(p => p.RefNumber == result.RefNumber);
            
            var item = new VisorQuotationItem
            {
                LineNumber = lineNum++,
                RefNumber = result.RefNumber,
                Description = result.Description ?? profile?.Description ?? "",
                Finish = result.Finish1 ?? "",
                Shade = result.Shade1 ?? "",
                Quantity = result.Quantity,
                Length = result.Length,
                UnitPrice = result.UnitPrice,
                TotalPrice = result.Amount
            };

            order.Items.Add(item);
            subtotal += result.Amount;
        }

        order.Subtotal = subtotal;
        order.CortizoTotal = subtotal;
        order.VatAmount = Math.Round(subtotal * (order.VatRate / 100), 2);
        order.Total = order.Subtotal + order.VatAmount;

        return order;
    }

    /// <summary>
    /// Create a quotation order directly from profile items (simplified version)
    /// </summary>
    public VisorQuotationOrder CreateQuotationFromProfiles(
        ParsedPdfResult parsedPdf,
        QuotationViewModel viewModel,
        decimal cortizoTotalAmount)
    {
        var order = new VisorQuotationOrder
        {
            QuotationNumber = GenerateQuotationNumber(),
            QuotationDate = DateTime.Now,
            ValidUntil = DateTime.Now.AddDays(30),
            ProjectName = parsedPdf.Header?.ProjectName ?? "N/A",
            ClientName = parsedPdf.Header?.CompanyName ?? "",
            CortizoReference = viewModel.ClientCode,
            CortizoTotal = cortizoTotalAmount
        };

        // Add items from selected profiles
        int lineNum = 1;

        foreach (var profile in parsedPdf.Profiles.Where(p => p.IsSelected))
        {
            var item = new VisorQuotationItem
            {
                LineNumber = lineNum++,
                RefNumber = profile.RefNumber,
                Description = profile.Description,
                Finish = !string.IsNullOrEmpty(profile.Finish1) ? profile.Finish1 : viewModel.GeneralFinish1,
                Shade = !string.IsNullOrEmpty(profile.Shade1) ? profile.Shade1 : viewModel.GeneralShade1,
                Quantity = profile.Amount,
                Length = profile.TotalLength,
                UnitPrice = 0,
                TotalPrice = 0
            };

            order.Items.Add(item);
        }

        // Add accessories/hardware items
        int accLineNum = 1;
        foreach (var acc in parsedPdf.Accessories.Where(a => a.IsSelected))
        {
            var item = new VisorQuotationItem
            {
                LineNumber = accLineNum++,
                RefNumber = acc.RefNumber,
                Description = acc.Description,
                Finish = acc.Finish ?? "",
                Shade = "",
                Quantity = acc.Amount,
                Length = 0,
                UnitPrice = 0,
                TotalPrice = 0
            };
            order.AccessoryItems.Add(item);
        }

        // Use Cortizo total
        order.Subtotal = cortizoTotalAmount;
        order.VatAmount = Math.Round(cortizoTotalAmount * (order.VatRate / 100), 2);
        order.Total = order.Subtotal + order.VatAmount;

        return order;
    }

    /// <summary>
    /// Generate PDF document for the quotation
    /// </summary>
    public byte[] GeneratePdf(VisorQuotationOrder order)
    {
        _logger.LogInformation($"Generating PDF for quotation {order.QuotationNumber}");

        var document = Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(40);
                page.DefaultTextStyle(x => x.FontSize(10));

                page.Header().Element(c => ComposeHeader(c, order));
                page.Content().Element(c => ComposeContent(c, order));
                page.Footer().Element(c => ComposeFooter(c, order));
            });
        });

        return document.GeneratePdf();
    }

    /// <summary>
    /// Save PDF to file
    /// </summary>
    public string SavePdf(VisorQuotationOrder order, string outputFolder)
    {
        var pdfBytes = GeneratePdf(order);
        var fileName = $"Visor_Quotation_{order.QuotationNumber}_{DateTime.Now:yyyyMMdd}.pdf";
        var filePath = Path.Combine(outputFolder, fileName);
        
        Directory.CreateDirectory(outputFolder);
        File.WriteAllBytes(filePath, pdfBytes);
        
        _logger.LogInformation($"Saved quotation PDF to: {filePath}");
        return filePath;
    }

    private void ComposeHeader(IContainer container, VisorQuotationOrder order)
    {
        container.Row(row =>
        {
            // Company logo/name on left
            row.RelativeItem().Column(col =>
            {
                col.Item().Text("VISOR").FontSize(28).Bold().FontColor(Colors.Blue.Darken3);
                col.Item().Text("Aluminum Systems").FontSize(12).FontColor(Colors.Grey.Darken1);
                col.Item().PaddingTop(5).Text(order.CompanyInfo.Address).FontSize(9);
                col.Item().Text(order.CompanyInfo.City).FontSize(9);
                col.Item().Text($"Tel: {order.CompanyInfo.Phone}").FontSize(9);
                col.Item().Text($"Email: {order.CompanyInfo.Email}").FontSize(9);
                col.Item().Text($"Web: {order.CompanyInfo.Website}").FontSize(9);
            });

            // Quotation info on right
            row.RelativeItem().AlignRight().Column(col =>
            {
                col.Item().Text("QUOTATION").FontSize(24).Bold().FontColor(Colors.Blue.Darken3);
                col.Item().PaddingTop(10).Row(r =>
                {
                    r.AutoItem().Text("Number: ").SemiBold();
                    r.AutoItem().Text(order.QuotationNumber);
                });
                col.Item().Row(r =>
                {
                    r.AutoItem().Text("Date: ").SemiBold();
                    r.AutoItem().Text(order.QuotationDate.ToString("dd/MM/yyyy"));
                });
                col.Item().Row(r =>
                {
                    r.AutoItem().Text("Valid Until: ").SemiBold();
                    r.AutoItem().Text(order.ValidUntil.ToString("dd/MM/yyyy"));
                });
                if (!string.IsNullOrEmpty(order.CortizoReference))
                {
                    col.Item().Row(r =>
                    {
                        r.AutoItem().Text("Cortizo Ref: ").SemiBold();
                        r.AutoItem().Text(order.CortizoReference);
                    });
                }
            });
        });

        container.PaddingVertical(15).LineHorizontal(1).LineColor(Colors.Blue.Darken3);
    }

    private void ComposeContent(IContainer container, VisorQuotationOrder order)
    {
        container.Column(col =>
        {
            // Client info
            col.Item().PaddingBottom(15).Row(row =>
            {
                row.RelativeItem().Column(c =>
                {
                    c.Item().Text("PROJECT / CLIENT").FontSize(11).SemiBold().FontColor(Colors.Blue.Darken2);
                    c.Item().PaddingTop(5).Text(order.ProjectName).FontSize(10);
                    if (!string.IsNullOrEmpty(order.ClientName))
                        c.Item().Text(order.ClientName).FontSize(10);
                    if (!string.IsNullOrEmpty(order.ClientAddress))
                        c.Item().Text(order.ClientAddress).FontSize(9);
                    if (!string.IsNullOrEmpty(order.ClientPhone))
                        c.Item().Text($"Tel: {order.ClientPhone}").FontSize(9);
                });
            });

            // Profiles section header
            if (order.Items.Count > 0)
            {
                col.Item().PaddingBottom(5).Text($"PROFILES ({order.Items.Count} items)")
                    .FontSize(11).SemiBold().FontColor(Colors.Blue.Darken2);
                col.Item().Element(c => ComposeTable(c, order));
            }

            // Accessories section
            if (order.AccessoryItems.Count > 0)
            {
                col.Item().PaddingTop(15).PaddingBottom(5)
                    .Text($"ACCESSORIES & HARDWARE ({order.AccessoryItems.Count} items)")
                    .FontSize(11).SemiBold().FontColor(Colors.Teal.Darken2);
                col.Item().Element(c => ComposeAccessoriesTable(c, order));
            }

            // Cortizo total reference
            if (order.CortizoTotal > 0)
            {
                col.Item().PaddingTop(10).Background(Colors.Blue.Lighten5).Padding(8).Row(r =>
                {
                    r.RelativeItem().Text("Cortizo Estimate Total:").SemiBold().FontSize(10);
                    r.AutoItem().Text($"{order.CortizoTotal:N2} {order.Currency}").Bold().FontSize(11)
                        .FontColor(Colors.Blue.Darken3);
                });
            }

            // Totals
            col.Item().PaddingTop(10).AlignRight().Width(250).Column(totals =>
            {
                totals.Item().Row(r =>
                {
                    r.RelativeItem().Text("Subtotal:").SemiBold();
                    r.AutoItem().Text($"{order.Subtotal:N2} {order.Currency}");
                });
                totals.Item().Row(r =>
                {
                    r.RelativeItem().Text($"VAT ({order.VatRate}%):").SemiBold();
                    r.AutoItem().Text($"{order.VatAmount:N2} {order.Currency}");
                });
                totals.Item().PaddingTop(5).BorderTop(1).BorderColor(Colors.Black).Row(r =>
                {
                    r.RelativeItem().Text("TOTAL:").Bold().FontSize(12);
                    r.AutoItem().Text($"{order.Total:N2} {order.Currency}").Bold().FontSize(12);
                });
            });

            // Unfilled items warning section (from automation)
            if (order.UnfilledProfiles.Count > 0 || order.UnfilledAccessories.Count > 0)
            {
                col.Item().PaddingTop(15).Column(unfilled =>
                {
                    unfilled.Item().Background(Colors.Red.Lighten4).Padding(10).Column(warning =>
                    {
                        warning.Item().Text("ITEMS REQUIRING MANUAL REVIEW").FontSize(11).Bold().FontColor(Colors.Red.Darken3);
                        warning.Item().PaddingTop(5).Text("The following items could not be automatically calculated and need manual verification:")
                            .FontSize(9).FontColor(Colors.Red.Darken2);
                        
                        if (order.UnfilledProfiles.Count > 0)
                        {
                            warning.Item().PaddingTop(8).Text($"PROFILES ({order.UnfilledProfiles.Count} items):").FontSize(9).SemiBold();
                            foreach (var item in order.UnfilledProfiles)
                            {
                                warning.Item().Text($"  - Row {item.RowNumber}: REF {item.RefNumber} x {item.Amount} - {item.Description}")
                                    .FontSize(8).FontColor(Colors.Red.Darken1);
                                warning.Item().Text($"    Reason: {item.Reason}").FontSize(7).Italic().FontColor(Colors.Grey.Darken1);
                            }
                        }
                        
                        if (order.UnfilledAccessories.Count > 0)
                        {
                            warning.Item().PaddingTop(8).Text($"ACCESSORIES ({order.UnfilledAccessories.Count} items):").FontSize(9).SemiBold();
                            foreach (var item in order.UnfilledAccessories)
                            {
                                warning.Item().Text($"  - Row {item.RowNumber}: REF {item.RefNumber} x {item.Amount} - {item.Description}")
                                    .FontSize(8).FontColor(Colors.Red.Darken1);
                                warning.Item().Text($"    Reason: {item.Reason}").FontSize(7).Italic().FontColor(Colors.Grey.Darken1);
                            }
                        }
                    });
                });
            }

            // Missing calculation items section (for manual price entry)
            if (order.MissingCalculationItems.Count > 0)
            {
                col.Item().PaddingTop(15).Element(c => ComposeMissingItemsTable(c, order));
            }

            // Terms and conditions
            col.Item().PaddingTop(20).Column(terms =>
            {
                terms.Item().Text("TERMS & CONDITIONS").FontSize(11).SemiBold().FontColor(Colors.Blue.Darken2);
                terms.Item().PaddingTop(5).Row(r =>
                {
                    r.AutoItem().Text("Delivery: ").SemiBold().FontSize(9);
                    r.AutoItem().Text($"{order.DeliveryTerms} - Estimated {order.DeliveryDays} working days").FontSize(9);
                });
                terms.Item().Row(r =>
                {
                    r.AutoItem().Text("Payment: ").SemiBold().FontSize(9);
                    r.AutoItem().Text(order.PaymentTerms).FontSize(9);
                });
                if (!string.IsNullOrEmpty(order.Notes))
                {
                    terms.Item().PaddingTop(5).Text("Notes:").SemiBold().FontSize(9);
                    terms.Item().Text(order.Notes).FontSize(9);
                }
            });
        });
    }

    private void ComposeTable(IContainer container, VisorQuotationOrder order)
    {
        container.Table(table =>
        {
            // Define columns
            table.ColumnsDefinition(cols =>
            {
                cols.ConstantColumn(30);  // #
                cols.ConstantColumn(60);  // Ref
                cols.RelativeColumn(3);   // Description
                cols.ConstantColumn(80);  // Finish/Shade
                cols.ConstantColumn(40);  // Qty
                cols.ConstantColumn(50);  // Length
                cols.ConstantColumn(60);  // Unit Price
                cols.ConstantColumn(70);  // Total
            });

            // Header
            table.Header(header =>
            {
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("#").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("REF").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("DESCRIPTION").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("FINISH").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("QTY").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("LENGTH").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("PRICE").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(Colors.Blue.Darken3).Padding(5)
                    .Text("TOTAL").FontColor(Colors.White).FontSize(9).SemiBold();
            });

            // Rows
            foreach (var item in order.Items)
            {
                var bgColor = item.LineNumber % 2 == 0 ? Colors.Grey.Lighten4 : Colors.White;

                table.Cell().Background(bgColor).Padding(4).Text(item.LineNumber.ToString()).FontSize(8);
                table.Cell().Background(bgColor).Padding(4).Text(item.RefNumber).FontSize(8).SemiBold();
                table.Cell().Background(bgColor).Padding(4).Text(item.Description).FontSize(8);
                table.Cell().Background(bgColor).Padding(4).Text($"{item.Finish}\n{item.Shade}").FontSize(7);
                table.Cell().Background(bgColor).Padding(4).AlignRight().Text(item.Quantity.ToString()).FontSize(8);
                table.Cell().Background(bgColor).Padding(4).AlignRight().Text($"{item.Length:N1}m").FontSize(8);
                table.Cell().Background(bgColor).Padding(4).AlignRight()
                    .Text(item.UnitPrice > 0 ? $"{item.UnitPrice:N2}" : "-").FontSize(8);
                table.Cell().Background(bgColor).Padding(4).AlignRight()
                    .Text(item.TotalPrice > 0 ? $"{item.TotalPrice:N2}" : "-").FontSize(8);
            }
        });
    }

    private void ComposeAccessoriesTable(IContainer container, VisorQuotationOrder order)
    {
        var headerBg = Colors.Teal.Darken2;

        container.Table(table =>
        {
            table.ColumnsDefinition(cols =>
            {
                cols.ConstantColumn(30);  // #
                cols.ConstantColumn(70);  // Ref
                cols.RelativeColumn(4);   // Description
                cols.ConstantColumn(70);  // Finish
                cols.ConstantColumn(45);  // Qty
                cols.ConstantColumn(65);  // Unit Price
                cols.ConstantColumn(75);  // Total
            });

            table.Header(header =>
            {
                header.Cell().Background(headerBg).Padding(5)
                    .Text("#").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(headerBg).Padding(5)
                    .Text("REF").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(headerBg).Padding(5)
                    .Text("DESCRIPTION").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(headerBg).Padding(5)
                    .Text("FINISH").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(headerBg).Padding(5)
                    .Text("QTY").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(headerBg).Padding(5)
                    .Text("PRICE").FontColor(Colors.White).FontSize(9).SemiBold();
                header.Cell().Background(headerBg).Padding(5)
                    .Text("TOTAL").FontColor(Colors.White).FontSize(9).SemiBold();
            });

            foreach (var item in order.AccessoryItems)
            {
                var bgColor = item.LineNumber % 2 == 0 ? Colors.Grey.Lighten4 : Colors.White;

                table.Cell().Background(bgColor).Padding(4).Text(item.LineNumber.ToString()).FontSize(8);
                table.Cell().Background(bgColor).Padding(4).Text(item.RefNumber).FontSize(8).SemiBold();
                table.Cell().Background(bgColor).Padding(4).Text(item.Description).FontSize(8);
                table.Cell().Background(bgColor).Padding(4).Text(item.Finish).FontSize(7);
                table.Cell().Background(bgColor).Padding(4).AlignRight().Text(item.Quantity.ToString()).FontSize(8);
                table.Cell().Background(bgColor).Padding(4).AlignRight()
                    .Text(item.UnitPrice > 0 ? $"{item.UnitPrice:N2}" : "-").FontSize(8);
                table.Cell().Background(bgColor).Padding(4).AlignRight()
                    .Text(item.TotalPrice > 0 ? $"{item.TotalPrice:N2}" : "-").FontSize(8);
            }
        });
    }

    private void ComposeMissingItemsTable(IContainer container, VisorQuotationOrder order)
    {
        var headerColor = Color.FromHex("E65100");
        var sectionBg = Color.FromHex("FFF3E0");
        var altRowBg = Color.FromHex("FFF8E1");

        container.Column(col =>
        {
            col.Item().Background(sectionBg).Padding(10).Column(section =>
            {
                section.Item().Text("MISSING ITEMS - MANUAL PRICE ENTRY REQUIRED")
                    .FontSize(11).Bold().FontColor(headerColor);
                section.Item().PaddingTop(3).Text("The following items were not found in the price list. Please fill in the prices manually.")
                    .FontSize(9).FontColor(Color.FromHex("BF360C"));

                section.Item().PaddingTop(8).Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.ConstantColumn(30);   // #
                        cols.ConstantColumn(65);   // Ref
                        cols.ConstantColumn(65);   // Category
                        cols.RelativeColumn(3);    // Description
                        cols.ConstantColumn(50);   // Finish
                        cols.ConstantColumn(40);   // Qty
                        cols.ConstantColumn(70);   // Price (blank for manual entry)
                    });

                    table.Header(header =>
                    {
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("#").FontColor(Colors.White).FontSize(8).SemiBold();
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("REF").FontColor(Colors.White).FontSize(8).SemiBold();
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("TYPE").FontColor(Colors.White).FontSize(8).SemiBold();
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("DESCRIPTION").FontColor(Colors.White).FontSize(8).SemiBold();
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("FINISH").FontColor(Colors.White).FontSize(8).SemiBold();
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("QTY").FontColor(Colors.White).FontSize(8).SemiBold();
                        header.Cell().Background(headerColor).Padding(4)
                            .Text("PRICE").FontColor(Colors.White).FontSize(8).SemiBold();
                    });

                    int rowNum = 1;
                    foreach (var item in order.MissingCalculationItems)
                    {
                        var bgColor = rowNum % 2 == 0 ? altRowBg : Colors.White;

                        table.Cell().Background(bgColor).Padding(3).Text(rowNum.ToString()).FontSize(7);
                        table.Cell().Background(bgColor).Padding(3).Text(item.RefNumber).FontSize(7).SemiBold();
                        table.Cell().Background(bgColor).Padding(3).Text(item.Category).FontSize(7);
                        table.Cell().Background(bgColor).Padding(3).Text(item.Description).FontSize(7);
                        table.Cell().Background(bgColor).Padding(3).Text(item.Finish).FontSize(7);
                        table.Cell().Background(bgColor).Padding(3).AlignRight().Text(item.Quantity.ToString()).FontSize(7);
                        table.Cell().Background(bgColor).Padding(3).AlignRight()
                            .Text(item.ManualPrice > 0 ? $"{item.ManualPrice:N2}" : "________").FontSize(7);

                        rowNum++;
                    }
                });

                section.Item().PaddingTop(5).Text($"Reason: {order.MissingCalculationItems.FirstOrDefault()?.Reason ?? "Not found in price list"}")
                    .FontSize(7).Italic().FontColor(Colors.Grey.Darken1);
            });
        });
    }

    private void ComposeFooter(IContainer container, VisorQuotationOrder order)
    {
        container.Column(col =>
        {
            col.Item().LineHorizontal(0.5f).LineColor(Colors.Grey.Medium);
            col.Item().PaddingTop(5).Row(row =>
            {
                row.RelativeItem().Column(c =>
                {
                    c.Item().Text($"NIPT: {order.CompanyInfo.Nipt}").FontSize(8).FontColor(Colors.Grey.Darken1);
                });
                row.RelativeItem().AlignCenter().Text(text =>
                {
                    text.Span("Page ").FontSize(8);
                    text.CurrentPageNumber().FontSize(8);
                    text.Span(" of ").FontSize(8);
                    text.TotalPages().FontSize(8);
                });
                row.RelativeItem().AlignRight().Text(order.CompanyInfo.Website).FontSize(8).FontColor(Colors.Grey.Darken1);
            });
        });
    }

    private string GenerateQuotationNumber()
    {
        // Format: VQ-YYYYMMDD-XXXX (e.g., VQ-20260217-0001)
        var datePart = DateTime.Now.ToString("yyyyMMdd");
        var randomPart = new Random().Next(1, 9999).ToString("D4");
        return $"VQ-{datePart}-{randomPart}";
    }
}

/// <summary>
/// Result from a single Cortizo line item
/// </summary>
public class CortizoLineResult
{
    public string RefNumber { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Finish1 { get; set; } = string.Empty;
    public string Shade1 { get; set; } = string.Empty;
    public string Finish2 { get; set; } = string.Empty;
    public string Shade2 { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Length { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal Amount { get; set; }
}
