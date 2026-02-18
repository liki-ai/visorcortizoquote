using System.Text.Json.Serialization;
using VisorQuotationWebApp.Hubs;
using VisorQuotationWebApp.Models;
using VisorQuotationWebApp.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddSignalR()
    .AddJsonProtocol(options =>
    {
        options.PayloadSerializerOptions.Converters.Add(new JsonStringEnumConverter());
    });

// Configure Cortizo settings
var cortizoConfig = builder.Configuration.GetSection("Cortizo").Get<AutomationConfig>() ?? new AutomationConfig();
builder.Services.AddSingleton(cortizoConfig);

// Register services
builder.Services.AddScoped<PdfParseService>();
builder.Services.AddScoped<CortizoAutomationService>();
builder.Services.AddScoped<VisorQuotationService>();
builder.Services.AddSingleton<ExcelPriceService>(); // Singleton to cache loaded prices

// Add session support for storing parsed PDF data
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
}
app.UseStaticFiles();

app.UseRouting();
app.UseSession();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapHub<AutomationHub>("/automationHub");

app.Run();
