using Microsoft.AspNetCore.SignalR;
using VisorQuotationWebApp.Models;

namespace VisorQuotationWebApp.Hubs;

/// <summary>
/// SignalR hub for real-time automation progress updates
/// </summary>
public class AutomationHub : Hub
{
    public async Task SendLog(AutomationLogEntry logEntry)
    {
        await Clients.All.SendAsync("ReceiveLog", logEntry);
    }

    public async Task SendProgress(int current, int total, string status)
    {
        await Clients.All.SendAsync("ReceiveProgress", current, total, status);
    }

    public async Task SendComplete(AutomationRunResult result)
    {
        await Clients.All.SendAsync("ReceiveComplete", result);
    }
}
