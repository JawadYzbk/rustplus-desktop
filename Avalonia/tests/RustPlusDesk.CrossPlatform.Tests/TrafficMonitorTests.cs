using System.Net;
using global::Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Diagnostics;
using RustPlusDesk.Core.Diagnostics;
using RustPlusDesk.Infrastructure.Diagnostics;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference NetworkTrafficMonitor and TrafficTrackingHttpMessageHandler: the app's own traffic, counted and shown,
/// with a rolling speed and a breakdown per category.
/// </summary>
[TestClass]
public sealed class TrafficMonitorTests
{
    [TestMethod]
    public void Recording_AddsTheRequest_AndTheTotals()
    {
        var monitor = new NetworkTrafficMonitor();
        monitor.Record("Cloud API", TrafficDirection.Both, "GET /api/v1/state", 1200, 300, 42, "200 OK");

        Assert.AreEqual(1, monitor.TotalRequestsCount);
        Assert.AreEqual(1200, monitor.TotalInboundBytes);
        Assert.AreEqual(300, monitor.TotalOutboundBytes);
        Assert.AreEqual(1500, monitor.TotalBytes);
        Assert.AreEqual("1.5 KB", monitor.FormattedTotalBytes);

        var entry = monitor.Entries.Single();
        Assert.AreEqual("Cloud API", entry.Category);
        Assert.AreEqual("⇅", entry.DirectionText);
        Assert.AreEqual("1.5 KB", entry.FormattedSize);
        Assert.AreEqual(42, entry.DurationMs);
    }

    [TestMethod]
    public void TheBreakdown_SaysWhereTheBytesWent()
    {
        var monitor = new NetworkTrafficMonitor();
        monitor.Record("Cloud API", TrafficDirection.Both, "GET /a", 900, 0);
        monitor.Record("Steam", TrafficDirection.Inbound, "GET /b", 100, 0);

        var categories = monitor.CategoryBreakdown;
        Assert.AreEqual(2, categories.Count);
        Assert.AreEqual("Cloud API", categories[0].Category, "The biggest category comes first.");
        Assert.AreEqual(1, categories[0].RequestCount);
        Assert.AreEqual(90.0, categories[0].PercentageOfTotal, 0.1);
        Assert.AreEqual("90.0%", categories[0].FormattedPercentage);
        Assert.AreEqual("10.0%", categories[1].FormattedPercentage);
    }

    [TestMethod]
    public void TheLogIsBounded()
    {
        var monitor = new NetworkTrafficMonitor();
        for (int i = 0; i < NetworkTrafficMonitor.MaxLogEntries + 10; i++)
            monitor.Record("HTTP", TrafficDirection.Inbound, "GET /x", 1, 0);
        Assert.AreEqual(NetworkTrafficMonitor.MaxLogEntries, monitor.Entries.Count);
    }

    [TestMethod]
    public void ClearingStartsTheCountersOver()
    {
        var monitor = new NetworkTrafficMonitor();
        monitor.Record("HTTP", TrafficDirection.Inbound, "GET /x", 100, 0);
        monitor.Clear();
        Assert.AreEqual(0, monitor.TotalRequestsCount);
        Assert.AreEqual(0, monitor.TotalBytes);
        Assert.AreEqual(0, monitor.Entries.Count);
        Assert.AreEqual(0, monitor.CategoryBreakdown.Count);
    }

    [TestMethod]
    public void TheSpeedIsARollingWindow()
    {
        var monitor = new NetworkTrafficMonitor();
        monitor.Record("HTTP", TrafficDirection.Inbound, "GET /x", 1500, 0);
        monitor.TickSpeed();
        Assert.AreEqual(1000, monitor.CurrentDownloadSpeedBps, 1, "1500 bytes over a second and a half.");
        Assert.AreEqual("1000 B/s", monitor.FormattedDownloadSpeed);
        Assert.AreEqual("8 Kbps", monitor.FormattedDownloadSpeedBits);
    }

    [TestMethod]
    public void SizesAreFormattedTheReferenceWay()
    {
        Assert.AreEqual("0 B", NetworkTrafficMonitor.FormatBytes(0));
        Assert.AreEqual("512 B", NetworkTrafficMonitor.FormatBytes(512));
        Assert.AreEqual("1 KB", NetworkTrafficMonitor.FormatBytes(1024));
        Assert.AreEqual("1.5 MB", NetworkTrafficMonitor.FormatBytes(1024 * 1536));
        Assert.AreEqual("0 bps", NetworkTrafficMonitor.FormatBits(0));
        Assert.AreEqual("8 bps", NetworkTrafficMonitor.FormatBits(1));
        Assert.AreEqual("8 Kbps", NetworkTrafficMonitor.FormatBits(1000));
        Assert.AreEqual("1 Mbps", NetworkTrafficMonitor.FormatBits(125000));
    }

    [TestMethod]
    public async Task TheHandlerRecordsEveryRequest()
    {
        var monitor = new NetworkTrafficMonitor();
        var handler = new TrafficTrackingHttpMessageHandler(new StubHandler(), monitor, "Cloud API");
        using var client = new HttpClient(handler);
        using var request = new HttpRequestMessage(HttpMethod.Get, "https://rustplusdesktop.cloud/api/v1/state");
        var response = await client.SendAsync(request);

        Assert.AreEqual(HttpStatusCode.OK, response.StatusCode);
        var entry = monitor.Entries.Single();
        Assert.AreEqual("Cloud API", entry.Category);
        Assert.AreEqual("200 OK", entry.Status);
        Assert.AreEqual(1, monitor.TotalRequestsCount);
    }

    [TestMethod]
    public async Task AFailedRequestIsRecordedAsAnError()
    {
        var monitor = new NetworkTrafficMonitor();
        var handler = new TrafficTrackingHttpMessageHandler(new FailingHandler(), monitor, "Cloud API");
        using var client = new HttpClient(handler);
        using var request = new HttpRequestMessage(HttpMethod.Get, "https://rustplusdesktop.cloud/api/v1/state");

        await Assert.ThrowsExceptionAsync<HttpRequestException>(() => client.SendAsync(request));
        var entry = monitor.Entries.Single();
        Assert.AreEqual("Error", entry.Status);
        Assert.AreEqual("↑ OUT", entry.DirectionText);
    }

    [TestMethod]
    public void TheEndpointIsShortenedForTheCloud()
    {
        Assert.AreEqual("GET /api/v1/state",
            TrafficTrackingHttpMessageHandler.FormatEndpoint("Cloud API", "GET https://rustplusdesktop.cloud/api/v1/state"));
        Assert.AreEqual("GET https://example.test/x",
            TrafficTrackingHttpMessageHandler.FormatEndpoint("HTTP", "GET https://example.test/x"));
    }

    [TestMethod]
    public void TheCategoryComesFromTheHost()
    {
        Assert.AreEqual("Cloud API", TrafficTrackingHttpMessageHandler.DetectCategoryFromUri(new Uri("https://rustplusdesktop.cloud/api")));
        Assert.AreEqual("Steam", TrafficTrackingHttpMessageHandler.DetectCategoryFromUri(new Uri("https://steamcommunity.com/x")));
        Assert.AreEqual("HTTP", TrafficTrackingHttpMessageHandler.DetectCategoryFromUri(new Uri("https://example.test/x")));
    }

    private sealed class StubHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) });
    }

    private sealed class FailingHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            throw new HttpRequestException("The host is unreachable.");
    }

    [TestMethod]
    public async Task TheTrafficWindow_RendersTheLog() => await UiTestSession.Current.Dispatch(async () =>
    {
        var monitor = new NetworkTrafficMonitor();
        monitor.Record("Cloud API", TrafficDirection.Both, "GET /api/v1/state", 1500, 300, 42, "200 OK");
        var model = new TrafficMonitorViewModel(monitor);
        model.Refresh();

        var window = new RustPlusDesk.Avalonia.Features.Diagnostics.TrafficMonitorWindow { DataContext = model, Width = 900, Height = 600 };
        try
        {
            window.Show();
            window.UpdateLayout();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-traffic.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        return true;
    }, default);
}
