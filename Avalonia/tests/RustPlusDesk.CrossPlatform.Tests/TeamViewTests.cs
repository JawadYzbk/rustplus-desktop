using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class TeamViewTests
{
    [TestMethod]
    public async Task MinimumLayout_TogglesLiveTeamWithoutCoveringActivity_AndDisconnectClearsRoster()
    {
        using var profiles = new DeviceProfileEditorTests.Fixture();
        await UiTestSession.Current.Dispatch(async () =>
        {
            var source = new DeviceImportTests.TeamSession
            {
                Members = new ServerTeamMember[] { new("76561198000000001", "Self", true, true), new("76561198000000002", new string('W', 500), false, false) }
                    .Concat(Enumerable.Range(3, 22).Select(index => new ServerTeamMember((76561198000000000ul + (ulong)index).ToString(), "Member " + index, false, true))).ToArray()
            };
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), profiles.Path, new Factory(source));
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark };
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                model.Connection.ConnectCommand.Execute(null);
                window.Show();
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                var team = window.GetVisualDescendants().OfType<TeamView>().Single();
                Assert.AreEqual(24, model.Connection.Team.Members.Count);
                team.GetVisualDescendants().OfType<Button>().Single().Command!.Execute(null);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-team-minimum.png"), PngBitmapEncoderOptions.Default);
                }
                var point = team.TranslatePoint(default, window)!.Value;
                var activity = window.GetVisualDescendants().OfType<Border>().Single(border => border.Name == "ActivityPanel");
                Assert.IsTrue(point.X >= 0 && point.X + team.Bounds.Width <= window.ClientSize.Width);
                Assert.IsTrue(point.Y >= 0 && point.Y + team.Bounds.Height <= activity.TranslatePoint(default, window)!.Value.Y);
                Assert.IsTrue(team.Bounds.Height > 50);
                var workspace = window.GetVisualDescendants().OfType<MapWorkspace>().Single();
                ((Grid)workspace.Content!).RowDefinitions[2].Height = new GridLength(400);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-team-large-activity.png"), PngBitmapEncoderOptions.Default);
                }
                point = team.TranslatePoint(default, window)!.Value;
                Assert.IsTrue(point.Y >= 0 && point.Y + team.Bounds.Height <= activity.TranslatePoint(default, window)!.Value.Y);
                await model.Connection.StopAsync();
                Assert.AreEqual(0, model.Connection.Team.Members.Count);
                Assert.IsTrue(source.Disposed);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
    }
    private sealed class Picker : IProfileFilePicker { public Task<string?> PickAsync() => Task.FromResult<string?>(null); }
    private sealed class Factory(IServerSession session) : IServerSessionFactory
    { public Task<IServerSession> ConnectAsync(string path, string key, CancellationToken ct) => Task.FromResult(session); }
}
