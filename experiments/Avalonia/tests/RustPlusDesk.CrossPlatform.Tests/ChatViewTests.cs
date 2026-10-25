using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ChatViewTests
{
    [TestMethod]
    public async Task Chat_RendersAtNarrowPanelSize_AndKeyboardSendAndSearchUseRealBindings()
    {
        await using var session = HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder));
        await session.Dispatch(async () =>
        {
            var model = new ChatLaneViewModel(ChatChannel.Team);
            int sent = 0;
            var client = new TestChatClient
            {
                Read = (_, _) => Task.FromResult<IReadOnlyList<ChatMessage>>([ChatViewModelTests.Message(1)]),
                Send = (_, text, _) =>
                {
                    sent++;
                    return Task.FromResult<ChatMessage?>(ChatViewModelTests.Message(2) with { Text = text });
                }
            };
            using var stop = new CancellationTokenSource();
            var running = model.RunAsync(client, stop.Token);
            var view = new ChatLaneView { DataContext = model };
            var window = new Window { Content = view, Width = 460, Height = 180 };
            try
            {
                window.Show();
                foreach (var theme in new[] { ThemeVariant.Light, ThemeVariant.Dark })
                {
                    window.RequestedThemeVariant = theme;
                    using var frame = window.CaptureRenderedFrame();
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, $"ui-chat-{theme}.png"), PngBitmapEncoderOptions.Default);
                    foreach (var control in view.GetVisualDescendants().OfType<Control>()
                        .Where(c => c is Button or TextBox && c.IsEffectivelyVisible))
                    {
                        var origin = control.TranslatePoint(default, view)!.Value;
                        Assert.IsTrue(origin.X >= 0 && origin.X + control.Bounds.Width <= view.Bounds.Width + .1);
                        Assert.IsTrue(origin.Y >= 0 && origin.Y + control.Bounds.Height <= view.Bounds.Height + .1);
                    }
                }
                var input = view.GetVisualDescendants().OfType<TextBox>().Single(c => AutomationProperties.GetName(c) == "Chat message");
                Assert.AreEqual(128, input.MaxLength);
                input.Text = "Keyboard send";
                input.Focus();
                window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);
                window.KeyRelease(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);
                if (model.SendCommand.ExecutionTask is { } task) await task;
                Assert.AreEqual(1, sent);
                Assert.AreEqual("", input.Text);
                Assert.IsTrue(model.Messages.Any(m => m.Text == "Keyboard send"));
                var search = view.GetVisualDescendants().OfType<TextBox>().Single(c => AutomationProperties.GetName(c) == "Search chat history");
                search.Text = "keyboard";
                using (var filtered = window.CaptureRenderedFrame()) Assert.IsNotNull(filtered);
                Assert.AreEqual(1, model.Messages.Count);
                stop.Cancel();
                await running;
                Assert.IsFalse(input.IsEffectivelyEnabled);
            }
            finally { stop.Cancel(); await running; window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
