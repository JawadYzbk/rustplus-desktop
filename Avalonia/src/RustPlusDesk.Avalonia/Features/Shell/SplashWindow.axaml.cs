using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.Media.Imaging;

namespace RustPlusDesk.Avalonia.Features.Shell;

/// <summary>
/// Reference the old app's splash window: the app's own artwork while it starts up.
///
/// The window says what the app is doing, so a slow start reads as work rather than as a hang.
/// </summary>
public sealed partial class SplashWindow : Window
{
    public SplashWindow()
    {
        InitializeComponent();
        VersionText.Text = $"v{RustPlusDesk.Core.Updates.AppVersion.Short(RustPlusDesk.Core.Updates.AppVersion.Raw())}";
        LoadArtwork();
    }

    /// <summary>
    /// The artwork, read from beside the app.
    ///
    /// It is a loose file rather than an embedded resource: the splash is drawn once, at start-up, and a three-megabyte
    /// picture in the assembly would be paid for on every run and on every install.
    /// </summary>
    private void LoadArtwork()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "Assets", "splash-screen.png");
        if (!File.Exists(path)) return;

        try
        {
            Artwork.Source = new Bitmap(path);
        }
        catch (Exception)
        {
            // A splash without its artwork is still a splash; the app must not fail to start over a picture.
        }
    }

    /// <summary>What the app is doing while it starts.</summary>
    public void SetStatus(string message)
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetStatus(message));
            return;
        }

        if (this.FindControl<TextBlock>("StatusText") is { } status) status.Text = message;
    }
}
