using System.Runtime.InteropServices;
using Avalonia.Controls;
using Avalonia.Interactivity;
using RustPlusDesk.Core.Crosshair;

namespace RustPlusDesk.Avalonia.Features.Crosshair;

/// <summary>
/// Reference CrosshairWindow: a small transparent window that lies over the game and shows the crosshair. On Windows
/// it is click-through, so the game keeps the mouse; the other platforms cannot make a click-through window without
/// platform code, so the crosshair is drawn but the mouse belongs to it there.
/// </summary>
public sealed class CrosshairWindow : Window
{
    private const int GwlExStyle = -20;
    private const int WsExTransparent = 0x00000020;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExNoActivate = 0x08000000;

    private CrosshairView view = new();

    public CrosshairWindow()
    {
        WindowDecorations = global::Avalonia.Controls.WindowDecorations.None;
        Background = null;
        TransparencyLevelHint = new[] { global::Avalonia.Controls.WindowTransparencyLevel.Transparent };
        ShowInTaskbar = false;
        Topmost = true;
        CanResize = false;
        ShowActivated = false;
        Content = view;
    }

    /// <summary>Reference SetStyle: the window takes the size of the style it shows.</summary>
    public void SetStyle(CrosshairStyle style, global::Avalonia.Media.Imaging.Bitmap? customImage = null)
    {
        view.Style = style;
        view.CustomImage = customImage;
        var (width, height) = CrosshairCatalogue.Size(style);
        Width = width;
        Height = height;
    }

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        if (!OperatingSystem.IsWindows()) return;
        try
        {
            IntPtr handle = TryGetNativeHandle();
            if (handle == IntPtr.Zero) return;
            long styles = GetWindowLongPtr(handle, GwlExStyle).ToInt64();
            SetWindowLongPtr(handle, GwlExStyle, new IntPtr(styles | WsExTransparent | WsExNoActivate | WsExToolWindow));
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException)
        {
            // Without the Win32 call the crosshair still shows, it just takes the mouse.
        }
    }

    private IntPtr TryGetNativeHandle()
    {
        if (TryGetPlatformHandle() is not { } handle) return IntPtr.Zero;
        return handle.Handle;
    }

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr(IntPtr handle, int index);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr(IntPtr handle, int index, IntPtr value);
}
