using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.MiniMap;

/// <summary>
/// Reference MiniMapWindow: the mini-map is a window of its own, holding the map in the shape the settings ask for.
///
/// It is deliberately not owned by the main window: an owned window hides when its owner is minimised, and the whole
/// point of the mini-map is to be up while the app is not. The cost of that is that nothing takes it down when the
/// app goes, which is why the shell closes it with the rest of the app.
/// </summary>
public sealed partial class MiniMapWindow : Window
{
    /// <summary>Reference MiniMapWindow: the mini-map keeps a margin from the screen edge the first time it opens.</summary>
    private const double EdgeMargin = 20;

    private MiniMapViewModel? model;

    /// <summary>Reference MiniMapWindow: the window is built, and the settings panel is put in its popup.</summary>
    public MiniMapWindow()
    {
        InitializeComponent();
        // Reference MiniMapWindow: no frame, no title bar and a see-through background, so the map is a shape on the
        // desktop rather than a window sitting on top of it.
        WindowDecorations = WindowDecorations.None;
        TransparencyLevelHint = new[] { WindowTransparencyLevel.Transparent };
        CanResize = false;
        ShowInTaskbar = true;
        // Not permanently on top: the reference's mini-map is an ordinary window, so it can be pushed behind another
        // program and does not sit over the game's own interface for the whole of a session.
        Topmost = false;
        var panel = new MiniMapSettingsPanel();
        panel.CloseRequested += (_, _) => SettingsPopup.Close();
        SettingsPopup.Child = panel;
        SettingsPopup.PlacementTarget = Shape;
        PositionChanged += (_, _) => SavePlace();
    }

    /// <summary>How often the open mini-map was brought forward, for tests.</summary>
    internal int FocusCount { get; private set; }

    /// <summary>
    /// Reference BtnToggleMiniMap_Click: asking for the mini-map again brings the open one forward instead of
    /// putting a second copy on the screen.
    /// </summary>
    public void BringToFront()
    {
        FocusCount++;
        Activate();
    }

    /// <summary>Reference MiniMapWindow: the window shows the model it is given, and follows its settings.</summary>
    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (model is not null) model.PropertyChanged -= ModelChanged;
        model = DataContext as MiniMapViewModel;
        if (model is null) return;
        model.PropertyChanged += ModelChanged;
        ApplySize();
        ApplyPlace();
    }

    /// <summary>Reference MiniMapWindow.UpdateSize: the window is the size of the map it shows.</summary>
    private void ApplySize()
    {
        if (model is null) return;
        Width = model.Size;
        Height = model.Height;
    }

    /// <summary>
    /// Reference MiniMapWindow: the mini-map comes back where it was left, and the first time it opens it sits at the
    /// top right of the working area.
    /// </summary>
    private void ApplyPlace()
    {
        if (model is null) return;
        if (model.SavedPosition is { } place)
        {
            Position = new PixelPoint(place.Left, place.Top);
            return;
        }
        PlaceFirstTime();
    }

    /// <summary>The first time it opens, the mini-map sits at the top right of the working area, a margin from the edge.</summary>
    private void PlaceFirstTime()
    {
        var screens = Screens.All;
        if (screens.Count == 0) return;
        var screen = screens.FirstOrDefault(item => item.IsPrimary) ?? screens[0];
        var area = screen.WorkingArea;
        double scaling = screen.Scaling <= 0 ? 1 : screen.Scaling;
        int width = (int)Math.Round(Width * scaling);
        int margin = (int)Math.Round(EdgeMargin * scaling);
        Position = new PixelPoint(area.Right - width - margin, area.Y + margin);
    }

    /// <summary>Reference MiniMapWindow.UpdateSize: the window follows the size and the shape the settings say.</summary>
    private void ModelChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MiniMapViewModel.Size) or nameof(MiniMapViewModel.ShapeIndex)) ApplySize();
    }

    /// <summary>Reference MiniMapWindow.SaveDockPosition: the place is written when the mini-map is moved.</summary>
    private void SavePlace()
    {
        if (model is null) return;
        _ = SavePlaceAsync();
    }

    /// <summary>The place the mini-map sits now is written to the cache file.</summary>
    private async Task SavePlaceAsync()
    {
        try { await model!.SavePositionAsync(Position.X, Position.Y).ConfigureAwait(false); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    /// <summary>Reference BtnMapSettings_Click: the settings open beside the map, and close again on a second press.</summary>
    private void OnSettingsClicked(object? sender, RoutedEventArgs e) => SettingsPopup.IsOpen = !SettingsPopup.IsOpen;

    /// <summary>
    /// Reference MiniMapWindow: the map itself is the handle, so the mini-map is dragged by pressing it.
    ///
    /// The map picture is not hit-testable, so the press reaches the border and the window moves - the reference does
    /// the same with its drag of the whole map.
    /// </summary>
    private void OnShapePointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (e.Handled || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        try { BeginMoveDrag(e); }
        catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException) { }
        e.Handled = true;
    }
}
