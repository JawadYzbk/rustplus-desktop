using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Map;

public sealed partial class LiveMapView : UserControl
{
    public LiveMapView()
    {
        InitializeComponent();
        TerrainView.NavigationStarted += (_, _) => StopFollowing();
    }
    private void StopFollowing() { if (DataContext is MapViewModel model) model.FollowSelectedMember = false; }
    private MapSurface TerrainView => this.FindControl<MapSurface>("Terrain")!;
    private void ZoomIn(object? sender, RoutedEventArgs e) { StopFollowing(); TerrainView.ZoomBy(1.25); }
    private void ZoomOut(object? sender, RoutedEventArgs e) { StopFollowing(); TerrainView.ZoomBy(.8); }
    private void Fit(object? sender, RoutedEventArgs e) { StopFollowing(); TerrainView.ResetView(); }
}
