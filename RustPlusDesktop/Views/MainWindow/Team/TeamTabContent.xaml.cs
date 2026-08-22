using System.Windows;
using System.Windows.Controls;

namespace RustPlusDesk.Views;

/// <summary>
/// Extracted module for the Team tab content (formerly inline in MainWindow.xaml).
/// Forwards UI events to the owning MainWindow partial implementations.
/// </summary>
public partial class TeamTabContent : UserControl
{
    /// <summary>Owning window; set by MainWindow after InitializeComponent.</summary>
    internal MainWindow? Main { get; set; }

    public TeamTabContent()
    {
        InitializeComponent();
    }

    private void Team_Center_Click(object sender, RoutedEventArgs e) => Main?.Team_Center_Click(sender, e);

    private void Team_Follow_Click(object sender, RoutedEventArgs e) => Main?.Team_Follow_Click(sender, e);

    private void Team_OpenProfile_Click(object sender, RoutedEventArgs e) => Main?.Team_OpenProfile_Click(sender, e);

    private void Team_Promote_Click(object sender, RoutedEventArgs e) => Main?.Team_Promote_Click(sender, e);

    private void TeamCheckBox_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        => Main?.TeamCheckBox_PreviewMouseLeftButtonDown(sender, e);

    private void TeamItem_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        => Main?.TeamItem_MouseLeftButtonUp(sender, e);

    private void ChkProfileMarkers_Toggled(object? sender, RoutedEventArgs e) => Main?.ChkProfileMarkers_Toggled(sender, e);

    private void ChkPlayerArrows_Toggled(object? sender, RoutedEventArgs e) => Main?.ChkPlayerArrows_Toggled(sender, e);

    private void ChkDeathMarkers_Toggled(object? sender, RoutedEventArgs e) => Main?.ChkDeathMarkers_Toggled(sender, e);

    private void BtnDeathMarkerSettings_Click(object sender, RoutedEventArgs e) => Main?.BtnDeathMarkerSettings_Click(sender, e);

    private void SliderPlayerIconSize_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        => Main?.SliderPlayerIconSize_ValueChanged(sender, e);

    private void BtnAbbreviateNames_Toggled(object sender, RoutedEventArgs e) => Main?.BtnAbbreviateNames_Toggled(sender, e);
}
