using System.Windows;
using System.Windows.Controls;

namespace RustPlusDesk.Views;

/// <summary>
/// Extracted module for the Clan tab content (formerly inline in MainWindow.xaml).
/// Forwards UI events to the owning MainWindow partial implementations.
/// </summary>
public partial class ClanTabContent : UserControl
{
    /// <summary>Owning window; set by MainWindow after InitializeComponent.</summary>
    internal MainWindow? Main { get; set; }

    public ClanTabContent()
    {
        InitializeComponent();
    }

    private void BtnOpenClanChat_Click(object sender, RoutedEventArgs e) => Main?.BtnOpenClanChat_Click(sender, e);

    private void BtnRefreshClan_Click(object sender, RoutedEventArgs e) => Main?.BtnRefreshClan_Click(sender, e);

    private void BtnToggleClanView_Click(object sender, RoutedEventArgs e) => Main?.BtnToggleClanView_Click(sender, e);

    private void Clan_CopySteamId_Click(object sender, RoutedEventArgs e) => Main?.Clan_CopySteamId_Click(sender, e);

    private void Clan_OpenProfile_Click(object sender, RoutedEventArgs e) => Main?.Clan_OpenProfile_Click(sender, e);
}
