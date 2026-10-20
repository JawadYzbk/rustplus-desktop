using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace RustPlusDesk.Avalonia.Features.Achievements;

/// <summary>
/// Reference the "all achievements unlocked" prompt: the app asks, once, whether to tell the developers. It does not
/// fade away, because it is a question and a prompt that vanishes is a question that is never answered; closing the
/// window counts as "no".
/// </summary>
public sealed partial class AchievementTicketPrompt : Window
{
    public AchievementTicketPrompt()
    {
        InitializeComponent();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void Yes(object? sender, global::Avalonia.Interactivity.RoutedEventArgs e)
    {
        Accepted = true;
        Close();
    }

    private void No(object? sender, global::Avalonia.Interactivity.RoutedEventArgs e)
    {
        Accepted = false;
        Close();
    }

    /// <summary>True when the user asked for the ticket to be sent.</summary>
    public bool Accepted { get; private set; }
}
