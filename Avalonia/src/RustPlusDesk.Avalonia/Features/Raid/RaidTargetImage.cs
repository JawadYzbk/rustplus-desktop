using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace RustPlusDesk.Avalonia.Features.Raid;

/// <summary>Raid target artwork from the bundled reference icon pack (Assets/RaidTargets).</summary>
public sealed class RaidTargetImage : Image
{
    public static readonly StyledProperty<string?> IconNameProperty = AvaloniaProperty.Register<RaidTargetImage, string?>(nameof(IconName));
    public string? IconName { get => GetValue(IconNameProperty); set => SetValue(IconNameProperty, value); }
    private static readonly Dictionary<string, Bitmap?> Cache = new(StringComparer.Ordinal);
    static RaidTargetImage() => IconNameProperty.Changed.AddClassHandler<RaidTargetImage>((image, _) => image.Source = Load(image.IconName));

    internal static Bitmap? Load(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        if (Cache.TryGetValue(name, out var ready)) return ready;
        var uri = new Uri("avares://RustPlusDesk.Avalonia/Assets/RaidTargets/" + Uri.EscapeDataString(name));
        if (!AssetLoader.Exists(uri)) return Cache[name] = null;
        using var stream = AssetLoader.Open(uri);
        return Cache[name] = new Bitmap(stream);
    }
}
