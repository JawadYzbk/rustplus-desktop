using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using RustPlusDesk.Infrastructure.Items;

namespace RustPlusDesk.Avalonia.Shared.Controls;

/// <summary>Shares the reference app's bundled item artwork without a network dependency.</summary>
public sealed class ItemIcon : Image
{
    public static readonly StyledProperty<string?> ShortNameProperty = AvaloniaProperty.Register<ItemIcon, string?>(nameof(ShortName));
    public string? ShortName { get => GetValue(ShortNameProperty); set => SetValue(ShortNameProperty, value); }
    public static readonly StyledProperty<int?> ItemIdProperty = AvaloniaProperty.Register<ItemIcon, int?>(nameof(ItemId));
    public int? ItemId { get => GetValue(ItemIdProperty); set => SetValue(ItemIdProperty, value); }
    public static readonly StyledProperty<string?> DefaultShortNameProperty = AvaloniaProperty.Register<ItemIcon, string?>(nameof(DefaultShortName));
    public string? DefaultShortName { get => GetValue(DefaultShortNameProperty); set => SetValue(DefaultShortNameProperty, value); }
    private static readonly Dictionary<string, Bitmap?> Cache = new(StringComparer.Ordinal);
    static ItemIcon()
    {
        ShortNameProperty.Changed.AddClassHandler<ItemIcon>((control, _) => control.Refresh());
        ItemIdProperty.Changed.AddClassHandler<ItemIcon>((control, _) => control.Refresh());
        DefaultShortNameProperty.Changed.AddClassHandler<ItemIcon>((control, _) => control.Refresh());
    }
    private void Refresh() => Source = (ItemId is { } id && BundledItemCatalog.Icons.TryGetValue(id, out var icon) ? Load(icon.ShortName) : null)
        ?? Load(ShortName) ?? Load(DefaultShortName);

    internal static Bitmap? Load(string? shortName)
    {
        if (string.IsNullOrWhiteSpace(shortName)) return null;
        if (Cache.TryGetValue(shortName, out var ready)) return ready;
        var uri = new Uri("avares://RustPlusDesk.Avalonia/Assets/Items/" + Uri.EscapeDataString(shortName) + ".png");
        if (!AssetLoader.Exists(uri)) return Cache[shortName] = null;
        using var stream = AssetLoader.Open(uri);
        return Cache[shortName] = Bitmap.DecodeToWidth(stream, 40);
    }
}
