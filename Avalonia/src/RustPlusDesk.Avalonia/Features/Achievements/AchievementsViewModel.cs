using System.Collections.ObjectModel;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Achievements;

namespace RustPlusDesk.Avalonia.Features.Achievements;

/// <summary>One achievement as the list shows it.</summary>
public sealed record AchievementRow(AchievementDef Definition, bool IsEarned, bool IsUnseen, DateTime? EarnedAt, IImage? Icon)
{
    /// <summary>A secret achievement shows as "???" until it is earned, then reveals its real name.</summary>
    public string Name => Definition.IsSecret && !IsEarned ? "???" : Definition.Name;
    public bool IsSecret => Definition.IsSecret;
    public string Detail => IsEarned
        ? $"Earned {EarnedAt?.ToLocalTime().ToString("g", System.Globalization.CultureInfo.CurrentCulture) ?? ""}".TrimEnd()
        : "Not earned yet";
}

/// <summary>
/// Reference AchievementsWindow and AchievementService: the list of every achievement, the badge for the ones that
/// have not been looked at, and the sheet the artwork is cut from.
/// </summary>
public sealed partial class AchievementsViewModel(AchievementTracker tracker) : ObservableObject
{
    private const string SheetUri = "avares://RustPlusDesk.Avalonia/Assets/achievements/achievements_sheet.png";

    private static Bitmap? sheet;

    /// <summary>The bundled sprite sheet. A failed load is not cached, so a later host with assets still works.</summary>
    private static Bitmap? Sheet => sheet ??= LoadSheet();

    public ObservableCollection<AchievementRow> Rows { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Summary))]
    private int earnedCount;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Summary))]
    private int unseenCount;
    [ObservableProperty] private string status = "";

    public int TotalCount => AchievementCatalog.All.Count;
    public string Summary => $"{EarnedCount} of {TotalCount} earned";

    /// <summary>True once every achievement is earned (reference AchievementService.AllEarned).</summary>
    public bool IsComplete => EarnedCount >= TotalCount && TotalCount > 0;

    /// <summary>Reference AchievementService.Unlocked: the toast and the badge follow it.</summary>
    public event Action<AchievementDef>? Unlocked;

    /// <summary>Raised when the last achievement is earned, so the app can offer to tell the developers.</summary>
    public event Action? Completed;

    public AchievementsViewModel() : this(new AchievementTracker(new NullAchievementStore())) { }

    /// <summary>Reads the saved state and fills the list.</summary>
    public void Load()
    {
        tracker.Load();
        tracker.Unlocked += OnUnlocked;
        tracker.UnseenChanged += count => UnseenCount = count;
        Refresh();
    }

    /// <summary>Reference MarkAllSeen: the badge is cleared when the list has actually been looked at.</summary>
    [RelayCommand]
    public void MarkSeen()
    {
        tracker.MarkAllSeen();
        Refresh();
    }

    /// <summary>Earns an achievement (the reference calls this from wherever the thing happens).</summary>
    public void Unlock(string id)
    {
        tracker.Unlock(id);
        Refresh();
    }

    private void OnUnlocked(AchievementDef definition) => Unlocked?.Invoke(definition);

    private void Refresh()
    {
        bool wasComplete = IsComplete;
        EarnedCount = tracker.EarnedCount;
        UnseenCount = tracker.UnseenCount;
        var sheet = Sheet;
        Rows.Clear();
        foreach (var definition in AchievementCatalog.All)
        {
            bool earned = tracker.IsEarned(definition.Id);
            Rows.Add(new AchievementRow(definition, earned, tracker.IsUnseen(definition.Id), tracker.EarnedAt(definition.Id),
                earned ? Crop(sheet, definition.SheetIndex) : null));
        }
        if (!wasComplete && IsComplete) Completed?.Invoke();
    }

    /// <summary>Reference the sprite sheet: the cell of the achievement, cut out of the sheet.</summary>
    private static IImage? Crop(Bitmap? sheet, int index)
    {
        if (sheet is null) return null;
        var (x, y) = AchievementCatalog.CellFor(index);
        if (x + AchievementCatalog.CellSize > sheet.PixelSize.Width
            || y + AchievementCatalog.CellSize > sheet.PixelSize.Height) return null;
        return new CroppedBitmap(sheet, new PixelRect(x, y, AchievementCatalog.CellSize, AchievementCatalog.CellSize));
    }

    private static Bitmap? LoadSheet()
    {
        try
        {
            var uri = new Uri(SheetUri);
            if (!AssetLoader.Exists(uri)) return null;
            using var stream = AssetLoader.Open(uri);
            return Bitmap.DecodeToWidth(stream, AchievementCatalog.CellSize * AchievementCatalog.Columns);
        }
        catch (Exception)
        {
            // Without an asset loader (a headless test, a trimmed host) the list still shows, just without artwork.
            return null;
        }
    }

    /// <summary>The tracker without a saved state, for a preview that has nowhere to write.</summary>
    private sealed class NullAchievementStore : IAchievementStore
    {
        public AchievementState Load() => new();
        public void Save(AchievementState state) { }
    }
}
