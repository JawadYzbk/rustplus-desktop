using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference CommandDockTemplates: one saved arrangement, as the list shows it.
///
/// The name is editable in place and is only written when it is committed, so a name that is being typed is not a name
/// that is being changed.
/// </summary>
public sealed partial class DockPresetRow : ObservableObject
{
    /// <summary>The arrangement itself.</summary>
    public CommandDockPreset Preset { get; }

    public DockPresetRow(CommandDockPreset preset)
    {
        Preset = preset;
        name = preset.Name;
    }

    [ObservableProperty] private string name;

    /// <summary>The arrangement's own id.</summary>
    public string Id => Preset.Id;

    /// <summary>Whether this is the built-in arrangement, which cannot be renamed or deleted.</summary>
    public bool IsBuiltIn => Preset.Id == DockViewModel.DefaultPresetId;

    /// <summary>How many tiles the arrangement holds.</summary>
    public int TileCount => Preset.Tiles.Count;

    /// <summary>When it was saved, or the fact that it was never saved because it is the built-in one.</summary>
    public string SavedText => IsBuiltIn
        ? "Always here"
        : Preset.SavedUtc.ToLocalTime().ToString("g");

    /// <summary>Where the arrangement's tiles sit, so the row can draw it.</summary>
    public IReadOnlyList<DockPresetPreviewRect> PreviewRects
    {
        get
        {
            var (rects, width, height) = DockPresetPreview.Build(Preset);
            if (rects.Count == 0 || width <= 0 || height <= 0) return rects;

            // The diagram is drawn in a fixed box, so the arrangement is scaled to fit it. The shapes stay the same;
            // only the size changes.
            double scale = Math.Min(PreviewBox / width, PreviewBox / height);
            return rects
                .Select(rect => rect with { X = rect.X * scale, Y = rect.Y * scale, Width = rect.Width * scale, Height = rect.Height * scale })
                .ToList();
        }
    }

    /// <summary>How wide the arrangement is, in its own pixels.</summary>
    public double PreviewWidth => DockPresetPreview.Build(Preset).Width;

    /// <summary>How high the arrangement is, in its own pixels.</summary>
    public double PreviewHeight => DockPresetPreview.Build(Preset).Height;

    /// <summary>The size of the box the diagram is drawn in.</summary>
    public const double PreviewBox = 88;
}

/// <summary>
/// Reference MiniMapWindow.CommandDock.Presets and CommandDockTemplates: the saved arrangements, and what can be done
/// with them.
///
/// The built-in arrangement is always the first row. It is not in the store - it is built on demand - so it cannot be
/// renamed away or deleted by accident, and it is the way back to the beginning.
/// </summary>
public sealed partial class DockTemplatesViewModel : ObservableObject
{
    private readonly DockViewModel dock;

    public DockTemplatesViewModel(DockViewModel dock)
    {
        this.dock = dock;
    }

    /// <summary>The arrangements, the built-in one first.</summary>
    public ObservableCollection<DockPresetRow> Presets { get; } = [];

    /// <summary>The name a new arrangement is saved under.</summary>
    [ObservableProperty] private string newName = "";

    /// <summary>What the window says after a save, a rename or a failure.</summary>
    [ObservableProperty] private string status = "";

    /// <summary>Reads the arrangements and shows them.</summary>
    [RelayCommand]
    public async Task RefreshAsync()
    {
        Presets.Clear();
        foreach (var preset in await dock.LoadPresetsAsync(CancellationToken.None).ConfigureAwait(true))
        {
            Presets.Add(new DockPresetRow(preset));
        }
    }

    /// <summary>Reference SavePreset: the arrangement on the dock is stored under a name.</summary>
    [RelayCommand]
    private async Task SaveAsync()
    {
        string name = NewName.Trim();
        if (name.Length == 0)
        {
            Status = "Give the arrangement a name first.";
            return;
        }

        await dock.SavePresetAsync(name, CancellationToken.None);
        NewName = "";
        Status = $"Saved \"{name}\".";
        await RefreshAsync();
    }

    /// <summary>Reference ApplyPreset: the arrangement is put in place.</summary>
    [RelayCommand]
    private async Task ApplyAsync(DockPresetRow? row)
    {
        if (row is null) return;

        await dock.ApplyPresetAsync(row.Id, CancellationToken.None);
        Status = $"Applied \"{row.Preset.Name}\".";
    }

    /// <summary>Reference RenamePreset: the arrangement is given another name.</summary>
    [RelayCommand]
    private async Task RenameAsync(DockPresetRow? row)
    {
        if (row is null || row.IsBuiltIn) return;

        string name = row.Name.Trim();
        if (name.Length == 0)
        {
            Status = "Give the arrangement a name first.";
            return;
        }

        await dock.RenamePresetAsync(row.Id, name, CancellationToken.None);
        Status = $"Renamed to \"{name}\".";
        await RefreshAsync();
    }

    /// <summary>Reference DeletePreset: the arrangement is removed.</summary>
    [RelayCommand]
    private async Task DeleteAsync(DockPresetRow? row)
    {
        if (row is null || row.IsBuiltIn) return;

        await dock.DeletePresetAsync(row.Id, CancellationToken.None);
        Status = $"Removed \"{row.Preset.Name}\".";
        await RefreshAsync();
    }
}
