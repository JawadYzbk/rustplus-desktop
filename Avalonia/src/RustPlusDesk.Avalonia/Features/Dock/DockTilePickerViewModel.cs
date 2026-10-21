using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Avalonia.Shared.Icons;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// One line in the picker: either a heading or an entry.
///
/// The two are the same type on purpose. A picker that draws a heading and an entry in two different lists has to
/// interleave them, and interleaving two lists is where the order goes wrong; one list of lines cannot get out of step
/// with itself.
/// </summary>
public sealed class DockTilePickerRow
{
    /// <summary>The heading of a section, or null for an entry.</summary>
    public string? Section { get; init; }

    /// <summary>The name of the section or of the entry.</summary>
    public string Title { get; init; } = "";

    /// <summary>What the entry does, under its name.</summary>
    public string Hint { get; init; } = "";

    /// <summary>The glyph beside the name.</summary>
    public FluentSymbol Icon { get; init; } = FluentSymbol.Grid24;

    /// <summary>The entry itself, or null for a heading.</summary>
    public DockTileOption? Option { get; init; }

    /// <summary>Whether this line is a heading rather than an entry.</summary>
    public bool IsHeading => Option is null;
}

/// <summary>
/// Reference CommandDockTilePicker: the window that lists what can be put on the dock.
///
/// The list is the catalogue's own, filtered by what this build can draw, so a tile that gains a builder turns up here
/// without anybody adding it to a second list.
/// </summary>
public sealed partial class DockTilePickerViewModel : ObservableObject
{
    public DockTilePickerViewModel(bool mapPresent, IDockHost? host = null)
    {
        var rows = new List<DockTilePickerRow>();
        foreach (var section in DockTileCatalogue.Sections(mapPresent, host))
        {
            rows.Add(new DockTilePickerRow { Section = section.Title, Title = section.Title, Icon = section.Icon });
            foreach (var option in section.Options)
            {
                rows.Add(new DockTilePickerRow
                {
                    Title = option.Title,
                    Hint = option.Hint,
                    Icon = option.Icon,
                    Option = option,
                });
            }
        }

        Rows = rows;
    }

    /// <summary>The lines of the picker, in the order they are shown.</summary>
    public IReadOnlyList<DockTilePickerRow> Rows { get; }

    /// <summary>Called with the entry that was chosen.</summary>
    public Action<DockTileOption>? Picked { get; set; }

    /// <summary>An entry was chosen: the tile is put on the dock.</summary>
    public void Pick(DockTileOption option) => Picked?.Invoke(option);
}
