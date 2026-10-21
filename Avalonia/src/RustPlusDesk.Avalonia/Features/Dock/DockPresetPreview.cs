using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>One rectangle of an arrangement, as the preview draws it.</summary>
/// <param name="X">Where it sits, in the arrangement's own pixels.</param>
/// <param name="Y">Where it sits, in the arrangement's own pixels.</param>
/// <param name="Width">How wide it is.</param>
/// <param name="Height">How high it is.</param>
/// <param name="IsMap">Whether it is the map, which the preview draws differently.</param>
public sealed record DockPresetPreviewRect(double X, double Y, double Width, double Height, bool IsMap);

/// <summary>
/// The shape of a saved arrangement, drawn small.
///
/// The reference fades an outline of the arrangement over the dock itself. This port's dock is a window of its own, and
/// a second transparent window over the game for a hover is a lot of machinery for a picture, so the arrangement is
/// drawn as a diagram instead: the same cells, the same pitch, at a size that fits in a list.
/// </summary>
public static class DockPresetPreview
{
    /// <summary>
    /// The rectangles of an arrangement and the box they fill.
    /// </summary>
    /// <param name="preset">The arrangement.</param>
    /// <param name="rects">Where each tile sits.</param>
    /// <param name="width">How wide the arrangement is.</param>
    /// <param name="height">How high the arrangement is.</param>
    public static (IReadOnlyList<DockPresetPreviewRect> Rects, double Width, double Height) Build(CommandDockPreset preset)
    {
        double zoom = CommandDockLayout.ClampZoom(preset.GridZoom ?? 1.0);
        double mapSize = preset.MapSize ?? 0;

        var rects = new List<DockPresetPreviewRect>();
        foreach (var tile in preset.Tiles)
        {
            bool isMap = tile.Kind == CommandDockTileKinds.Map;

            double x = CommandDockLayout.CellOffset(tile.Col, zoom);
            double y = CommandDockLayout.CellOffset(tile.Row, zoom);
            double width = isMap ? mapSize : CommandDockLayout.CellsToPixels(tile.ColSpan, zoom);
            double height = isMap ? mapSize : CommandDockLayout.CellsToPixels(tile.RowSpan, zoom);

            if (width <= 0 || height <= 0) continue;
            rects.Add(new DockPresetPreviewRect(x, y, width, height, isMap));
        }

        if (rects.Count == 0) return ([], 0, 0);

        // Normalised to its own top-left, which is what loading does: the first occupied cell is (0,0).
        double shiftX = -rects.Min(rect => rect.X);
        double shiftY = -rects.Min(rect => rect.Y);

        double totalWidth = rects.Max(rect => rect.X + rect.Width) + shiftX;
        double totalHeight = rects.Max(rect => rect.Y + rect.Height) + shiftY;

        var normalised = rects
            .Select(rect => rect with { X = rect.X + shiftX, Y = rect.Y + shiftY })
            .ToList();

        return (normalised, totalWidth, totalHeight);
    }
}
