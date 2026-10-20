using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;
using RustPlusDesk.Avalonia.Shared.Icons;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MiniMapWindow.CommandDock.Style.TileStyle: everything a tile builder needs to draw itself in the user's
/// chosen style.
///
/// Every tile asks here rather than reaching for a theme resource directly, so "transparent with readable text" is one
/// decision made once instead of a rule each builder has to remember. The three values resolve tile-first, dock-default
/// second - see <see cref="CommandDockTile.Opacity"/> for why the global ones are defaults and not a second set of
/// switches.
/// </summary>
public sealed class DockStyle
{
    /// <summary>Reference TileStyle.Opacity: how far the tile's background and border are faded.</summary>
    public double Opacity { get; init; } = 1.0;

    /// <summary>Reference TileStyle.FontScale: the multiplier on every font size in the tile.</summary>
    public double FontScale { get; init; } = 1.0;

    /// <summary>Reference TileStyle.TextMain: the colour of the tile's own text.</summary>
    public IBrush TextMain { get; init; } = Brushes.White;

    /// <summary>Reference TileStyle.TextSub: the colour of the tile's captions.</summary>
    public IBrush TextSub { get; init; } = Brushes.Gray;

    /// <summary>Reference TileStyle.IconScale: how much bigger the tile's icon is drawn than the text alone makes it.</summary>
    public double IconScale { get; init; } = 1.0;

    /// <summary>
    /// Reference TileStyle.Zoom: the grid's pitch on its own.
    ///
    /// <see cref="FontScale"/> already has it folded in, which is right for text - but an icon must not follow the text
    /// scale, and it must still follow the zoom. So the zoom is kept separately rather than divided back out.
    /// </summary>
    public double Zoom { get; init; } = 1.0;

    /// <summary>Reference TileStyle.TextShadow: whether the tile's text carries a shadow over the game.</summary>
    public bool TextShadow { get; init; }

    /// <summary>The colour key the tile resolved to, so a builder can tell "auto" from a colour the user picked.</summary>
    public string TextColorKey { get; init; } = CommandDockTextColors.Auto;

    /// <summary>Reference TileStyle.Size: a font size with the tile's scale applied, rounded to a whole pixel.</summary>
    public double Size(double baseSize) => Math.Round(Math.Max(7, baseSize * FontScale));

    /// <summary>
    /// Reference TileStyle.Icon: an icon's size, which the text scale does not touch.
    ///
    /// Only the grid's zoom and the tile's own icon scale apply. A device or an event is recognised by its picture rather
    /// than by the word under it, so the two are set against each other.
    /// </summary>
    public double Icon(double baseSize) => Math.Round(Math.Max(8, baseSize * Zoom * IconScale));

    /// <summary>Reference TileStyle.Chrome: fades a colour by the tile's opacity - for backgrounds and borders only.</summary>
    public IBrush Chrome(Color color) =>
        new SolidColorBrush(Color.FromArgb((byte)Math.Round(color.A * Opacity), color.R, color.G, color.B));

    /// <summary>Reference TileStyle.Chrome: the same, starting from a brush the theme supplied.</summary>
    public IBrush Chrome(IBrush brush) => brush is ISolidColorBrush solid ? Chrome(solid.Color) : brush;

    /// <summary>
    /// Reference TileStyle.TextShadow: the soft shadow that keeps a faded tile's text readable over the game.
    ///
    /// Below roughly half opacity the tile stops carrying its own background and the game shows through behind it, so
    /// the text needs a shadow of its own: white over snow and black over water.
    /// </summary>
    public void Shadow(TextBlock text)
    {
        if (!TextShadow) return;

        text.Effect = new DropShadowDirectionEffect
        {
            BlurRadius = 4,
            ShadowDepth = 1,
            Direction = 270,
            Opacity = 0.85,
            Color = Colors.Black,
        };
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.StyleFor: the tile's own appearance, resolved against the dock's defaults.
    ///
    /// The dock's value where it is set to override, the tile's own otherwise, and the dock's again as the fallback for a
    /// tile that has never been given one.
    /// </summary>
    public static DockStyle For(CommandDockLayout layout, CommandDockTile tile)
    {
        double opacity = Math.Clamp(
            layout.OverrideOpacity ? layout.DefaultOpacity : tile.Opacity ?? layout.DefaultOpacity, 0, 1);

        double scale = Math.Clamp(
            layout.OverrideFontScale ? layout.DefaultFontScale : tile.FontScale ?? layout.DefaultFontScale, 0.7, 2.0);

        // The grid's zoom rides on the font scale rather than being a second knob the builders have to remember, so a
        // bigger cell holds bigger contents instead of the same small ones in more empty space. Clamped after multiplying,
        // not before: the tile's scale is a preference and the zoom is the pitch of the grid.
        double zoom = CommandDockLayout.ClampZoom(layout.GridZoom);
        scale = Math.Clamp(scale * zoom, 0.7, 4.0);

        string key = layout.OverrideTextColor
            ? layout.DefaultTextColorKey ?? CommandDockTextColors.Auto
            : tile.TextColorKey ?? layout.DefaultTextColorKey ?? CommandDockTextColors.Auto;

        var (main, sub) = TextBrushes(key);

        return new DockStyle
        {
            Opacity = opacity,
            FontScale = scale,
            IconScale = Math.Clamp(tile.IconScale ?? 1.0, 0.6, 3.0),
            Zoom = zoom,
            TextMain = main,
            TextSub = sub,
            TextColorKey = key,

            // Below roughly half opacity the tile stops carrying its own text and the game shows through behind it. A
            // soft shadow is what keeps white legible over snow and black legible over water.
            TextShadow = opacity < 0.55,
        };
    }

    /// <summary>Reference MiniMapWindow.CommandDock.Style.TextBrushes: the palette behind the dock's colour swatches.</summary>
    public static (IBrush Main, IBrush Sub) TextBrushes(string key) => key switch
    {
        CommandDockTextColors.White => (Frozen(Colors.White), Frozen(Color.FromRgb(0xDC, 0xDC, 0xDC))),
        CommandDockTextColors.Black => (Frozen(Color.FromRgb(0x10, 0x12, 0x14)), Frozen(Color.FromRgb(0x36, 0x3B, 0x3F))),
        CommandDockTextColors.Cyan => (Frozen(Color.FromRgb(0x3F, 0xD7, 0xFF)), Frozen(Color.FromRgb(0x86, 0xE4, 0xFF))),
        CommandDockTextColors.Amber => (Frozen(Color.FromRgb(0xFF, 0xC2, 0x46)), Frozen(Color.FromRgb(0xFF, 0xD9, 0x8C))),
        CommandDockTextColors.Red => (Frozen(Color.FromRgb(0xFF, 0x6B, 0x5E)), Frozen(Color.FromRgb(0xFF, 0x9D, 0x94))),
        CommandDockTextColors.Green => (Frozen(Color.FromRgb(0x5B, 0xD9, 0x8A)), Frozen(Color.FromRgb(0x95, 0xE8, 0xB4))),
        _ => (Theme("TextPrimary", Colors.White), Theme("TextSubtle", Colors.Gray)),
    };

    /// <summary>Reference MiniMapWindow.CommandDock.Brush: the theme's brush, or the fallback the reference carries.</summary>
    public static IBrush Theme(string key, Color fallback)
    {
        if (Application.Current is { } app
            && app.TryFindResource(key, app.ActualThemeVariant, out var found)
            && found is IBrush brush)
            return brush;
        return new SolidColorBrush(fallback);
    }

    private static IBrush Frozen(Color color) => new SolidColorBrush(color);
}

/// <summary>
/// Reference MiniMapWindow.CommandDock.BuildTile: what a tile is made of, and the refresher that keeps it up to date.
///
/// The reference keeps its refreshers in a list the window's timer walks. The port keeps the same pairing, so a tile
/// that fails to refresh cannot stop the others from ticking.
/// </summary>
public sealed class DockTileContent
{
    /// <summary>Reference the tile's element: what the tile draws itself with.</summary>
    public DockTileContent(Control content, Action refresh, string? tip = null)
    {
        Content = content;
        Refresh = refresh;
        Tip = tip;
    }

    /// <summary>The element the tile draws itself with.</summary>
    public Control Content { get; }

    /// <summary>What the tile re-reads once a second.</summary>
    public Action Refresh { get; }

    /// <summary>The text the tile shows when the pointer rests on it, if it has any.</summary>
    public string? Tip { get; }
}

/// <summary>
/// Reference MiniMapWindow.CommandDock.BuildTile: the tiles this slice builds.
///
/// Only the clock, the server information, the two chats and the collapse button are here; the other kinds the reference
/// knows are left for their own slices and are skipped rather than shown as empty boxes.
/// </summary>
public static class DockTiles
{
    /// <summary>Reference MiniMapWindow.CommandDock.Minimal.ClockStyleMinimal: the clock style that means "just the time".</summary>
    public const int ClockStyleMinimal = 3;

    /// <summary>Reference MiniMapWindow.CommandDock.Minimal.ClockStyleMinimal: the clock style that draws a dial.</summary>
    public const int ClockStyleAnalogue = 1;

    /// <summary>Reference MiniMapWindow.CommandDock.Minimal.ClockStyleMinimal: the clock style that uses the game's own red.</summary>
    public const int ClockStyleRust = 2;

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BuildTile: whether this build draws the kind at all.
    ///
    /// A kind from a newer build is skipped rather than shown as an empty box, and it keeps its place on the grid.
    /// </summary>
    public static bool IsSupported(string kind) => kind switch
    {
        CommandDockTileKinds.Clock => true,
        CommandDockTileKinds.ServerInfo => true,
        CommandDockTileKinds.TeamChat => true,
        CommandDockTileKinds.ClanChat => true,
        CommandDockTileKinds.Collapse => true,
        CommandDockTileKinds.Map => true,
        _ => false,
    };

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BuildTile: the tile's content, or null for a kind this build does not draw.
    ///
    /// A kind from a newer build is skipped rather than shown as an empty box, and the tile keeps its place on the grid.
    /// </summary>
    public static DockTileContent? Build(CommandDockTile tile, IDockHost host, DockStyle style, Func<bool> isCollapsed) =>
        tile.Kind switch
        {
            CommandDockTileKinds.Clock => tile.ClockStyle == ClockStyleMinimal
                ? BuildMinimalClock(tile, host, style)
                : BuildClock(tile, host, style),
            CommandDockTileKinds.ServerInfo => tile.ServerInfoMinimal
                ? BuildMinimalServerInfo(tile, host, style)
                : BuildServerInfo(tile, host, style),
            CommandDockTileKinds.TeamChat => BuildChat(tile, host, style, clan: false),
            CommandDockTileKinds.ClanChat => BuildChat(tile, host, style, clan: true),
            CommandDockTileKinds.Collapse => BuildCollapse(tile, style, isCollapsed),
            CommandDockTileKinds.Map => BuildMap(tile, style),
            _ => null,
        };

    // ── Clock ─────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BuildClockTile: the clock, the day and night glyph and the phase line.
    ///
    /// A bigger clock tile gets a bigger clock, driven by the smaller of the two spans - a clock three cells wide and one
    /// tall has no room to grow, which is what "as far as the height allows" means.
    /// </summary>
    private static DockTileContent BuildClock(CommandDockTile tile, IDockHost host, DockStyle style)
    {
        var stack = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
        };

        double grow = Math.Min(3.0, 1 + (Math.Min(tile.ColSpan, tile.RowSpan) - 1) * 0.8);
        double Sized(double baseSize) => style.Size(baseSize * grow);

        if (tile.ClockStyle == ClockStyleAnalogue)
        {
            var face = BuildDial(style, grow, out var dial, out var hourHand, out var minuteHand);
            stack.Children.Add(face);

            var phase = SubtleText(style);
            phase.FontSize = Sized(10);
            stack.Children.Add(phase);

            return new DockTileContent(stack, () =>
            {
                if (TryParseServerTime(host.ServerTime, out int h, out int m))
                {
                    hourHand.Angle = (h % 12) * 30 + m * 0.5;
                    minuteHand.Angle = m * 6;
                }

                bool isDay = IsDay(host);
                dial.Stroke = new SolidColorBrush(isDay ? DayColour : NightColour);

                phase.Text = tile.ClockShowDayNight ? host.TimeUntilNextPhase ?? "" : "";
                phase.IsVisible = !string.IsNullOrEmpty(phase.Text);
            });
        }

        bool rustStyle = tile.ClockStyle == ClockStyleRust;

        var glyph = new Icon
        {
            Symbol = FluentSymbol.WeatherSunny24,
            FontSize = Sized(14),
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 0, 0, 2),
        };

        var clock = new TextBlock
        {
            FontSize = Sized(rustStyle ? 20 : 19),
            FontWeight = FontWeight.Bold,
            HorizontalAlignment = HorizontalAlignment.Center,
            FontFamily = new FontFamily(rustStyle ? "Impact, Segoe UI" : "Consolas, Segoe UI"),

            // The Rust clock keeps the game's own red unless a colour was picked; anything else would make "Rust style"
            // just mean a different typeface.
            Foreground = rustStyle && style.TextColorKey == CommandDockTextColors.Auto
                ? new SolidColorBrush(Color.FromRgb(0xCD, 0x41, 0x2B))
                : style.TextMain,
        };

        var phaseText = SubtleText(style);
        phaseText.FontSize = Sized(10);

        style.Shadow(clock);

        stack.Children.Add(glyph);
        stack.Children.Add(clock);
        stack.Children.Add(phaseText);

        return new DockTileContent(stack, () =>
        {
            clock.Text = tile.Clock12Hour ? To12Hour(host.ServerTime) : host.ServerTime;

            bool isDay = IsDay(host);
            glyph.Symbol = isDay ? FluentSymbol.WeatherSunny24 : FluentSymbol.WeatherMoon24;
            glyph.Foreground = new SolidColorBrush(isDay ? DayColour : NightColour);

            phaseText.Text = tile.ClockShowDayNight ? host.TimeUntilNextPhase ?? "" : "";
            phaseText.IsVisible = !string.IsNullOrEmpty(phaseText.Text);
        });
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Minimal.BuildMinimalClockTile: the clock and the day and night glyph, in one cell.
    ///
    /// The same two readings from the same two sources as the full clock, so a dock with both on it never shows two
    /// different times.
    /// </summary>
    private static DockTileContent BuildMinimalClock(CommandDockTile tile, IDockHost host, DockStyle style)
    {
        double grow = Math.Min(2.6, 1 + (Math.Min(tile.ColSpan, tile.RowSpan) - 1) * 0.7);

        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var glyph = new Icon
        {
            Symbol = FluentSymbol.WeatherSunny24,
            FontSize = style.Size(13 * grow),
            Margin = new Thickness(0, 0, 5, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };

        var time = new TextBlock
        {
            FontSize = style.Size(14 * grow),
            FontWeight = FontWeight.SemiBold,
            Foreground = style.TextMain,
            VerticalAlignment = VerticalAlignment.Center,
        };

        row.Children.Add(glyph);
        row.Children.Add(time);

        style.Shadow(time);

        return new DockTileContent(row, () =>
        {
            time.Text = tile.Clock12Hour ? To12Hour(host.ServerTime) : host.ServerTime;

            // Sun or moon, in the colour the mini-map uses for each, because the colour is half the reading: at a glance
            // it says night before the digits are read.
            bool isDay = IsDay(host);
            glyph.Symbol = isDay ? FluentSymbol.WeatherSunny24 : FluentSymbol.WeatherMoon24;
            glyph.Foreground = new SolidColorBrush(isDay ? MinimalDayColour : NightColour);
        });
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BuildClockTile's analogue face: an Ellipse, twelve marks and two hands.
    ///
    /// Avalonia turns an element about its own origin, so every swinging part sits in a canvas the size of the face and
    /// is turned about the face's centre. The face is drawn at the size the tile asks for rather than scaled afterwards,
    /// which is the same picture without a layout transform.
    /// </summary>
    private static Canvas BuildDial(DockStyle style, double grow, out Ellipse dial,
        out RotateTransform hourHand, out RotateTransform minuteHand)
    {
        double size = 46 * grow;
        double centre = size / 2;

        var face = new Canvas { Width = size, Height = size, HorizontalAlignment = HorizontalAlignment.Center };

        dial = new Ellipse
        {
            Width = size,
            Height = size,
            Stroke = new SolidColorBrush(Colors.Gray),
            StrokeThickness = 1.5 * grow,
            Fill = Brushes.Transparent,
        };
        face.Children.Add(dial);

        // Twelve marks, the quarters longer. Without them a bare ring gives the hands nothing to be read against.
        for (int hour = 0; hour < 12; hour++)
        {
            bool quarter = hour % 3 == 0;
            double width = (quarter ? 1.6 : 1.0) * grow;
            double height = (quarter ? 5.0 : 2.5) * grow;
            double top = (quarter ? 3.5 : 4.5) * grow;

            var swinger = Swinger(size);
            var mark = new Border
            {
                Width = width,
                Height = height,
                CornerRadius = new CornerRadius(width / 2),
                Background = style.TextSub,
            };
            Canvas.SetLeft(mark, centre - width / 2);
            Canvas.SetTop(mark, top);
            swinger.Children.Add(mark);
            swinger.RenderTransform = new RotateTransform(hour * 30, centre, centre);
            face.Children.Add(swinger);
        }

        var hourCanvas = Swinger(size);
        var hourLine = new Border
        {
            Width = 2.5 * grow,
            Height = 12 * grow,
            CornerRadius = new CornerRadius(1.25 * grow),
            Background = style.TextMain,
        };
        Canvas.SetLeft(hourLine, centre - 1.25 * grow);
        Canvas.SetTop(hourLine, centre - 12 * grow);
        hourCanvas.Children.Add(hourLine);
        hourHand = new RotateTransform(0, centre, centre);
        hourCanvas.RenderTransform = hourHand;
        face.Children.Add(hourCanvas);

        var minuteCanvas = Swinger(size);
        var minuteLine = new Border
        {
            Width = 1.5 * grow,
            Height = 17 * grow,
            CornerRadius = new CornerRadius(0.75 * grow),
            Background = DockStyle.Theme("Accent", Color.FromRgb(0x3F, 0xD7, 0xFF)),
        };
        Canvas.SetLeft(minuteLine, centre - 0.75 * grow);
        Canvas.SetTop(minuteLine, centre - 17 * grow);
        minuteCanvas.Children.Add(minuteLine);
        minuteHand = new RotateTransform(0, centre, centre);
        minuteCanvas.RenderTransform = minuteHand;
        face.Children.Add(minuteCanvas);

        return face;
    }

    /// <summary>A canvas the size of the dial's face, with its origin at the face's corner so a turn is about the centre.</summary>
    private static Canvas Swinger(double size) =>
        new() { Width = size, Height = size, RenderTransformOrigin = RelativePoint.TopLeft };

    // ── Server information ───────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.ServerInfo.BuildServerInfoTile: how busy the server is, now and over the session.
    ///
    /// The number answers "can I get in"; the line answers "is it filling up or emptying out", which is the question
    /// behind it and the one a single figure cannot settle.
    /// </summary>
    private static DockTileContent BuildServerInfo(CommandDockTile tile, IDockHost host, DockStyle style)
    {
        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };

        var figures = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 0),
        };

        // The count and the cap belong on one line: at a single cell, "/ 50 · 0 waiting" underneath was wider than the
        // tile and left the queue nowhere to go. The cap stays small, so the reading is still what the eye lands on.
        var headline = new StackPanel { Orientation = Orientation.Horizontal };

        var players = new TextBlock
        {
            FontSize = style.Size(18),
            FontWeight = FontWeight.Bold,
            Foreground = style.TextMain,
        };

        var capacity = new TextBlock
        {
            FontSize = style.Size(9),
            Margin = new Thickness(2, 0, 0, 2),
            VerticalAlignment = VerticalAlignment.Bottom,
            Foreground = style.TextSub,
        };

        headline.Children.Add(players);
        headline.Children.Add(capacity);

        var caption = new TextBlock
        {
            FontSize = style.Size(9),
            Foreground = style.TextSub,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };

        figures.Children.Add(headline);
        figures.Children.Add(caption);
        Grid.SetColumn(figures, 0);
        row.Children.Add(figures);

        style.Shadow(players);
        style.Shadow(capacity);
        style.Shadow(caption);

        // Drawn by hand rather than through a charting library: it is one polyline, an area under it and a dot on the
        // end, and none of that is worth a dependency.
        var plot = new Canvas { ClipToBounds = true, Margin = new Thickness(0, 2, 0, 2) };
        var accent = DockStyle.Theme("Accent", Color.FromRgb(0x3F, 0xD7, 0xFF));
        var area = new Polygon { Fill = new SolidColorBrush(Color.FromArgb(0x30, 0x3F, 0xD7, 0xFF)) };
        var line = new Polyline { Stroke = accent, StrokeThickness = 1.5, StrokeJoin = PenLineJoin.Round };
        var head = new Ellipse { Width = 4, Height = 4, Fill = accent, IsVisible = false };
        plot.Children.Add(area);
        plot.Children.Add(line);
        plot.Children.Add(head);
        Grid.SetColumn(plot, 1);
        row.Children.Add(plot);

        // The readings the tile has seen. The reference keeps them for the whole session; the port keeps the ones this
        // dock has watched, which is the same line over the same period.
        var samples = new List<int>();
        int lastCount = -1;
        double lastWidth = -1;

        return new DockTileContent(row, () =>
        {
            var (current, max) = ParsePopulation(host.Population);
            string queue = host.Queue;

            players.Text = current > 0 || max > 0 ? current.ToString() : "—";
            capacity.Text = max > 0 ? $"/ {max}" : "";

            // The second line is the queue alone, and only when there is a queue to report. A server with nobody waiting
            // says "players" instead of an empty row.
            caption.Text = string.IsNullOrEmpty(queue)
                ? Loc.Text("CommandDockServerPlayers", "players")
                : string.Format(Loc.Text("CommandDockServerQueue", "{0} waiting"), queue);

            plot.IsVisible = tile.ShowGraph;
            if (!tile.ShowGraph) return;

            if (samples.Count == 0 || samples[^1] != current)
            {
                samples.Add(current);
                if (samples.Count > 60) samples.RemoveAt(0);
            }

            // Redrawn when there is something new to draw or the tile changed size, not on every tick - the shape only
            // moves once every five minutes.
            if (samples.Count == lastCount && Math.Abs(plot.Bounds.Width - lastWidth) < 0.5) return;
            lastCount = samples.Count;
            lastWidth = plot.Bounds.Width;

            DrawPopulationLine(plot, area, line, head, samples);
        });
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.ServerInfo.DrawPopulationLine: the readings across the plot area.
    ///
    /// The vertical scale is the range actually seen plus a little headroom, not zero to the server cap: a server sitting
    /// between 180 and 200 would otherwise be a flat line at the top, and the whole point is to show it moving.
    /// </summary>
    private static void DrawPopulationLine(Canvas plot, Polygon area, Polyline line, Ellipse head, IReadOnlyList<int> samples)
    {
        line.Points.Clear();
        area.Points.Clear();
        head.IsVisible = false;

        double w = plot.Bounds.Width, h = plot.Bounds.Height;
        if (w <= 1 || h <= 1 || samples.Count < 2) return;

        int low = samples.Min();
        int high = samples.Max();
        double span = Math.Max(1, high - low);

        var points = new Points();
        for (int i = 0; i < samples.Count; i++)
        {
            double x = samples.Count == 1 ? w : w * i / (samples.Count - 1.0);
            double y = h - 2 - (samples[i] - low) / span * (h - 4);
            points.Add(new Point(x, y));
        }

        line.Points = points;

        var filled = new Points();
        foreach (var point in points) filled.Add(point);
        filled.Add(new Point(w, h));
        filled.Add(new Point(0, h));
        area.Points = filled;

        var last = points[points.Count - 1];
        Canvas.SetLeft(head, last.X - head.Width / 2);
        Canvas.SetTop(head, last.Y - head.Height / 2);
        head.IsVisible = true;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Minimal.BuildMinimalPopulationTile: the player count and the figure, in one cell.
    ///
    /// The same reading the mini-map has always shown along its bottom edge - a glyph and a figure, nothing else.
    /// </summary>
    private static DockTileContent BuildMinimalServerInfo(CommandDockTile tile, IDockHost host, DockStyle style)
    {
        double grow = Math.Min(2.6, 1 + (Math.Min(tile.ColSpan, tile.RowSpan) - 1) * 0.7);

        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        row.Children.Add(new Icon
        {
            Symbol = FluentSymbol.People12,
            FontSize = style.Size(13 * grow),
            Margin = new Thickness(0, 0, 5, 0),
            Foreground = style.TextSub,
            VerticalAlignment = VerticalAlignment.Center,
        });

        var count = new TextBlock
        {
            FontSize = style.Size(14 * grow),
            FontWeight = FontWeight.SemiBold,
            Foreground = style.TextMain,
            VerticalAlignment = VerticalAlignment.Center,
        };

        row.Children.Add(count);

        style.Shadow(count);

        return new DockTileContent(row, () =>
        {
            var (current, max) = ParsePopulation(host.Population);

            // The capacity as well, where the server reports one: "128" alone says nothing about whether that is a full
            // server or an empty one.
            count.Text = current > 0 || max > 0 ? (max > 0 ? $"{current}/{max}" : current.ToString()) : "—";
        });
    }

    // ── Chat ─────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BuildChatTile: the last lines of the team chat or of the clan chat.
    ///
    /// Chat needs room for a line of text; one cell would only ever show an ellipsis. The height is free - a single row
    /// still shows the last message, which is often all anyone wants from it.
    /// </summary>
    private static DockTileContent BuildChat(CommandDockTile tile, IDockHost host, DockStyle style, bool clan)
    {
        if (tile.ColSpan < 2) tile.ColSpan = 2;

        var grid = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        var accent = DockStyle.Theme("Accent", Color.FromRgb(0x3F, 0xD7, 0xFF));

        var header = new TextBlock
        {
            FontSize = style.Size(10),
            FontWeight = FontWeight.SemiBold,
            Margin = new Thickness(0, 0, 0, 4),
            Foreground = accent,
            Text = clan ? "Clan chat" : Loc.Text("TeamChat", "Team chat"),
        };
        Grid.SetRow(header, 0);
        grid.Children.Add(header);

        style.Shadow(header);

        var lines = new StackPanel();
        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Hidden,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = lines,
        };
        Grid.SetRow(scroll, 1);
        grid.Children.Add(scroll);

        int lastCount = -1;
        string lastSignature = "";

        return new DockTileContent(grid, () =>
        {
            var feed = host.Chat(clan, 30);
            string signature = feed.Count == 0 ? "" : $"{feed.Count}|{feed[^1].At:o}|{feed[^1].Message}";
            if (feed.Count == lastCount && signature == lastSignature) return;
            lastCount = feed.Count;
            lastSignature = signature;

            lines.Children.Clear();
            foreach (var line in feed)
            {
                var para = new TextBlock
                {
                    FontSize = style.Size(11),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 2),
                };
                style.Shadow(para);

                var span = new Span();
                AddInline(span, new Run(tile.ChatAbbreviateNames ? Abbreviate(line.Author, 10) : line.Author)
                {
                    FontWeight = FontWeight.SemiBold,
                    Foreground = accent,
                });
                AddInline(span, new Run("  " + line.Message) { Foreground = style.TextMain });
                AddInline(para, span);

                ToolTip.SetTip(para, $"{line.Author} · {line.At:HH:mm}");
                lines.Children.Add(para);
            }

            scroll.ScrollToEnd();
        });
    }

    // ── Collapse ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Collapse.BuildCollapseTile: the one-cell button that puts the dock away.
    ///
    /// A dock worth building is a dock in the way - of the game behind it, mostly. Dragging it off screen loses the
    /// arrangement, and closing it loses the dock.
    /// </summary>
    private static DockTileContent BuildCollapse(CommandDockTile tile, DockStyle style, Func<bool> isCollapsed)
    {
        var glyph = new Icon
        {
            Symbol = FluentSymbol.ArrowCollapseAll24,
            FontSize = style.Size(18),
            Foreground = style.TextMain,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        void Update()
        {
            bool collapsed = isCollapsed();
            glyph.Symbol = collapsed ? FluentSymbol.ArrowExpand24 : FluentSymbol.ArrowCollapseAll24;
        }

        Update();

        return new DockTileContent(glyph, Update,
            isCollapsed() ? Loc.Text("CommandDockCollapseExpand", "Show the dock again")
                : Loc.Text("CommandDockCollapseHide", "Hide every other tile"));
    }

    // ── Map ───────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Reference MiniMapWindow.CommandDock's map tile: the map itself, as one tile among the widgets.
    ///
    /// The mini-map's own view lands with the map slice; until then the tile holds its place on the grid and its cells,
    /// which is what the arrangement is measured against.
    /// </summary>
    private static DockTileContent BuildMap(CommandDockTile tile, DockStyle style)
    {
        var box = new Border
        {
            CornerRadius = new CornerRadius(12),
            Background = DockStyle.Theme("MapBackground", Color.FromRgb(0x0B, 0x39, 0x48)),
            BorderBrush = style.Chrome(DockStyle.Theme("CardBorder", Color.FromArgb(0x55, 0xFF, 0xFF, 0xFF))),
            BorderThickness = new Thickness(1),
        };

        var label = new TextBlock
        {
            Text = Loc.Text("UiMap", "Map"),
            FontSize = style.Size(12),
            Foreground = style.TextSub,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        box.Child = label;
        return new DockTileContent(box, () => { });
    }

    // ── Shared ───────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Reference the chat's paragraphs: the author in the accent colour and the message in the tile's own.
    ///
    /// Avalonia creates a text block's inline collection on demand, so it is asked for rather than assumed.
    /// </summary>
    private static void AddInline(TextBlock text, Inline inline)
    {
        text.Inlines ??= new InlineCollection();
        text.Inlines.Add(inline);
    }

    /// <summary>Reference the chat's paragraphs: the runs of a line of chat, in the tile's colours.</summary>
    private static void AddInline(Span span, Inline inline)
    {
        span.Inlines ??= new InlineCollection();
        span.Inlines.Add(inline);
    }

    /// <summary>Reference MiniMapWindow.CommandDock.SubtleText: a caption: small, quiet and trimmed rather than wrapped.</summary>
    private static TextBlock SubtleText(DockStyle style)
    {
        var text = new TextBlock
        {
            FontSize = style.Size(10),
            HorizontalAlignment = HorizontalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Foreground = style.TextSub,
        };
        style.Shadow(text);
        return text;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.To12Hour: "15:20" as "3:20 PM".
    ///
    /// Returns the input untouched when it is not a clock reading - the server time is a dash until the first status
    /// arrives.
    /// </summary>
    public static string To12Hour(string time)
    {
        if (!TryParseServerTime(time, out int h, out int m)) return time;

        string suffix = h >= 12 ? "PM" : "AM";
        int hour12 = h % 12;
        if (hour12 == 0) hour12 = 12;

        return $"{hour12}:{m:D2} {suffix}";
    }

    /// <summary>Reference MiniMapWindow.CommandDock.TryParseServerTime: the server's clock as two numbers.</summary>
    public static bool TryParseServerTime(string text, out int hours, out int minutes)
    {
        hours = minutes = 0;
        if (string.IsNullOrWhiteSpace(text)) return false;

        var parts = text.Split(':');
        return parts.Length >= 2
            && int.TryParse(parts[0].Trim(), out hours)
            && int.TryParse(parts[1].Trim(), out minutes);
    }

    /// <summary>
    /// Reference MainWindow.DockPopulation's parsing: "128 / 200" as the players on the server and the server's cap.
    /// </summary>
    public static (int Current, int Max) ParsePopulation(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return (0, 0);

        var parts = text.Split('/');
        int current = parts.Length > 0 && int.TryParse(parts[0].Trim(), out int c) ? c : 0;
        int max = parts.Length > 1 && int.TryParse(parts[1].Trim(), out int m) ? m : 0;
        return (current, max);
    }

    /// <summary>
    /// Reference DockServerTime's day flag: whether the server is in daylight.
    ///
    /// The reference's host hands the dock the flag with the clock. The port's host contract is smaller, so the answer is
    /// read from the phase line the dock already shows: "12m until night" means the sun is still up and "12m until day"
    /// means it has gone down.
    /// </summary>
    public static bool IsDay(IDockHost host)
    {
        string phase = host.TimeUntilNextPhase ?? "";
        return !phase.Contains("until day", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Reference MiniMapWindow.CommandDock.Abbreviate: a name cut short so the message still fits.</summary>
    private static string Abbreviate(string? text, int max)
    {
        if (string.IsNullOrEmpty(text)) return "";
        return text.Length <= max ? text : text.Substring(0, Math.Max(1, max - 1)) + "…";
    }

    /// <summary>Reference the clock tile's daylight colour.</summary>
    private static Color DayColour => Color.FromRgb(0xFF, 0xD1, 0x66);

    /// <summary>Reference the minimal clock's daylight colour, which the mini-map uses for its own sun.</summary>
    private static Color MinimalDayColour => Color.FromRgb(0xFF, 0xD5, 0x4F);

    /// <summary>Reference the clock tile's night colour.</summary>
    private static Color NightColour => Color.FromRgb(0x90, 0xCA, 0xF9);
}
