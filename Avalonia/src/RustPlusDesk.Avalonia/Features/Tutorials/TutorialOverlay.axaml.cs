using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Reference the tutorial overlay: the dimming, the spotlight over the target and the card with the step.
///
/// The dimming is drawn with a hole in it rather than as a layer over the target, so the target stays crisp and the
/// pointer still reaches it when the step allows it.
/// </summary>
public sealed partial class TutorialOverlay : UserControl
{
    /// <summary>How much the spotlight is grown by on each side of the target.</summary>
    private const double SpotlightPadding = 6;

    /// <summary>The gap between the card and the target it belongs to.</summary>
    private const double CardGap = 12;

    /// <summary>The service that moves the tutorial on. The shell sets it.</summary>
    public ITutorialService? Service { get; set; }

    public TutorialOverlay()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => Attach();
        SizeChanged += (_, _) => Layout();
    }

    private TutorialOverlayPresenter? presenter;

    /// <summary>Follows the presenter's step and draws it.</summary>
    private void Attach()
    {
        if (presenter is not null) presenter.PropertyChanged -= OnPresenterChanged;

        presenter = DataContext as TutorialOverlayPresenter;
        if (presenter is not null) presenter.PropertyChanged += OnPresenterChanged;

        ShowImage(presenter?.Presentation?.ImagePath);
        Layout();
    }

    private void OnPresenterChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(TutorialOverlayPresenter.Presentation) or nameof(TutorialOverlayPresenter.IsVisible))
        {
            ShowImage(presenter?.Presentation?.ImagePath);
            Layout();
        }
    }

    /// <summary>
    /// The illustration under the description, for the steps that explain something outside the app.
    ///
    /// A picture that is not there is left out rather than shown as a broken frame.
    /// </summary>
    private void ShowImage(string? path)
    {
        StepImage.Source = null;
        StepImage.IsVisible = false;
        if (string.IsNullOrWhiteSpace(path)) return;

        try
        {
            StepImage.Source = new global::Avalonia.Media.Imaging.Bitmap(path);
            StepImage.IsVisible = true;
        }
        catch (Exception)
        {
            // A picture that cannot be read is no picture, and the step is still worth showing.
        }
    }

    /// <summary>
    /// Places the spotlight over the target and the card beside it.
    ///
    /// The card goes where the step asks for it, and where the step does not say, it goes where there is the most
    /// room - which is the one place a card that covers the thing it is talking about never goes.
    /// </summary>
    private void Layout()
    {
        double width = Bounds.Width;
        double height = Bounds.Height;
        if (width <= 0 || height <= 0) return;

        var presentation = presenter?.Presentation;
        var target = presentation?.Target ?? TutorialTarget.Empty;
        bool hasTarget = !target.Bounds.IsEmpty;

        if (hasTarget)
        {
            var bounds = target.Bounds.Inflate(new TutorialPadding(SpotlightPadding));

            Canvas.SetLeft(Spotlight, bounds.X);
            Canvas.SetTop(Spotlight, bounds.Y);
            Spotlight.Width = bounds.Width;
            Spotlight.Height = bounds.Height;
            Spotlight.IsVisible = true;
        }
        else
        {
            Spotlight.IsVisible = false;
        }

        // The dimming: the whole area, with the target's box taken out of it.
        var geometry = new GeometryGroup { FillRule = FillRule.EvenOdd };
        geometry.Children.Add(new RectangleGeometry(new Rect(0, 0, width, height)));
        if (hasTarget)
        {
            var bounds = target.Bounds.Inflate(new TutorialPadding(SpotlightPadding));
            geometry.Children.Add(new RectangleGeometry(new Rect(bounds.X, bounds.Y, bounds.Width, bounds.Height)));
        }
        Dim.Data = geometry;

        double cardWidth = Card.Width > 0 ? Card.Width : 380;
        double cardHeight = Card.Bounds.Height > 0 ? Card.Bounds.Height : 220;

        double x;
        double y;
        if (!hasTarget || presentation?.Placement == TutorialPlacement.Center)
        {
            x = (width - cardWidth) / 2;
            y = (height - cardHeight) / 2;
        }
        else
        {
            var bounds = target.Bounds;
            var placement = presentation?.Placement ?? TutorialPlacement.Auto;

            (x, y) = placement switch
            {
                TutorialPlacement.Top => ((width - cardWidth) / 2, bounds.Y - cardHeight - CardGap),
                TutorialPlacement.Bottom => ((width - cardWidth) / 2, bounds.Bottom + CardGap),
                TutorialPlacement.Left => (bounds.X - cardWidth - CardGap, (height - cardHeight) / 2),
                TutorialPlacement.Right => (bounds.Right + CardGap, (height - cardHeight) / 2),
                _ => Auto(bounds, width, height, cardWidth, cardHeight),
            };

            // A card that would hang off the screen is put back inside it rather than lost.
            x = Math.Clamp(x, 8, Math.Max(8, width - cardWidth - 8));
            y = Math.Clamp(y, 8, Math.Max(8, height - cardHeight - 8));
        }

        Canvas.SetLeft(Card, x);
        Canvas.SetTop(Card, y);
    }

    /// <summary>The side with the most room in it, which is where the card goes when the step does not say.</summary>
    private static (double X, double Y) Auto(TutorialRect bounds, double width, double height, double cardWidth, double cardHeight)
    {
        double above = bounds.Y;
        double below = height - bounds.Bottom;
        double left = bounds.X;
        double right = width - bounds.Right;

        double best = Math.Max(Math.Max(above, below), Math.Max(left, right));

        if (best == above) return ((width - cardWidth) / 2, bounds.Y - cardHeight - CardGap);
        if (best == below) return ((width - cardWidth) / 2, bounds.Bottom + CardGap);
        if (best == left) return (bounds.X - cardWidth - CardGap, (height - cardHeight) / 2);
        return (bounds.Right + CardGap, (height - cardHeight) / 2);
    }

    private async void OnNext(object? sender, RoutedEventArgs e)
    {
        if (Service is not null) await Service.NextAsync();
    }

    private async void OnBack(object? sender, RoutedEventArgs e)
    {
        if (Service is not null) await Service.BackAsync();
    }

    private async void OnSkip(object? sender, RoutedEventArgs e)
    {
        if (Service is not null) await Service.SkipAsync();
    }

    private async void OnFinish(object? sender, RoutedEventArgs e)
    {
        if (Service is not null) await Service.FinishAsync();
    }
}
