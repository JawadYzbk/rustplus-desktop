using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Tutorials;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Reference the tutorial centre's card: one tutorial, what has been done of it, and what can be done next.
/// </summary>
public sealed partial class TutorialRow : ObservableObject
{
    private readonly TutorialDefinition definition;
    private TutorialProgress progress;

    public TutorialRow(TutorialDefinition definition, TutorialProgress progress)
    {
        this.definition = definition;
        this.progress = progress;
    }

    /// <summary>The tutorial this row is about.</summary>
    public TutorialDefinition Definition => definition;

    /// <summary>The tutorial's own id.</summary>
    public string Id => definition.Id;

    /// <summary>The tutorial's name, in the app's language.</summary>
    public string Title => Loc.Text(definition.TitleKey, definition.Id);

    /// <summary>What the tutorial is about, in the app's language.</summary>
    public string Description => Loc.Text(definition.DescriptionKey, definition.Id);

    /// <summary>The category the tutorial sits under.</summary>
    public string Category => Loc.Text($"Tutorials.Category.{definition.Category}", definition.Category);

    /// <summary>The glyph beside the name.</summary>
    public string Icon => definition.Icon;

    /// <summary>Whether the app suggests this one first.</summary>
    public bool IsRecommended => definition.IsRecommended;

    /// <summary>Whether this is a new feature's tutorial.</summary>
    public bool IsNewFeature => definition.IsNewFeature;

    /// <summary>How many steps the tutorial has.</summary>
    public int StepCount => definition.Steps.Count;

    /// <summary>How much of it has been done.</summary>
    public int CompletedCount => progress.CompletedStepIds.Count;

    /// <summary>How far through the tutorial the player is, as a percentage.</summary>
    public double ProgressPercent => StepCount == 0 ? 0 : Math.Clamp(CompletedCount * 100.0 / StepCount, 0, 100);

    /// <summary>What the progress line says.</summary>
    public string ProgressLabel => $"{CompletedCount} / {StepCount}";

    /// <summary>What the status says.</summary>
    public string StatusLabel => Loc.Text($"Tutorials.Status.{progress.Status}", progress.Status.ToString());

    /// <summary>Whether the tutorial has been started but not finished.</summary>
    public bool CanContinue => progress.Status is TutorialStatus.InProgress or TutorialStatus.Updated
        && progress.CompletedStepIds.Count > 0 && progress.CompletedStepIds.Count < StepCount;

    /// <summary>Whether there is anything saved to continue.</summary>
    public bool HasProgress => progress.Status is not TutorialStatus.NotStarted;

    /// <summary>What the button on the card says.</summary>
    public string PrimaryActionLabel => progress.Status switch
    {
        TutorialStatus.Completed => Loc.Text("Tutorials.Common.Restart", "Restart"),
        TutorialStatus.InProgress => Loc.Text("Tutorials.Common.Continue", "Continue"),
        _ => Loc.Text("Tutorials.Common.Start", "Start"),
    };

    /// <summary>The saved progress, for the service to write back.</summary>
    public TutorialProgress Progress => progress;

    /// <summary>Takes the progress the service has written.</summary>
    public void Update(TutorialProgress updated)
    {
        progress = updated;
        OnPropertyChanged(nameof(CompletedCount));
        OnPropertyChanged(nameof(ProgressPercent));
        OnPropertyChanged(nameof(ProgressLabel));
        OnPropertyChanged(nameof(StatusLabel));
        OnPropertyChanged(nameof(CanContinue));
        OnPropertyChanged(nameof(HasProgress));
        OnPropertyChanged(nameof(PrimaryActionLabel));
    }
}

/// <summary>
/// Reference the tutorial centre: every tutorial the app offers, what has been done of each, and the buttons that
/// start, continue and reset them.
/// </summary>
public sealed partial class TutorialsViewModel : ObservableObject
{
    private readonly ITutorialRegistry registry;
    private readonly ITutorialProgressStore progressStore;
    private readonly ITutorialService service;

    public TutorialsViewModel(ITutorialRegistry registry, ITutorialProgressStore progressStore, ITutorialService service)
    {
        this.registry = registry;
        this.progressStore = progressStore;
        this.service = service;
    }

    /// <summary>The tutorials, grouped by the category they sit under.</summary>
    public ObservableCollection<TutorialCategoryGroup> Categories { get; } = [];

    /// <summary>What the window says after a start, a reset or a failure.</summary>
    [ObservableProperty] private string status = "";

    /// <summary>Whether the app starts the basic tour by itself.</summary>
    [ObservableProperty]
    private bool autoStartBasic = true;

    /// <summary>Whether the app offers the new features' tours by itself.</summary>
    [ObservableProperty]
    private bool autoStartNewFeatures = true;

    /// <summary>The tutorial that was last run, if there was one.</summary>
    [ObservableProperty] private string lastTutorialId = "";

    /// <summary>Whether there is a tutorial to go back to.</summary>
    public bool CanContinueLast => !string.IsNullOrEmpty(LastTutorialId);

    /// <summary>Whether there is a tutorial that has been started.</summary>
    public bool HasAnyProgress => Categories.SelectMany(group => group.Tutorials).Any(tutorial => tutorial.HasProgress);

    /// <summary>Reads the catalogue and the saved progress, and shows them.</summary>
    [RelayCommand]
    public async Task RefreshAsync()
    {
        var all = await progressStore.GetAllAsync(CancellationToken.None);
        var preferences = await progressStore.GetPreferencesAsync(CancellationToken.None);

        Categories.Clear();
        foreach (var group in registry.Tutorials
            .OrderBy(tutorial => tutorial.DisplayOrder)
            .GroupBy(tutorial => tutorial.Category))
        {
            var rows = new List<TutorialRow>();
            foreach (var definition in group)
            {
                var progress = all.TryGetValue(definition.Id, out var saved)
                    ? saved
                    : new TutorialProgress { TutorialId = definition.Id, TutorialVersion = definition.Version };
                rows.Add(new TutorialRow(definition, progress));
            }

            Categories.Add(new TutorialCategoryGroup(group.Key, rows));
        }

        AutoStartBasic = preferences.AutoStartBasicTutorial;
        AutoStartNewFeatures = preferences.AutoStartNewFeatureTutorials;
        LastTutorialId = preferences.LastTutorialId ?? "";

        OnPropertyChanged(nameof(CanContinueLast));
        OnPropertyChanged(nameof(HasAnyProgress));
    }

    /// <summary>Reference Start_Click: the tutorial is started, or continued where it was left.</summary>
    [RelayCommand]
    private async Task StartAsync(TutorialRow? row)
    {
        if (row is null) return;

        try
        {
            if (row.CanContinue) await service.ContinueAsync(row.Id);
            else await service.StartAsync(row.Id);
            Status = "";
        }
        catch (Exception ex)
        {
            Status = ex.Message;
        }

        await RefreshAsync();
    }

    /// <summary>Reference Restart_Click: the tutorial starts again from its first step.</summary>
    [RelayCommand]
    private async Task RestartAsync(TutorialRow? row)
    {
        if (row is null) return;

        try
        {
            await service.ResetAsync(row.Id);
            await service.StartAsync(row.Id);
            Status = "";
        }
        catch (Exception ex)
        {
            Status = ex.Message;
        }

        await RefreshAsync();
    }

    /// <summary>Reference Reset_Click: the progress of one tutorial is forgotten.</summary>
    [RelayCommand]
    private async Task ResetAsync(TutorialRow? row)
    {
        if (row is null) return;

        await service.ResetAsync(row.Id);
        await RefreshAsync();
    }

    /// <summary>Reference ContinueLast_Click: the tutorial that was last run is picked up where it was left.</summary>
    [RelayCommand]
    private async Task ContinueLastAsync()
    {
        if (!CanContinueLast) return;

        try
        {
            await service.ContinueAsync(LastTutorialId);
            Status = "";
        }
        catch (Exception ex)
        {
            Status = ex.Message;
        }

        await RefreshAsync();
    }

    /// <summary>Reference StartRecommended_Click: the tour the app recommends is started.</summary>
    [RelayCommand]
    private async Task StartRecommendedAsync()
    {
        var recommended = Categories.SelectMany(group => group.Tutorials).FirstOrDefault(tutorial => tutorial.IsRecommended);
        if (recommended is null) return;

        await StartAsync(recommended);
    }

    /// <summary>Reference ResetAll_Click: every tutorial's progress is forgotten.</summary>
    [RelayCommand]
    private async Task ResetAllAsync()
    {
        await service.ResetAllAsync();
        await RefreshAsync();
    }

    /// <summary>Reference the two switches in the tutorial centre: whether the app offers a tour by itself.</summary>
    partial void OnAutoStartBasicChanged(bool value) => SavePreferencesAsync().ConfigureAwait(false);

    partial void OnAutoStartNewFeaturesChanged(bool value) => SavePreferencesAsync().ConfigureAwait(false);

    /// <summary>The switches are written as they are flipped.</summary>
    private async Task SavePreferencesAsync()
    {
        var preferences = await progressStore.GetPreferencesAsync(CancellationToken.None);
        preferences.AutoStartBasicTutorial = AutoStartBasic;
        preferences.AutoStartNewFeatureTutorials = AutoStartNewFeatures;
        await progressStore.SavePreferencesAsync(preferences, CancellationToken.None);
    }
}

/// <summary>One category of the tutorial centre and the tutorials under it.</summary>
/// <param name="Name">The category's name, in the app's language.</param>
/// <param name="Tutorials">The tutorials in it.</param>
public sealed record TutorialCategoryGroup(string Name, IReadOnlyList<TutorialRow> Tutorials);
