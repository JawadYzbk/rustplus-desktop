using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Infrastructure.Ai;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the OpenRouter model browser: the catalogue of models, a search over it, and a way to pick one.
///
/// The catalogue comes from the service, which reads the live list where it can and the cache or the built-in list
/// where it cannot, so the browser always has something to show.
/// </summary>
public sealed partial class AiModelBrowserViewModel : ObservableObject
{
    private readonly IOpenRouterModelService service;
    private IReadOnlyList<OpenRouterModelInfo> all = [];

    public AiModelBrowserViewModel(IOpenRouterModelService? service = null) =>
        this.service = service ?? OpenRouterModelService.Shared;

    /// <summary>The models on show, after the search and the free-only switch.</summary>
    public ObservableCollection<OpenRouterModelInfo> Models { get; } = [];

    /// <summary>What is being looked for. Every word has to match somewhere in the model.</summary>
    [ObservableProperty]
    private string search = "";

    /// <summary>Whether only the free models are shown.</summary>
    [ObservableProperty]
    private bool freeOnly;

    /// <summary>The model that is highlighted.</summary>
    [ObservableProperty]
    private OpenRouterModelInfo? selected;

    /// <summary>What the card says about the catalogue: how many models, and where they came from.</summary>
    [ObservableProperty] private string status = "";

    /// <summary>Whether the catalogue is being fetched.</summary>
    [ObservableProperty] private bool isLoading;

    /// <summary>Whether there is anything to show.</summary>
    public bool IsEmpty => Models.Count == 0;

    /// <summary>Reads the catalogue and shows it.</summary>
    [RelayCommand]
    private async Task LoadAsync()
    {
        await LoadAsync(forceRefresh: false);
    }

    /// <summary>Reads the catalogue again from OpenRouter, ignoring the cache.</summary>
    [RelayCommand]
    private async Task RefreshAsync()
    {
        await LoadAsync(forceRefresh: true);
    }

    private async Task LoadAsync(bool forceRefresh)
    {
        IsLoading = true;
        try
        {
            all = await service.GetModelsAsync(forceRefresh);
            Apply();
        }
        catch (Exception ex)
        {
            Status = ex.Message;
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>The search narrows the list as it is typed.</summary>
    partial void OnSearchChanged(string value) => Apply();

    /// <summary>The free-only switch narrows the list too.</summary>
    partial void OnFreeOnlyChanged(bool value) => Apply();

    /// <summary>Applies the filters to the catalogue.</summary>
    private void Apply()
    {
        var filtered = OpenRouterModelSearch.Filter(all, Search, FreeOnly ? true : null).ToList();

        Models.Clear();
        foreach (var model in filtered) Models.Add(model);

        Status = filtered.Count == all.Count
            ? $"{all.Count} models."
            : $"Showing {filtered.Count} of {all.Count} models.";

        OnPropertyChanged(nameof(IsEmpty));

        if (Selected is not null && !filtered.Contains(Selected)) Selected = null;
    }
}
