using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Console;

namespace RustPlusDesk.Avalonia.Features.Console;

/// <summary>
/// The reference Console Helper: client and admin tabs, a "Most useful" block and alphabetical groups, search over
/// title/description/command/group, parameter editors, copy command, key binds in Rust notation, copy bind and copy
/// every bind. Binds and changed values persist in the reference console-helper.json; the popout window shares
/// this model, so both views stay in step.
/// </summary>
public sealed partial class ConsoleHelperViewModel : ObservableObject
{
    private readonly IConsoleCommandStore store;
    private readonly IReadOnlyList<ConsoleItemChoice> items;
    private readonly Dictionary<int, string> itemNames;
    private Task pendingSave = Task.CompletedTask;
    private bool loading;

    public ConsoleHelperViewModel(IConsoleCommandStore store, IEnumerable<ConsoleItemChoice> items)
    {
        this.store = store;
        this.items = items.Where(item => !string.IsNullOrWhiteSpace(item.Name)).OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase).ToList();
        itemNames = [];
        foreach (var item in this.items) itemNames.TryAdd(item.Id, item.Name);
    }

    public IReadOnlyList<ConsoleCommandViewModel> Commands { get; private set; } = [];
    [ObservableProperty] private string searchText = "";
    [ObservableProperty] private bool isAdminTab;
    [ObservableProperty] private IReadOnlyList<ConsoleCommandViewModel> featured = [];
    [ObservableProperty] private IReadOnlyList<ConsoleCommandGroup> groups = [];
    [ObservableProperty] private string emptyMessage = "";
    /// <summary>Short confirmation or hint ("Copied", "Pick a key first …").</summary>
    [ObservableProperty] private string notice = "";
    [ObservableProperty] private bool canPopOut = true;

    public bool HasFeatured => Featured.Count > 0;
    public bool HasEmptyMessage => EmptyMessage.Length > 0;
    partial void OnFeaturedChanged(IReadOnlyList<ConsoleCommandViewModel> value) => OnPropertyChanged(nameof(HasFeatured));
    partial void OnEmptyMessageChanged(string value) => OnPropertyChanged(nameof(HasEmptyMessage));
    partial void OnSearchTextChanged(string value) => ApplyFilter();
    partial void OnIsAdminTabChanged(bool value) => ApplyFilter();

    /// <summary>Clipboard writer supplied by the view.</summary>
    public Func<string, Task>? CopyText { get; set; }
    /// <summary>Raised by Pop out; the shell owns the window.</summary>
    public event EventHandler? PopOutRequested;
    /// <summary>Activity log sink.</summary>
    public Action<string>? Log { get; set; }

    /// <summary>Loads the catalogue and layers the stored binds and values on top (reference ConsoleCommandLibrary.Load).</summary>
    public async Task LoadAsync(CancellationToken ct)
    {
        try
        {
            var catalog = store.LoadCatalog();
            var state = await store.LoadStateAsync(ct);
            loading = true;
            Commands = catalog.Commands.Select(spec => new ConsoleCommandViewModel(spec, this)).ToList();
            foreach (var command in Commands)
                if (state.TryGetValue(command.Spec.Id, out var saved) && saved is not null) command.Apply(saved);
            Log?.Invoke($"[console-helper] loaded {Commands.Count} commands");
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or System.Text.Json.JsonException)
        {
            Commands = [];
            Log?.Invoke("[console-helper] The command list could not be read.");
        }
        finally { loading = false; }
        OnPropertyChanged(nameof(Commands));
        ApplyFilter();
    }

    private void ApplyFilter()
    {
        if (Commands.Count == 0)
        {
            Featured = [];
            Groups = [];
            EmptyMessage = "Could not load the command list. See the log for details.";
            return;
        }
        string category = IsAdminTab ? ConsoleCommandRules.Admin : ConsoleCommandRules.Client;
        var visible = Commands.Where(command => command.Spec.Category.Equals(category, StringComparison.OrdinalIgnoreCase)
            && ConsoleCommandRules.Matches(command.Spec, SearchText, command.Title, command.Description)).ToList();
        Featured = visible.Where(command => command.Spec.Featured).ToList();
        Groups = ConsoleCommandRules.Grouped(visible, command => command.Spec, command => command.Title)
            .Select(group => new ConsoleCommandGroup(group.Key, group.ToList())).ToList();
        EmptyMessage = Featured.Count == 0 && Groups.Count == 0 ? "No command matches your search." : "";
    }

    [RelayCommand] private void ShowClient() => IsAdminTab = false;
    [RelayCommand] private void ShowAdmin() => IsAdminTab = true;
    [RelayCommand] private void ClearSearch() => SearchText = "";
    [RelayCommand] private void PopOut() => PopOutRequested?.Invoke(this, EventArgs.Empty);

    [RelayCommand]
    private async Task CopyCommandAsync(ConsoleCommandViewModel? command)
    {
        if (command is null) return;
        await CopyAsync(command.ResolvedCommand, "Command copied.");
    }

    /// <summary>Reference: without a key the line would error in the console, so the user is asked to pick one.</summary>
    [RelayCommand]
    private async Task CopyBindAsync(ConsoleCommandViewModel? command)
    {
        if (command is null) return;
        if (!command.HasBind) { Notice = "Pick a key first, then the bind can be copied."; return; }
        await CopyAsync(command.BindLine, "Bind copied.");
    }

    /// <summary>Reference AllBindLines: every bindable command with a key, one line each.</summary>
    [RelayCommand]
    private async Task CopyAllBindsAsync()
    {
        string all = AllBindLines();
        if (all.Length == 0) { Notice = "No keys bound yet. Pick a key on any command first."; return; }
        await CopyAsync(all, "Every bind copied.");
    }

    public string AllBindLines() =>
        string.Join(Environment.NewLine, Commands.Where(command => command.Bindable && command.HasBind).Select(command => command.BindLine));

    private async Task CopyAsync(string text, string done)
    {
        if (string.IsNullOrWhiteSpace(text) || CopyText is null) return;
        try { await CopyText(text); Notice = done; }
        catch (Exception) { Notice = "The clipboard is busy. Try again."; }
    }

    internal string ItemName(string idText) =>
        int.TryParse(idText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int id)
        && itemNames.TryGetValue(id, out var name) ? name : idText;

    /// <summary>Reference SearchItems: the query anywhere in "Name (id)".</summary>
    internal IEnumerable<ConsoleItemChoice> SearchItems(string query) =>
        items.Where(item => item.Label.Contains(query, StringComparison.OrdinalIgnoreCase));

    /// <summary>Writes entries that carry a bind or changed values, in catalogue order (reference SaveUserState).</summary>
    internal void QueueSave()
    {
        if (loading) return;
        var snapshot = new Dictionary<string, ConsoleCommandUserState>();
        foreach (var command in Commands)
        {
            var state = command.ToState();
            if (state.Bind is not null || state.Values is { Count: > 0 }) snapshot[command.Spec.Id] = state;
        }
        pendingSave = SaveAfterAsync(pendingSave, snapshot);
    }

    private async Task SaveAfterAsync(Task previous, IReadOnlyDictionary<string, ConsoleCommandUserState> snapshot)
    {
        await previous;
        try { await store.SaveStateAsync(snapshot, CancellationToken.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Notice = "Your binds could not be saved. Check file permissions and disk space."; }
    }

    public Task FlushAsync() => pendingSave;
}
