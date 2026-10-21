using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Updates;

namespace RustPlusDesk.Avalonia.Features.Updates;

/// <summary>
/// Reference the patch notes window: what each release changed.
///
/// The old app carries its notes in the window itself; the port reads the published releases and their notes, so the
/// list is the one the releases actually have.
/// </summary>
public sealed partial class PatchNotesViewModel : ObservableObject
{
    private readonly Func<Task<IReadOnlyList<ReleaseNote>>>? releases;

    public PatchNotesViewModel(Func<Task<IReadOnlyList<ReleaseNote>>>? releases = null) => this.releases = releases;

    /// <summary>The published releases, newest first.</summary>
    public ObservableCollection<ReleaseNote> Releases { get; } = [];

    /// <summary>What the window says while it is loading, and afterwards.</summary>
    [ObservableProperty] private string status = "";

    /// <summary>Whether the notes are being read.</summary>
    [ObservableProperty] private bool isLoading;

    /// <summary>The release whose notes are on show.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelected))]
    private ReleaseNote? selected;

    /// <summary>Whether a release is on show.</summary>
    public bool HasSelected => Selected is not null;

    /// <summary>The running version of the app, as the old app's window shows it.</summary>
    public string CurrentVersion => $"v{RustPlusDesk.Core.Updates.AppVersion.Short(RustPlusDesk.Core.Updates.AppVersion.Raw())}";

    /// <summary>Reads the releases and shows them.</summary>
    [RelayCommand]
    public async Task LoadAsync()
    {
        if (releases is not { } read) { Status = "The release service is not available."; return; }

        IsLoading = true;
        Status = "Reading the releases…";
        try
        {
            var notes = await read().ConfigureAwait(true);
            Releases.Clear();
            foreach (var note in notes) Releases.Add(note);

            Selected = Releases.FirstOrDefault();
            Status = Releases.Count == 0
                ? "No release could be read. Check the connection and try again."
                : $"{Releases.Count} releases.";
        }
        catch (Exception ex)
        {
            Status = $"The releases could not be read: {ex.Message}";
        }
        finally { IsLoading = false; }
    }

    /// <summary>Opens a link in the user's browser; the window owns it, because only a window can.</summary>
    public Func<string, Task<bool>>? OpenUrl { get; set; }

    /// <summary>Reference the release's own page, opened when the user asks for it.</summary>
    [RelayCommand]
    private async Task OpenReleaseAsync(ReleaseNote? note)
    {
        if (note is null || OpenUrl is not { } open || note.Url.Length == 0) return;
        await open(note.Url);
    }
}
