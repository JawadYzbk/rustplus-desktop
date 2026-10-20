using System.Collections.ObjectModel;
using System.Globalization;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Ai;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Ai;

/// <summary>One exchange as the window shows it: the words, not the record behind them.</summary>
public sealed class AiHistoryRow
{
    /// <summary>When the question was asked, in the local time the player remembers it in.</summary>
    public string When { get; init; } = "";

    /// <summary>Which provider was asked.</summary>
    public string Provider { get; init; } = "";

    /// <summary>Which model was asked for.</summary>
    public string Model { get; init; } = "";

    /// <summary>What went with the question: a screenshot, the game's own sound, or both.</summary>
    public string Badges { get; init; } = "";

    /// <summary>The question, where the recording was turned into words on this machine.</summary>
    public string? Transcript { get; init; }

    /// <summary>The answer, or the reason it failed.</summary>
    public string Answer { get; init; } = "";

    /// <summary>Whether the exchange failed, so the reason is shown instead of an answer.</summary>
    public bool Failed { get; init; }

    /// <summary>What a click on the row puts on the clipboard.</summary>
    public string CopyText { get; init; } = "";
}

/// <summary>
/// Reference AiHistoryWindow: the last few exchanges, so a failure that happened while the answer panel was switched
/// off is still readable afterwards.
///
/// It exists for the failures. "Failed" on a tile the size of a stamp is not a reason, and the reason - a retired
/// model, a rejected key, a quota - is the one thing that tells the player what to change.
/// </summary>
public sealed partial class AiHistoryViewModel : ObservableObject, IDisposable
{
    /// <summary>The exchanges, newest first.</summary>
    public ObservableCollection<AiHistoryRow> Rows { get; } = [];

    /// <summary>How many of the last twenty are kept on this PC.</summary>
    [ObservableProperty]
    private string countText = "";

    /// <summary>Whether there is anything to clear.</summary>
    [ObservableProperty]
    private bool hasEntries;

    /// <summary>Whether nothing has been asked yet.</summary>
    [ObservableProperty]
    private bool isEmpty = true;

    public AiHistoryViewModel()
    {
        AiHistory.Changed += OnHistoryChanged;
        Loc.LanguageChanged += OnLanguageChanged;
        Render();
    }

    /// <summary>Reference BtnClear_Click: the history is emptied.</summary>
    [RelayCommand]
    private void Clear() => AiHistory.Clear();

    /// <summary>Reference Render: the exchanges are listed, newest first.</summary>
    private void Render()
    {
        var history = AiHistory.Entries;
        Rows.Clear();
        foreach (var exchange in history) Rows.Add(Row(exchange));
        CountText = string.Format(CultureInfo.CurrentCulture,
            Loc.Text("AiHistoryCount", "{0} of the last {1} kept on this PC"), history.Count, AiHistory.MaxEntries);
        HasEntries = history.Count > 0;
        IsEmpty = history.Count == 0;
    }

    /// <summary>Reference Card: what the window shows of one exchange.</summary>
    private static AiHistoryRow Row(AiExchange exchange)
    {
        var badges = new List<string>();
        if (exchange.HadScreenshot) badges.Add(Loc.Text("AiHistoryWithScreenshot", "screenshot"));
        if (exchange.HadGameAudio) badges.Add(Loc.Text("AiHistoryWithGameAudio", "game audio"));
        string answer = exchange.Failed ? exchange.Error ?? "" : exchange.Answer;
        return new AiHistoryRow
        {
            // Local time: the entry is compared against what the player remembers doing, and they do not remember
            // it in UTC.
            When = exchange.WhenUtc.ToLocalTime().ToString("g", CultureInfo.CurrentCulture),
            Provider = AiProviders.Label(exchange.Provider),
            Model = exchange.Model,
            Badges = string.Join(", ", badges),
            // Only where the recording was turned into words on this machine. Where the provider heard it directly
            // there is nothing to show, and an empty "you asked:" would read as though the question had been lost.
            Transcript = string.IsNullOrWhiteSpace(exchange.Transcript) ? null : exchange.Transcript.Trim(),
            Answer = answer,
            Failed = exchange.Failed,
            CopyText = answer
        };
    }

    /// <summary>The history is changed from wherever the answer came back, so the window is drawn on its own thread.</summary>
    private void OnHistoryChanged(object? sender, EventArgs e) => OnUiThread(Render);

    /// <summary>The labels are read again when the language changes, so the window says it in the new one.</summary>
    private void OnLanguageChanged(object? sender, EventArgs e) => OnUiThread(Render);

    private static void OnUiThread(Action action)
    {
        if (Dispatcher.UIThread.CheckAccess()) action();
        else Dispatcher.UIThread.Post(action);
    }

    /// <summary>Lets go of the history, so a closed window is not called back into.</summary>
    public void Dispose()
    {
        AiHistory.Changed -= OnHistoryChanged;
        Loc.LanguageChanged -= OnLanguageChanged;
    }
}
