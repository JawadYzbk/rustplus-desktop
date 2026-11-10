using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Features.Settings;

public sealed partial class SettingsView : UserControl
{
    private SettingsViewModel? model;

    /// <summary>The backup card whose confirmation is being listened to (the page replaces its model on a restore).</summary>
    private DataBackupViewModel? backup;

    public SettingsView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) =>
        {
            if (model is not null) { model.TemplatesFocusRequested -= ShowTemplates; model.WebhookFocusRequested -= ShowWebhook; }
            if (backup is not null) backup.RestoreConfirmed -= OnRestoreConfirmed;
            model = DataContext as SettingsViewModel;
            if (model is not null) { model.TemplatesFocusRequested += ShowTemplates; model.WebhookFocusRequested += ShowWebhook; }
            // The restore's confirmation is answered here: only a window can open the picker that follows it.
            backup = model?.Backup;
            if (backup is not null) backup.RestoreConfirmed += OnRestoreConfirmed;
            // The model browser is a window, and only a window can open one: the card asks for it and this view lends
            // it the window it lives in.
            if (model?.Ai is { } ai) ai.BrowseModels = ShowModelBrowser;
            // Only a window can open a link, so the update card borrows the one this view lives in.
            if (model?.Update is { } update)
            {
                update.OpenUrl = url => TopLevel.GetTopLevel(this)?.Launcher?.LaunchUriAsync(new Uri(url)) ?? Task.FromResult(false);
            }
        };
    }

    /// <summary>Opens the model browser over the card that asked for it.</summary>
    private void ShowModelBrowser()
    {
        if (model?.Ai is not { } card) return;

        var window = new AiModelBrowserWindow { DataContext = new AiModelBrowserViewModel() };
        window.Picked = card.ApplyModel;
        if (TopLevel.GetTopLevel(this) is Window owner) window.Show(owner);
        else window.Show();
    }

    /// <summary>
    /// Reference BtnResetAppData_Click: the reset window, with the actions the settings page was given.
    ///
    /// The window owns the checkboxes; the shell owns the connection and the account, so the actions arrive here and go
    /// straight through.
    /// </summary>
    private void OnResetData(object? sender, global::Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (model is not { } page) return;

        var view = new ResetDataViewModel
        {
            Disconnect = page.Disconnect,
            SignOut = page.SignOut,
            ClearCacheAndDrawings = page.ClearCacheAndDrawings,
            ClearCrosshairs = page.ClearCrosshairs,
            ClearProfiles = page.ClearProfiles,
        };

        var window = new ResetDataWindow { DataContext = view };
        if (TopLevel.GetTopLevel(this) is Window owner) window.Show(owner);
        else window.Show();
    }

    /// <summary>
    /// Reference BtnBackupData_Click: the password the archive is protected with, then where it is written, then the
    /// archive itself. An empty password is the old app's way of leaving the backup unencrypted.
    /// </summary>
    private async void OnBackupData(object? sender, global::Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (backup is not { } data) return;
        if (TopLevel.GetTopLevel(this) is not Window owner) return;

        var protection = new BackupPasswordViewModel(isDecrypting: false);
        var prompt = new BackupPasswordWindow { DataContext = protection };
        if (await prompt.ShowDialog<bool>(owner) != true) return;

        string? path = await PickZipAsync(owner, Loc.Text("BackupApplicationDataTitle", "Backup Application Data"),
            data.SuggestedFileName);
        if (path is null) return;
        await data.BackupAsync(path, protection.EffectivePassword);
    }

    /// <summary>Reference BtnRestoreData_Click: the confirmation first, then the file it names.</summary>
    private void OnRestoreData(object? sender, global::Avalonia.Interactivity.RoutedEventArgs e) =>
        backup?.RequestRestoreCommand.Execute(null);

    /// <summary>
    /// The confirmed restore: the archive is picked, its password asked for when it carries the reference's
    /// signature, and the service puts it back. The reference stops the restore when the password prompt is closed.
    /// </summary>
    private async void OnRestoreConfirmed()
    {
        if (backup is not { } data) return;
        if (TopLevel.GetTopLevel(this) is not Window owner) return;

        string? path = await PickZipAsync(owner, Loc.Text("RestoreApplicationDataTitle", "Restore Application Data"), null);
        if (path is null) return;

        string? password = null;
        if (DataBackupViewModel.IsEncrypted(path))
        {
            var protection = new BackupPasswordViewModel(isDecrypting: true);
            var prompt = new BackupPasswordWindow { DataContext = protection };
            if (await prompt.ShowDialog<bool>(owner) != true) return;
            password = protection.EffectivePassword;
        }

        await data.RestoreAsync(path, password);
    }

    /// <summary>
    /// Reference the SaveFileDialog and OpenFileDialog with the ZIP filter. Only a local file can be written or read,
    /// so a location the host cannot name on disk is not offered to the service.
    /// </summary>
    private static async Task<string?> PickZipAsync(Window owner, string title, string? suggestedName)
    {
        if (suggestedName is null)
        {
            var files = await owner.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = title,
                AllowMultiple = false,
                FileTypeFilter = [new FilePickerFileType("ZIP Archives") { Patterns = ["*.zip"] }]
            });
            if (files.Count == 0) return null;
            using var file = files[0];
            return file.TryGetLocalPath();
        }

        var saved = await owner.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = title,
            SuggestedFileName = suggestedName,
            DefaultExtension = "zip",
            FileTypeChoices = [new FilePickerFileType("ZIP Archives") { Patterns = ["*.zip"] }],
            ShowOverwritePrompt = true
        });
        if (saved is null) return null;
        using (saved) return saved.TryGetLocalPath();
    }

    /// <summary>Scrolls the templates panel to the top of the page.</summary>
    private void ShowTemplates() => ScrollTo(TemplatesPanel);
    /// <summary>Scrolls the Discord webhook card to the top of the page.</summary>
    private void ShowWebhook() => ScrollTo(WebhookPanel);

    private void ScrollTo(Control panel)
    {
        UpdateLayout();
        if (panel.TranslatePoint(default, PageScroll.Content as Control ?? PageScroll) is { } top)
            PageScroll.Offset = new(PageScroll.Offset.X, top.Y);
    }
}
