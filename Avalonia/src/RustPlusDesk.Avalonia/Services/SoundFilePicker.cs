using Avalonia.Controls;
using Avalonia.Platform.Storage;

namespace RustPlusDesk.Avalonia.Services;

/// <summary>
/// Picks the sound a timer makes (reference BtnSelectTimerCountdownAudio / BtnSelectTimerAlarmAudio). Only a local
/// file can be played, so a file the host cannot name on disk is refused.
/// </summary>
public interface ISoundFilePicker
{
    Task<string?> PickAsync();
}

internal sealed class SoundFilePicker(Window owner) : ISoundFilePicker
{
    public async Task<string?> PickAsync()
    {
        var files = await owner.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Choose a timer sound",
            AllowMultiple = false,
            FileTypeFilter =
            [
                new FilePickerFileType("Sounds") { Patterns = ["*.wav", "*.mp3", "*.ogg", "*.flac"] },
                new FilePickerFileType("All files") { Patterns = ["*"] }
            ]
        });
        if (files.Count == 0) return null;
        using var file = files[0];
        return file.TryGetLocalPath();
    }
}
