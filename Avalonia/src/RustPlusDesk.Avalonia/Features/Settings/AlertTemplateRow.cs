using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Alerts;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// One row of the reference CustomAlertsEditor: name, description, the editable template and its variables, with Save
/// (an empty text removes the override) and Reset (back to the default).
/// </summary>
public sealed partial class AlertTemplateRow(AlertTemplateInfo info, AlertTemplates templates) : ObservableObject
{
    public string Key => info.Key;
    public string Name => info.Name;
    public string Description => info.Description;
    public string Variables => string.Format(CultureInfo.InvariantCulture, "Variables: {0}", info.Variables);
    [ObservableProperty] private string text = templates.Get(info.Key);
    /// <summary>Short feedback after Save/Reset (the reference flashed the text box).</summary>
    [ObservableProperty] private string status = "";

    [RelayCommand]
    private void Save()
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            templates.RemoveOverride(Key);
            Text = templates.Get(Key);
            Status = "Default restored";
            return;
        }
        templates.SetOverride(Key, Text);
        Status = "Saved";
    }

    [RelayCommand]
    private void Reset()
    {
        templates.RemoveOverride(Key);
        Text = templates.Get(Key);
        Status = "Reset to default";
    }

    partial void OnTextChanged(string value) => Status = "";
}
