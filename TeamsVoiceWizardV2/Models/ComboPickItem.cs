namespace TeamsVoiceWizardV2.Models;

/// <summary>
/// ComboBox row for policy / user pickers (display text + backing id for Graph / PowerShell).
/// </summary>
public sealed class ComboPickItem
{
    public ComboPickItem(string label, string? value)
    {
        Label = label;
        Value = value;
    }

    public string Label { get; }
    public string? Value { get; }

    public override string ToString() => Label;
}
