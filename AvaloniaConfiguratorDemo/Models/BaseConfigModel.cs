using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace BackupConfiguratorDemo.Models;

public partial class BaseConfigModel : ObservableObject
{
    [ObservableProperty]
    private string tag = "new-base";

    [ObservableProperty]
    private string title = "Новая база";

    [ObservableProperty]
    private string description = string.Empty;

    [ObservableProperty]
    private OptionItem<string>? backupTypeOption;

    [ObservableProperty]
    private string sourcePath = string.Empty;

    [ObservableProperty]
    private string destinationPath = string.Empty;

    [ObservableProperty]
    private string exePath = string.Empty;

    [ObservableProperty]
    private int keepCopies = 5;

    [ObservableProperty]
    private OptionItem<string>? cloudTypeOption;

    [ObservableProperty]
    private int cloudKeep = 3;

    [ObservableProperty]
    private string stopServicesText = string.Empty;

    [ObservableProperty]
    private bool disabled;

    public bool RequiresDesigner => string.Equals(BackupTypeOption?.Value, "DT", StringComparison.OrdinalIgnoreCase);

    public bool UsesCloud => string.Equals(CloudTypeOption?.Value, "Yandex.Disk", StringComparison.OrdinalIgnoreCase);

    public string CloudTokenKey => $"{Tag}__YADiskToken";

    public string DesignerLoginKey => $"{Tag}__DT_Login";

    public string DesignerPasswordKey => $"{Tag}__DT_Password";

    partial void OnBackupTypeOptionChanged(OptionItem<string>? value)
    {
        OnPropertyChanged(nameof(RequiresDesigner));
    }

    partial void OnCloudTypeOptionChanged(OptionItem<string>? value)
    {
        OnPropertyChanged(nameof(UsesCloud));
    }

    partial void OnTagChanged(string value)
    {
        OnPropertyChanged(nameof(CloudTokenKey));
        OnPropertyChanged(nameof(DesignerLoginKey));
        OnPropertyChanged(nameof(DesignerPasswordKey));
    }
}
