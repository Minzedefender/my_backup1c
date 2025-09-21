namespace BackupConfiguratorDemo.Models;

public record OptionItem<T>(T Value, string Title)
{
    public override string ToString() => Title;
}
