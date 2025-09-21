using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using BackupConfiguratorDemo.Models;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace BackupConfiguratorDemo.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private static readonly HttpClient HttpClient = new();

    public ObservableCollection<OptionItem<string>> BackupTypeOptions { get; } = new()
    {
        new OptionItem<string>("DT", "Выгрузка .dt (DESIGNER)"),
        new OptionItem<string>("1CD", "Копия файла .1CD")
    };

    public ObservableCollection<OptionItem<string>> CloudOptions { get; } = new()
    {
        new OptionItem<string>("None", "Не использовать облако"),
        new OptionItem<string>("Yandex.Disk", "Yandex.Disk")
    };

    public ObservableCollection<OptionItem<int>> AfterBackupOptions { get; } = new()
    {
        new OptionItem<int>(1, "Выключить компьютер после копирования"),
        new OptionItem<int>(2, "Перезагрузить компьютер"),
        new OptionItem<int>(3, "Ничего не делать")
    };

    public ObservableCollection<BaseConfigModel> BaseConfigs { get; }

    [ObservableProperty]
    private BaseConfigModel? selectedBase;

    [ObservableProperty]
    private OptionItem<int>? selectedAfterBackupOption;

    [ObservableProperty]
    private bool telegramEnabled = true;

    [ObservableProperty]
    private bool telegramNotifyOnlyOnErrors;

    [ObservableProperty]
    private string telegramBotToken = string.Empty;

    [ObservableProperty]
    private string telegramChatId = string.Empty;

    [ObservableProperty]
    private string telegramSecondaryToken = string.Empty;

    [ObservableProperty]
    private string telegramSecondaryChatId = string.Empty;

    [ObservableProperty]
    private string telegramMessage = "Тестовое сообщение от демо-конфигуратора.";

    [ObservableProperty]
    private string statusMessage = "Готов к отправке тестового сообщения.";

    [ObservableProperty]
    private bool isSending;

    public RelayCommand CreateConfigCommand { get; }
    public RelayCommand DuplicateConfigCommand { get; }
    public RelayCommand DeleteConfigCommand { get; }
    public RelayCommand SaveConfigCommand { get; }
    public RelayCommand SaveGlobalSettingsCommand { get; }
    public RelayCommand SaveSecretsCommand { get; }
    public AsyncRelayCommand SendTelegramTestCommand { get; }

    public MainWindowViewModel()
    {
        BaseConfigs = new ObservableCollection<BaseConfigModel>
        {
            new BaseConfigModel
            {
                Tag = "main-prod",
                Title = "Основная база продаж",
                Description = "Боевой контур для рабочей команды",
                BackupTypeOption = BackupTypeOptions[0],
                SourcePath = @"C:\\1C\\Bases\\Main",
                DestinationPath = @"D:\\Backups\\Main",
                ExePath = @"C:\\Program Files\\1cv8\\bin\\1cv8.exe",
                KeepCopies = 7,
                CloudTypeOption = CloudOptions[1],
                CloudKeep = 14,
                StopServicesText = "Apache2.4\r\nRagent",
                Disabled = false
            },
            new BaseConfigModel
            {
                Tag = "analytics",
                Title = "Аналитика",
                Description = "Экспериментальная копия для BI",
                BackupTypeOption = BackupTypeOptions[1],
                SourcePath = @"C:\\1C\\Bases\\Analytics\\1Cv8.1CD",
                DestinationPath = @"D:\\Backups\\Analytics",
                KeepCopies = 10,
                CloudTypeOption = CloudOptions[0],
                CloudKeep = 0,
                StopServicesText = string.Empty,
                Disabled = false
            },
            new BaseConfigModel
            {
                Tag = "demo-standby",
                Title = "Резерв",
                Description = "Горячий стенд, поднимается по требованию",
                BackupTypeOption = BackupTypeOptions[0],
                SourcePath = @"C:\\1C\\Bases\\Standby",
                DestinationPath = @"E:\\Backups\\Standby",
                ExePath = @"C:\\Program Files\\1cv8\\bin\\1cv8.exe",
                KeepCopies = 3,
                CloudTypeOption = CloudOptions[1],
                CloudKeep = 6,
                StopServicesText = "Apache2.4",
                Disabled = true
            }
        };

        SelectedBase = BaseConfigs.FirstOrDefault();
        SelectedAfterBackupOption = AfterBackupOptions.Last();

        CreateConfigCommand = new RelayCommand(() =>
        {
            StatusMessage = "Демо: создание новой базы отключено. Полный функционал появится в релизе.";
        });

        DuplicateConfigCommand = new RelayCommand(() =>
        {
            StatusMessage = SelectedBase is null
                ? "Демо: выберите базу для копирования."
                : $"Демо: копирование '{SelectedBase.Title}' недоступно.";
        }, () => SelectedBase is not null);

        DeleteConfigCommand = new RelayCommand(() =>
        {
            StatusMessage = SelectedBase is null
                ? "Демо: выберите базу для удаления."
                : $"Демо: удаление '{SelectedBase.Title}' отключено.";
        }, () => SelectedBase is not null);

        SaveConfigCommand = new RelayCommand(() =>
        {
            StatusMessage = SelectedBase is null
                ? "Демо: нет выбранной записи для сохранения."
                : $"Демо: изменения '{SelectedBase.Title}' не сохраняются на диск.";
        }, () => SelectedBase is not null);

        SaveGlobalSettingsCommand = new RelayCommand(() =>
        {
            var action = SelectedAfterBackupOption?.Title ?? "Ничего не делать";
            StatusMessage = $"Демо: сохранение глобальных настроек отключено (выбрано: {action}).";
        });

        SaveSecretsCommand = new RelayCommand(() =>
        {
            StatusMessage = "Демо: обновление ключей и секретов будет добавлено в полной версии.";
        });

        SendTelegramTestCommand = new AsyncRelayCommand(SendTelegramAsync, CanSendTelegram);

        SelectedBase = BaseConfigs.FirstOrDefault();\r\n        SelectedAfterBackupOption = AfterBackupOptions.Last();
    }

    partial void OnSelectedBaseChanged(BaseConfigModel? value)
    {
        DuplicateConfigCommand.NotifyCanExecuteChanged();
        DeleteConfigCommand.NotifyCanExecuteChanged();
        SaveConfigCommand.NotifyCanExecuteChanged();
    }

    partial void OnTelegramBotTokenChanged(string value)
    {
        SendTelegramTestCommand.NotifyCanExecuteChanged();
    }

    partial void OnTelegramChatIdChanged(string value)
    {
        SendTelegramTestCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsSendingChanged(bool value)
    {
        SendTelegramTestCommand.NotifyCanExecuteChanged();
    }

    private bool CanSendTelegram()
    {
        return !IsSending && !string.IsNullOrWhiteSpace(TelegramBotToken) && !string.IsNullOrWhiteSpace(TelegramChatId);
    }

    private async Task SendTelegramAsync()
    {
        try
        {
            IsSending = true;
            StatusMessage = "Отправка тестового сообщения...";

            var token = TelegramBotToken.Trim();
            var chatId = TelegramChatId.Trim();
            var message = string.IsNullOrWhiteSpace(TelegramMessage)
                ? "Тестовое сообщение от демо-конфигуратора."
                : TelegramMessage.Trim();

            using var content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                { "chat_id", chatId },
                { "text", message }
            });

            var requestUri = $"https://api.telegram.org/bot{token}/sendMessage";
            using var response = await HttpClient.PostAsync(requestUri, content).ConfigureAwait(false);

            if (response.IsSuccessStatusCode)
            {
                StatusMessage = "Сообщение успешно отправлено.";
            }
            else
            {
                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                StatusMessage = $"Telegram вернул ошибку {(int)response.StatusCode}: {body}";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Не удалось отправить сообщение: {ex.Message}";
        }
        finally
        {
            IsSending = false;
        }
    }
}
