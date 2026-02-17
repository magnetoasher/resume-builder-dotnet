using System;
using System.IO;
using System.Text.Json;

namespace ResumeBuilder.Services;

public class SettingsService
{
    private readonly string _settingsPath;

    public AppSettings Settings { get; private set; } = new();

    public SettingsService()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var baseDir = Path.Combine(appData, "ResumeBuilder");
        _settingsPath = Path.Combine(baseDir, "settings.json");
        Load();
    }

    public void Load()
    {
        try
        {
            if (!File.Exists(_settingsPath))
            {
                Settings = new AppSettings();
                return;
            }

            var json = File.ReadAllText(_settingsPath);
            Settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            Settings = new AppSettings();
        }
    }

    public void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(_settingsPath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            File.WriteAllText(_settingsPath, json);
        }
        catch
        {
            // Intentionally ignore settings save failures.
        }
    }
}
