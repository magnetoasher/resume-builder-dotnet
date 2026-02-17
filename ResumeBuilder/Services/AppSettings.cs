namespace ResumeBuilder.Services;

public class AppSettings
{
    public string ProfilesPath { get; set; } = string.Empty;
    public string OutputDirectory { get; set; } = string.Empty;
    public string Model { get; set; } = "gpt-5.2";
    public double Temperature { get; set; } = 0.6;
    public double TopP { get; set; } = 0.9;
}
