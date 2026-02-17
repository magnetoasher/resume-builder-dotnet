namespace ResumeBuilder.Models;

public class ApplicationLogEntry
{
    public string Timestamp { get; set; } = string.Empty;
    public string Company { get; set; } = string.Empty;
    public string Role { get; set; } = string.Empty;
    public string JobUrl { get; set; } = string.Empty;
    public string JobDescription { get; set; } = string.Empty;
    public string ResumePath { get; set; } = string.Empty;
    public string ProfileName { get; set; } = string.Empty;
}
