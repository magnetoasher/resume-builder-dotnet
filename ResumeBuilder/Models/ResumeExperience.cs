using System.Collections.Generic;

namespace ResumeBuilder.Models;

public class ResumeExperience
{
    public string Company { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Dates { get; set; } = string.Empty;
    public List<string> Bullets { get; set; } = new();
}
