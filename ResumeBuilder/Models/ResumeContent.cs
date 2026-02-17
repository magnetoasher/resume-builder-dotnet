using System.Collections.Generic;

namespace ResumeBuilder.Models;

public class ResumeContent
{
    public string Summary { get; set; } = string.Empty;
    public List<string> Skills { get; set; } = new();
    public List<Education> Education { get; set; } = new();
    public List<ResumeExperience> Experience { get; set; } = new();
}
