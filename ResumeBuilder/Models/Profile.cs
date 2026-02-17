using System.Collections.Generic;

namespace ResumeBuilder.Models;

public class Profile
{
    public string Id { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string ContactLine { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string LinkedIn { get; set; } = string.Empty;
    public List<Education> Education { get; set; } = new();
    public List<Experience> Experience { get; set; } = new();

    public int RolesCount => Experience?.Count ?? 0;
    public int EducationCount => Education?.Count ?? 0;
    public bool HasLinkedIn => !string.IsNullOrWhiteSpace(LinkedIn) || (ContactLine?.Contains("linkedin", StringComparison.OrdinalIgnoreCase) ?? false);
}
