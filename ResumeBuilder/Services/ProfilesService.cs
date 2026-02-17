using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using ResumeBuilder.Models;

namespace ResumeBuilder.Services;

public class ProfilesService
{
    public List<Profile> LoadProfiles(string path)
    {
        var json = File.ReadAllText(path);
        var dict = JsonSerializer.Deserialize<Dictionary<string, ProfileDto>>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }) ?? new Dictionary<string, ProfileDto>();

        var results = new List<Profile>();
        foreach (var pair in dict)
        {
            var dto = pair.Value ?? new ProfileDto();
            var profile = new Profile
            {
                Id = pair.Key,
                DisplayName = dto.DisplayName ?? pair.Key,
                ContactLine = dto.ContactLine ?? string.Empty,
                Email = dto.Email ?? string.Empty,
                Phone = dto.Phone ?? string.Empty,
                Address = dto.Address ?? string.Empty,
                LinkedIn = dto.LinkedIn ?? string.Empty,
                Education = dto.Education ?? new List<Education>(),
                Experience = dto.Experience ?? new List<Experience>()
            };
            results.Add(profile);
        }

        return results;
    }

    private class ProfileDto
    {
        [JsonPropertyName("display_name")]
        public string? DisplayName { get; set; }

        [JsonPropertyName("contact_line")]
        public string? ContactLine { get; set; }

        [JsonPropertyName("email")]
        public string? Email { get; set; }

        [JsonPropertyName("phone")]
        public string? Phone { get; set; }

        [JsonPropertyName("address")]
        public string? Address { get; set; }

        [JsonPropertyName("linkedin")]
        public string? LinkedIn { get; set; }

        [JsonPropertyName("education")]
        public List<Education>? Education { get; set; }

        [JsonPropertyName("experience")]
        public List<Experience>? Experience { get; set; }
    }
}
