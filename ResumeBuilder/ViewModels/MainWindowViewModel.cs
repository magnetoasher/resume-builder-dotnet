using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ResumeBuilder.Models;
using ResumeBuilder.Services;

namespace ResumeBuilder.ViewModels;

public class MainWindowViewModel : ViewModelBase
{
    private readonly SettingsService _settingsService;
    private readonly ProfilesService _profilesService;
    private readonly ResumeGenerationService _resumeGenerationService;
    private readonly DocumentBuilder _documentBuilder;
    private readonly ApplicationsLogService _applicationsLogService;

    private AppStep _step;
    private string _setupProfilesPath = string.Empty;
    private ObservableCollection<Profile> _profiles = new();
    private Profile? _selectedProfile;
    private string _companyName = string.Empty;
    private string _jobUrl = string.Empty;
    private string _jobDescription = string.Empty;
    private string _apiKeyInput = string.Empty;
    private string _validatedApiKey = string.Empty;
    private bool _isBusy;
    private string _busyMessage = string.Empty;
    private string _toastMessage = string.Empty;
    private string _toastKind = "info";
    private bool _isToastVisible;
    private string _outputDirectory = string.Empty;
    private double _profileCardWidth = 200;
    private double _profileCardHeight = 308;
    private double _profileCardSlotWidth = 228;
    private double _profileCardSlotHeight = 336;
    private double _profileDeckMaxWidth = 1120;
    private double _profileNameFontSize = 22;

    public MainWindowViewModel()
    {
        _settingsService = new SettingsService();
        _profilesService = new ProfilesService();
        _resumeGenerationService = new ResumeGenerationService();
        _documentBuilder = new DocumentBuilder();
        _applicationsLogService = new ApplicationsLogService();

        BrowseSetupProfilesCommand = new RelayCommand(_ => BrowseSetupProfiles());
        ContinueSetupCommand = new RelayCommand(async _ => await ContinueSetupAsync(), _ => CanContinueSetup());
        ChangeProfilesCommand = new RelayCommand(_ => ChangeProfiles());
        SelectProfileCommand = new RelayCommand(p => SelectProfile(p as Profile));
        GenerateCommand = new RelayCommand(async _ => await GenerateAsync(), _ => CanGenerateNew());
        OpenOutputFolderCommand = new RelayCommand(_ => OpenOutputFolder());
        BackToProfilesCommand = new RelayCommand(_ => Step = AppStep.SelectProfileCard);
        OpenApplicationsCommand = new RelayCommand(_ => OpenApplications());

        Initialize();
    }

    public ICommand BrowseSetupProfilesCommand { get; }
    public ICommand ContinueSetupCommand { get; }
    public ICommand ChangeProfilesCommand { get; }
    public ICommand SelectProfileCommand { get; }
    public ICommand GenerateCommand { get; }
    public ICommand OpenOutputFolderCommand { get; }
    public ICommand BackToProfilesCommand { get; }
    public ICommand OpenApplicationsCommand { get; }

    public AppStep Step
    {
        get => _step;
        set
        {
            _step = value;
            OnPropertyChanged();
        }
    }

    public string SetupProfilesPath
    {
        get => _setupProfilesPath;
        set
        {
            _setupProfilesPath = value;
            OnPropertyChanged();
            RaiseCommandChanges();
        }
    }

    public ObservableCollection<Profile> Profiles
    {
        get => _profiles;
        set
        {
            _profiles = value;
            OnPropertyChanged();
        }
    }

    public double ProfileCardWidth
    {
        get => _profileCardWidth;
        private set
        {
            _profileCardWidth = value;
            OnPropertyChanged();
        }
    }

    public double ProfileCardHeight
    {
        get => _profileCardHeight;
        private set
        {
            _profileCardHeight = value;
            OnPropertyChanged();
        }
    }

    public double ProfileCardSlotWidth
    {
        get => _profileCardSlotWidth;
        private set
        {
            _profileCardSlotWidth = value;
            OnPropertyChanged();
        }
    }

    public double ProfileCardSlotHeight
    {
        get => _profileCardSlotHeight;
        private set
        {
            _profileCardSlotHeight = value;
            OnPropertyChanged();
        }
    }

    public double ProfileDeckMaxWidth
    {
        get => _profileDeckMaxWidth;
        private set
        {
            _profileDeckMaxWidth = value;
            OnPropertyChanged();
        }
    }

    public double ProfileNameFontSize
    {
        get => _profileNameFontSize;
        private set
        {
            _profileNameFontSize = value;
            OnPropertyChanged();
        }
    }

    public Profile? SelectedProfile
    {
        get => _selectedProfile;
        set
        {
            _selectedProfile = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(SelectedProfileName));
            RaiseCommandChanges();
        }
    }

    public string SelectedProfileName => SelectedProfile?.DisplayName ?? "";

    public string CompanyName
    {
        get => _companyName;
        set
        {
            _companyName = value;
            OnPropertyChanged();
            RefreshOutputDirectoryPreview();
            RaiseCommandChanges();
        }
    }

    public string JobUrl
    {
        get => _jobUrl;
        set
        {
            _jobUrl = value;
            OnPropertyChanged();
        }
    }

    public string JobDescription
    {
        get => _jobDescription;
        set
        {
            _jobDescription = value;
            OnPropertyChanged();
            RaiseCommandChanges();
        }
    }

    public string ApiKeyInput
    {
        get => _apiKeyInput;
        set
        {
            _apiKeyInput = value;
            OnPropertyChanged();
            RaiseCommandChanges();
        }
    }

    public bool IsApiKeySet => !string.IsNullOrWhiteSpace(_validatedApiKey);

    public bool IsBusy
    {
        get => _isBusy;
        set
        {
            _isBusy = value;
            OnPropertyChanged();
            RaiseCommandChanges();
        }
    }

    public string BusyMessage
    {
        get => _busyMessage;
        set
        {
            _busyMessage = value;
            OnPropertyChanged();
        }
    }

    public string ToastMessage
    {
        get => _toastMessage;
        set
        {
            _toastMessage = value;
            OnPropertyChanged();
        }
    }

    public string ToastKind
    {
        get => _toastKind;
        set
        {
            _toastKind = value;
            OnPropertyChanged();
        }
    }

    public bool IsToastVisible
    {
        get => _isToastVisible;
        set
        {
            _isToastVisible = value;
            OnPropertyChanged();
        }
    }

    public string OutputDirectory
    {
        get => _outputDirectory;
        set
        {
            _outputDirectory = value;
            OnPropertyChanged();
        }
    }

    private void Initialize()
    {
        OutputDirectory = ResolveOutputDirectory();
        SetupProfilesPath = _settingsService.Settings.ProfilesPath;
        UpdateProfileCardLayout(0);
        Step = AppStep.SelectProfiles;
        OnPropertyChanged(nameof(IsApiKeySet));
    }

    private void BrowseSetupProfiles()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON Files (*.json)|*.json",
            Title = "Select profiles"
        };

        if (dialog.ShowDialog() == true)
        {
            SetupProfilesPath = dialog.FileName;
        }
    }

    private async Task ContinueSetupAsync()
    {
        if (string.IsNullOrWhiteSpace(ApiKeyInput))
        {
            ShowToast("OpenAI API key is required.", "warning");
            return;
        }

        if (string.IsNullOrWhiteSpace(SetupProfilesPath) || !File.Exists(SetupProfilesPath))
        {
            ShowToast("Select a valid json file.", "warning");
            return;
        }

        IsBusy = true;
        BusyMessage = "Validating OpenAI API key...";
        RaiseCommandChanges();

        try
        {
            var (isValid, errorMessage) = await OpenAiService.ValidateApiKeyAsync(ApiKeyInput.Trim());
            if (!isValid)
            {
                _validatedApiKey = string.Empty;
                OnPropertyChanged(nameof(IsApiKeySet));
                ShowToast(errorMessage, "error");
                return;
            }

            BusyMessage = "Loading profiles...";
            if (!LoadProfiles(SetupProfilesPath))
            {
                return;
            }

            _validatedApiKey = ApiKeyInput.Trim();
            OnPropertyChanged(nameof(IsApiKeySet));
            Step = AppStep.SelectProfileCard;
            ShowToast("API key validated. Select a profile.", "success");
        }
        finally
        {
            IsBusy = false;
            BusyMessage = string.Empty;
            RaiseCommandChanges();
        }
    }

    private void ChangeProfiles()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON Files (*.json)|*.json",
            Title = "Select new profiles"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        SetupProfilesPath = dialog.FileName;
        if (LoadProfiles(SetupProfilesPath))
        {
            Step = AppStep.SelectProfileCard;
            ShowToast("Profiles updated.", "success");
        }
    }

    private bool LoadProfiles(string path)
    {
        try
        {
            var list = _profilesService.LoadProfiles(path);
            if (list.Count == 0)
            {
                ShowToast("This json does not contain any profiles.", "error");
                return false;
            }

            Profiles = new ObservableCollection<Profile>(list.OrderBy(p => p.DisplayName));
            UpdateProfileCardLayout(Profiles.Count);
            SetupProfilesPath = path;
            _settingsService.Settings.ProfilesPath = path;
            _settingsService.Save();
            return true;
        }
        catch (Exception ex)
        {
            ShowToast($"Failed to load profiles: {ex.Message}", "error");
            return false;
        }
    }

    private void UpdateProfileCardLayout(int profileCount)
    {
        if (profileCount <= 2)
        {
            ProfileCardWidth = 248;
            ProfileCardHeight = 378;
            ProfileCardSlotWidth = 312;
            ProfileCardSlotHeight = 412;
            ProfileDeckMaxWidth = 720;
            ProfileNameFontSize = 30;
            return;
        }

        if (profileCount == 3)
        {
            ProfileCardWidth = 224;
            ProfileCardHeight = 344;
            ProfileCardSlotWidth = 276;
            ProfileCardSlotHeight = 378;
            ProfileDeckMaxWidth = 860;
            ProfileNameFontSize = 25;
            return;
        }

        ProfileCardWidth = 200;
        ProfileCardHeight = 308;
        ProfileCardSlotWidth = 228;
        ProfileCardSlotHeight = 336;
        ProfileDeckMaxWidth = 1120;
        ProfileNameFontSize = 22;
    }

    private void SelectProfile(Profile? profile)
    {
        if (profile == null)
        {
            return;
        }

        SelectedProfile = profile;
        RefreshOutputDirectoryPreview();
        Step = AppStep.Builder;
    }

    private bool CanContinueSetup()
    {
        return !IsBusy && !string.IsNullOrWhiteSpace(ApiKeyInput) && !string.IsNullOrWhiteSpace(SetupProfilesPath);
    }

    private bool CanGenerateNew()
    {
        return SelectedProfile != null
            && !string.IsNullOrWhiteSpace(CompanyName)
            && !string.IsNullOrWhiteSpace(JobDescription)
            && IsApiKeySet
            && !IsBusy;
    }

    private async Task GenerateAsync()
    {
        if (SelectedProfile == null)
        {
            ShowToast("Select a profile first.", "warning");
            return;
        }

        if (string.IsNullOrWhiteSpace(JobDescription))
        {
            ShowToast("Paste a job description.", "warning");
            return;
        }

        if (string.IsNullOrWhiteSpace(CompanyName))
        {
            ShowToast("Enter company name.", "warning");
            return;
        }

        if (!IsApiKeySet)
        {
            ShowToast("Validate your OpenAI API key on the first screen.", "warning");
            return;
        }

        var outputRootDir = ResolveOutputDirectory();
        var profileOutputDir = ResolveProfileOutputDirectory(outputRootDir, SelectedProfile.DisplayName);
        var companyOutputDir = ResolveCompanyOutputDirectory(profileOutputDir, CompanyName);
        if (!FileHelpers.CanWriteToDirectory(companyOutputDir))
        {
            ShowToast("Cannot create output folder. Check write permissions.", "error");
            return;
        }

        OutputDirectory = companyOutputDir;
        var logPath = Path.Combine(outputRootDir, "applications.xlsx");
        if (FileHelpers.IsFileLocked(logPath))
        {
            ShowToast("applications.xlsx is already open. Close it and try again.", "warning");
            return;
        }

        IsBusy = true;
        BusyMessage = "Generating new resume...";
        RaiseCommandChanges();

        try
        {
            var candidateCorpus = BuildCandidateCorpus(SelectedProfile);
            var resume = await _resumeGenerationService.GenerateResumeAsync(
                SelectedProfile,
                JobDescription,
                candidateCorpus,
                _validatedApiKey,
                _settingsService.Settings);

            var baseName = ResolveUniqueResumeBaseName(companyOutputDir, SelectedProfile.DisplayName);
            var docxPath = Path.Combine(companyOutputDir, $"{baseName}.docx");
            var pdfPath = Path.Combine(companyOutputDir, $"{baseName}.pdf");

            _documentBuilder.BuildResumeDocx(resume, SelectedProfile, docxPath);
            var generatedResumePath = docxPath;
            var toastKind = "warning";
            var toastMessage = "PDF export unavailable. Saved DOCX instead.";

            try
            {
                await PdfConverter.ConvertDocxToPdfAsync(docxPath, pdfPath);
                generatedResumePath = pdfPath;
                toastKind = "success";
                toastMessage = "Resume generated successfully.";

                try
                {
                    File.Delete(docxPath);
                }
                catch
                {
                    // Keep docx if deletion fails.
                }
            }
            catch (Exception pdfEx)
            {
                var reason = pdfEx.Message;
                if (reason.Length > 240)
                {
                    reason = $"{reason[..240]}...";
                }

                toastMessage = $"PDF export unavailable. Saved DOCX instead. {reason}";
            }

            var parsedRole = CompanyNameResolver.InferRoleName(JobUrl, JobDescription);
            var entry = new ApplicationLogEntry
            {
                Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                Company = CompanyName.Trim(),
                Role = parsedRole,
                JobUrl = JobUrl,
                JobDescription = JobDescription,
                ResumePath = generatedResumePath,
                ProfileName = SelectedProfile.DisplayName
            };

            try
            {
                _applicationsLogService.AppendEntry(logPath, entry);
            }
            catch (IOException)
            {
                toastKind = "warning";
                toastMessage = $"{toastMessage} Resume file was created, but applications.xlsx is open. Close it and generate again to log this entry.";
            }
            catch (Exception logEx)
            {
                toastKind = "warning";
                var reason = logEx.Message;
                if (reason.Length > 180)
                {
                    reason = $"{reason[..180]}...";
                }

                toastMessage = $"{toastMessage} Resume file was created, but updating applications.xlsx failed: {reason}";
            }

            ShowToast(toastMessage, toastKind);
        }
        catch (Exception ex)
        {
            ShowToast($"Generation failed: {ex.Message}", "error");
        }
        finally
        {
            IsBusy = false;
            BusyMessage = string.Empty;
            RaiseCommandChanges();
        }
    }

    private static string BuildCandidateCorpus(Profile profile)
    {
        var builder = new StringBuilder();
        builder.AppendLine(profile.DisplayName);
        builder.AppendLine(profile.ContactLine);

        foreach (var role in profile.Experience)
        {
            builder.AppendLine($"{role.Title} at {role.Company} ({role.Dates})");
        }

        foreach (var education in profile.Education)
        {
            builder.AppendLine($"{education.Degree} - {education.School} ({education.Dates})");
        }

        return builder.ToString();
    }

    private string ResolveOutputDirectory()
    {
        var exeDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
        var resumesDir = Path.Combine(exeDir, "Resumes");
        if (FileHelpers.CanWriteToDirectory(resumesDir))
        {
            _settingsService.Settings.OutputDirectory = resumesDir;
            _settingsService.Save();
            return resumesDir;
        }

        if (!string.IsNullOrWhiteSpace(_settingsService.Settings.OutputDirectory) &&
            FileHelpers.CanWriteToDirectory(_settingsService.Settings.OutputDirectory))
        {
            return _settingsService.Settings.OutputDirectory;
        }

        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var fallback = Path.Combine(appData, "ResumeBuilder", "Resumes");
        FileHelpers.CanWriteToDirectory(fallback);
        _settingsService.Settings.OutputDirectory = fallback;
        _settingsService.Save();
        return fallback;
    }

    private void RefreshOutputDirectoryPreview()
    {
        var outputRootDir = ResolveOutputDirectory();
        if (SelectedProfile == null)
        {
            OutputDirectory = outputRootDir;
            return;
        }

        var profileOutputDir = ResolveProfileOutputDirectory(outputRootDir, SelectedProfile.DisplayName);
        if (string.IsNullOrWhiteSpace(CompanyName))
        {
            OutputDirectory = profileOutputDir;
            return;
        }

        OutputDirectory = ResolveCompanyOutputDirectory(profileOutputDir, CompanyName);
    }

    private static string ResolveProfileOutputDirectory(string outputRootDir, string profileName)
    {
        var safeProfileName = FileHelpers.SanitizeFileName(profileName);
        return Path.Combine(outputRootDir, safeProfileName);
    }

    private static string ResolveCompanyOutputDirectory(string profileOutputDir, string companyName)
    {
        var safeCompanyName = FileHelpers.SanitizeFileName(companyName);
        return Path.Combine(profileOutputDir, safeCompanyName);
    }

    private static string ResolveUniqueResumeBaseName(string targetDirectory, string profileName)
    {
        var safeProfileName = FileHelpers.SanitizeFileName(profileName);
        var candidate = safeProfileName;
        var sequence = 2;

        while (File.Exists(Path.Combine(targetDirectory, $"{candidate}.pdf")) ||
               File.Exists(Path.Combine(targetDirectory, $"{candidate}.docx")))
        {
            candidate = $"{safeProfileName} ({sequence})";
            sequence++;
        }

        return candidate;
    }

    private void OpenOutputFolder()
    {
        if (string.IsNullOrWhiteSpace(OutputDirectory) || !Directory.Exists(OutputDirectory))
        {
            return;
        }

        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = OutputDirectory,
                UseShellExecute = true
            });
        }
        catch
        {
            // Ignore open folder errors.
        }
    }

    private void OpenApplications()
    {
        try
        {
            var outputRootDir = ResolveOutputDirectory();
            var logPath = Path.Combine(outputRootDir, "applications.xlsx");
            var window = new ApplicationsWindow(logPath);
            if (Application.Current?.MainWindow != null)
            {
                window.Owner = Application.Current.MainWindow;
            }

            window.Show();
        }
        catch (Exception ex)
        {
            ShowToast($"Unable to open applications: {ex.Message}", "error");
        }
    }

    private void ShowToast(string message, string kind)
    {
        ToastMessage = message;
        ToastKind = kind;
        IsToastVisible = true;

        var timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(4)
        };
        timer.Tick += (_, _) =>
        {
            IsToastVisible = false;
            timer.Stop();
        };
        timer.Start();
    }

    private void RaiseCommandChanges()
    {
        if (ContinueSetupCommand is RelayCommand continueSetup)
        {
            continueSetup.RaiseCanExecuteChanged();
        }

        if (GenerateCommand is RelayCommand generate)
        {
            generate.RaiseCanExecuteChanged();
        }
    }
}

