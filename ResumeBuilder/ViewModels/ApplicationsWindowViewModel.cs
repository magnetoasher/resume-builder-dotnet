using System;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using ResumeBuilder.Models;
using ResumeBuilder.Services;

namespace ResumeBuilder.ViewModels;

public class ApplicationsWindowViewModel : ViewModelBase
{
    private const string AllOption = "All";

    private readonly string _logPath;
    private readonly ApplicationsQueryService _queryService;
    private readonly ObservableCollection<ApplicationRecord> _records;
    private readonly ICollectionView _recordsView;
    private string _searchText = string.Empty;
    private string _selectedCompany = AllOption;
    private string _selectedProfile = AllOption;
    private string _statusMessage = string.Empty;
    private ApplicationRecord? _selectedRecord;

    public ApplicationsWindowViewModel(string logPath)
    {
        _logPath = logPath;
        _queryService = new ApplicationsQueryService();
        _records = new ObservableCollection<ApplicationRecord>();
        _recordsView = CollectionViewSource.GetDefaultView(_records);
        _recordsView.Filter = FilterRecord;

        Companies = new ObservableCollection<string> { AllOption };
        Profiles = new ObservableCollection<string> { AllOption };

        RefreshCommand = new RelayCommand(_ => Refresh());
        OpenResumeCommand = new RelayCommand(_ => OpenResume(), _ => CanOpenResume());
        OpenJobUrlCommand = new RelayCommand(url => OpenJobUrl(url as string), url => CanOpenJobUrl(url as string));
        OpenLogFolderCommand = new RelayCommand(_ => OpenLogFolder());

        Refresh();
    }

    public ObservableCollection<string> Companies { get; }
    public ObservableCollection<string> Profiles { get; }
    public ICollectionView RecordsView => _recordsView;

    public ICommand RefreshCommand { get; }
    public ICommand OpenResumeCommand { get; }
    public ICommand OpenJobUrlCommand { get; }
    public ICommand OpenLogFolderCommand { get; }

    public string SearchText
    {
        get => _searchText;
        set
        {
            _searchText = value ?? string.Empty;
            OnPropertyChanged();
            _recordsView.Refresh();
        }
    }

    public string SelectedCompany
    {
        get => _selectedCompany;
        set
        {
            _selectedCompany = string.IsNullOrWhiteSpace(value) ? AllOption : value;
            OnPropertyChanged();
            _recordsView.Refresh();
        }
    }

    public string SelectedProfile
    {
        get => _selectedProfile;
        set
        {
            _selectedProfile = string.IsNullOrWhiteSpace(value) ? AllOption : value;
            OnPropertyChanged();
            _recordsView.Refresh();
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set
        {
            _statusMessage = value;
            OnPropertyChanged();
        }
    }

    public ApplicationRecord? SelectedRecord
    {
        get => _selectedRecord;
        set
        {
            _selectedRecord = value;
            OnPropertyChanged();
            if (OpenResumeCommand is RelayCommand openResume)
            {
                openResume.RaiseCanExecuteChanged();
            }
        }
    }

    private void Refresh()
    {
        _records.Clear();

        var loaded = _queryService.Load(_logPath);
        foreach (var record in loaded)
        {
            _records.Add(record);
        }

        UpdateFilterOptions();
        _recordsView.Refresh();

        if (!File.Exists(_logPath))
        {
            StatusMessage = "applications.xlsx not found yet. Generate at least one resume first.";
            return;
        }

        StatusMessage = $"{_records.Count} application(s) loaded.";
    }

    private void UpdateFilterOptions()
    {
        var companyBefore = SelectedCompany;
        var profileBefore = SelectedProfile;

        Companies.Clear();
        Companies.Add(AllOption);
        foreach (var company in _records
                     .Select(record => record.Company?.Trim())
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
        {
            Companies.Add(company!);
        }

        Profiles.Clear();
        Profiles.Add(AllOption);
        foreach (var profile in _records
                     .Select(record => record.ProfileName?.Trim())
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
        {
            Profiles.Add(profile!);
        }

        SelectedCompany = Companies.Contains(companyBefore) ? companyBefore : AllOption;
        SelectedProfile = Profiles.Contains(profileBefore) ? profileBefore : AllOption;
    }

    private bool FilterRecord(object obj)
    {
        if (obj is not ApplicationRecord record)
        {
            return false;
        }

        if (!string.Equals(SelectedCompany, AllOption, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(record.Company, SelectedCompany, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!string.Equals(SelectedProfile, AllOption, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(record.ProfileName, SelectedProfile, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var search = SearchText.Trim();
        if (search.Length == 0)
        {
            return true;
        }

        return Contains(record.TimestampDisplay, search)
               || Contains(record.Company, search)
               || Contains(record.Role, search)
               || Contains(record.ProfileName, search)
               || Contains(record.JobUrl, search)
               || Contains(record.ResumePath, search);
    }

    private static bool Contains(string? text, string term)
    {
        return !string.IsNullOrWhiteSpace(text) &&
               text.Contains(term, StringComparison.OrdinalIgnoreCase);
    }

    private bool CanOpenResume()
    {
        return SelectedRecord != null &&
               !string.IsNullOrWhiteSpace(SelectedRecord.ResumePath) &&
               File.Exists(SelectedRecord.ResumePath);
    }

    private void OpenResume()
    {
        if (!CanOpenResume())
        {
            StatusMessage = "Resume file not found for the selected row.";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = SelectedRecord!.ResumePath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to open resume: {ex.Message}";
        }
    }

    private void OpenLogFolder()
    {
        var folder = Path.GetDirectoryName(_logPath);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            StatusMessage = "Log folder is not available yet.";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = folder,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to open log folder: {ex.Message}";
        }
    }

    private static bool CanOpenJobUrl(string? jobUrl)
    {
        return !string.IsNullOrWhiteSpace(jobUrl);
    }

    private void OpenJobUrl(string? jobUrl)
    {
        if (!CanOpenJobUrl(jobUrl))
        {
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = jobUrl!,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to open job URL: {ex.Message}";
        }
    }
}
