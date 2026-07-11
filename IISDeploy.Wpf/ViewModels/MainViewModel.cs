using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using IISDeploy.Core;
using IISDeploy.Wpf.Models;
using IISDeploy.Wpf.Mvvm;

namespace IISDeploy.Wpf.ViewModels;

public enum WizardStep
{
    Package = 0,
    Configure = 1,
    Deploy = 2,
}

public sealed class MainViewModel : ObservableObject
{
    public MainViewModel()
    {
        BrowseCommand = new RelayCommand(BrowseForFolder, () => !IsBusy);
        BackCommand = new RelayCommand(GoBack, CanGoBack);
        NextCommand = new RelayCommand(GoNext, CanGoNext);
        StartOverCommand = new RelayCommand(StartOver, () => IsFinished);
        ExitCommand = new RelayCommand(() => Application.Current.Shutdown());

        // Default to the folder the tool is launched from - the same "drop the tool
        // next to the ZIP and run it" workflow the CLI uses.
        WorkingFolder = AppContext.BaseDirectory;
    }

    // -----------------------------------------------------------------
    // Commands
    // -----------------------------------------------------------------
    public RelayCommand BrowseCommand { get; }
    public RelayCommand BackCommand { get; }
    public RelayCommand NextCommand { get; }
    public RelayCommand StartOverCommand { get; }
    public RelayCommand ExitCommand { get; }

    // -----------------------------------------------------------------
    // Step / navigation state
    // -----------------------------------------------------------------
    private WizardStep _currentStep = WizardStep.Package;
    public WizardStep CurrentStep
    {
        get => _currentStep;
        private set
        {
            if (SetProperty(ref _currentStep, value))
            {
                OnPropertyChanged(nameof(IsPackageStep));
                OnPropertyChanged(nameof(IsConfigureStep));
                OnPropertyChanged(nameof(IsDeployStep));
                OnPropertyChanged(nameof(StepNumber));
                OnPropertyChanged(nameof(HeaderTitle));
                OnPropertyChanged(nameof(HeaderSubtitle));
                OnPropertyChanged(nameof(NextButtonText));
                RefreshCommands();
            }
        }
    }

    public bool IsPackageStep => CurrentStep == WizardStep.Package;
    public bool IsConfigureStep => CurrentStep == WizardStep.Configure;
    public bool IsDeployStep => CurrentStep == WizardStep.Deploy;
    public int StepNumber => (int)CurrentStep + 1;

    public string HeaderTitle => CurrentStep switch
    {
        WizardStep.Package => "Select deployment package",
        WizardStep.Configure => "Choose a target",
        WizardStep.Deploy => IsFinished ? "Finished" : "Deploying",
        _ => "IIS Deploy",
    };

    public string HeaderSubtitle => CurrentStep switch
    {
        WizardStep.Package => "Pick the ZIP file that contains the Publish folder to deploy.",
        WizardStep.Configure => "Deploy into an existing IIS site or create a brand new one.",
        WizardStep.Deploy => IsFinished ? "The operation has completed." : "Please wait while the site is being updated.",
        _ => string.Empty,
    };

    public string NextButtonText => CurrentStep == WizardStep.Configure
        ? (IsCreateNewSelected ? "Create & deploy" : "Deploy")
        : "Next";

    // -----------------------------------------------------------------
    // Step 1 - package selection
    // -----------------------------------------------------------------
    private string _workingFolder = string.Empty;
    public string WorkingFolder
    {
        get => _workingFolder;
        set
        {
            if (SetProperty(ref _workingFolder, value))
                RefreshZipFiles();
        }
    }

    public ObservableCollection<string> ZipFiles { get; } = new();

    private string? _selectedZipPath;
    public string? SelectedZipPath
    {
        get => _selectedZipPath;
        set
        {
            if (SetProperty(ref _selectedZipPath, value))
            {
                OnPropertyChanged(nameof(SelectedZipName));
                RefreshCommands();
            }
        }
    }

    public string SelectedZipName =>
        string.IsNullOrEmpty(SelectedZipPath) ? string.Empty : Path.GetFileName(SelectedZipPath);

    public bool HasZips => ZipFiles.Count > 0;
    public bool HasSingleZip => ZipFiles.Count == 1;
    public bool HasMultipleZips => ZipFiles.Count > 1;

    public string ZipStatusMessage => ZipFiles.Count switch
    {
        0 => "No ZIP files found in this folder. Browse to the folder that contains your deployment package.",
        1 => "One deployment package found.",
        _ => $"{ZipFiles.Count} deployment packages found - select the one to deploy.",
    };

    // -----------------------------------------------------------------
    // Step 2 - target selection
    // -----------------------------------------------------------------
    public ObservableCollection<TargetOption> Targets { get; } = new();

    private TargetOption? _selectedTarget;
    public TargetOption? SelectedTarget
    {
        get => _selectedTarget;
        set
        {
            if (SetProperty(ref _selectedTarget, value))
            {
                OnPropertyChanged(nameof(IsCreateNewSelected));
                OnPropertyChanged(nameof(NextButtonText));
                RefreshCommands();
            }
        }
    }

    public bool IsCreateNewSelected => SelectedTarget?.IsCreateNew ?? false;

    private bool _isLoadingSites;
    public bool IsLoadingSites
    {
        get => _isLoadingSites;
        set => SetProperty(ref _isLoadingSites, value);
    }

    private string? _sitesError;
    public string? SitesError
    {
        get => _sitesError;
        set
        {
            if (SetProperty(ref _sitesError, value))
                OnPropertyChanged(nameof(HasSitesError));
        }
    }
    public bool HasSitesError => !string.IsNullOrEmpty(SitesError);

    private string _newSiteName = string.Empty;
    public string NewSiteName
    {
        get => _newSiteName;
        set
        {
            if (SetProperty(ref _newSiteName, value))
                RefreshCommands();
        }
    }

    private string _newSitePort = "443";
    public string NewSitePort
    {
        get => _newSitePort;
        set => SetProperty(ref _newSitePort, value);
    }

    // -----------------------------------------------------------------
    // Step 3 - deployment / log
    // -----------------------------------------------------------------
    public ObservableCollection<LogEntry> LogEntries { get; } = new();

    private bool _isBusy;
    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
                RefreshCommands();
        }
    }

    private bool _isFinished;
    public bool IsFinished
    {
        get => _isFinished;
        private set
        {
            if (SetProperty(ref _isFinished, value))
            {
                OnPropertyChanged(nameof(HeaderTitle));
                OnPropertyChanged(nameof(HeaderSubtitle));
                RefreshCommands();
            }
        }
    }

    private bool _hadError;
    public bool HadError
    {
        get => _hadError;
        private set
        {
            if (SetProperty(ref _hadError, value))
                OnPropertyChanged(nameof(ResultSummary));
        }
    }

    public string ResultSummary => !IsFinished
        ? string.Empty
        : HadError
            ? "Completed with errors - review the log above."
            : "Everything completed successfully.";

    // -----------------------------------------------------------------
    // Command implementations
    // -----------------------------------------------------------------
    private bool CanGoBack() => CurrentStep == WizardStep.Configure && !IsBusy;

    private bool CanGoNext()
    {
        return CurrentStep switch
        {
            WizardStep.Package => !string.IsNullOrEmpty(SelectedZipPath) && !IsBusy,
            WizardStep.Configure => SelectedTarget != null && !IsBusy
                && (!IsCreateNewSelected || !string.IsNullOrWhiteSpace(NewSiteName)),
            _ => false,
        };
    }

    private void GoBack()
    {
        if (CurrentStep == WizardStep.Configure)
            CurrentStep = WizardStep.Package;
    }

    private async void GoNext()
    {
        switch (CurrentStep)
        {
            case WizardStep.Package:
                CurrentStep = WizardStep.Configure;
                await LoadSitesAsync();
                break;

            case WizardStep.Configure:
                CurrentStep = WizardStep.Deploy;
                await RunDeploymentAsync();
                break;
        }
    }

    private void StartOver()
    {
        LogEntries.Clear();
        SelectedTarget = null;
        NewSiteName = string.Empty;
        NewSitePort = "443";
        IsFinished = false;
        HadError = false;
        CurrentStep = WizardStep.Package;
        RefreshZipFiles();
    }

    private void BrowseForFolder()
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "Select the folder that contains the deployment ZIP",
            InitialDirectory = Directory.Exists(WorkingFolder) ? WorkingFolder : AppContext.BaseDirectory,
        };

        if (dialog.ShowDialog() == true)
            WorkingFolder = dialog.FolderName;
    }

    // -----------------------------------------------------------------
    // Data loading
    // -----------------------------------------------------------------
    public void RefreshZipFiles()
    {
        var previous = SelectedZipPath;

        ZipFiles.Clear();
        foreach (var zip in DeploymentService.FindZipFiles(WorkingFolder))
            ZipFiles.Add(zip);

        OnPropertyChanged(nameof(HasZips));
        OnPropertyChanged(nameof(HasSingleZip));
        OnPropertyChanged(nameof(HasMultipleZips));
        OnPropertyChanged(nameof(ZipStatusMessage));

        // Preserve the previous selection when possible, otherwise auto-select the
        // only file (so the single-file case needs no interaction).
        if (previous != null && ZipFiles.Contains(previous))
            SelectedZipPath = previous;
        else
            SelectedZipPath = ZipFiles.FirstOrDefault();
    }

    private async Task LoadSitesAsync()
    {
        IsLoadingSites = true;
        SitesError = null;
        Targets.Clear();
        SelectedTarget = null;

        try
        {
            var sites = await Task.Run(() => new DeploymentService(_ => { }).GetSites());

            foreach (var site in sites)
                Targets.Add(TargetOption.FromSite(site));

            Targets.Add(TargetOption.CreateNewOption());
        }
        catch (Exception ex)
        {
            SitesError = $"Could not read the IIS configuration: {ex.Message}";
            // Still allow creating a new site.
            Targets.Add(TargetOption.CreateNewOption());
        }
        finally
        {
            IsLoadingSites = false;
            RefreshCommands();
        }
    }

    // -----------------------------------------------------------------
    // Deployment
    // -----------------------------------------------------------------
    private async Task RunDeploymentAsync()
    {
        if (SelectedTarget == null || string.IsNullOrEmpty(SelectedZipPath))
            return;

        LogEntries.Clear();
        HadError = false;
        IsFinished = false;
        IsBusy = true;

        var zip = SelectedZipPath;
        var target = SelectedTarget;
        var newSiteName = NewSiteName;
        var port = int.TryParse(NewSitePort, out var parsedPort) ? parsedPort : 443;

        var dispatcher = Application.Current.Dispatcher;
        void Log(string message) => dispatcher.Invoke(() => AppendLog(message));

        try
        {
            await Task.Run(() =>
            {
                var service = new DeploymentService(Log);

                if (target.IsCreateNew)
                    service.CreateNewSite(zip, newSiteName, port);
                else
                    service.DeployToSite(target.Site!.Name, zip);
            });
        }
        catch (Exception ex)
        {
            AppendLog($"ERROR: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
            IsFinished = true;
        }
    }

    private void AppendLog(string message)
    {
        var entry = new LogEntry(message);
        if (entry.Severity == LogSeverity.Error)
            HadError = true;
        LogEntries.Add(entry);
    }

    private void RefreshCommands()
    {
        BrowseCommand.RaiseCanExecuteChanged();
        BackCommand.RaiseCanExecuteChanged();
        NextCommand.RaiseCanExecuteChanged();
        StartOverCommand.RaiseCanExecuteChanged();
    }
}
