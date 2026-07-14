using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using IISDeploy.Core;
using IISDeploy.Wpf.Models;
using IISDeploy.Wpf.Mvvm;
using IISDeploy.Wpf.Views;

namespace IISDeploy.Wpf.ViewModels;

public enum WizardStep
{
    Package = 0,
    Configure = 1,
    Deploy = 2,
}

public enum DeployMode
{
    UpdateExisting,
    InstallNew,
}

public sealed class MainViewModel : ObservableObject
{
    public MainViewModel()
    {
        BrowseCommand = new RelayCommand(BrowseForFolder, () => !IsBusy);
        BackCommand = new RelayCommand(GoBack, CanGoBack);
        NextCommand = new RelayCommand(GoNext, CanGoNext);
        StartOverCommand = new RelayCommand(StartOver, () => IsFinished);
        EditAppSettingsCommand = new RelayCommand(EditAppSettings, () => ConfigureSite && HasAppSettingsTemplate);
        EditWebConfigCommand = new RelayCommand(EditWebConfig, () => ConfigureSite && HasWebConfigTemplate);
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
    public RelayCommand EditAppSettingsCommand { get; }
    public RelayCommand EditWebConfigCommand { get; }
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
        WizardStep.Configure => "Install or update",
        WizardStep.Deploy => IsFinished ? "Finished" : "Deploying",
        _ => "IIS Deploy",
    };

    public string HeaderSubtitle => CurrentStep switch
    {
        WizardStep.Package => "Pick the ZIP file that contains the Publish folder to deploy.",
        WizardStep.Configure => "Install a brand new IIS site, or update an existing one.",
        WizardStep.Deploy => IsFinished ? "The operation has completed." : "Please wait while the site is being updated.",
        _ => string.Empty,
    };

    public string NextButtonText => CurrentStep == WizardStep.Configure
        ? (IsInstallMode ? "Create & deploy" : "Deploy")
        : "Next";

    // -----------------------------------------------------------------
    // Deployment mode (Install new vs. Update existing)
    // -----------------------------------------------------------------
    private DeployMode? _mode;
    public DeployMode? Mode
    {
        get => _mode;
        set
        {
            if (SetProperty(ref _mode, value))
            {
                OnPropertyChanged(nameof(IsUpdateMode));
                OnPropertyChanged(nameof(IsInstallMode));
                OnPropertyChanged(nameof(NextButtonText));
                OnPropertyChanged(nameof(ConfigurationIncomplete));
                RefreshCommands();
            }
        }
    }

    public bool IsUpdateMode => Mode == DeployMode.UpdateExisting;
    public bool IsInstallMode => Mode == DeployMode.InstallNew;

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
                _templatesLoaded = false; // reload config templates from the new package
                ProbeConfigTemplates();
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
                RefreshCommands();
        }
    }

    public bool HasSites => Targets.Count > 0;

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

    // Configuring a new site in-wizard: optionally seed appsettings.json / web.config from
    // the deployment templates (appsettings.deployment.json / web.deployment.config) and set
    // a custom application pool identity. When the site is configured here it is also started
    // automatically after installation. The file editors are offered only when the matching
    // template is present; the identity option is always available.
    private bool _templatesLoaded;

    private bool _hasAppSettingsTemplate;
    public bool HasAppSettingsTemplate
    {
        get => _hasAppSettingsTemplate;
        private set
        {
            if (SetProperty(ref _hasAppSettingsTemplate, value))
            {
                OnPropertyChanged(nameof(CanConfigureFiles));
                OnPropertyChanged(nameof(HasBothTemplates));
                OnPropertyChanged(nameof(ConfigurationIncomplete));
            }
        }
    }

    private bool _hasWebConfigTemplate;
    public bool HasWebConfigTemplate
    {
        get => _hasWebConfigTemplate;
        private set
        {
            if (SetProperty(ref _hasWebConfigTemplate, value))
            {
                OnPropertyChanged(nameof(CanConfigureFiles));
                OnPropertyChanged(nameof(HasBothTemplates));
                OnPropertyChanged(nameof(ConfigurationIncomplete));
            }
        }
    }

    // Whether the package ships config-file templates (controls the per-file edit links).
    public bool CanConfigureFiles => HasAppSettingsTemplate || HasWebConfigTemplate;
    public bool HasBothTemplates => HasAppSettingsTemplate && HasWebConfigTemplate;

    // Master toggle: configure the new site in this wizard (and start it afterwards).
    private bool _configureSite;
    public bool ConfigureSite
    {
        get => _configureSite;
        set
        {
            if (SetProperty(ref _configureSite, value))
            {
                if (value)
                    LoadConfigTemplates();
                OnPropertyChanged(nameof(ConfigurationIncomplete));
                RefreshCommands();
            }
        }
    }

    private string _appSettingsContent = string.Empty;
    public string AppSettingsContent
    {
        get => _appSettingsContent;
        set
        {
            if (SetProperty(ref _appSettingsContent, value))
            {
                OnPropertyChanged(nameof(ConfigurationIncomplete));
                RefreshCommands();
            }
        }
    }

    private string _webConfigContent = string.Empty;
    public string WebConfigContent
    {
        get => _webConfigContent;
        set
        {
            if (SetProperty(ref _webConfigContent, value))
            {
                OnPropertyChanged(nameof(ConfigurationIncomplete));
                RefreshCommands();
            }
        }
    }

    // Custom application pool identity (IIS: App Pool > Advanced Settings > Identity).
    private bool _useCustomIdentity;
    public bool UseCustomIdentity
    {
        get => _useCustomIdentity;
        set
        {
            if (SetProperty(ref _useCustomIdentity, value))
            {
                OnPropertyChanged(nameof(ConfigurationIncomplete));
                RefreshCommands();
            }
        }
    }

    private string _appPoolUserName = string.Empty;
    public string AppPoolUserName
    {
        get => _appPoolUserName;
        set
        {
            if (SetProperty(ref _appPoolUserName, value))
            {
                OnPropertyChanged(nameof(ConfigurationIncomplete));
                RefreshCommands();
            }
        }
    }

    private string _appPoolPassword = string.Empty;
    public string AppPoolPassword
    {
        get => _appPoolPassword;
        set => SetProperty(ref _appPoolPassword, value);
    }

    // Blocks installation while a configured item is still incomplete: an offered config
    // file left blank, or a custom identity with no user name.
    public bool ConfigurationIncomplete =>
        IsInstallMode && ConfigureSite
        && ((HasAppSettingsTemplate && string.IsNullOrWhiteSpace(AppSettingsContent))
            || (HasWebConfigTemplate && string.IsNullOrWhiteSpace(WebConfigContent))
            || (UseCustomIdentity && string.IsNullOrWhiteSpace(AppPoolUserName)));

    // Determines which config templates the selected package ships, controlling whether
    // each per-file edit link is shown.
    private void ProbeConfigTemplates()
    {
        string? zip = SelectedZipPath;
        HasAppSettingsTemplate = !string.IsNullOrEmpty(zip)
            && DeploymentService.PublishFileExists(zip, DeploymentService.AppSettingsTemplateName);
        HasWebConfigTemplate = !string.IsNullOrEmpty(zip)
            && DeploymentService.PublishFileExists(zip, DeploymentService.WebConfigTemplateName);

        // Reload the seed content from the new package if the user is already configuring.
        if (ConfigureSite)
            LoadConfigTemplates();
    }

    private void LoadConfigTemplates()
    {
        if (_templatesLoaded || string.IsNullOrEmpty(SelectedZipPath))
            return;

        AppSettingsContent = HasAppSettingsTemplate
            ? DeploymentService.ReadPublishFileText(SelectedZipPath, DeploymentService.AppSettingsTemplateName) ?? string.Empty
            : string.Empty;
        WebConfigContent = HasWebConfigTemplate
            ? DeploymentService.ReadPublishFileText(SelectedZipPath, DeploymentService.WebConfigTemplateName) ?? string.Empty
            : string.Empty;
        _templatesLoaded = true;
    }

    private void EditAppSettings()
    {
        LoadConfigTemplates();
        var window = new FileEditorWindow("appsettings.json", AppSettingsContent)
        {
            Owner = Application.Current.MainWindow,
        };
        if (window.ShowDialog() == true)
            AppSettingsContent = window.EditedText;
    }

    private void EditWebConfig()
    {
        LoadConfigTemplates();
        var window = new FileEditorWindow("web.config", WebConfigContent)
        {
            Owner = Application.Current.MainWindow,
        };
        if (window.ShowDialog() == true)
            WebConfigContent = window.EditedText;
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
            WizardStep.Configure => Mode != null && !IsBusy
                && ((IsUpdateMode && SelectedTarget != null)
                    || (IsInstallMode && !string.IsNullOrWhiteSpace(NewSiteName) && !ConfigurationIncomplete)),
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
        Mode = null;
        SelectedTarget = null;
        NewSiteName = string.Empty;
        NewSitePort = "443";
        ConfigureSite = false;
        AppSettingsContent = string.Empty;
        WebConfigContent = string.Empty;
        UseCustomIdentity = false;
        AppPoolUserName = string.Empty;
        AppPoolPassword = string.Empty;
        _templatesLoaded = false;
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
        }
        catch (Exception ex)
        {
            SitesError = $"Could not read the IIS configuration: {ex.Message}";
        }
        finally
        {
            OnPropertyChanged(nameof(HasSites));

            // Preselect the most likely action: update when sites exist, otherwise
            // install. The user can still switch with the mode cards.
            if (Mode == null)
                Mode = HasSites ? DeployMode.UpdateExisting : DeployMode.InstallNew;

            IsLoadingSites = false;
            RefreshCommands();
        }
    }

    // -----------------------------------------------------------------
    // Deployment
    // -----------------------------------------------------------------
    private async Task RunDeploymentAsync()
    {
        if (string.IsNullOrEmpty(SelectedZipPath) || Mode == null)
            return;
        if (IsUpdateMode && SelectedTarget == null)
            return;

        LogEntries.Clear();
        HadError = false;
        IsFinished = false;
        IsBusy = true;

        var zip = SelectedZipPath;
        var install = IsInstallMode;
        var siteName = SelectedTarget?.Site.Name;
        var newSiteName = NewSiteName;
        var port = int.TryParse(NewSitePort, out var parsedPort) ? parsedPort : 443;

        // Only pass config content when installing a new site, the user is configuring it,
        // and the package ships that template. Otherwise leave it null: appsettings.json is
        // then not created, and web.config falls back to the web.deployment.config default.
        var configuring = install && ConfigureSite;
        var appSettings = configuring && HasAppSettingsTemplate ? AppSettingsContent : null;
        var webConfig = configuring && HasWebConfigTemplate ? WebConfigContent : null;

        // Custom application pool identity, and whether to start the site after install.
        var appPoolUser = configuring && UseCustomIdentity ? AppPoolUserName : null;
        var appPoolPassword = configuring && UseCustomIdentity ? AppPoolPassword : null;
        var startAfterInstall = configuring;

        var dispatcher = Application.Current.Dispatcher;
        void Log(string message) => dispatcher.Invoke(() => AppendLog(message));

        try
        {
            await Task.Run(() =>
            {
                var service = new DeploymentService(Log);

                if (install)
                    service.CreateNewSite(zip, newSiteName, port, appSettings, webConfig,
                        appPoolUser, appPoolPassword, startAfterInstall);
                else
                    service.DeployToSite(siteName!, zip);
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
        EditAppSettingsCommand.RaiseCanExecuteChanged();
        EditWebConfigCommand.RaiseCanExecuteChanged();
    }
}
