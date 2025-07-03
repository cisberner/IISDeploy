using IISDeploy.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace IISDeploy.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly IISDeployManager _deployManager;
        private string? _selectedZipFilePath;

        public MainWindow()
        {
            InitializeComponent();
            _deployManager = new IISDeployManager();
            // Load existing sites on startup
            RefreshSitesList();
        }

        private void AppendLog(string message, bool isError = false)
        {
            Dispatcher.Invoke(() =>
            {
                var run = new Run($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}")
                {
                    Foreground = isError ? Brushes.Red : Brushes.Black
                };
                LogOutputText.Inlines.Add(run);
                // Auto-scroll to the end
                (LogOutputText.Parent as ScrollViewer)?.ScrollToEnd();
            });
        }

        private void ProcessDeploymentResult(DeploymentResult result, string successMessagePrefix = "Operation")
        {
            foreach (var log in result.LogMessages)
            {
                AppendLog(log);
            }

            if (result.Success)
            {
                AppendLog($"{successMessagePrefix} completed successfully: {result.Message}");
            }
            else
            {
                AppendLog($"{successMessagePrefix} failed: {result.Message}", isError: true);
                if (result.Exception != null)
                {
                    AppendLog($"Exception: {result.Exception.Message}", isError: true);
                    AppendLog($"Stack Trace: {result.Exception.StackTrace}", isError: true);
                }
            }
        }

        private void SelectZipButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "ZIP Files (*.zip)|*.zip",
                Title = "Select Deployment ZIP File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedZipFilePath = openFileDialog.FileName;
                SelectedZipFileText.Text = _selectedZipFilePath;
                AppendLog($"Selected ZIP file: {_selectedZipFilePath}");
            }
        }

        private async void CreateNewSiteButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedZipFilePath))
            {
                AppendLog("Please select a deployment ZIP file first.", isError: true);
                MessageBox.Show("Please select a deployment ZIP file first.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string siteName = NewSiteNameText.Text.Trim();
            if (string.IsNullOrWhiteSpace(siteName))
            {
                AppendLog("Site Name cannot be empty.", isError: true);
                MessageBox.Show("Site Name cannot be empty.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!int.TryParse(NewSitePortText.Text, out int port) || port <= 0 || port > 65535)
            {
                AppendLog("Invalid Port Number. Must be between 1 and 65535.", isError: true);
                MessageBox.Show("Invalid Port Number. Must be between 1 and 65535.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string pfxPassword = NewSitePfxPassword.Password;
            if (string.IsNullOrEmpty(pfxPassword)) // PFX password can be empty, but let's require one for this tool's self-signed cert
            {
                 AppendLog("PFX Password cannot be empty for new self-signed certificate generation.", isError: true);
                 MessageBox.Show("PFX Password cannot be empty for new self-signed certificate generation.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                 return;
            }

            string basePath = NewSiteBasePathText.Text.Trim();
            if (string.IsNullOrWhiteSpace(basePath)) basePath = @"C:\inetpub"; // Default from core lib

            string certsPath = NewSiteCertsPathText.Text.Trim();
            if (string.IsNullOrWhiteSpace(certsPath)) certsPath = @"C:\Certs"; // Default from core lib

            AppendLog($"Attempting to create new site '{siteName}' on port {port}...");
            CreateNewSiteButton.IsEnabled = false;

            try
            {
                DeploymentResult result = await Task.Run(() =>
                    _deployManager.CreateNewSite(siteName, _selectedZipFilePath, port, pfxPassword, baseSitePath: basePath, certsDirectory: certsPath)
                );
                ProcessDeploymentResult(result, $"Site creation for '{siteName}'");
                if(result.Success) RefreshSitesList(); // Refresh site list if creation was successful
            }
            catch (Exception ex)
            {
                AppendLog($"An unexpected error occurred during site creation: {ex.Message}", isError: true);
            }
            finally
            {
                CreateNewSiteButton.IsEnabled = true;
            }
        }

        private async void DeployToExistingSiteButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedZipFilePath))
            {
                AppendLog("Please select a deployment ZIP file first.", isError: true);
                MessageBox.Show("Please select a deployment ZIP file first.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (ExistingSitesCombo.SelectedItem == null)
            {
                AppendLog("Please select an existing site from the list.", isError: true);
                MessageBox.Show("Please select an existing site from the list.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string selectedSiteName = ExistingSitesCombo.SelectedItem.ToString()!;
            AppendLog($"Attempting to deploy to existing site '{selectedSiteName}'...");
            DeployToExistingSiteButton.IsEnabled = false;
            RefreshSitesButton.IsEnabled = false;

            try
            {
                DeploymentResult result = await Task.Run(() =>
                    _deployManager.DeployToSite(selectedSiteName, _selectedZipFilePath)
                );
                ProcessDeploymentResult(result, $"Deployment to '{selectedSiteName}'");
            }
            catch (Exception ex)
            {
                AppendLog($"An unexpected error occurred during deployment: {ex.Message}", isError: true);
            }
            finally
            {
                DeployToExistingSiteButton.IsEnabled = true;
                RefreshSitesButton.IsEnabled = true;
            }
        }

        private void RefreshSitesButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshSitesList();
        }

        private async void RefreshSitesList()
        {
            AppendLog("Refreshing IIS site list...");
            ExistingSitesCombo.IsEnabled = false;
            RefreshSitesButton.IsEnabled = false;

            var result = new DeploymentResult(); // For GetIISsites logging
            List<string> sites = new List<string>();
            try
            {
                 sites = await Task.Run(() => _deployManager.GetIISsites(result));
                 ProcessDeploymentResult(result, "Site list retrieval"); // Process logs from GetIISsites
            }
            catch (Exception ex)
            {
                AppendLog($"Error refreshing site list: {ex.Message}", isError: true);
                result.Success = false; // Ensure it's marked as failed
                result.Message = ex.Message; // Populate message for ProcessDeploymentResult
                ProcessDeploymentResult(result, "Site list retrieval");
            }
            finally
            {
                 Dispatcher.Invoke(() =>
                {
                    ExistingSitesCombo.ItemsSource = sites;
                    if (sites.Any())
                    {
                        ExistingSitesCombo.SelectedIndex = 0;
                        AppendLog($"Found {sites.Count} sites.");
                    }
                    else
                    {
                        AppendLog("No IIS sites found or accessible.");
                    }
                    ExistingSitesCombo.IsEnabled = true;
                    RefreshSitesButton.IsEnabled = true;
                });
            }
        }
    }
}
