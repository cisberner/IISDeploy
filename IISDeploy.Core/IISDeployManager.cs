using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading;

namespace IISDeploy.Core
{
    public class IISDeployManager
    {
        public string FindDeploymentZip(string directory, DeploymentResult result)
        {
            result.AddLog($"Searching for ZIP deployment file in '{directory}'...");
            try
            {
                var zipFiles = Directory.GetFiles(directory, "*.zip");

                if (zipFiles.Length == 0)
                {
                    result.Success = false;
                    result.Message = "No ZIP file found in the specified directory.";
                    result.AddLog(result.Message);
                    return string.Empty;
                }

                if (zipFiles.Length > 1)
                {
                    result.Success = false;
                    result.Message = $"Expected exactly one ZIP file, but found {zipFiles.Length}.";
                    result.AddLog(result.Message);
                    return string.Empty;
                }

                var zipFile = zipFiles[0];
                result.AddLog($"Found ZIP: {Path.GetFileName(zipFile)}");
                // Success will be true by default if we reach here with a valid file.
                // Caller should check result.Success if they use a pre-existing result object.
                // For a freshly instantiated result, it's false by default.
                // Let's ensure it's set.
                result.Success = true;
                result.Message = $"Successfully found ZIP file: {Path.GetFileName(zipFile)}";
                return zipFile;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error finding deployment ZIP: {ex.Message}";
                result.AddLog(result.Message);
                result.Exception = ex;
                return string.Empty;
            }
        }

        public List<string> GetIISsites(DeploymentResult result)
        {
            var siteNames = new List<string>();
            result.AddLog("Listing installed IIS sites...");
            try
            {
                using (var serverManager = new ServerManager())
                {
                    siteNames = serverManager.Sites.Select(s => s.Name).ToList();
                    result.AddLog($"Found {siteNames.Count} sites.");
                    foreach (var siteName in siteNames)
                    {
                        result.AddLog($"- {siteName}");
                    }
                }
                result.Success = true;
                result.Message = "Successfully retrieved IIS site list.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error listing IIS sites: {ex.Message}";
                result.AddLog(result.Message);
                result.Exception = ex;
            }
            return siteNames;
        }

        private static bool IsProtectedFile(string fileName)
        {
            var lower = fileName.ToLowerInvariant();
            return lower == "appsettings.json"
                || lower == "web.config";
        }

        private static byte[] StringToByteArray(string hex)
        {
            // Remove all non-hex characters (including invisible Unicode)
            hex = new string(hex.Where(c => Uri.IsHexDigit(c)).ToArray());

            if (hex.Length % 2 != 0)
                throw new FormatException("Invalid hex string length.");

            return Enumerable.Range(0, hex.Length / 2)
                .Select(x => Convert.ToByte(hex.Substring(x * 2, 2), 16))
                .ToArray();
        }

        public DeploymentResult DeployToSite(string siteName, string zipFilePath)
        {
            var result = new DeploymentResult(false, $"Starting deployment to site: {siteName}"); // Default to false, set true on full success
            result.AddLog($"Deployment initiated for site '{siteName}' using ZIP '{Path.GetFileName(zipFilePath)}'.");

            try
            {
                using (var serverManager = new ServerManager())
                {
                    Site selectedSite = serverManager.Sites[siteName];
                    if (selectedSite == null)
                    {
                        result.Message = $"Site '{siteName}' not found in IIS.";
                        result.AddLog(result.Message);
                        return result;
                    }

                    var physicalPath = selectedSite.Applications["/"].VirtualDirectories["/"].PhysicalPath;
                    result.AddLog($"Site physical path: {physicalPath}");

                    // 1. Stop the site
                    result.AddLog("Stopping the IIS site...");
                    try
                    {
                        if (selectedSite.State == ObjectState.Started || selectedSite.State == ObjectState.Starting)
                        {
                            selectedSite.Stop();
                            // serverManager.CommitChanges(); // Commit after all operations or critical steps
                            result.AddLog($"Site '{selectedSite.Name}' stopped.");
                        }
                        else
                        {
                            result.AddLog($"Site '{selectedSite.Name}' already stopped.");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue if stopping fails, as it might be due to various reasons not critical for deployment attempt
                        result.AddLog($"Warning: Could not definitively stop site '{selectedSite.Name}': {ex.Message}");
                    }


                    // 2. Stop the application pool
                    string appPoolName = selectedSite.Applications["/"].ApplicationPoolName;
                    ApplicationPool? appPool = serverManager.ApplicationPools[appPoolName];

                    if (appPool != null)
                    {
                        result.AddLog($"Stopping App Pool: {appPoolName}...");
                        if (appPool.State == ObjectState.Started || appPool.State == ObjectState.Starting)
                        {
                            appPool.Stop();
                            result.AddLog($"App Pool '{appPoolName}' stopped.");
                        }
                        else
                        {
                            result.AddLog($"App Pool '{appPoolName}' already stopped.");
                        }
                    }
                    else
                    {
                        result.AddLog($"Warning: App Pool '{appPoolName}' not found.");
                    }

                    serverManager.CommitChanges(); // Commit site and app pool state changes
                    result.AddLog("Committed site and app pool state changes.");

                    // Optionally wait for shutdown
                    result.AddLog("Waiting for 3 seconds for services to shut down...");
                    Thread.Sleep(3000);

                    // 3. Create backup
                    var currentDirectory = Path.GetDirectoryName(zipFilePath) ?? Directory.GetCurrentDirectory(); // Base backup near zip or CWD
                    string backupFolder = Path.Combine(currentDirectory, "Backups", siteName); // Site-specific backup folder
                    Directory.CreateDirectory(backupFolder);
                    string backupZip = Path.Combine(backupFolder, $"{selectedSite.Name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}.zip");

                    result.AddLog("Creating backup...");
                    if (Directory.Exists(physicalPath) && Directory.EnumerateFileSystemEntries(physicalPath).Any())
                    {
                        ZipFile.CreateFromDirectory(physicalPath, backupZip, CompressionLevel.Optimal, includeBaseDirectory: false);
                        result.AddLog($"Backup created: {backupZip}");
                    }
                    else
                    {
                        result.AddLog($"Skipping backup for '{physicalPath}' as it's empty or doesn't exist.");
                    }


                    // 4. Delete old files/folders except protected ones
                    result.AddLog("Cleaning up site folder...");
                    if (Directory.Exists(physicalPath))
                    {
                        foreach (var file in Directory.GetFiles(physicalPath))
                        {
                            var fileName = Path.GetFileName(file);
                            if (!IsProtectedFile(fileName))
                            {
                                try
                                {
                                    File.Delete(file);
                                    result.AddLog($"Deleted file: {fileName}");
                                }
                                catch (Exception ex)
                                {
                                    result.AddLog($"⚠️ Could not delete file {fileName}: {ex.Message}");
                                }
                            }
                            else
                            {
                                result.AddLog($"Protected file skipped: {fileName}");
                            }
                        }

                        foreach (var dir in Directory.GetDirectories(physicalPath))
                        {
                            try
                            {
                                Directory.Delete(dir, recursive: true);
                                result.AddLog($"Deleted directory: {Path.GetFileName(dir)}");
                            }
                            catch (Exception ex)
                            {
                                result.AddLog($"⚠️ Could not delete directory {Path.GetFileName(dir)}: {ex.Message}");
                            }
                        }
                        result.AddLog("Site folder cleanup complete.");
                    }
                    else
                    {
                        result.AddLog($"Physical path '{physicalPath}' does not exist. Creating it.");
                        Directory.CreateDirectory(physicalPath);
                    }


                    // 5. Extract ZIP to site folder, skip protected files
                    result.AddLog("Extracting new deployment...");
                    using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            // Original logic expects files under a "Publish/" folder in the zip.
                            if (!entry.FullName.StartsWith("Publish/", StringComparison.OrdinalIgnoreCase))
                            {
                                // result.AddLog($"Skipping entry not in Publish/ folder: {entry.FullName}");
                                continue;
                            }

                            var relativePath = entry.FullName.Substring("Publish/".Length);

                            if (string.IsNullOrWhiteSpace(relativePath)) // Skip the Publish folder itself
                                continue;

                            var targetPath = Path.Combine(physicalPath, relativePath);
                            var targetDir = Path.GetDirectoryName(targetPath);
                            if(targetDir != null && !Directory.Exists(targetDir))
                            {
                                Directory.CreateDirectory(targetDir);
                            }

                            if (string.IsNullOrWhiteSpace(entry.Name)) // It's a directory entry
                            {
                                Directory.CreateDirectory(targetPath);
                                result.AddLog($"Created directory: {targetPath}");
                                continue;
                            }

                            string entryFileName = Path.GetFileName(relativePath);
                            if (IsProtectedFile(entryFileName))
                            {
                                result.AddLog($"Skipping protected file from ZIP: {relativePath}");
                                continue;
                            }

                            try
                            {
                                entry.ExtractToFile(targetPath, overwrite: true);
                                result.AddLog($"Extracted: {relativePath}");
                            }
                            catch (Exception ex)
                            {
                                result.AddLog($"⚠️ Failed to extract {relativePath}: {ex.Message}");
                                // Potentially collect these errors and decide if deployment is a failure
                            }
                        }
                    }
                    result.AddLog("Extraction complete.");

                    // 6. Start App Pool
                    if (appPool != null)
                    {
                         if (appPool.State == ObjectState.Stopped)
                        {
                            result.AddLog($"Starting App Pool '{appPool.Name}'...");
                            appPool.Start();
                            result.AddLog("App Pool started.");
                        }
                        else
                        {
                            result.AddLog($"App Pool '{appPool.Name}' was not stopped or already started.");
                        }
                    }


                    // 7. Start site
                    if (selectedSite.State == ObjectState.Stopped)
                    {
                        result.AddLog($"Starting site '{selectedSite.Name}'...");
                        selectedSite.Start();
                        result.AddLog("Site started.");
                    }
                    else
                    {
                         result.AddLog($"Site '{selectedSite.Name}' was not stopped or already started.");
                    }

                    serverManager.CommitChanges(); // Commit final state changes
                    result.AddLog("Committed final site and app pool state changes.");

                    result.Success = true;
                    result.Message = $"Deployment to site '{siteName}' completed successfully.";
                    result.AddLog("🎉 Done.");
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error deploying to site {siteName}: {ex.Message}";
                result.AddLog($"Critical Error: {result.Message}");
                result.Exception = ex;
            }

            return result;
        }

        public DeploymentResult CreateNewSite(string siteName, string zipFilePath, int port, string pfxPassword,
                                    string appPoolRuntimeVersion = "v4.0",
                                    string baseSitePath = @"C:\inetpub",
                                    string certsDirectory = @"C:\Certs")
        {
            var result = new DeploymentResult(false, $"Attempting to create new site: {siteName}");
            result.AddLog($"Initiating creation of new IIS site: '{siteName}'. Port: {port}, AppPoolRuntime: {appPoolRuntimeVersion}");
            result.AddLog($"Base site path: {baseSitePath}, Certs directory: {certsDirectory}");

            try
            {
                if (string.IsNullOrWhiteSpace(siteName))
                {
                    result.Message = "Site name cannot be empty.";
                    result.AddLog(result.Message);
                    return result;
                }

                string siteFolder = Path.Combine(baseSitePath, siteName);
                result.AddLog($"Site physical folder will be: {siteFolder}");

                if (Directory.Exists(siteFolder))
                {
                    result.Message = $"Site folder '{siteFolder}' already exists. Choose another site name or clean up previous installation first.";
                    result.AddLog(result.Message);
                    return result; // Cannot proceed if folder exists, to prevent accidental overwrite
                }

                Directory.CreateDirectory(siteFolder);
                result.AddLog($"Created site folder: {siteFolder}");

                // Copy appsettings.json.sample and web.config.sample if they exist next to the ZIP file
                string? zipDirectory = Path.GetDirectoryName(zipFilePath);
                if (!string.IsNullOrEmpty(zipDirectory))
                {
                    string appSettingsSampleFile = Path.Combine(zipDirectory, "appsettings.json.sample");
                    if (File.Exists(appSettingsSampleFile))
                    {
                        File.Copy(appSettingsSampleFile, Path.Combine(siteFolder, "appsettings.json"), overwrite: false);
                        result.AddLog("Copied appsettings.json.sample to site folder as appsettings.json.");
                    }
                    else
                    {
                        result.AddLog("appsettings.json.sample not found, skipping copy.");
                    }

                    string webConfigSampleFile = Path.Combine(zipDirectory, "web.config.sample");
                    if (File.Exists(webConfigSampleFile))
                    {
                        File.Copy(webConfigSampleFile, Path.Combine(siteFolder, "web.config"), overwrite: false);
                        result.AddLog("Copied web.config.sample to site folder as web.config.");
                    }
                    else
                    {
                        result.AddLog("web.config.sample not found, skipping copy.");
                    }
                } else {
                    result.AddLog("Could not determine ZIP file directory, skipping copy of sample config files.");
                }


                string certSubject = $"CN={siteName}.local"; // As per original logic
                X509Certificate2? cert = null;

                result.AddLog($"Checking for existing certificate with subject: {certSubject}");
                using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
                {
                    store.Open(OpenFlags.ReadOnly);
                    var existingCerts = store.Certificates
                        .Find(X509FindType.FindBySubjectDistinguishedName, certSubject, false); // exact match
                    if (existingCerts.Count > 0)
                    {
                        cert = existingCerts[0];
                        result.AddLog($"Using existing certificate: {cert.Subject} (Thumbprint: {cert.Thumbprint})");
                    }
                    store.Close();
                }

                if (cert == null)
                {
                    result.AddLog("No existing certificate found. Creating a new self-signed certificate.");
                    if (!Directory.Exists(certsDirectory))
                    {
                        Directory.CreateDirectory(certsDirectory);
                        result.AddLog($"Created certificate output directory: {certsDirectory}");
                    }
                    string certPath = Path.Combine(certsDirectory, siteName + ".pfx");

                    // Use the CertificateGenerator class
                    cert = Core.CertificateGenerator.CreateSelfSignedCertificate(
                        certName: $"{siteName}.local", // Original logic used siteName.local
                        outputPfxPath: certPath,
                        password: pfxPassword,
                        result: result); // Pass the result object for logging within the generator

                    if (cert == null || !result.LogMessages.Contains("PFX file created.")) // Check if cert creation was successful by checking logs or a direct status if available
                    {
                        // CreateSelfSignedCertificate already logs errors to 'result'
                        result.Message = result.Message актуальнее // "Failed to create self-signed certificate. Check logs for details.";
                        return result; // Stop if certificate creation failed
                    }
                    result.AddLog($"Self-signed certificate PFX created at: {certPath}");

                    if (!Core.CertificateGenerator.InstallCertificate(cert, result))
                    {
                        // InstallCertificate logs errors to 'result'
                         result.Message = result.Message актуальнее // "Failed to install the new certificate. Check logs for details.";
                        return result; // Stop if certificate installation failed
                    }
                    result.AddLog($"New certificate created and installed: {cert.Subject} (Thumbprint: {cert.Thumbprint})");
                }

                if (cert == null) // Should not happen if logic above is correct, but as a safeguard
                {
                    result.Message = "Certificate is null after generation/retrieval process. Cannot proceed.";
                    result.AddLog(result.Message);
                    return result;
                }


                using (var serverManager = new ServerManager())
                {
                    if (serverManager.Sites.Any(s => s.Name.Equals(siteName, StringComparison.OrdinalIgnoreCase)))
                    {
                        result.Message = $"A site with the name '{siteName}' already exists in IIS.";
                        result.AddLog(result.Message);
                        return result;
                    }

                    result.AddLog($"Creating application pool '{siteName}' with runtime '{appPoolRuntimeVersion}'.");
                    ApplicationPool newAppPool = serverManager.ApplicationPools.Add(siteName);
                    newAppPool.ManagedRuntimeVersion = appPoolRuntimeVersion;
                    // Consider other app pool settings if needed (e.g., identity)

                    result.AddLog($"Creating new IIS site '{siteName}'.");
                    // The binding is created with protocol, IP, port, and physical path.
                    // For HTTPS, the certificate needs to be associated.
                    // The Add method with "https" and a thumbprint in the binding info is one way.
                    // Another is to add HTTP binding then add HTTPS binding separately with cert hash.
                    // Original console app adds site then binds the certificate.
                    // Let's use the direct HTTPS binding method with thumbprint.

                    var bindingInformation = $"*:{port}:"; // Hostname can be added here if needed: "*:{port}:yourhost.com"
                    Site newSite = serverManager.Sites.Add(siteName, siteFolder, port); // Add basic site first
                    newSite.ServerAutoStart = true;

                    // Remove default HTTP binding if only HTTPS is desired, or configure as needed.
                    var defaultHttpBinding = newSite.Bindings.FirstOrDefault(b => b.Protocol == "http");
                    if (defaultHttpBinding != null)
                    {
                        result.AddLog($"Removing default HTTP binding for site '{siteName}'.");
                        newSite.Bindings.Remove(defaultHttpBinding);
                    }

                    // Add HTTPS binding
                    result.AddLog($"Adding HTTPS binding for site '{siteName}' on port {port} with certificate thumbprint '{cert.Thumbprint}'.");
                    var httpsBinding = newSite.Bindings.CreateElement("binding");
                    httpsBinding["protocol"] = "https";
                    httpsBinding["bindingInformation"] = bindingInformation;
                    httpsBinding["certificateStoreName"] = StoreName.My.ToString(); // Standard store for machine certs
                    httpsBinding["certificateHash"] = StringToByteArray(cert.Thumbprint); // Uses the helper method
                    newSite.Bindings.Add(httpsBinding);

                    newSite.ApplicationDefaults.ApplicationPoolName = siteName;
                    result.AddLog($"Site '{siteName}' configured to use application pool '{siteName}'.");

                    serverManager.CommitChanges();
                    result.AddLog($"IIS changes committed. Site '{siteName}' created with HTTPS on port {port}.");

                    // After site creation, deploy the application from the ZIP file
                    result.AddLog($"Deploying application content from '{Path.GetFileName(zipFilePath)}' to new site '{siteName}'.");
                    // Call DeployToSite without backup and without start/stop of site/appool as it's a fresh setup
                    // However, DeployToSite is designed for existing sites.
                    // We need to adapt the extraction logic here or make DeployToSite more flexible.
                    // For now, let's replicate the extraction part of DeployToSite.

                    result.AddLog("Extracting application content...");
                    using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            if (!entry.FullName.StartsWith("Publish/", StringComparison.OrdinalIgnoreCase))
                                continue;

                            var relativePath = entry.FullName.Substring("Publish/".Length);
                            if (string.IsNullOrWhiteSpace(relativePath))
                                continue;

                            var targetPath = Path.Combine(siteFolder, relativePath);
                             var targetDir = Path.GetDirectoryName(targetPath);
                            if(targetDir != null && !Directory.Exists(targetDir))
                            {
                                Directory.CreateDirectory(targetDir);
                            }

                            if (string.IsNullOrWhiteSpace(entry.Name)) // Directory
                            {
                                Directory.CreateDirectory(targetPath);
                                result.AddLog($"Created directory from ZIP: {targetPath}");
                                continue;
                            }

                            // Do not overwrite appsettings.json/web.config if they were copied from samples
                            string entryFileName = Path.GetFileName(relativePath);
                            if (IsProtectedFile(entryFileName) && File.Exists(targetPath))
                            {
                                result.AddLog($"Skipping extraction of protected file '{relativePath}' as it already exists (likely from sample).");
                                continue;
                            }

                            entry.ExtractToFile(targetPath, overwrite: true);
                            result.AddLog($"Extracted from ZIP: {relativePath}");
                        }
                    }
                    result.AddLog("Application content extraction complete.");

                    // Start the site and app pool
                    result.AddLog($"Starting application pool '{newAppPool.Name}'.");
                    if(newAppPool.State == ObjectState.Stopped) newAppPool.Start();

                    result.AddLog($"Starting site '{newSite.Name}'.");
                    if(newSite.State == ObjectState.Stopped) newSite.Start();

                    serverManager.CommitChanges(); // Commit start states
                    result.AddLog($"Site and AppPool started. Final commit.");

                    result.Success = true;
                    result.Message = $"Successfully created and deployed to new site '{siteName}'.";
                    result.AddLog("🎉 New site creation and deployment process complete.");
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error creating new site {siteName}: {ex.Message}";
                result.AddLog($"Critical Error: {result.Message}");
                result.Exception = ex;
            }

            return result;
        }
    }

    // CertificateGenerator class and its methods will be added next.
    // For now, keeping them separate for clarity during refactoring.
}
