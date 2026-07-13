using Microsoft.Web.Administration;
using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography.X509Certificates;

namespace IISDeploy.Core;

/// <summary>
/// Core IIS deployment logic shared by the command line tool and the WPF UI.
/// All progress and diagnostic output is reported through the <see cref="Log"/>
/// callback supplied at construction time, so the exact same messages the CLI
/// prints to the console can be streamed into the UI log view.
/// </summary>
public sealed class DeploymentService
{
    private readonly Action<string> _log;

    public DeploymentService(Action<string> log)
    {
        _log = log ?? (_ => { });
    }

    private void Log(string message) => _log(message);

    /// <summary>
    /// Returns the ZIP files found in <paramref name="folder"/>, matching the
    /// CLI's discovery in the current working directory.
    /// </summary>
    public static IReadOnlyList<string> FindZipFiles(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            return Array.Empty<string>();

        return Directory.GetFiles(folder, "*.zip");
    }

    // Template files shipped inside the deployment ZIP (under Publish/) used to
    // seed a new site's appsettings.json and web.config.
    public const string AppSettingsTemplateName = "appsettings.deployment.json";
    public const string WebConfigTemplateName = "web.deployment.config";

    /// <summary>
    /// Reads the text of a file stored under the "Publish/" folder of the ZIP,
    /// or null when the ZIP or entry does not exist. Used to load the config
    /// templates for a new site.
    /// </summary>
    public static string? ReadPublishFileText(string zipFile, string fileName)
    {
        if (string.IsNullOrWhiteSpace(zipFile) || !File.Exists(zipFile))
            return null;

        try
        {
            using var archive = ZipFile.OpenRead(zipFile);
            var entry = FindPublishEntry(archive, fileName);
            if (entry == null)
                return null;

            using var reader = new StreamReader(entry.Open());
            return reader.ReadToEnd();
        }
        catch
        {
            return null; // not a readable ZIP - treat as no template
        }
    }

    /// <summary>
    /// Returns whether a file exists under the "Publish/" folder of the ZIP. Used to
    /// decide which configuration templates a new site can offer.
    /// </summary>
    public static bool PublishFileExists(string zipFile, string fileName)
    {
        if (string.IsNullOrWhiteSpace(zipFile) || !File.Exists(zipFile))
            return false;

        try
        {
            using var archive = ZipFile.OpenRead(zipFile);
            return FindPublishEntry(archive, fileName) != null;
        }
        catch
        {
            return false; // not a readable ZIP
        }
    }

    private static System.IO.Compression.ZipArchiveEntry? FindPublishEntry(ZipArchive archive, string fileName)
    {
        string target = "Publish/" + fileName;
        return archive.Entries.FirstOrDefault(e =>
            e.FullName.Replace('\\', '/').Equals(target, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Enumerates the installed IIS sites.
    /// </summary>
    public IReadOnlyList<SiteInfo> GetSites()
    {
        using var serverManager = new ServerManager();
        var result = new List<SiteInfo>();

        foreach (var site in serverManager.Sites)
        {
            string? physicalPath = null;
            string? state = null;

            try
            {
                physicalPath = site.Applications["/"]?.VirtualDirectories["/"]?.PhysicalPath;
            }
            catch { /* some sites have no root application */ }

            try
            {
                state = site.State.ToString();
            }
            catch { /* state can throw when the app pool is misconfigured */ }

            result.Add(new SiteInfo
            {
                Name = site.Name,
                PhysicalPath = physicalPath,
                State = state,
            });
        }

        return result;
    }

    // ---------------------------------------------------------------------
    // Create a brand new site
    // ---------------------------------------------------------------------

    /// <summary>
    /// Creates a new IIS site (app pool, HTTPS binding, self-signed certificate)
    /// and extracts the deployment ZIP into it. Returns false when a validation
    /// check prevents the site from being created.
    /// </summary>
    public bool CreateNewSite(string zipFile, string siteName, int port = 443,
        string? appSettingsContent = null, string? webConfigContent = null,
        string? appPoolUserName = null, string? appPoolPassword = null,
        bool startAfterInstall = false)
    {
        siteName = siteName?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(siteName))
        {
            Log("ERROR: Site name cannot be empty.");
            return false;
        }

        string baseFolder = @"C:\inetpub";
        string siteFolder = Path.Combine(baseFolder, siteName);

        if (Directory.Exists(siteFolder))
        {
            Log("WARNING: Folder already exists. Choose another site name or clean up previous installation first.");
            return false;
        }

        Directory.CreateDirectory(siteFolder);

        // appsettings.json is never taken from the ZIP payload for a new site: write it from
        // the content the user entered, or fall back to an appsettings.json.sample next to
        // the ZIP. It stays protected during extraction so it is never overwritten.
        string? zipDir = Path.GetDirectoryName(zipFile);
        WriteInitialConfigFile(siteFolder, "appsettings.json", appSettingsContent,
            zipDir != null ? Path.Combine(zipDir, "appsettings.json.sample") : null);

        // web.config: use the content the user entered, otherwise seed it from the
        // web.deployment.config template in the ZIP so a new site gets a working default.
        string? webContent = webConfigContent ?? ReadPublishFileText(zipFile, WebConfigTemplateName);
        if (webContent != null)
            WriteInitialConfigFile(siteFolder, "web.config", webContent, null);

        string certSubject = $"CN={siteName}.local";

        // We only need the thumbprint to bind. Never keep an open X509Certificate2
        // (and its private-key handle) alive during the IIS binding commit below:
        // HTTP.sys opens the key from the store itself, and a competing open handle
        // makes AddSslCertificate fail with 0x80070520 ("logon session does not exist").
        string? certThumbprint = null;

        // Check if certificate already exists in LocalMachine\My
        using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
        {
            store.Open(OpenFlags.ReadOnly);
            var existingCerts = store.Certificates
                .Find(X509FindType.FindBySubjectDistinguishedName, certSubject, false);
            if (existingCerts.Count > 0)
            {
                using var existing = existingCerts[0];
                certThumbprint = existing.Thumbprint;
                Log($"Using existing certificate: {existing.Subject} (Thumbprint: {certThumbprint})");
            }
            store.Close();
        }

        // If not found, create and install a new certificate
        if (certThumbprint == null)
        {
            Log("Creating self-signed certificate...");
            string certDirectory = @"C:\Certs";
            if (!Directory.Exists(certDirectory))
            {
                Directory.CreateDirectory(certDirectory);
            }
            string certPath = Path.Combine(certDirectory, siteName + ".pfx");

            // Dispose the cert (and its private-key handle) as soon as it is
            // persisted to the machine store; only the thumbprint is kept.
            using var cert = CertificateGenerator.CreateSelfSignedCertificate(
                certName: $"{siteName}.local",
                outputPfxPath: certPath,
                password: "IFMAdmin123");
            CertificateGenerator.InstallCertificate(cert);
            certThumbprint = cert.Thumbprint;
            Log($"Created and installed new certificate: {cert.Subject} (Thumbprint: {certThumbprint})");
        }

        // 1) Create the app pool + site with the HTTPS binding, but WITHOUT the SSL
        //    certificate yet. Committing an https binding with no cert hash does not
        //    touch HTTP.sys, so it cannot fail with 0x80070520.
        using (var serverManager = new ServerManager())
        {
            if (serverManager.Sites.Any(s => s.Name.Equals(siteName, StringComparison.OrdinalIgnoreCase)))
            {
                Log("ERROR: A site with that name already exists in IIS.");
                return false;
            }

            Log($"Creating application pool '{siteName}'...");
            var appPool = serverManager.ApplicationPools.Add(siteName);
            appPool.ManagedRuntimeVersion = "v4.0";

            // Optional custom application pool identity (IIS: App Pool > Advanced
            // Settings > Identity > Custom account).
            if (!string.IsNullOrWhiteSpace(appPoolUserName))
            {
                appPool.ProcessModel.IdentityType = ProcessModelIdentityType.SpecificUser;
                appPool.ProcessModel.UserName = appPoolUserName;
                appPool.ProcessModel.Password = appPoolPassword ?? string.Empty;
                Log($"Application pool identity set to '{appPoolUserName}'.");
            }

            Log($"Creating new site '{siteName}'...");
            var newSite = serverManager.Sites.Add(siteName, "https", $"*:{port}:", siteFolder);
            newSite.ApplicationDefaults.ApplicationPoolName = siteName;

            // Auto-start only when the site was configured in this tool; otherwise keep it
            // stopped so it can be configured before its first start.
            newSite.ServerAutoStart = startAfterInstall;

            serverManager.CommitChanges();
            Log($"Created new site '{siteName}' with HTTPS on port {port}.");
        }

        // 2) Bind the certificate as a separate step (this is the call that touches
        //    HTTP.sys and can fail with 0x80070520 right after the cert was created).
        Log("Binding certificate to HTTPS...");
        BindCertificate(siteName, port, certThumbprint);

        // Extract ZIP contents to the new site folder. appsettings.json and web.config are
        // protected (never taken from the payload); their content comes from the templates
        // above. The *.deployment.* templates themselves are skipped silently.
        ExtractZipToFolder(zipFile, siteFolder);

        if (startAfterInstall)
        {
            StartNewSite(siteName);
            Log("Done.");
            Log($"The new site '{siteName}' is installed and started.");
        }
        else
        {
            // Leave the new site stopped so it can be configured before first start.
            Log("Done.");
            Log($"The new site '{siteName}' is installed, but needs to be configured before starting.");
        }
        return true;
    }

    private void StartNewSite(string siteName)
    {
        try
        {
            using var serverManager = new ServerManager();
            var site = serverManager.Sites[siteName];
            if (site == null)
                return;

            string poolName = site.Applications["/"].ApplicationPoolName;
            var pool = serverManager.ApplicationPools[poolName];
            if (pool != null && pool.State == ObjectState.Stopped)
            {
                pool.Start();
                Log("App Pool started.");
            }

            if (site.State == ObjectState.Stopped)
            {
                site.Start();
                Log("Site started.");
            }
        }
        catch (Exception ex)
        {
            Log($"WARNING: Could not start the site automatically: {ex.Message}");
        }
    }

    // ---------------------------------------------------------------------
    // Deploy into an existing site
    // ---------------------------------------------------------------------

    /// <summary>
    /// Stops the target site, backs up its current content, cleans the folder and
    /// extracts the deployment ZIP, then restarts the site and app pool.
    /// </summary>
    public void DeployToSite(string siteName, string zipFile)
    {
        using var serverManager = new ServerManager();

        var selectedSite = serverManager.Sites[siteName];
        if (selectedSite == null)
        {
            Log($"ERROR: Site '{siteName}' not found.");
            return;
        }

        var physicalPath = selectedSite.Applications["/"].VirtualDirectories["/"].PhysicalPath;
        Log($"Site physical path: {physicalPath}");

        // 1. Stop the site
        Log("Stopping the IIS site...");

        try
        {
            if (selectedSite.State == ObjectState.Started || selectedSite.State == ObjectState.Starting)
            {
                selectedSite.Stop();
                serverManager.CommitChanges();
                Log("Site stopped.");
            }
            else
            {
                Log("Site already stopped.");
            }
        }
        catch { }

        // 2. Stop the application pool
        string appPoolName = selectedSite.Applications["/"].ApplicationPoolName;
        ApplicationPool appPool = serverManager.ApplicationPools[appPoolName];

        if (appPool != null && (appPool.State == ObjectState.Started || appPool.State == ObjectState.Starting))
        {
            Log($"Stopping App Pool: {appPoolName}");
            appPool.Stop();
            serverManager.CommitChanges();
            Log("App Pool stopped.");
        }

        // Optionally wait for shutdown
        Thread.Sleep(3000);

        // 3. Create backup
        string baseDirectory = Path.GetDirectoryName(zipFile) ?? Directory.GetCurrentDirectory();
        string backupFolder = Path.Combine(baseDirectory, "Backups");
        Directory.CreateDirectory(backupFolder);
        string backupZip = Path.Combine(backupFolder, $"{selectedSite.Name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}.zip");

        Log("Creating backup...");
        ZipFile.CreateFromDirectory(physicalPath, backupZip, CompressionLevel.Optimal, includeBaseDirectory: false);
        Log($"Backup created: {backupZip}");

        // 4. Delete old files/folders except protected ones
        Log("Cleaning up site folder...");
        foreach (var file in Directory.GetFiles(physicalPath))
        {
            var fileName = Path.GetFileName(file);
            if (!IsProtectedFile(fileName))
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    Log($"WARNING: Could not delete {fileName}: {ex.Message}");
                }
            }
        }

        foreach (var dir in Directory.GetDirectories(physicalPath))
        {
            try
            {
                Directory.Delete(dir, recursive: true);
            }
            catch (Exception ex)
            {
                Log($"WARNING: Could not delete directory {dir}: {ex.Message}");
            }
        }

        // 5. Extract ZIP to site folder, skipping protected files so the site's own
        //    appsettings.json and web.config are never overwritten.
        ExtractZipToFolder(zipFile, physicalPath);

        Log("Deployment complete.");

        // 6. Start App Pool
        if (appPool != null && appPool.State == ObjectState.Stopped)
        {
            appPool.Start();
            Log("App Pool started.");
        }

        // 7. Start site
        if (selectedSite.State == ObjectState.Stopped)
        {
            selectedSite.Start();
            Log("Site started.");
        }

        Log("Done.");
    }

    // ---------------------------------------------------------------------
    // Shared helpers
    // ---------------------------------------------------------------------

    private void ExtractZipToFolder(string zipFile, string targetFolder)
    {
        Log("Extracting deployment files...");
        using (ZipArchive archive = ZipFile.OpenRead(zipFile))
        {
            foreach (var entry in archive.Entries)
            {
                if (!entry.FullName.StartsWith("Publish/", StringComparison.OrdinalIgnoreCase))
                    continue; // Skip anything outside the Publish folder

                var relativePath = entry.FullName.Substring("Publish/".Length);

                if (string.IsNullOrWhiteSpace(relativePath))
                    continue; // Skip the Publish folder itself

                var targetPath = Path.Combine(targetFolder, relativePath);

                if (string.IsNullOrWhiteSpace(entry.Name)) // It's a directory
                {
                    Directory.CreateDirectory(targetPath);
                    continue;
                }

                string fileName = Path.GetFileName(relativePath);

                // The deployment templates (appsettings.deployment.json / web.deployment.config)
                // only seed a new site's config files and must never land in the deployed
                // folder. Skip them silently - logging would just confuse the administrator
                // with files they never chose to deploy.
                if (IsDeploymentTemplate(fileName))
                    continue;

                if (IsProtectedFile(fileName))
                {
                    Log($"WARNING: Skipping protected file: {relativePath}");
                    continue;
                }

                try
                {
                    var dir = Path.GetDirectoryName(targetPath);
                    if (dir != null)
                        Directory.CreateDirectory(dir);
                    entry.ExtractToFile(targetPath, overwrite: true);
                }
                catch (Exception ex)
                {
                    Log($"WARNING: Failed to extract {relativePath}: {ex.Message}");
                }
            }
        }
        Log("Deployment files extracted.");
    }

    private void WriteInitialConfigFile(string siteFolder, string fileName, string? content, string? samplePath)
    {
        string target = Path.Combine(siteFolder, fileName);

        if (content != null)
        {
            File.WriteAllText(target, content);
            Log($"Created {fileName}.");
        }
        else if (samplePath != null && File.Exists(samplePath))
        {
            File.Copy(samplePath, target, overwrite: false);
        }
    }

    // The runtime config files: never overwritten during extraction, and preserved by
    // the update cleanup step.
    private static bool IsProtectedFile(string fileName)
    {
        var lower = fileName.ToLowerInvariant();
        return lower == "appsettings.json"
            || lower == "web.config";
    }

    // The deployment templates used only to seed a new site's config files. They must
    // never appear in the deployed folder.
    private static bool IsDeploymentTemplate(string fileName)
    {
        return string.Equals(fileName, AppSettingsTemplateName, StringComparison.OrdinalIgnoreCase)
            || string.Equals(fileName, WebConfigTemplateName, StringComparison.OrdinalIgnoreCase);
    }

    // Binds the certificate (already in LocalMachine\My) to the site's HTTPS binding.
    // Tries IIS/Microsoft.Web.Administration first with a couple of retries, then falls
    // back to netsh - which runs in a separate process and therefore reads the freshly
    // created private key cleanly, exactly like a manual re-run of this tool would.
    private void BindCertificate(string siteName, int port, string certThumbprint)
    {
        const int maxAttempts = 3;
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                using var serverManager = new ServerManager();
                var site = serverManager.Sites[siteName];
                if (site == null)
                {
                    Log($"ERROR: Site '{siteName}' not found while binding certificate.");
                    return;
                }

                CertificateGenerator.BindCertificateToIIS(site, siteName, port: port, certThumbprint: certThumbprint);
                serverManager.CommitChanges();
                Log($"Certificate bound to HTTPS via IIS (attempt {attempt}).");
                return;
            }
            catch (Exception ex)
            {
                Log($"WARNING: IIS binding attempt {attempt} failed (0x{(uint)ex.HResult:X8}): {ex.Message}");
                if (attempt < maxAttempts)
                    Thread.Sleep(1500);
            }
        }

        // Fallback: register the certificate with HTTP.sys directly. netsh is a fresh
        // process, which sidesteps the 0x80070520 ("logon session does not exist") error.
        Log("Falling back to 'netsh http add sslcert'...");
        BindCertificateViaNetsh(port, certThumbprint);
    }

    private void BindCertificateViaNetsh(int port, string certThumbprint)
    {
        string appId = "{" + Guid.NewGuid() + "}";

        // Remove any stale registration on this port first (ignore failures), then add.
        RunNetsh($"http delete sslcert ipport=0.0.0.0:{port}", ignoreFailure: true);

        bool ok = RunNetsh(
            $"http add sslcert ipport=0.0.0.0:{port} certhash={certThumbprint} appid={appId} certstorename=MY",
            ignoreFailure: false);

        if (ok)
            Log($"Certificate registered with HTTP.sys on port {port} via netsh.");
        else
            Log($"ERROR: Failed to register the certificate on port {port} via netsh.");
    }

    private bool RunNetsh(string arguments, bool ignoreFailure)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using var proc = Process.Start(psi);
            if (proc == null)
                return false;

            string stdout = proc.StandardOutput.ReadToEnd();
            string stderr = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            if (!string.IsNullOrWhiteSpace(stdout))
                Log(stdout.Trim());
            if (!string.IsNullOrWhiteSpace(stderr))
                Log(stderr.Trim());

            return proc.ExitCode == 0;
        }
        catch (Exception ex)
        {
            if (!ignoreFailure)
                Log($"WARNING: netsh {arguments} failed: {ex.Message}");
            return false;
        }
    }
}
