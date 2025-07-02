using Microsoft.Web.Administration;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace IISDeploymentHelper;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("🔍 Searching for ZIP deployment file...");

        var currentDirectory = Directory.GetCurrentDirectory();
        var zipFiles = Directory.GetFiles(currentDirectory, "*.zip");

        if (zipFiles.Length != 1)
        {
            Console.WriteLine($"❌ Expected exactly one ZIP file, but found {zipFiles.Length}.");
            return;
        }

        var zipFile = zipFiles[0];
        Console.WriteLine($"✅ Found ZIP: {Path.GetFileName(zipFile)}");

        Console.WriteLine("\n🌐 Listing installed IIS sites:");
        Console.WriteLine($"");
        using (var serverManager = new ServerManager())
        {

            var sites = serverManager.Sites.ToList();

            for (int i = 0; i < sites.Count; i++)
            {
                Console.WriteLine($"{i + 1}: {sites[i].Name}");
            }

            Console.WriteLine($"");
            Console.WriteLine($"{sites.Count + 1}: Create New Site");
            Console.WriteLine($"");
            Console.WriteLine($"{sites.Count + 2}: {"Cancel"}");

            Console.Write("\nEnter the number of the site to deploy to (or create new): ");
            var input = Console.ReadLine();

            if (int.TryParse(input, out int selectedIndex))
            {
                if (selectedIndex == sites.Count + 2)
                    return;

                if (selectedIndex == sites.Count + 1)
                {
                    CreateNewSite(zipFile);
                    return;
                }

                if (selectedIndex < 1 || selectedIndex > sites.Count)
                {
                    Console.WriteLine("❌ Invalid selection.");
                    return;
                }

                // Continue with your existing deployment logic...
            }

            var selectedSite = sites[selectedIndex - 1];
            DeployToSite(selectedSite, zipFile);

        }
    }

    static void CreateNewSite(string zipFile)
    {
        Console.Write("📝 Enter the name for the new IIS site: ");
        var siteName = Console.ReadLine()?.Trim();

        if (string.IsNullOrWhiteSpace(siteName))
        {
            Console.WriteLine("❌ Site name cannot be empty.");
            return;
        }

        string baseFolder = @"C:\inetpub";
        string siteFolder = Path.Combine(baseFolder, siteName);

        if (Directory.Exists(siteFolder))
        {
            Console.WriteLine("⚠️ Folder already exists. Choose another site name or clean up previous installation first.");
            return;
        }

        Directory.CreateDirectory(siteFolder);

        // Copy appsettings.json.sample and web.config if they exist
        string appSettingsFile = Path.Combine(Path.GetDirectoryName(zipFile)!, "appsettings.json.sample");
        if (File.Exists(appSettingsFile))
        {
            File.Copy(appSettingsFile, Path.Combine(siteFolder, "appsettings.json"), overwrite: false);
        }
        string webConfigFile = Path.Combine(Path.GetDirectoryName(zipFile)!, "web.config.sample");
        if (File.Exists(webConfigFile))
        {
            File.Copy(webConfigFile, Path.Combine(siteFolder, "web.config"), overwrite: false);
        }

        Console.Write("🔢 Enter port number to bind the site to (default 443): ");
        var portInput = Console.ReadLine();
        int port = int.TryParse(portInput, out int parsedPort) ? parsedPort : 443;

        string certSubject = $"CN={siteName}.local";
        X509Certificate2? cert = null;

        // Check if certificate already exists in LocalMachine\My
        using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
        {
            store.Open(OpenFlags.ReadOnly);
            var existingCerts = store.Certificates
                .Find(X509FindType.FindBySubjectDistinguishedName, certSubject, false);
            if (existingCerts.Count > 0)
            {
                cert = existingCerts[0];
                Console.WriteLine($"🔑 Using existing certificate: {cert.Subject} (Thumbprint: {cert.Thumbprint})");
            }
            store.Close();
        }

        // If not found, create and install a new certificate
        if (cert == null)
        {
            Console.WriteLine("🔐 Creating self-signed certificate...");
            string certDirectory = @"C:\Certs";
            if (!Directory.Exists(certDirectory))
            {
                Directory.CreateDirectory(certDirectory);
            }
            string certPath = Path.Combine(certDirectory, siteName + ".pfx");
            cert = CertificateGenerator.CreateSelfSignedCertificate(
                certName: $"{siteName}.local",
                outputPfxPath: certPath,
                password: "IFMAdmin123");
            CertificateGenerator.InstallCertificate(cert);
            Console.WriteLine($"✅ Created and installed new certificate: {cert.Subject} (Thumbprint: {cert.Thumbprint})");
        }

        // Add site and bind HTTPS with certificate
        using (var serverManager = new ServerManager())
        {
            if (serverManager.Sites.Any(s => s.Name.Equals(siteName, StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("❌ A site with that name already exists in IIS.");
                return;
            }

            Console.WriteLine($"⚙️ Creating application pool '{siteName}'...");
            var appPool = serverManager.ApplicationPools.Add(siteName);
            appPool.ManagedRuntimeVersion = "v4.0";

            Console.WriteLine($"🌐 Creating new site '{siteName}'...");
            var newSite = serverManager.Sites.Add(siteName, "https", $"*:{port}:", siteFolder);
            newSite.ApplicationDefaults.ApplicationPoolName = siteName;

            serverManager.CommitChanges();
            Console.WriteLine($"✅ Created new site '{siteName}' with HTTPS on port {port}.");
            Thread.Sleep(3000);
        }
    }

    static void DeployToSite(Site selectedSite, string zipFile)
    {
        var physicalPath = selectedSite.Applications["/"].VirtualDirectories["/"].PhysicalPath;
        Console.WriteLine($"📁 Site physical path: {physicalPath}");

        using (var serverManager = new ServerManager())
        {
            // 1. Stop the site
            Console.WriteLine("🛑 Stopping the IIS site...");

            try
            {
                if (selectedSite.State == ObjectState.Started || selectedSite.State == ObjectState.Starting)
                {
                    selectedSite.Stop();
                    serverManager.CommitChanges();
                    Console.WriteLine("✅ Site stopped.");
                }
                else
                {
                    Console.WriteLine("ℹ️ Site already stopped.");
                }
            }
            catch { }

            // 2. Stop the application pool
            string appPoolName = selectedSite.Applications["/"].ApplicationPoolName;
            ApplicationPool appPool = serverManager.ApplicationPools[appPoolName];

            if (appPool != null && (appPool.State == ObjectState.Started || appPool.State == ObjectState.Starting))
            {
                Console.WriteLine($"🛑 Stopping App Pool: {appPoolName}");
                appPool.Stop();
                serverManager.CommitChanges();
                Console.WriteLine("✅ App Pool stopped.");
            }

            // Optionally wait for shutdown
            Thread.Sleep(3000);

            // 3. Create backup
            var currentDirectory = Directory.GetCurrentDirectory();
            string backupFolder = Path.Combine(currentDirectory, "Backups");
            Directory.CreateDirectory(backupFolder);
            string backupZip = Path.Combine(backupFolder, $"{selectedSite.Name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}.zip");

            Console.WriteLine("💾 Creating backup...");
            ZipFile.CreateFromDirectory(physicalPath, backupZip, CompressionLevel.Optimal, includeBaseDirectory: false);
            Console.WriteLine($"✅ Backup created: {backupZip}");

            // 4. Delete old files/folders except protected ones
            Console.WriteLine("🧹 Cleaning up site folder...");
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
                        Console.WriteLine($"⚠️ Could not delete {fileName}: {ex.Message}");
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
                    Console.WriteLine($"⚠️ Could not delete directory {dir}: {ex.Message}");
                }
            }

            // 5. Extract ZIP to site folder, skip protected files
            Console.WriteLine("📦 Extracting new deployment...");
            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                foreach (var entry in archive.Entries)
                {
                    if (!entry.FullName.StartsWith("Publish/", StringComparison.OrdinalIgnoreCase))
                        continue; // Skip anything outside the Publish folder

                    var relativePath = entry.FullName.Substring("Publish/".Length);

                    if (string.IsNullOrWhiteSpace(relativePath))
                        continue; // Skip the Publish folder itself

                    var targetPath = Path.Combine(physicalPath, relativePath);

                    if (string.IsNullOrWhiteSpace(entry.Name)) // It's a directory
                    {
                        Directory.CreateDirectory(targetPath);
                        continue;
                    }

                    string fileName = Path.GetFileName(relativePath);
                    if (IsProtectedFile(fileName))
                    {
                        Console.WriteLine($"⚠️ Skipping protected file: {relativePath}");
                        continue;
                    }

                    try
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
                        entry.ExtractToFile(targetPath, overwrite: true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ Failed to extract {relativePath}: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("✅ Deployment complete.");

            // 6. Start App Pool
            if (appPool != null && appPool.State == ObjectState.Stopped)
            {
                appPool.Start();
                Console.WriteLine("🚀 App Pool started.");
            }

        }

        // 7. Start site
        if (selectedSite.State == ObjectState.Stopped)
        {
            selectedSite.Start();
            Console.WriteLine("🚀 Site started.");
        }

        Console.WriteLine("🎉 Done.");

    }

    static byte[] StringToByteArray(string hex)
    {
        // Remove all non-hex characters (including invisible Unicode)
        hex = new string(hex.Where(c => Uri.IsHexDigit(c)).ToArray());

        if (hex.Length % 2 != 0)
            throw new FormatException("Invalid hex string length.");

        return Enumerable.Range(0, hex.Length / 2)
            .Select(x => Convert.ToByte(hex.Substring(x * 2, 2), 16))
            .ToArray();
    }

    static bool IsProtectedFile(string fileName)
    {
        var lower = fileName.ToLowerInvariant();
        return lower == "appsettings.json"
            || lower == "web.config";
    }

    public class CertificateGenerator
    {
        public static X509Certificate2 CreateSelfSignedCertificate(string certName, string outputPfxPath, string password)
        {
            using (RSA rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(
                    $"CN={certName}",
                    rsa,
                    HashAlgorithmName.SHA256,
                    RSASignaturePadding.Pkcs1);

                request.CertificateExtensions.Add(
                    new X509BasicConstraintsExtension(false, false, 0, false));
                request.CertificateExtensions.Add(
                    new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment, false));
                request.CertificateExtensions.Add(
                    new X509SubjectKeyIdentifierExtension(request.PublicKey, false));

                // Valid for 5 years
                var cert = request.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddYears(5));

                // Export to PFX
                File.WriteAllBytes(outputPfxPath, cert.Export(X509ContentType.Pfx, password));

                // return cert;
                return new X509Certificate2(
                    outputPfxPath,
                    password,
                    X509KeyStorageFlags.MachineKeySet |
                    X509KeyStorageFlags.PersistKeySet |
                    X509KeyStorageFlags.Exportable);

            }
        }

        public static void InstallCertificate(X509Certificate2 cert)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(cert);
                store.Close();
            }
        }

        public static void BindCertificateToIIS(Site site, string siteName, string ip = "*", int port = 443, string certThumbprint = "")
        {

            // Remove any existing binding on 443 if needed (optional)
            var existingBinding = site.Bindings
                .FirstOrDefault(b => b.Protocol == "https" && b.EndPoint.Port == port);
            if (existingBinding != null)
            {
                site.Bindings.Remove(existingBinding);
            }

            // Create the new HTTPS binding
            var binding = site.Bindings.CreateElement("binding");

            Console.WriteLine($"Bind certificate (4).");

            binding["protocol"] = "https";
            binding["bindingInformation"] = $"{ip}:{port}:"; // hostname can go after the last colon if needed

            // Certificate settings
            binding["certificateStoreName"] = "My";

            Console.WriteLine($"Bind certificate (4.1).");
            binding["certificateHash"] = StringToByteArray(certThumbprint);

            Console.WriteLine($"Bind certificate (5).");

            site.Bindings.Add(binding);

        }
    }
}
