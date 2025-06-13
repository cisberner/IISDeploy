using Microsoft.Web.Administration;
using System.IO.Compression;

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
        using (var serverManager = new ServerManager())
        {

            var sites = serverManager.Sites.ToList();

            for (int i = 0; i < sites.Count; i++)
            {
                Console.WriteLine($"{i + 1}: {sites[i].Name}");
            }

            Console.WriteLine($"{sites.Count + 1}: {"Cancel"}");

            Console.Write("\nEnter the number of the site to deploy to: ");
            var input = Console.ReadLine();

            if (int.TryParse(input, out int testIndex) && testIndex == sites.Count + 1)
            {
                return;
            }

            if (!int.TryParse(input, out int selectedIndex) || selectedIndex < 1 || selectedIndex > sites.Count)
            {
                Console.WriteLine("❌ Invalid selection.");
                return;
            }

            var selectedSite = sites[selectedIndex - 1];
            var physicalPath = selectedSite.Applications["/"].VirtualDirectories["/"].PhysicalPath;

            Console.WriteLine($"📁 Site physical path: {physicalPath}");

            // 1. Stop the site
            Console.WriteLine("🛑 Stopping the IIS site...");
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

            // 7. Start site
            if (selectedSite.State == ObjectState.Stopped)
            {
                selectedSite.Start();
                Console.WriteLine("🚀 Site started.");
            }

            Console.WriteLine("🎉 Done.");
        }
    }

    static bool IsProtectedFile(string fileName)
    {
        var lower = fileName.ToLowerInvariant();
        return lower == "appsettings.json"
            || lower == "web.config";
    }
}
