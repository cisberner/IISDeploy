using IISDeploy.Core;
using Microsoft.Web.Administration;

namespace IISDeploymentHelper;

class Program
{
    static void Main(string[] args)
    {
        var service = new DeploymentService(Console.WriteLine);

        Console.WriteLine("Searching for ZIP deployment file...");

        var currentDirectory = Directory.GetCurrentDirectory();
        var zipFiles = DeploymentService.FindZipFiles(currentDirectory).ToList();

        if (zipFiles.Count != 1)
        {
            Console.WriteLine($"ERROR: Expected exactly one ZIP file, but found {zipFiles.Count}.");
            return;
        }

        var zipFile = zipFiles[0];
        Console.WriteLine($"Found ZIP: {Path.GetFileName(zipFile)}");

        Console.WriteLine("\nListing installed IIS sites:");
        Console.WriteLine($"");

        var sites = service.GetSites();

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

        if (!int.TryParse(input, out int selectedIndex))
        {
            Console.WriteLine("ERROR: Invalid selection. Please enter a number.");
            return;
        }

        if (selectedIndex == sites.Count + 2)
            return;

        if (selectedIndex == sites.Count + 1)
        {
            CreateNewSite(service, zipFile);
            return;
        }

        if (selectedIndex < 1 || selectedIndex > sites.Count)
        {
            Console.WriteLine("ERROR: Invalid selection.");
            return;
        }

        var selectedSite = sites[selectedIndex - 1];
        service.DeployToSite(selectedSite.Name, zipFile, ReadHostingSettings());
    }

    static void CreateNewSite(DeploymentService service, string zipFile)
    {
        Console.Write("Enter the name for the new IIS site: ");
        var siteName = Console.ReadLine()?.Trim();

        if (string.IsNullOrWhiteSpace(siteName))
        {
            Console.WriteLine("ERROR: Site name cannot be empty.");
            return;
        }

        Console.Write("Enter port number to bind the site to (default 443): ");
        var portInput = Console.ReadLine();
        int port = int.TryParse(portInput, out int parsedPort) ? parsedPort : 443;

        service.CreateNewSite(zipFile, siteName, port, hostingSettings: ReadHostingSettings());
    }

    // Asks for the application pool's Start Mode and Idle Time-out and the site's Preload
    // Enabled; Enter keeps the default shown, which keeps the application warm.
    static HostingSettings ReadHostingSettings()
    {
        Console.Write("Application pool start mode - AlwaysRunning or OnDemand (default AlwaysRunning): ");
        var startModeInput = Console.ReadLine()?.Trim();
        var startMode = string.Equals(startModeInput, nameof(StartMode.OnDemand), StringComparison.OrdinalIgnoreCase)
            ? StartMode.OnDemand
            : StartMode.AlwaysRunning;

        Console.Write("Application pool idle time-out in minutes, 0 = never (default 0): ");
        var idleInput = Console.ReadLine();
        int idleTimeout = int.TryParse(idleInput, out int parsedIdle) ? parsedIdle : 0;

        Console.Write("Site preload enabled - Y or N (default Y): ");
        var preloadInput = Console.ReadLine()?.Trim();
        bool preload = !string.Equals(preloadInput, "N", StringComparison.OrdinalIgnoreCase);

        return new HostingSettings
        {
            StartMode = startMode,
            IdleTimeoutMinutes = idleTimeout,
            PreloadEnabled = preload,
        };
    }
}
