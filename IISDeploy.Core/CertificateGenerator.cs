using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace IISDeploy.Core
{
    internal class CertificateGenerator
    {
        public static X509Certificate2? CreateSelfSignedCertificate(string certName, string outputPfxPath, string password, DeploymentResult result)
        {
            try
            {
                result.AddLog($"Creating self-signed certificate: {certName}");
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
                    result.AddLog($"Certificate created in memory. Exporting to PFX: {outputPfxPath}");

                    // Export to PFX
                    File.WriteAllBytes(outputPfxPath, cert.Export(X509ContentType.Pfx, password));
                    result.AddLog("PFX file created.");

                    // Return a new X509Certificate2 object with appropriate storage flags
                    var installedCert = new X509Certificate2(
                        outputPfxPath,
                        password,
                        X509KeyStorageFlags.MachineKeySet |
                        X509KeyStorageFlags.PersistKeySet |
                        X509KeyStorageFlags.Exportable);

                    result.AddLog($"Certificate loaded from PFX for installation. Subject: {installedCert.Subject}, Thumbprint: {installedCert.Thumbprint}");
                    return installedCert;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error creating self-signed certificate: {ex.Message}";
                result.AddLog(result.Message);
                result.Exception = ex;
                return null;
            }
        }

        public static bool InstallCertificate(X509Certificate2 cert, DeploymentResult result)
        {
            if (cert == null)
            {
                result.Success = false;
                result.Message = "Certificate to install cannot be null.";
                result.AddLog(result.Message);
                return false;
            }

            try
            {
                result.AddLog($"Installing certificate: {cert.Subject} (Thumbprint: {cert.Thumbprint}) to LocalMachine/My store.");
                using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
                {
                    store.Open(OpenFlags.ReadWrite);
                    store.Add(cert);
                    store.Close();
                }
                result.AddLog("Certificate installed successfully.");
                return true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error installing certificate: {ex.Message}";
                result.AddLog(result.Message);
                result.Exception = ex;
                return false;
            }
        }

        // The BindCertificateToIIS method from the original Program.cs was not directly used in the main flows.
        // It seems to be an alternative way to bind a certificate if not done during site creation.
        // For now, I'll include it here, refactored to use DeploymentResult.
        // It might require a Site object, which means it might fit better in IISDeployManager or be called from there.
        // For now, let's assume it's a utility that might be called with a Site object.
        public static bool BindCertificateToIIS(Site site, string siteName, DeploymentResult result, string ip = "*", int port = 443, string certThumbprint = "")
        {
            if (site == null)
            {
                result.Success = false;
                result.Message = "Site object cannot be null for binding certificate.";
                result.AddLog(result.Message);
                return false;
            }
            if (string.IsNullOrEmpty(certThumbprint))
            {
                result.Success = false;
                result.Message = "Certificate thumbprint cannot be empty for binding.";
                result.AddLog(result.Message);
                return false;
            }

            result.AddLog($"Binding certificate for site '{siteName}' on port {port} using thumbprint {certThumbprint}.");
            try
            {
                // Remove any existing binding on the specified port if needed
                var existingBinding = site.Bindings.FirstOrDefault(b => b.Protocol.Equals("https", StringComparison.OrdinalIgnoreCase) && b.EndPoint.Port == port);
                if (existingBinding != null)
                {
                    result.AddLog($"Removing existing HTTPS binding on port {port}.");
                    site.Bindings.Remove(existingBinding);
                }

                result.AddLog("Creating new HTTPS binding.");
                var binding = site.Bindings.CreateElement("binding");
                binding["protocol"] = "https";
                binding["bindingInformation"] = $"{ip}:{port}:"; // Hostname can go after the last colon if needed

                binding["certificateStoreName"] = StoreName.My.ToString(); // Ensure this matches the store it was installed to
                binding["certificateHash"] = StringToByteArray(certThumbprint); // Assuming StringToByteArray is accessible or moved

                site.Bindings.Add(binding);
                result.AddLog("New HTTPS binding added to site configuration. CommitChanges on ServerManager is required to apply.");
                // Note: serverManager.CommitChanges() must be called by the caller (e.g., within CreateNewSite or a dedicated method in IISDeployManager)
                return true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error binding certificate to IIS for site {siteName}: {ex.Message}";
                result.AddLog(result.Message);
                result.Exception = ex;
                return false;
            }
        }

        // Helper method, also present in IISDeployManager. Could be moved to a shared utility class.
        private static byte[] StringToByteArray(string hex)
        {
            hex = new string(hex.Where(c => Uri.IsHexDigit(c)).ToArray());
            if (hex.Length % 2 != 0)
                throw new FormatException("Invalid hex string length.");
            return Enumerable.Range(0, hex.Length / 2)
                .Select(x => Convert.ToByte(hex.Substring(x * 2, 2), 16))
                .ToArray();
        }
    }
}
