using Microsoft.Web.Administration;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace IISDeploy.Core;

public static class CertificateGenerator
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
        // Remove any existing HTTPS binding on this port first.
        var existingBinding = site.Bindings
            .FirstOrDefault(b => b.Protocol == "https" && b.EndPoint != null && b.EndPoint.Port == port);
        if (existingBinding != null)
        {
            site.Bindings.Remove(existingBinding);
        }

        // Use the strongly-typed SSL overload: it sets protocol=https and the
        // certificate hash/store correctly. Setting certificateHash by hand on a
        // raw element mis-marshals the byte[] and throws 0x80070459.
        byte[] certHash = StringToByteArray(certThumbprint);
        site.Bindings.Add($"{ip}:{port}:", certHash, "My");
    }

    internal static byte[] StringToByteArray(string hex)
    {
        // Remove all non-hex characters (including invisible Unicode)
        hex = new string(hex.Where(c => Uri.IsHexDigit(c)).ToArray());

        if (hex.Length % 2 != 0)
            throw new FormatException("Invalid hex string length.");

        return Enumerable.Range(0, hex.Length / 2)
            .Select(x => Convert.ToByte(hex.Substring(x * 2, 2), 16))
            .ToArray();
    }
}
